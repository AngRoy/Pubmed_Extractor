import streamlit as st
import requests
import urllib.parse
import xml.etree.ElementTree as ET
import pandas as pd
import datetime
import calendar
import re
import math
import time
import logging
from io import BytesIO

# Suppress excessive logging
import warnings
warnings.filterwarnings('ignore')

# Configure logging to file instead of console
logging.basicConfig(
    level=logging.WARNING,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='pubmed_extractor.log',
    filemode='a'
)
logger = logging.getLogger("pubmed_extractor")

# ------------------ Page Config & Styling ------------------

st.set_page_config(
    page_title="PubMed Month-to-Daily Extractor",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS to hide streamlit messages
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 1rem;
    }
    .text-centered {
        text-align: center;
    }
    .footer {
        text-align: center;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #1565C0;
        color: #BBDEFB;
    }
    .stApp {
        background-color: #121212;
        color: #FFFFFF;
    }
    .stDownloadButton button {
        background-color: #1976D2 !important;
        color: white !important;
        font-weight: bold !important;
        border: 1px solid #1565C0 !important;
        padding: 10px 20px !important;
    }
    .stDownloadButton button:hover {
        background-color: #1565C0 !important;
        border-color: #0D47A1 !important;
    }
    [data-testid="stSidebar"] {
        background-color: #1A1A1A;
    }
    div[data-testid="stStatusWidget"] {
        display: none;
    }
    .stProgress .st-bo {
        background-color: #1976D2 !important;
    }
    .status-message {
        margin-top: 8px;
        margin-bottom: 8px;
        font-style: italic;
        color: #90CAF9;
    }
    /* Hide warning/error messages unless specifically shown */
    [data-testid="stException"] {
        display: none !important;
    }
    /* Style warning */
    .warning-note {
        background-color: #FFA000;
        color: #000;
        border-radius: 8px;
        padding: 10px;
        margin-bottom: 16px;
    }
    /* Style day cards container */
    .day-cards-container {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
        gap: 16px;
        margin-top: 20px;
    }
    /* Style each day card */
    .day-card {
        background-color: #1A237E;
        border-radius: 8px;
        padding: 16px;
        display: flex;
        flex-direction: column;
        height: 100%;
    }
    .day-title {
        font-size: 1.1rem;
        margin-bottom: 8px;
        color: #BBDEFB;
        font-weight: bold;
    }
    .count-badge {
        background-color: #1976D2;
        border-radius: 12px;
        padding: 4px 10px;
        font-size: 0.9rem;
        margin-top: 8px;
        display: inline-block;
    }
    /* Progress indicator */
    .processing-header {
        margin-top: 1.5rem;
        margin-bottom: 0.5rem;
        color: #64B5F6;
        font-size: 1.2rem;
    }
</style>

<h1 class="main-header">PubMed Month-to-Daily Extractor</h1>
<p class="text-centered">Select a month & year to automatically fetch all PubMed articles published each day</p>
""", unsafe_allow_html=True)

# ------------------ PubMed API Config ------------------

ESearch_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
EFetch_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
BATCH_SIZE = 500  # maximum records per EFetch call
TOOL_NAME = "PubMedExtractor"
CONTACT_EMAIL = "your-email@example.com"  # Replace with your email
MAX_RETRIES = 3  # Maximum number of retry attempts for API calls

# The consistent API key to use for all requests
API_KEY = "29d53109210e6c9cb4249426d56ada159108"

# Store article data in session state to avoid resetting
if 'day_data' not in st.session_state:
    st.session_state.day_data = {}
if 'selected_month' not in st.session_state:
    st.session_state.selected_month = None
if 'days_processed' not in st.session_state:
    st.session_state.days_processed = 0

# ------------------ Helper Functions ------------------

def mkquery(base_url, params):
    """Create a formatted URL from base and parameters"""
    return base_url + "?" + "&".join(f"{k}={urllib.parse.quote(str(v))}" for k, v in params.items())

def get_xml(base_url, params, retry_count=0):
    """Make API request with retries and error handling"""
    api_params = params.copy()
    
    # Add API key to all requests
    api_params["api_key"] = API_KEY
    
    # Add tool parameters for better API identification
    api_params["tool"] = TOOL_NAME
    api_params["email"] = CONTACT_EMAIL
    
    try:
        logger.debug(f"Making request to {base_url} with params: {api_params}")
        response = requests.get(mkquery(base_url, api_params))
        response.raise_for_status()
        return ET.fromstring(response.text)
    except requests.HTTPError as e:
        if retry_count < MAX_RETRIES:
            # Exponential backoff
            wait_time = 2 ** retry_count
            logger.warning(f"API request failed: {str(e)}. Retrying in {wait_time} seconds...")
            time.sleep(wait_time)
            return get_xml(base_url, params, retry_count + 1)
        else:
            logger.error(f"API request failed after {MAX_RETRIES} retries: {str(e)}")
            raise
    except ET.ParseError:
        if retry_count < MAX_RETRIES:
            wait_time = 2 ** retry_count
            logger.warning(f"XML parsing failed. Retrying in {wait_time} seconds...")
            time.sleep(wait_time)
            return get_xml(base_url, params, retry_count + 1)
        else:
            logger.error(f"XML parsing failed after {MAX_RETRIES} retries")
            raise ValueError("Failed to parse XML response from PubMed API")

def get_text(node, xpath, default=""):
    """Extract text from XML node with default value"""
    n = node.find(xpath)
    return n.text.strip() if (n is not None and n.text) else default

def parse_month(mstr):
    """Parse month from string (name or number)"""
    if not mstr:
        return 1
    
    # If already numeric
    if mstr.isdigit():
        m = int(mstr)
        return m if 1 <= m <= 12 else 1
    
    # Handle month names and abbreviations
    mstr = mstr.strip().lower()
    month_map = {
        'jan': 1, 'january': 1,
        'feb': 2, 'february': 2,
        'mar': 3, 'march': 3,
        'apr': 4, 'april': 4,
        'may': 5,
        'jun': 6, 'june': 6,
        'jul': 7, 'july': 7,
        'aug': 8, 'august': 8,
        'sep': 9, 'september': 9, 'sept': 9,
        'oct': 10, 'october': 10,
        'nov': 11, 'november': 11,
        'dec': 12, 'december': 12
    }
    
    for abbr, num in month_map.items():
        if mstr.startswith(abbr):
            return num
    
    return 1  # Default to January if no match

def is_future_date(date_str):
    """Check if a date string is in the future"""
    try:
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
        today = datetime.date.today()
        return date_obj > today
    except (ValueError, TypeError):
        return False

def validate_article_date(article, year, month, day):
    """
    Validate that the article's publication date is exactly the specified day.
    Returns True if valid, False otherwise.
    """
    # Get the publication date from the article
    pub_date = article.get("PublicationDate", "")
    
    # Skip if no publication date
    if not pub_date:
        return False
    
    try:
        # Parse the date
        pub_date_obj = datetime.datetime.strptime(pub_date, "%Y-%m-%d").date()
        
        # Create the exact day we want
        target_date = datetime.date(year, month, day)
        
        # Check if date is in the future (beyond today)
        today = datetime.date.today()
        if pub_date_obj > today:
            return False
        
        # Check if publication date is exactly the target date
        return pub_date_obj == target_date
    except ValueError:
        # If date parsing fails, exclude the article
        return False

# ------------------ Core Fetch Functions ------------------

def get_day_article_count(year, month, day):
    """
    Get the count of PubMed articles for a specific day.
    """
    # Create the exact date string
    date_obj = datetime.date(year, month, day)
    ds = date_obj.strftime("%Y/%m/%d")
    term = f'"{ds}"[pdat]'
    
    try:
        # ESearch to get count
        root = get_xml(ESearch_URL, {
            "db": "pubmed", 
            "term": term, 
            "retmax": 0,
        })
        return int(get_text(root, "./Count", "0"))
    except Exception as e:
        logger.error(f"Error getting article count: {str(e)}")
        raise

def fetch_pubmed_for_day(year, month, day, progress_callback=None, status_callback=None):
    """
    Fetch PubMed articles published on a specific day.
    """
    # Check if date is in the future
    target_date = datetime.date(year, month, day)
    today = datetime.date.today()
    if target_date > today:
        if status_callback:
            status_callback(f"The date {target_date.strftime('%d/%m/%Y')} is in the future. No articles available.")
        return []
    
    # Check if we already fetched this exact date
    date_key = f"{day:02d}_{month:02d}_{year}"
    if date_key in st.session_state.day_data:
        if status_callback:
            status_callback(f"Using previously fetched data for {date_key}")
        return st.session_state.day_data[date_key]
        
    # Create the exact date string for query
    date_str = target_date.strftime("%Y/%m/%d")
    term = f'"{date_str}"[pdat]'
    
    if status_callback:
        status_callback(f"Fetching articles for {target_date.strftime('%d/%m/%Y')}")
    
    try:
        # Get a fresh WebEnv for this day
        search_params = {
            "db": "pubmed", 
            "term": term, 
            "usehistory": "y",
            "retmax": 0,
        }
        
        search_root = get_xml(ESearch_URL, search_params)
        count = int(get_text(search_root, "./Count", "0"))
        
        if count == 0:
            if status_callback:
                status_callback(f"No records found for {target_date.strftime('%d/%m/%Y')}")
            # Store empty list in session state
            st.session_state.day_data[date_key] = []
            return []
            
        webenv = get_text(search_root, "./WebEnv")
        query_key = get_text(search_root, "./QueryKey")
        
        if not webenv or not query_key:
            if status_callback:
                status_callback("Error: WebEnv or QueryKey not found in search results")
            # Store empty list in session state
            st.session_state.day_data[date_key] = []
            return []
        
        all_records = []
        batch_count = math.ceil(count / BATCH_SIZE)
        
        # Fetch records in batches
        for batch_idx in range(0, count, BATCH_SIZE):
            batch_start = batch_idx
            batch_end = min(batch_start + BATCH_SIZE, count)
            batch_size = batch_end - batch_start
            current_batch = batch_idx // BATCH_SIZE + 1
            
            if status_callback:
                overall_progress = (batch_idx + batch_size) / count
                status_callback(f"Fetching batch {current_batch}/{batch_count}: records {batch_start+1}-{batch_end} of {count}")
            
            try:
                fetch_root = get_xml(EFetch_URL, {
                    "db": "pubmed",
                    "query_key": query_key,
                    "WebEnv": webenv,
                    "retstart": batch_start,
                    "retmax": batch_size,
                    "retmode": "xml"
                })
                
                for art in fetch_root.findall(".//PubmedArticle"):
                    pmid = get_text(art, "./MedlineCitation/PMID")
                    title = get_text(art, "./MedlineCitation/Article/ArticleTitle")
                    journal = get_text(art, "./MedlineCitation/Article/Journal/Title")

                    # Publication date parsing with better fallbacks
                    # First try ArticleDate - most accurate electronic publication date
                    y = get_text(art, "./MedlineCitation/Article/ArticleDate/Year")
                    m = get_text(art, "./MedlineCitation/Article/ArticleDate/Month")
                    d = get_text(art, "./MedlineCitation/Article/ArticleDate/Day")
                    
                    # If not available, try JournalIssue/PubDate
                    if not y:
                        y = get_text(art, "./MedlineCitation/Article/Journal/JournalIssue/PubDate/Year")
                        m = get_text(art, "./MedlineCitation/Article/Journal/JournalIssue/PubDate/Month")
                        d = get_text(art, "./MedlineCitation/Article/Journal/JournalIssue/PubDate/Day")
                    
                    # If still no month, try MedlineDate
                    if not m:
                        medline_date = get_text(art, "./MedlineCitation/Article/Journal/JournalIssue/PubDate/MedlineDate", "")
                        if medline_date:
                            # Try to extract year and month from strings like "2014 Mar-Apr" or "2023 Jan"
                            parts = medline_date.split()
                            if len(parts) >= 2 and parts[0].isdigit():
                                y = parts[0]
                                month_part = parts[1].split('-')[0]  # Take first month if range
                                m = str(parse_month(month_part))
                    
                    # Convert month string to number if needed
                    m_parsed = parse_month(m)
                    d_parsed = int(d) if d and d.isdigit() else 1
                    
                    # Format the final date
                    try:
                        if y and y.isdigit():
                            pubdate = f"{int(y):04d}-{m_parsed:02d}-{d_parsed:02d}"
                        else:
                            pubdate = ""
                    except (ValueError, TypeError):
                        pubdate = ""

                    # Extract abstract - handle multiple AbstractText elements
                    abs_nodes = art.findall("./MedlineCitation/Article/Abstract/AbstractText")
                    abstract = " ".join(a.text.strip() for a in abs_nodes if a.text) if abs_nodes else ""

                    # Author extraction
                    authors = []
                    for author in art.findall("./MedlineCitation/Article/AuthorList/Author"):
                        last_name = get_text(author, "./LastName")
                        fore_name = get_text(author, "./ForeName")
                        
                        # If CollectiveName exists, use that
                        collective = get_text(author, "./CollectiveName")
                        if collective:
                            authors.append(collective)
                        elif last_name and fore_name:
                            authors.append(f"{last_name} {fore_name}")
                        elif last_name:
                            authors.append(last_name)
                        elif fore_name:
                            authors.append(fore_name)
                    
                    author_str = ", ".join(authors)

                    all_records.append({
                        "PMID": pmid,
                        "Journal": journal,
                        "Title": title,
                        "Authors": author_str,
                        "PublicationDate": pubdate,
                        "Abstract": abstract
                    })

                # Update progress
                if progress_callback:
                    progress_callback(min((batch_idx + batch_size) / count, 1.0))
                
                # Small delay to respect rate limits
                time.sleep(0.1)
                    
            except Exception as e:
                logger.error(f"Error in batch {current_batch}: {str(e)}")
                if status_callback:
                    status_callback(f"Error in batch {current_batch}: {str(e)}")
                # Continue with next batch instead of failing the entire day
                continue
        
        # Filter records by actual publication date to ensure they're exactly on the requested day
        valid_records = []
        future_dates = 0
        wrong_dates = 0
        
        for record in all_records:
            if is_future_date(record.get("PublicationDate", "")):
                future_dates += 1
                continue
                
            if validate_article_date(record, year, month, day):
                valid_records.append(record)
            else:
                wrong_dates += 1
        
        # Report filtering results
        if status_callback:
            if future_dates > 0:
                status_callback(f"Removed {future_dates} articles with future publication dates")
            if wrong_dates > 0:
                status_callback(f"Removed {wrong_dates} articles with dates other than {target_date.strftime('%d/%m/%Y')}")
            status_callback(f"Successfully retrieved {len(valid_records)} valid articles for {target_date.strftime('%d/%m/%Y')}")
        
        # Store in session state
        st.session_state.day_data[date_key] = valid_records
        
        return valid_records
        
    except Exception as e:
        logger.error(f"Error in fetch_pubmed_for_day: {str(e)}")
        if status_callback:
            status_callback(f"Error: {str(e)}")
        raise

def get_valid_days(year, month):
    """
    Get a list of valid days for the given month.
    Excludes future days.
    """
    days_in_month = calendar.monthrange(year, month)[1]
    
    # Check if requested month is in the future
    today = datetime.date.today()
    if year > today.year or (year == today.year and month > today.month):
        # Future month, return empty list
        return []
    
    # If the requested month is current month, limit to today's date
    if year == today.year and month == today.month:
        days_in_month = min(days_in_month, today.day)
    
    return list(range(1, days_in_month + 1))

def create_excel(df, include_styling=True):
    """Convert dataframe to Excel file in memory with optional styling"""
    if df.empty:
        # Return an empty Excel file with just headers if df is empty
        output = BytesIO()
        pd.DataFrame(columns=['No Data Available']).to_excel(output, index=False)
        return output.getvalue()
    
    output = BytesIO()
    
    try:
        if include_styling:
            # Use xlsxwriter for better styling control
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Articles')
                
                # Get workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets['Articles']
                
                # Define formats
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#1976D2',
                    'font_color': 'white',
                    'border': 1,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'align': 'center'
                })
                
                row_format = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'top',
                    'border': 1
                })
                
                # Apply header format
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Set column widths based on content
                for i, col in enumerate(df.columns):
                    # Determine column width based on content
                    max_width = 0
                    for j in range(min(10, len(df))):  # Check first 10 rows for speed
                        cell_value = str(df.iloc[j, i]) if not pd.isna(df.iloc[j, i]) else ""
                        # Calculate width based on content length
                        width = min(len(cell_value), 100) / 1.5  # Divide by constant to account for font width
                        max_width = max(max_width, width)
                    
                    # Set default widths for specific columns
                    if col == 'Abstract':
                        max_width = 80
                    elif col == 'Title':
                        max_width = 50
                    elif col == 'Journal':
                        max_width = 30
                    elif col == 'Authors':
                        max_width = 40
                    elif col == 'PMID':
                        max_width = 12
                    
                    # Ensure a minimum width and cap maximum width
                    max_width = max(10, min(max_width, 100))
                    worksheet.set_column(i, i, max_width)
                
                # Apply row formatting
                for row_num in range(1, len(df) + 1):
                    worksheet.set_row(row_num, None, row_format)
                
                # Freeze top row
                worksheet.freeze_panes(1, 0)
        else:
            # Simple Excel without styling
            df.to_excel(output, index=False)
        
        output.seek(0)
        return output.getvalue()
        
    except Exception as e:
        logger.error(f"Error creating Excel file: {str(e)}")
        # Create a simple Excel file without formatting as fallback
        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        return output.getvalue()

# ------------------ Sidebar Form ------------------

with st.sidebar:
    st.markdown("### Select Month & Year")
    
    # Add API key indicator only once
    st.success("✅ Using API key: Higher rate limits enabled")
    
    with st.form("month_form"):
        # Month and year selection only
        current_date = datetime.date.today()
        
        year = st.number_input(
            "Year",
            min_value=1900,
            max_value=current_date.year,
            value=current_date.year
        )
        
        month = st.selectbox(
            "Month",
            list(range(1, 13)),
            format_func=lambda x: calendar.month_name[x],
            index=current_date.month - 1  # Default to current month
        )
        
        # Additional options
        include_styling = st.checkbox("Include Excel styling", value=True, 
                                     help="Better formatting, but may be slower for large datasets")
        include_authors = st.checkbox("Include Authors", value=True)
        
        submitted = st.form_submit_button("Fetch All Days")
    
    # Reset button to clear session data
    if st.button("Reset All Data"):
        st.session_state.day_data = {}
        st.session_state.selected_month = None
        st.session_state.days_processed = 0
        st.success("All cached data has been cleared!")
    
    # PubMed API Info
    st.markdown("---")
    st.markdown("### About Month Processing")
    st.markdown("""
    - All days in the selected month will be processed automatically
    - Each file contains articles from a single day
    - Files are named as DD_MM_YYYY.xlsx (e.g., 07_05_2023.xlsx)
    - All days remain available for download after processing
    """)

# ------------------ Main App Logic ------------------

if submitted:
    # Check if month is in the future
    selected_date = datetime.date(year, month, 1)
    today = datetime.date.today()
    
    month_key = f"{month:02d}_{year}"
    
    if year > today.year or (year == today.year and month > today.month):
        st.warning(f"⚠️ {calendar.month_name[month]} {year} is in the future. No articles are available yet.")
    else:
        # Store the selected month for session state
        st.session_state.selected_month = month_key
        
        try:
            # Get valid days for this month
            valid_days = get_valid_days(year, month)
            
            if not valid_days:
                st.warning(f"No valid days for {calendar.month_name[month]} {year}.")
            else:
                # Fetch data for all days
                st.markdown(f"## Processing {calendar.month_name[month]} {year}")
                
                if today.year == year and today.month == month:
                    st.markdown(f"""
                    <div class="warning-note">
                    ⚠️ Note: Processing current month up to today ({today.strftime('%d/%m/%Y')}).
                    </div>
                    """, unsafe_allow_html=True)
                
                # Create processing area with progress
                st.markdown('<div class="processing-header">Processing each day:</div>', unsafe_allow_html=True)
                progress_bar = st.progress(0.0)
                status_text = st.empty()
                
                # Track total articles found
                total_articles = 0
                
                # Process each day in the month
                for i, day in enumerate(valid_days):
                    day_key = f"{day:02d}_{month:02d}_{year}"
                    
                    # Update overall progress
                    progress_bar.progress(i / len(valid_days))
                    
                    try:
                        # Fetch data for this day
                        status_text.markdown(f'<p class="status-message">Processing day {day:02d}/{month:02d}/{year}</p>', unsafe_allow_html=True)
                        
                        records = fetch_pubmed_for_day(
                            year, month, day,
                            status_callback=lambda msg: status_text.markdown(f'<p class="status-message">{msg}</p>', unsafe_allow_html=True)
                        )
                        
                        total_articles += len(records)
                        
                    except Exception as e:
                        st.error(f"🛑 Error processing {day:02d}/{month:02d}/{year}: {str(e)}")
                        logger.error(f"Error processing {day:02d}/{month:02d}/{year}: {str(e)}")
                        continue
                
                # Update final progress
                progress_bar.progress(1.0)
                status_text.markdown(f'<p class="status-message">All {len(valid_days)} days processed successfully!</p>', unsafe_allow_html=True)
                
                # Show summary of processing
                st.success(f"Found {total_articles:,} articles across {len(valid_days)} days in {calendar.month_name[month]} {year}")
                
                # Display all days in a grid for downloading
                st.markdown("## Download Files by Day")
                st.markdown('<div class="day-cards-container">', unsafe_allow_html=True)
                
                # Create a card for each day with a download button
                for day in valid_days:
                    day_key = f"{day:02d}_{month:02d}_{year}"
                    
                    if day_key in st.session_state.day_data:
                        records = st.session_state.day_data[day_key]
                        
                        # Create DataFrame
                        df = pd.DataFrame(records)
                        
                        # Remove authors column if not requested
                        if not df.empty and not include_authors and "Authors" in df.columns:
                            df = df.drop(columns=["Authors"])
                        
                        # Create day card
                        st.markdown(f"""
                        <div class="day-card">
                            <div class="day-title">{day:02d}/{month:02d}/{year}</div>
                            <div class="count-badge">{len(records)} articles</div>
                        """, unsafe_allow_html=True)
                        
                        # Create Excel file for download
                        excel_data = create_excel(df, include_styling)
                        
                        # Download button
                        st.download_button(
                            label=f"⬇️ Download",
                            data=excel_data,
                            file_name=f"{day:02d}_{month:02d}_{year}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_btn_{day}"
                        )
                        
                        st.markdown("</div>", unsafe_allow_html=True)
                    else:
                        # Create empty card for days with no data
                        st.markdown(f"""
                        <div class="day-card">
                            <div class="day-title">{day:02d}/{month:02d}/{year}</div>
                            <div class="count-badge">0 articles</div>
                            <p>No data available</p>
                        </div>
                        """, unsafe_allow_html=True)
                
                st.markdown('</div>', unsafe_allow_html=True)
                
        except Exception as e:
            st.error(f"🛑 An error occurred: {str(e)}")
            logger.error(f"Error in main app: {str(e)}")

# Display previously processed data if available but not newly submitted
elif st.session_state.selected_month is not None:
    # Parse out the month and year from the session key
    try:
        month, year = map(int, st.session_state.selected_month.split('_'))
        valid_days = get_valid_days(year, month)
        
        if valid_days:
            # Display the download grid for previously processed data
            st.markdown(f"## Download Files for {calendar.month_name[month]} {year}")
            st.markdown('<div class="day-cards-container">', unsafe_allow_html=True)
            
            for day in valid_days:
                day_key = f"{day:02d}_{month:02d}_{year}"
                
                if day_key in st.session_state.day_data:
                    records = st.session_state.day_data[day_key]
                    
                    # Create DataFrame
                    df = pd.DataFrame(records)
                    
                    # Remove authors column if needed
                    include_authors = True  # Default value since we don't know user preference now
                    if not df.empty and not include_authors and "Authors" in df.columns:
                        df = df.drop(columns=["Authors"])
                    
                    # Create day card
                    st.markdown(f"""
                    <div class="day-card">
                        <div class="day-title">{day:02d}/{month:02d}/{year}</div>
                        <div class="count-badge">{len(records)} articles</div>
                    """, unsafe_allow_html=True)
                    
                    # Create Excel file for download
                    excel_data = create_excel(df, include_styling=True)
                    
                    # Download button
                    st.download_button(
                        label=f"⬇️ Download",
                        data=excel_data,
                        file_name=f"{day:02d}_{month:02d}_{year}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_btn_{day}"
                    )
                    
                    st.markdown("</div>", unsafe_allow_html=True)
            
            st.markdown('</div>', unsafe_allow_html=True)
    except Exception as e:
        logger.error(f"Error displaying previous data: {str(e)}")
        # Don't show an error to the user, just silently fail

# ------------------ Footer ------------------

st.markdown("""
<div class="footer">
    <p>Developed by Rhenix Life Sciences</p>
    <p>© 2025 Rhenix Life Sciences. All rights reserved.</p>
    <p>Data sourced from PubMed via E-utilities API</p>
</div>
""", unsafe_allow_html=True)