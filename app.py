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
    page_title="PubMed Month Extractor",
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
    /* Style chunk cards */
    .chunk-card {
        background-color: #1A237E;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 16px;
    }
    .chunk-title {
        font-size: 1.2rem;
        margin-bottom: 8px;
        color: #BBDEFB;
    }
</style>

<h1 class="main-header">PubMed Month Extractor</h1>
<p class="text-centered">Select a year &amp; month, fetch all PubMed articles published then, and download as multiple .xlsx files</p>
""", unsafe_allow_html=True)

# ------------------ PubMed API Config ------------------

ESearch_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
EFetch_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
BATCH_SIZE = 500  # maximum records per EFetch call
TOOL_NAME = "PubMedMonthExtractor"
CONTACT_EMAIL = "angshuman@rhenix.org"  # Replace with your email
MAX_RETRIES = 3  # Maximum number of retry attempts for API calls

# The consistent API key to use for all requests
API_KEY = "29d53109210e6c9cb4249426d56ada159108"

# Number of days per chunk
DAYS_PER_CHUNK = 5

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

# ------------------ Core Fetch Functions ------------------

def get_total_record_count(year, month):
    """
    Get the total number of PubMed articles for the given year/month.
    """
    # Build date range for the entire month
    start = datetime.date(year, month, 1)
    last_day = calendar.monthrange(year, month)[1]
    end = datetime.date(year, month, last_day)
    ds, de = start.strftime("%Y/%m/%d"), end.strftime("%Y/%m/%d")
    term = f'"{ds}"[pdat] : "{de}"[pdat]'
    
    try:
        # ESearch to get total count
        root = get_xml(ESearch_URL, {
            "db": "pubmed", 
            "term": term, 
            "retmax": 0,
        })
        return int(get_text(root, "./Count", "0"))
    except Exception as e:
        logger.error(f"Error getting total record count: {str(e)}")
        raise

def fetch_pubmed_by_date_range(year, month, start_day, end_day, progress_callback=None, status_callback=None):
    """
    Fetch PubMed articles published in a specific date range within a month.
    This allows retrieving different subsets of articles that don't overlap.
    """
    # Build date range for this specific period
    start = datetime.date(year, month, start_day)
    end = datetime.date(year, month, end_day)
    ds, de = start.strftime("%Y/%m/%d"), end.strftime("%Y/%m/%d")
    term = f'"{ds}"[pdat] : "{de}"[pdat]'
    
    if status_callback:
        status_callback(f"Fetching articles from {ds} to {de}")
    
    try:
        # Get a fresh WebEnv for this date range
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
                status_callback(f"No records found from {ds} to {de}")
            return []
            
        webenv = get_text(search_root, "./WebEnv")
        query_key = get_text(search_root, "./QueryKey")
        
        if not webenv or not query_key:
            if status_callback:
                status_callback("Error: WebEnv or QueryKey not found in search results")
            return []
        
        records = []
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
                    # First try ArticleDate
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

                    records.append({
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
                
                # Small delay to respect rate limits - with API key we can make 10 requests/sec
                time.sleep(0.1)
                    
            except Exception as e:
                logger.error(f"Error in batch {current_batch}: {str(e)}")
                if status_callback:
                    status_callback(f"Error in batch {current_batch}: {str(e)}")
                # Continue with next batch instead of failing the entire chunk
                continue
        
        if status_callback:
            status_callback(f"Successfully retrieved {len(records)} articles from {ds} to {de}")
        
        return records
        
    except Exception as e:
        logger.error(f"Error in fetch_pubmed_by_date_range: {str(e)}")
        if status_callback:
            status_callback(f"Error: {str(e)}")
        raise

def calculate_date_ranges(year, month):
    """
    Calculate date ranges for the given month, splitting into chunks of DAYS_PER_CHUNK days.
    Returns a list of (start_day, end_day) tuples.
    """
    days_in_month = calendar.monthrange(year, month)[1]
    date_ranges = []
    
    for start_day in range(1, days_in_month + 1, DAYS_PER_CHUNK):
        end_day = min(start_day + DAYS_PER_CHUNK - 1, days_in_month)
        date_ranges.append((start_day, end_day))
    
    return date_ranges

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
        year = st.number_input(
            "Year",
            min_value=1900,
            max_value=datetime.date.today().year,
            value=datetime.date.today().year
        )
        
        month = st.selectbox(
            "Month",
            list(range(1, 13)),
            format_func=lambda x: calendar.month_name[x]
        )
        
        # Additional options
        include_styling = st.checkbox("Include Excel styling", value=True, 
                                     help="Better formatting, but may be slower for large datasets")
        include_authors = st.checkbox("Include Authors", value=True)
        
        submitted = st.form_submit_button("Fetch Articles")
    
    # PubMed API Info
    st.markdown("---")
    st.markdown("### About Date Range Downloads")
    st.markdown(f"""
    - This app divides the month into {DAYS_PER_CHUNK}-day periods
    - Each chunk processes a different date range (no overlap)
    - 30-day months will generate 6 chunks
    - 31-day months will generate 7 chunks
    - This ensures all articles are retrieved reliably
    """)

# ------------------ Main App Logic ------------------

if submitted:
    # Get total record count first
    try:
        with st.spinner("Checking total article count..."):
            total_count = get_total_record_count(year, month)
            
        if total_count == 0:
            st.info(f"📭 No publications found for {calendar.month_name[month]} {year}.")
        else:
            st.markdown("---")
            st.success(f"Found {total_count:,} articles for {calendar.month_name[month]} {year}")
            
            # Calculate date ranges for this month
            date_ranges = calculate_date_ranges(year, month)
            
            st.info(f"Articles will be downloaded in {len(date_ranges)} separate files by date ranges")
            
            all_chunks_article_count = 0
            
            for chunk_idx, (start_day, end_day) in enumerate(date_ranges):
                st.markdown(f"### Processing Date Range {start_day}-{end_day} {calendar.month_name[month]} {year}")
                
                progress_bar = st.progress(0.0)
                status_text = st.empty()
                
                # Fetch articles for this date range
                try:
                    records = fetch_pubmed_by_date_range(
                        year, month, start_day, end_day,
                        progress_callback=lambda frac: progress_bar.progress(frac),
                        status_callback=lambda msg: status_text.markdown(f'<p class="status-message">{msg}</p>', unsafe_allow_html=True)
                    )
                    
                    # Convert to DataFrame
                    df = pd.DataFrame(records)
                    
                    if not df.empty:
                        # Remove authors column if not requested
                        if not include_authors and "Authors" in df.columns:
                            df = df.drop(columns=["Authors"])
                        
                        all_chunks_article_count += len(df)
                        
                        # Create Excel file
                        excel_data = create_excel(df, include_styling)
                        
                        # Format date range for display
                        date_format = lambda d: f"{year}-{month:02d}-{d:02d}"
                        start_date = date_format(start_day)
                        end_date = date_format(end_day)
                        
                        # Create download section
                        st.markdown(f"""
                        <div class="chunk-card">
                            <div class="chunk-title">Date Range: {start_date} to {end_date}</div>
                            <p>Contains {len(df):,} articles</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        st.download_button(
                            label=f"⬇️ Download {start_day}-{end_day} {calendar.month_name[month]} ({len(df):,} articles)",
                            data=excel_data,
                            file_name=f"pubmed_{calendar.month_name[month].lower()}{year}_{start_day}-{end_day}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_btn_{chunk_idx}"
                        )
                        
                        # Preview data
                        with st.expander(f"Preview {start_day}-{end_day} {calendar.month_name[month]}"):
                            df_display = df.copy()
                            st.dataframe(df_display.head(5), use_container_width=True)
                    else:
                        st.warning(f"No articles found for days {start_day}-{end_day}")
                
                except Exception as e:
                    st.error(f"🛑 Error processing days {start_day}-{end_day}: {str(e)}")
                    logger.error(f"Error processing days {start_day}-{end_day}: {str(e)}")
                    continue
            
            # Show summary
            st.markdown("---")
            percent_retrieved = min(100, (all_chunks_article_count / total_count) * 100)
            st.success(f"Successfully retrieved {all_chunks_article_count:,} articles ({percent_retrieved:.1f}% of total {total_count:,})")
            
    except Exception as e:
        st.error(f"🛑 An error occurred: {str(e)}")
        logger.error(f"Error in main app: {str(e)}")

# ------------------ Footer ------------------

st.markdown("""
<div class="footer">
    <p>Developed by Rhenix Life Sciences</p>
    <p>© 2025 Rhenix Life Sciences. All rights reserved.</p>
    <p>Data sourced from PubMed via E-utilities API</p>
</div>
""", unsafe_allow_html=True)