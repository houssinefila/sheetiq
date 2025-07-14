import streamlit as st
import pandas as pd
import duckdb
import io
from datetime import datetime
import re
import time

# Configure the page with enhanced styling
st.set_page_config(
    page_title="Sheetiq - Excel Data Analysis",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for enhanced UI styling
def load_custom_css():
    """
    Load custom CSS for modern, professional styling.
    This function applies consistent theming across all components.
    """
    st.markdown("""
    <style>
    /* Import modern font family */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Global font styling - Inter is a modern, readable font */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    /* Main container styling with subtle background */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }
    
    /* Header styling with gradient background */
    .main-header {
        background: linear-gradient(90deg, #1e293b 0%, #334155 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
    }
    
    /* Section dividers with subtle styling */
    .section-divider {
        height: 2px;
        background: linear-gradient(90deg, transparent 0%, #3b82f6 50%, transparent 100%);
        margin: 2rem 0;
        border-radius: 1px;
    }
    
    /* Enhanced button styling */
    .stButton > button {
        background: linear-gradient(90deg, #3b82f6 0%, #1d4ed8 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 1.5rem;
        font-weight: 500;
        font-size: 0.9rem;
        transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(59, 130, 246, 0.3);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4);
        background: linear-gradient(90deg, #1d4ed8 0%, #1e40af 100%);
    }
    
    /* Primary button special styling */
    .stButton > button[kind="primary"] {
        background: linear-gradient(90deg, #059669 0%, #047857 100%);
        box-shadow: 0 2px 8px rgba(5, 150, 105, 0.3);
    }
    
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(90deg, #047857 0%, #065f46 100%);
        box-shadow: 0 4px 12px rgba(5, 150, 105, 0.4);
    }
    
    /* Enhanced input field styling */
    .stTextArea textarea {
        background-color: #1e293b;
        border: 2px solid #334155;
        border-radius: 8px;
        color: #f1f5f9;
        font-family: 'Monaco', 'Menlo', 'Ubuntu Mono', monospace;
        padding: 1rem;
        transition: border-color 0.3s ease;
    }
    
    .stTextArea textarea:focus {
        border-color: #3b82f6;
        box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background-color: #0f172a;
        border-right: 2px solid #1e293b;
    }
    
    /* File uploader styling */
    .stFileUploader {
        background-color: #1e293b;
        border: 2px dashed #3b82f6;
        border-radius: 12px;
        padding: 2rem;
        text-align: center;
        transition: all 0.3s ease;
    }
    
    .stFileUploader:hover {
        border-color: #60a5fa;
        background-color: #334155;
    }
    
    /* Success/Error message styling */
    .stSuccess {
        background-color: rgba(16, 185, 129, 0.1);
        border-left: 4px solid #10b981;
        border-radius: 6px;
        padding: 1rem;
    }
    
    .stError {
        background-color: rgba(239, 68, 68, 0.1);
        border-left: 4px solid #ef4444;
        border-radius: 6px;
        padding: 1rem;
    }
    
    .stWarning {
        background-color: rgba(245, 158, 11, 0.1);
        border-left: 4px solid #f59e0b;
        border-radius: 6px;
        padding: 1rem;
    }
    
    /* Loading spinner customization */
    .stSpinner > div {
        border-color: #3b82f6 transparent #3b82f6 transparent;
    }
    
    /* Data table styling */
    .stDataFrame {
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    }
    
    /* Expander styling */
    .streamlit-expanderHeader {
        background-color: #1e293b;
        border-radius: 8px;
        border: 1px solid #334155;
    }
    
    /* Card-like sections */
    .info-card {
        background-color: #1e293b;
        padding: 1.5rem;
        border-radius: 12px;
        border: 1px solid #334155;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
    }
    
    /* Query example buttons */
    .example-button {
        background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.75rem 1rem;
        margin: 0.5rem;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.3s ease;
        box-shadow: 0 2px 6px rgba(99, 102, 241, 0.3);
    }
    
    .example-button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(99, 102, 241, 0.4);
    }
    
    /* Tooltip styling */
    .tooltip {
        position: relative;
        cursor: help;
    }
    
    .tooltip:hover::after {
        content: attr(data-tooltip);
        position: absolute;
        bottom: 125%;
        left: 50%;
        transform: translateX(-50%);
        background-color: #1e293b;
        color: #f1f5f9;
        padding: 0.5rem;
        border-radius: 6px;
        font-size: 0.8rem;
        white-space: nowrap;
        z-index: 1000;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.3);
    }
    
    /* Animation classes */
    .fade-in {
        animation: fadeIn 0.5s ease-in;
    }
    
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    .pulse {
        animation: pulse 2s infinite;
    }
    
    @keyframes pulse {
        0% { opacity: 1; }
        50% { opacity: 0.7; }
        100% { opacity: 1; }
    }
    </style>
    """, unsafe_allow_html=True)

def init_session_state():
    """Initialize session state variables for multi-table support"""
    if 'query_history' not in st.session_state:
        st.session_state.query_history = []
    if 'uploaded_tables' not in st.session_state:
        st.session_state.uploaded_tables = {}  # Dict: {table_name: dataframe}
    if 'current_query' not in st.session_state:
        st.session_state.current_query = ""
    if 'selected_preview_table' not in st.session_state:
        st.session_state.selected_preview_table = None

def is_select_query(query):
    """
    Check if the query is a SELECT statement only.
    This function blocks potentially harmful SQL operations.
    """
    # Remove comments and normalize whitespace
    cleaned_query = re.sub(r'--.*$', '', query, flags=re.MULTILINE)
    cleaned_query = re.sub(r'/\*.*?\*/', '', cleaned_query, flags=re.DOTALL)
    cleaned_query = cleaned_query.strip().upper()
    
    # Check if query starts with SELECT
    if not cleaned_query.startswith('SELECT'):
        return False
    
    # Block dangerous keywords
    dangerous_keywords = [
        'INSERT', 'UPDATE', 'DELETE', 'DROP', 'CREATE', 'ALTER', 
        'TRUNCATE', 'REPLACE', 'MERGE', 'EXEC', 'EXECUTE', 'CALL'
    ]
    
    for keyword in dangerous_keywords:
        if keyword in cleaned_query:
            return False
    
    return True

def load_excel_data(uploaded_file):
    """Load Excel file and return the first sheet as a pandas DataFrame"""
    try:
        # Read only the first sheet
        df = pd.read_excel(uploaded_file, sheet_name=0, engine='openpyxl')
        return df, None
    except Exception as e:
        return None, f"Error reading Excel file: {str(e)}"

def get_table_name_from_filename(filename):
    """Extract a clean table name from filename (remove extension and special chars)"""
    import re
    # Remove file extension
    table_name = filename.rsplit('.', 1)[0]
    # Replace spaces and special characters with underscores
    table_name = re.sub(r'[^a-zA-Z0-9_]', '_', table_name)
    # Ensure it starts with a letter or underscore
    if table_name and not table_name[0].isalpha() and table_name[0] != '_':
        table_name = 'table_' + table_name
    return table_name or 'unnamed_table'

def delete_table(table_name):
    """Remove a table from uploaded_tables and update session state"""
    if table_name in st.session_state.uploaded_tables:
        del st.session_state.uploaded_tables[table_name]
        # Clear preview selection if it was the deleted table
        if st.session_state.selected_preview_table == table_name:
            st.session_state.selected_preview_table = None
        st.rerun()

def execute_sql_query_multi_table(uploaded_tables, query):
    """Execute SQL query on multiple tables using DuckDB"""
    try:
        # Check if query is a SELECT statement
        if not is_select_query(query):
            return None, "Only SELECT statements are allowed for security reasons."
        
        if not uploaded_tables:
            return None, "No tables available. Please upload at least one Excel file."
        
        # Create DuckDB connection
        conn = duckdb.connect(':memory:')
        
        # Register all tables with their respective names
        for table_name, df in uploaded_tables.items():
            conn.register(table_name, df)
        
        # Execute the query
        result = conn.execute(query).fetchdf()
        conn.close()
        
        return result, None
    except Exception as e:
        return None, f"SQL execution error: {str(e)}"

# Keep the old function for backward compatibility
def execute_sql_query(df, query):
    """Execute SQL query on a single DataFrame using DuckDB (legacy function)"""
    try:
        if not is_select_query(query):
            return None, "Only SELECT statements are allowed for security reasons."
        
        conn = duckdb.connect(':memory:')
        conn.register('df', df)
        result = conn.execute(query).fetchdf()
        conn.close()
        
        return result, None
    except Exception as e:
        return None, f"SQL execution error: {str(e)}"

def add_to_query_history(query, result_count=None, error=None):
    """Add executed query to session history"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    history_item = {
        'timestamp': timestamp,
        'query': query,
        'result_count': result_count,
        'error': error,
        'success': error is None
    }
    st.session_state.query_history.append(history_item)

def create_excel_download(df, filename="query_result.xlsx"):
    """Create Excel file from DataFrame for download"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Query Result', index=False)
    return output.getvalue()

def display_table_management():
    """Display table management interface in sidebar"""
    with st.sidebar:
        st.subheader("📊 Your Tables")
        
        if not st.session_state.uploaded_tables:
            st.caption("No tables uploaded yet")
            return
        
        st.caption(f"{len(st.session_state.uploaded_tables)} table(s) available")
        
        # Display each table with delete option and preview
        for table_name, df in st.session_state.uploaded_tables.items():
            with st.expander(f"📋 {table_name}", expanded=False):
                st.caption(f"📊 {len(df):,} rows × {len(df.columns)} columns")
                
                # Table actions
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("👁️ Preview", key=f"preview_{table_name}"):
                        st.session_state.selected_preview_table = table_name
                        st.rerun()
                with col2:
                    if st.button("🗑️ Delete", key=f"delete_{table_name}"):
                        delete_table(table_name)

def display_example_queries():
    """Display example SQL queries for multiple tables"""
    st.subheader("💡 Quick Start Examples")
    
    # Get table names for examples
    table_names = list(st.session_state.uploaded_tables.keys())
    
    if not table_names:
        st.info("Upload Excel files to see query examples")
        return
    
    # Dynamic examples based on available tables
    first_table = table_names[0]
    examples = [
        {
            "title": "Basic Selection",
            "query": f"SELECT * FROM {first_table} LIMIT 5;",
            "description": "Select first 5 rows from a table"
        },
        {
            "title": "Count Records",
            "query": f"SELECT COUNT(*) as total_rows FROM {first_table};",
            "description": "Count total rows in table"
        }
    ]
    
    # Add multi-table examples if we have multiple tables
    if len(table_names) >= 2:
        table1, table2 = table_names[0], table_names[1]
        examples.extend([
            {
                "title": "Join Tables",
                "query": f"SELECT * FROM {table1} a JOIN {table2} b ON a.id = b.id LIMIT 10;",
                "description": "Join two tables (adjust column names)"
            },
            {
                "title": "Union Tables",
                "query": f"SELECT * FROM {table1} UNION ALL SELECT * FROM {table2};",
                "description": "Combine data from multiple tables"
            }
        ])
    
    # Add general examples
    examples.extend([
        {
            "title": "Window Functions",
            "query": f"SELECT *, ROW_NUMBER() OVER (ORDER BY column_name) as rank FROM {first_table};",
            "description": "Add row numbers (replace column_name)"
        }
    ])
    
    # Display examples in columns
    num_cols = min(len(examples), 5)
    cols = st.columns(num_cols)
    for i, example in enumerate(examples):
        with cols[i % num_cols]:
            if st.button(f"📋 {example['title']}", key=f"example_{i}"):
                st.session_state.current_query = example['query']
                st.rerun()
            st.caption(example['description'])

def display_query_history():
    """Display enhanced query history in sidebar"""
    if st.session_state.query_history:
        with st.sidebar:
            st.subheader("📚 Query History")
            st.caption("Recent queries and results")
        
        for i, item in enumerate(reversed(st.session_state.query_history)):
            with st.sidebar.expander(
                f"🔍 Query {len(st.session_state.query_history) - i}",
                expanded=False
            ):
                # Query timestamp
                st.caption(f"🕒 {item['timestamp']}")
                
                # Query code with syntax highlighting
                st.code(item['query'], language='sql')
                
                # Results summary
                if item['success']:
                    st.success(f"✅ Success: {item['result_count']} rows")
                    
                    # Action buttons
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("🔄 Re-run", key=f"rerun_{i}", use_container_width=True):
                            st.session_state.current_query = item['query']
                            st.rerun()
                    with col2:
                        if st.button("✏️ Edit", key=f"edit_{i}", use_container_width=True):
                            st.session_state.current_query = item['query']
                            st.rerun()
                else:
                    error_msg = item['error'][:50] + '...' if len(item['error']) > 50 else item['error']
                    st.error(f"❌ Error: {error_msg}")

def create_enhanced_header():
    """
    Create an enhanced header with modern styling and helpful information.
    Uses custom HTML for better visual appeal and user guidance.
    """
    st.markdown("""
    <div class="main-header fade-in">
        <h1 style="margin: 0; font-size: 3rem; font-weight: 700; color: #f1f5f9;">
            🔍 Sheetiq
        </h1>
        <p style="margin: 0.5rem 0 0 0; font-size: 1.2rem; color: #94a3b8; font-weight: 300;">
            Smart Excel Analysis Made Simple
        </p>
        <div style="margin-top: 1rem; font-size: 0.9rem; color: #64748b;">
            Upload your Excel file and query it with SQL — no database setup required
        </div>
    </div>
    """, unsafe_allow_html=True)

def create_section_divider(title=None):
    """
    Create a visual section divider with optional title.
    Helps organize content and improve visual flow.
    """
    if title:
        st.markdown(f"""
        <div style="margin: 2rem 0 1rem 0;">
            <h3 style="color: #e2e8f0; margin-bottom: 0.5rem; font-weight: 600;">{title}</h3>
            <div class="section-divider"></div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

def show_loading_animation(message="Processing..."):
    """
    Display a loading animation with custom message.
    Provides visual feedback during longer operations.
    """
    with st.spinner(message):
        time.sleep(0.1)  # Brief pause for visual effect

def main():
    """
    Main application function with enhanced UI and user experience.
    Organized into clear sections with proper styling and animations.
    """
    # Load custom CSS first for proper styling
    load_custom_css()
    
    # Initialize session state
    init_session_state()
    
    # Enhanced header section
    create_enhanced_header()
    
    # Enhanced sidebar with multiple file upload support
    with st.sidebar:
        st.title("⚡ Controls")
        st.caption("Manage your data and queries")
        
        # Multiple file upload section
        st.subheader("📁 Upload Excel Files")
        st.caption("Upload multiple .xlsx files (up to 100GB each)")
        
        uploaded_files = st.file_uploader(
            "Choose Excel files",
            type=['xlsx'],
            accept_multiple_files=True,
            help="🔍 Tips: Each file's first sheet will be loaded as a separate table. Use filename (without extension) as table name in SQL queries.",
            label_visibility="collapsed"
        )
    
    # Process uploaded files
    if uploaded_files:
        for uploaded_file in uploaded_files:
            table_name = get_table_name_from_filename(uploaded_file.name)
            
            # Skip if already loaded
            if table_name in st.session_state.uploaded_tables:
                continue
            
            # Show loading animation
            with st.spinner(f"🔄 Loading {uploaded_file.name}..."):
                df, error = load_excel_data(uploaded_file)
            
            if error:
                st.error(f"❌ **Failed to load {uploaded_file.name}**: {error}")
                continue
            
            # Store the table
            st.session_state.uploaded_tables[table_name] = df
            
            # Show success message
            st.success(f"✅ Loaded **{table_name}** ({len(df):,} rows × {len(df.columns)} columns)")
    
    # Display table management interface
    display_table_management()
    
    # Display query history
    display_query_history()
    
    # Main content area with multi-table support
    if st.session_state.uploaded_tables:
        # Show table preview if selected
        if st.session_state.selected_preview_table:
            selected_table = st.session_state.selected_preview_table
            if selected_table in st.session_state.uploaded_tables:
                df = st.session_state.uploaded_tables[selected_table]
                
                create_section_divider(f"📊 Preview: {selected_table}")
                
                st.subheader("👀 Data Preview")
                st.caption(f"Showing the first 10 rows of table '{selected_table}'")
                
                st.dataframe(
                    df.head(10), 
                    use_container_width=True,
                    height=400
                )
                
                # Column information
                with st.expander("📋 **Column Details & Statistics**", expanded=False):
                    col_info = pd.DataFrame({
                        'Column Name': df.columns,
                        'Data Type': df.dtypes.astype(str),
                        'Non-Null Values': df.count(),
                        'Missing Values': df.isnull().sum(),
                        'Sample Values': [str(df[col].dropna().iloc[0]) if not df[col].dropna().empty else 'N/A' for col in df.columns]
                    })
                    st.dataframe(col_info, use_container_width=True)
                    
                    # Quick insights for selected table
                    st.subheader("📈 Quick Insights")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Total Rows", f"{len(df):,}")
                        st.metric("Missing Values", f"{df.isnull().sum().sum():,}")
                    with col2:
                        st.metric("Columns", len(df.columns))
                        st.metric("Numeric Columns", len(df.select_dtypes(include=['number']).columns))
                
                # Button to clear preview
                if st.button("❌ Close Preview"):
                    st.session_state.selected_preview_table = None
                    st.rerun()
        
        else:
            # Overview of all tables
            create_section_divider("📊 Tables Overview")
            
            st.subheader("📋 Available Tables")
            st.caption(f"You have {len(st.session_state.uploaded_tables)} table(s) loaded")
            
            # Display summary of all tables
            table_summary = []
            for table_name, df in st.session_state.uploaded_tables.items():
                table_summary.append({
                    'Table Name': table_name,
                    'Rows': f"{len(df):,}",
                    'Columns': len(df.columns),
                    'Memory Usage': f"{df.memory_usage(deep=True).sum() / 1024 / 1024:.1f} MB"
                })
            
            if table_summary:
                summary_df = pd.DataFrame(table_summary)
                st.dataframe(summary_df, use_container_width=True)
        
        # Enhanced example queries section
        create_section_divider("💡 Get Started with Examples")
        display_example_queries()
        
        # Enhanced SQL Query Interface for multiple tables
        create_section_divider("💬 Query Your Data")
        
        st.subheader("🔍 Multi-Table SQL Query Builder")
        
        # Display available table names
        table_names = list(st.session_state.uploaded_tables.keys())
        if table_names:
            st.caption(f"Available tables: {', '.join(f'`{name}`' for name in table_names)}")
        else:
            st.warning("No tables available. Upload Excel files first.")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.success("✓ SELECT statements")
        with col2:
            st.success("✓ JOIN multiple tables")
        with col3:
            st.success("✓ Window functions")
        with col4:
            st.warning("⚠ Read-only (secure)")
        
        # Enhanced query input with dynamic placeholder
        if table_names:
            first_table = table_names[0]
            if len(table_names) >= 2:
                second_table = table_names[1]
                placeholder_text = f"""-- Example queries for your tables:
-- SELECT * FROM {first_table} LIMIT 10;
-- SELECT COUNT(*) FROM {first_table};
-- SELECT * FROM {first_table} a JOIN {second_table} b ON a.id = b.id;
-- SELECT * FROM {first_table} UNION ALL SELECT * FROM {second_table};"""
            else:
                placeholder_text = f"""-- Example queries:
-- SELECT * FROM {first_table} LIMIT 10;
-- SELECT COUNT(*) FROM {first_table};
-- SELECT column_name, COUNT(*) FROM {first_table} GROUP BY column_name;"""
        else:
            placeholder_text = "-- Upload Excel files first to start querying"
        
        query = st.text_area(
            "✏️ **Write your SQL query here:**",
            value=st.session_state.current_query,
            height=250,
            placeholder=placeholder_text,
            help="💡 Tip: Use the table names shown above. You can JOIN, UNION, or query multiple tables in a single query."
        )
        
        # Update session state
        st.session_state.current_query = query
        
        # Enhanced execution controls
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns([2, 1.5, 1, 3])
        
        with col1:
            execute_button = st.button(
                "🚀 **Run Query**", 
                type="primary",
                use_container_width=True,
                help="Execute your SQL query on the uploaded data"
            )
        with col2:
            clear_button = st.button(
                "🗑️ Clear Query", 
                use_container_width=True,
                help="Clear the current query"
            )
        with col3:
            format_button = st.button(
                "✨ Format", 
                use_container_width=True,
                help="Auto-format your SQL query"
            )
        
        # Handle button actions
        if clear_button:
            st.session_state.current_query = ""
            st.rerun()
            
        if format_button and query.strip():
            # Simple SQL formatting
            formatted_query = query.strip()
            formatted_query = formatted_query.replace(' select ', ' SELECT ')
            formatted_query = formatted_query.replace(' from ', ' FROM ')
            formatted_query = formatted_query.replace(' where ', ' WHERE ')
            formatted_query = formatted_query.replace(' group by ', ' GROUP BY ')
            formatted_query = formatted_query.replace(' order by ', ' ORDER BY ')
            st.session_state.current_query = formatted_query
            st.rerun()
        
        # Enhanced query execution with multi-table support
        if execute_button and query.strip():
            if not st.session_state.uploaded_tables:
                st.warning("⚠️ **No tables available!** Please upload Excel files first.")
            else:
                # Show loading with query preview
                with st.spinner(f"🔄 Executing query on {len(st.session_state.uploaded_tables)} table(s)..."):
                    start_time = time.time()
                    result_df, error = execute_sql_query_multi_table(st.session_state.uploaded_tables, query)
                    execution_time = time.time() - start_time
                
                if error:
                    st.error(f"❌ **Query Failed**: {error}")
                    
                    # Enhanced error guidance
                    with st.expander("🔧 **Troubleshooting Help**", expanded=True):
                        st.write("**Common Solutions:**")
                        st.write("• **Table not found**: Use correct table names (check sidebar)")
                        st.write("• **Column not found**: Check column names with `SELECT * FROM table_name LIMIT 5`")
                        st.write("• **JOIN issues**: Verify that join columns exist in both tables")
                        st.write("• **Syntax error**: Ensure proper SQL syntax (commas, quotes, etc.)")
                        
                        st.info("💡 **Pro Tip**: Start with simple queries like `SELECT * FROM table_name LIMIT 5` to explore your data structure first.")
                    
                    add_to_query_history(query, error=error)
                else:
                    # Enhanced success message with execution stats
                    st.success(f"🎉 **Query Executed Successfully!** Returned {len(result_df):,} rows in {execution_time:.2f} seconds")
                    
                    add_to_query_history(query, result_count=len(result_df))
                    
                    # Enhanced results display
                    create_section_divider("📊 Query Results")
                    
                    if len(result_df) > 0:
                        # Results summary
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("📊 Rows Returned", f"{len(result_df):,}")
                        with col2:
                            st.metric("📋 Columns", len(result_df.columns))
                        with col3:
                            st.metric("⏱️ Execution Time", f"{execution_time:.2f}s")
                        
                        # Display results with enhanced styling
                        st.subheader("📋 Results Table")
                        st.dataframe(
                            result_df, 
                            use_container_width=True,
                            height=min(600, len(result_df) * 40 + 100)
                        )
                        
                        # Enhanced download section
                        st.subheader("💾 Export Results")
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            excel_data = create_excel_download(result_df)
                            st.download_button(
                                label="📊 Download as Excel",
                                data=excel_data,
                                file_name=f"sheetiq_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                help="Download your query results as an Excel file"
                            )
                        
                        with col2:
                            csv_data = result_df.to_csv(index=False)
                            st.download_button(
                                label="📄 Download as CSV",
                                data=csv_data,
                                file_name=f"sheetiq_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                mime="text/csv",
                                help="Download your query results as a CSV file"
                            )
                    else:
                        st.info("🔍 Query executed successfully but returned no results. Try adjusting your query conditions.")
        
        elif execute_button and not query.strip():
            st.warning("⚠️ **Please write a SQL query first!** Use the examples above to get started.")
    
    else:
        # Enhanced welcome section when no files are uploaded
        st.markdown("### 📊 Ready to Analyze Your Excel Data?")
        st.write("Upload multiple Excel files using the sidebar to get started with powerful multi-table SQL analysis")
        
        st.markdown("---")
        
        # Feature highlights
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.info("🚀 **Multi-File Upload**\n\nUpload multiple Excel files at once (up to 100GB each)")
        
        with col2:
            st.info("🔍 **Multi-Table SQL**\n\nJoin, union, and query across multiple tables with advanced SQL")
        
        with col3:
            st.info("📥 **Easy Management**\n\nPreview, delete, and export data with intuitive controls")
        
        st.success("💡 **Pro Tip:** Each Excel file becomes a table named after its filename. You can join multiple tables in a single SQL query!")

if __name__ == "__main__":
    main()
