import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.units import inch
from reportlab.lib import colors

# Page config
st.set_page_config(
    page_title="Open-AudIT ROI Calculator",
    page_icon="💰",
    layout="wide"
)

# Constants (from your VBA module)
SAVING_PCT_LICENSE = 0.6
LICENCE_SPEND_REDUCTION_PCT = 0.05
MIN_PER_DEVICE_DISCOVERY = 10
HOURS_PER_ASSET_REPORT = 0.5
SAVING_PCT_ASSET = 0.8
CRITICAL_DEVICE_PCT = 0.3
MIN_PER_CHECK_MANUAL = 5
MIN_PER_CHECK_AUTOMATED = 1
MIN_PER_DEVICE_VULN_PER_YEAR = 10

# Custom CSS with navy blue theme
st.markdown("""
    <style>
    .main-header {
        font-size: 4rem;
        color: #1f4788;
        font-weight: bold;
        margin-bottom: 1rem;
        line-height: 1.2;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .savings-total {
        background-color: #4a6fa5;
        padding: 20px;
        border-radius: 10px;
        border: 2px solid #2c4a7c;
        color: white;
    }
    .savings-total h3 {
        color: white;
    }
    .savings-total p {
        color: white;
    }
    /* Increase sidebar font sizes */
    [data-testid="stSidebar"] label {
        font-size: 32px !important;
        font-weight: bold !important;
    }
    [data-testid="stSidebar"] input {
        font-size: 16px !important;
    }
    [data-testid="stSidebar"] .stNumberInput label {
        font-size: 32px !important;
        font-weight: bold !important;
    }
    /* Reduce sidebar spacing */
    [data-testid="stSidebar"] .stNumberInput {
        margin-bottom: 0.2rem !important;
        margin-top: 0 !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        margin-bottom: 0.1rem !important;
        margin-top: 0.1rem !important;
    }
    [data-testid="stSidebar"] hr {
        margin-top: 0.3rem !important;
        margin-bottom: 0.3rem !important;
    }
    [data-testid="stSidebar"] h2 {
        margin-top: 0.3rem !important;
        margin-bottom: 0.3rem !important;
        font-size: 1.1rem !important;
    }
    /* Reduce caption spacing */
    [data-testid="stSidebar"] .stCaption {
        margin-top: 0 !important;
        margin-bottom: 0 !important;
    }
    /* Tighter button spacing */
    [data-testid="stSidebar"] .stButton {
        margin-top: 0.5rem !important;
    }
    /* Reduce overall sidebar padding */
    [data-testid="stSidebar"] > div:first-child {
        padding-top: 1rem !important;
    }
    /* Green clear button */
    .stButton button[kind="secondary"] {
        background-color: #28a745 !important;
        color: white !important;
    }
    .stButton button[kind="secondary"]:hover {
        background-color: #218838 !important;
    }
    /* Hero section spacing */
    .hero-section {
        padding: 2rem 0;
        margin-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

# Header with logo - HERO SECTION
st.markdown('<div class="hero-section">', unsafe_allow_html=True)
col1, col2 = st.columns([1, 3])
with col1:
    try:
        st.image("firstwave_logo.png", width=300)
    except:
        st.markdown('<h1 style="font-size: 3rem; color: #1f4788;">FirstWave</h1>', unsafe_allow_html=True)

with col2:
    st.markdown('<p class="main-header">Open-AudIT ROI Calculator</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Calculate your return on investment with Open-AudIT automation</p>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# Sidebar for inputs
with st.sidebar:
    st.header("📊 Your IT Environment")
    
    col_input, col_display = st.columns([2, 1])
    with col_input:
        num_employees = st.number_input(
            "Number of Employees", 
            min_value=0, 
            value=1000, 
            step=100,
            format="%d"
        )
    with col_display:
        st.markdown(f"<p style='margin-top:28px; font-size:14px; color:#666;'>{num_employees:,}</p>", unsafe_allow_html=True)
    
    col_input, col_display = st.columns([2, 1])
    with col_input:
        num_devices = st.number_input(
            "Number of IT Devices", 
            min_value=0, 
            value=10000, 
            step=1000,
            format="%d"
        )
    with col_display:
        st.markdown(f"<p style='margin-top:28px; font-size:14px; color:#666;'>{num_devices:,}</p>", unsafe_allow_html=True)
    
    col_input, col_display = st.columns([2, 1])
    with col_input:
        hourly_rate = st.number_input(
            "Average Hourly Rate of IT Staff ($/hr)", 
            min_value=0.0, 
            value=50.0, 
            step=5.0,
            format="%.2f"
        )
    with col_display:
        st.markdown(f"<p style='margin-top:28px; font-size:14px; color:#666;'>${hourly_rate:,.2f}</p>", unsafe_allow_html=True)
    
    st.divider()
    st.header("🔧 Current Processes")
    
    col_input, col_display = st.columns([2, 1])
    with col_input:
        licence_requests = st.number_input(
            "Warranty/Licence Requests per Year", 
            min_value=0, 
            value=1000, 
            step=50,
            format="%d"
        )
    with col_display:
        st.markdown(f"<p style='margin-top:28px; font-size:14px; color:#666;'>{licence_requests:,}</p>", unsafe_allow_html=True)
    
    licence_hours = st.number_input(
        "Avg Processing Time per Licence Request (hrs)", 
        min_value=0.0, 
        value=0.5, 
        step=0.1,
        format="%.2f"
    )
    
    col_input, col_display = st.columns([2, 1])
    with col_input:
        licence_spend = st.number_input(
            "Total Current Licence Spend ($/yr)", 
            min_value=0, 
            value=5000000, 
            step=100000,
            format="%d"
        )
    with col_display:
        st.markdown(f"<p style='margin-top:28px; font-size:14px; color:#666;'>${licence_spend:,}</p>", unsafe_allow_html=True)
    
    reports_per_year = st.number_input(
        "Asset & Inventory Reports per Year", 
        min_value=0, 
        value=12, 
        step=1,
        format="%d"
    )
    
    checks_per_year = st.number_input(
        "Change Detection & Config Mgmt Checks per Year", 
        min_value=0, 
        value=12, 
        step=1,
        format="%d"
    )
    
    st.divider()
    st.header("💵 Investment")
    
    col_input, col_display = st.columns([2, 1])
    with col_input:
        sub_cost = st.number_input(
            "Open-AudIT Annual Subscription Cost ($)", 
            min_value=0, 
            value=100000, 
            step=10000,
            format="%d"
        )
    with col_display:
        st.markdown(f"<p style='margin-top:28px; font-size:14px; color:#666;'>${sub_cost:,}</p>", unsafe_allow_html=True)

# Initialize session state
if 'show_calculations' not in st.session_state:
    st.session_state.show_calculations = False

# Calculate button
if st.sidebar.button("Calculate ROI", type="primary", use_container_width=True):
    st.session_state.show_calculations = True

# Calculations
if st.session_state.show_calculations:
    
    # Calculate all values first
    warranty_hours_calc = licence_requests * licence_hours * SAVING_PCT_LICENSE
    warranty_dollars_calc = warranty_hours_calc * hourly_rate
    
    licence_spend_savings_calc = licence_spend * LICENCE_SPEND_REDUCTION_PCT
    
    asset_hours_calc = ((num_devices * MIN_PER_DEVICE_DISCOVERY / 60.0) + 
                   (reports_per_year * HOURS_PER_ASSET_REPORT)) * SAVING_PCT_ASSET
    asset_dollars_calc = asset_hours_calc * hourly_rate
    
    critical_devices = num_devices * CRITICAL_DEVICE_PCT
    change_hours_calc = critical_devices * checks_per_year * ((MIN_PER_CHECK_MANUAL - MIN_PER_CHECK_AUTOMATED) / 60.0)
    change_hours_calc = max(0, change_hours_calc)
    change_dollars_calc = change_hours_calc * hourly_rate
    
    vuln_hours_calc = num_devices * MIN_PER_DEVICE_VULN_PER_YEAR / 60.0
    vuln_dollars_calc = vuln_hours_calc * hourly_rate
    
    report_hours_calc = reports_per_year * 312
    report_dollars_calc = report_hours_calc * hourly_rate
    
    # Display Results
    st.header("💡 Your ROI Results")
    
    # Detailed breakdown with checkboxes IN THE TABLE
    st.subheader("📋 Detailed Savings Breakdown")
    
    # Create table with checkboxes
    table_col1, table_col2, table_col3, table_col4 = st.columns([1, 3, 2, 2])
    
    with table_col1:
        st.write("**Include**")
    with table_col2:
        st.write("**Automation Item**")
    with table_col3:
        st.write("**Hours Saved**")
    with table_col4:
        st.write("**$ Saved**")
    
    # Row 1: Warranty
    r1_col1, r1_col2, r1_col3, r1_col4 = st.columns([1, 3, 2, 2])
    with r1_col1:
        chk_warranty = st.checkbox("", value=True, key="chk1", label_visibility="collapsed")
    with r1_col2:
        st.write("Warranty Requests Response Automation")
    with r1_col3:
        st.write(f"{warranty_hours_calc if chk_warranty else 0:,.1f}")
    with r1_col4:
        st.write(f"${warranty_dollars_calc if chk_warranty else 0:,.0f}")
    
    # Row 2: Licence Spend
    r2_col1, r2_col2, r2_col3, r2_col4 = st.columns([1, 3, 2, 2])
    with r2_col1:
        chk_licence_spend = st.checkbox("", value=True, key="chk2", label_visibility="collapsed")
    with r2_col2:
        st.write("Enterprise Software Licence Spend Optimisation")
    with r2_col3:
        st.write("-")
    with r2_col4:
        st.write(f"${licence_spend_savings_calc if chk_licence_spend else 0:,.0f}")
    
    # Row 3: Asset
    r3_col1, r3_col2, r3_col3, r3_col4 = st.columns([1, 3, 2, 2])
    with r3_col1:
        chk_asset = st.checkbox("", value=True, key="chk3", label_visibility="collapsed")
    with r3_col2:
        st.write("Asset Discovery & Inventory")
    with r3_col3:
        st.write(f"{asset_hours_calc if chk_asset else 0:,.1f}")
    with r3_col4:
        st.write(f"${asset_dollars_calc if chk_asset else 0:,.0f}")
    
    # Row 4: Change
    r4_col1, r4_col2, r4_col3, r4_col4 = st.columns([1, 3, 2, 2])
    with r4_col1:
        chk_change = st.checkbox("", value=True, key="chk4", label_visibility="collapsed")
    with r4_col2:
        st.write("Change Detection & Config Management")
    with r4_col3:
        st.write(f"{change_hours_calc if chk_change else 0:,.1f}")
    with r4_col4:
        st.write(f"${change_dollars_calc if chk_change else 0:,.0f}")
    
    # Row 5: Vuln
    r5_col1, r5_col2, r5_col3, r5_col4 = st.columns([1, 3, 2, 2])
    with r5_col1:
        chk_vuln = st.checkbox("", value=True, key="chk5", label_visibility="collapsed")
    with r5_col2:
        st.write("Vulnerability Identification")
    with r5_col3:
        st.write(f"{vuln_hours_calc if chk_vuln else 0:,.1f}")
    with r5_col4:
        st.write(f"${vuln_dollars_calc if chk_vuln else 0:,.0f}")
    
    # Row 6: Reports
    r6_col1, r6_col2, r6_col3, r6_col4 = st.columns([1, 3, 2, 2])
    with r6_col1:
        chk_reports = st.checkbox("", value=True, key="chk6", label_visibility="collapsed")
    with r6_col2:
        st.write("Report Generation & Distribution")
    with r6_col3:
        st.write(f"{report_hours_calc if chk_reports else 0:,.1f}")
    with r6_col4:
        st.write(f"${report_dollars_calc if chk_reports else 0:,.0f}")
    
    st.divider()
    
    # Calculate totals based on checked items
    warranty_hours = warranty_hours_calc if chk_warranty else 0
    warranty_dollars = warranty_dollars_calc if chk_warranty else 0
    licence_spend_savings = licence_spend_savings_calc if chk_licence_spend else 0
    asset_hours = asset_hours_calc if chk_asset else 0
    asset_dollars = asset_dollars_calc if chk_asset else 0
    change_hours = change_hours_calc if chk_change else 0
    change_dollars = change_dollars_calc if chk_change else 0
    vuln_hours = vuln_hours_calc if chk_vuln else 0
    vuln_dollars = vuln_dollars_calc if chk_vuln else 0
    report_hours = report_hours_calc if chk_reports else 0
    report_dollars = report_dollars_calc if chk_reports else 0
    
    total_hours = warranty_hours + asset_hours + change_hours + vuln_hours + report_hours
    total_dollars = warranty_dollars + licence_spend_savings + asset_dollars + change_dollars + vuln_dollars + report_dollars
    roi_percentage = ((total_dollars - sub_cost) / sub_cost * 100) if sub_cost > 0 else 0
    
    # Top-level metrics AFTER calculating based on checkboxes
    metric_col1, metric_col2, metric_col3 = st.columns(3)
    with metric_col1:
        st.metric("Total Annual Savings", f"${total_dollars:,.0f}", delta="vs current state")
    with metric_col2:
        st.metric("Annual Investment", f"${sub_cost:,.0f}")
    with metric_col3:
        st.metric("ROI", f"{roi_percentage:.0f}%", delta=f"${total_dollars - sub_cost:,.0f} net savings")
    
    st.divider()
    
    # Create DataFrame for CSV/PDF export
    results_data = {
        'Automation Item': [
            'Warranty Requests Response Automation',
            'Enterprise Software Licence Spend Optimisation',
            'Asset Discovery & Inventory',
            'Change Detection & Config Management',
            'Vulnerability Identification',
            'Report Generation & Distribution'
        ],
        'Hours Saved': [
            f"{warranty_hours:,.1f}",
            '-',
            f"{asset_hours:,.1f}",
            f"{change_hours:,.1f}",
            f"{vuln_hours:,.1f}",
            f"{report_hours:,.1f}"
        ],
        '$ Saved': [
            f"${warranty_dollars:,.0f}",
            f"${licence_spend_savings:,.0f}",
            f"${asset_dollars:,.0f}",
            f"${change_dollars:,.0f}",
            f"${vuln_dollars:,.0f}",
            f"${report_dollars:,.0f}"
        ]
    }
    
    df = pd.DataFrame(results_data)
    
    # Summary box with navy blue background
    st.markdown(f"""
    <div class="savings-total">
        <h3>Total Annual Savings: ${total_dollars:,.0f}</h3>
        <p><strong>Total Hours Saved:</strong> {total_hours:,.0f} hours/year</p>
        <p><strong>Annual Investment:</strong> ${sub_cost:,}</p>
        <p><strong>Net Savings:</strong> ${total_dollars - sub_cost:,.0f}</p>
        <p><strong>Return on Investment:</strong> {roi_percentage:.0f}%</p>
    </div>
    """, unsafe_allow_html=True)

else:
    # Instructions when calculator hasn't been run
    st.info("👈 Enter your IT environment details in the sidebar and click 'Calculate ROI' to see your results.")
    
    with st.expander("ℹ️ How to Use This Calculator"):
        st.markdown("""
        ### Instructions
        
        1. **Enter Your IT Environment Details** in the sidebar:
           - Number of employees and IT devices
           - Average hourly rate for IT staff
           
        2. **Describe Your Current Processes**:
           - How many warranty/licence requests you handle annually
           - Time spent per request
           - Current licence spending
           - Frequency of reports and config checks
           
        3. **Enter Your Investment**:
           - Annual Open-AudIT subscription cost
           
        4. **Click 'Calculate ROI'** to see your results
        
        5. **Select which automation items to include** using the checkboxes
        
        6. **Download** your results as CSV or a formatted PDF report
        
        ### What This Calculator Shows You
        
        This calculator quantifies the time and cost savings from automating key IT operations tasks:
        - **Warranty & Licence Management**: Reduced manual processing time
        - **Software Licence Optimisation**: Eliminated overspending on unused licences
        - **Asset Discovery**: Automated inventory management
        - **Change Detection**: Streamlined configuration monitoring
        - **Vulnerability Management**: Faster identification and remediation
        - **Report Generation**: Automated compliance and audit reporting
        """)

# Footer
st.divider()
st.markdown("""
    <div style='text-align: center; color: #666; padding: 20px;'>
        <p>Questions about Open-AudIT? Contact us at sales@firstwave.com</p>
        <p style='font-size: 0.8rem;'>© 2025 FirstWave. All calculations are estimates based on industry averages.</p>
    </div>
""", unsafe_allow_html=True)
