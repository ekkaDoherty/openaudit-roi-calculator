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

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f4788;
        font-weight: bold;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f8ff;
        padding: 20px;
        border-radius: 10px;
        border: 2px solid #1f4788;
    }
    .savings-total {
        background-color: #d4edda;
        padding: 20px;
        border-radius: 10px;
        border: 2px solid #28a745;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown('<p class="main-header">Open-AudIT ROI Calculator</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Calculate your return on investment with Open-AudIT automation</p>', unsafe_allow_html=True)

# Sidebar for inputs
with st.sidebar:
    st.header("📊 Your IT Environment")
    
    num_employees = st.number_input("Number of Employees", min_value=0, value=4500, step=100)
    num_devices = st.number_input("Number of IT Devices", min_value=0, value=45000, step=1000)
    hourly_rate = st.number_input("Average Hourly Rate of IT Staff ($/hr)", min_value=0.0, value=50.0, step=5.0)
    
    st.divider()
    st.header("🔧 Current Processes")
    
    licence_requests = st.number_input("Warranty/Licence Requests per Year", min_value=0, value=1000, step=50)
    licence_hours = st.number_input("Avg Processing Time per Licence Request (hrs)", min_value=0.0, value=0.5, step=0.1)
    licence_spend = st.number_input("Total Current Licence Spend ($/yr)", min_value=0, value=5000000, step=100000)
    reports_per_year = st.number_input("Asset & Inventory Reports per Year", min_value=0, value=12, step=1)
    checks_per_year = st.number_input("Change Detection & Config Mgmt Checks per Year", min_value=0, value=12, step=1)
    
    st.divider()
    st.header("💵 Investment")
    
    sub_cost = st.number_input("Open-AudIT Annual Subscription Cost ($)", min_value=0, value=1000000, step=10000)

# Initialize session state for checkboxes
if 'show_calculations' not in st.session_state:
    st.session_state.show_calculations = False

# Calculate button
if st.sidebar.button("Calculate ROI", type="primary", use_container_width=True):
    st.session_state.show_calculations = True

# Calculations
if st.session_state.show_calculations:
    
    # Warranty Requests Response Automation
    warranty_hours = licence_requests * licence_hours * SAVING_PCT_LICENSE
    warranty_dollars = warranty_hours * hourly_rate
    
    # Enterprise Software Licence Spend Optimisation
    licence_spend_savings = licence_spend * LICENCE_SPEND_REDUCTION_PCT
    
    # Asset Discovery & Inventory
    asset_hours = ((num_devices * MIN_PER_DEVICE_DISCOVERY / 60.0) + 
                   (reports_per_year * HOURS_PER_ASSET_REPORT)) * SAVING_PCT_ASSET
    asset_dollars = asset_hours * hourly_rate
    
    # Change Detection & Config Management
    critical_devices = num_devices * CRITICAL_DEVICE_PCT
    change_hours = critical_devices * checks_per_year * ((MIN_PER_CHECK_MANUAL - MIN_PER_CHECK_AUTOMATED) / 60.0)
    change_hours = max(0, change_hours)
    change_dollars = change_hours * hourly_rate
    
    # Vulnerability Identification
    vuln_hours = num_devices * MIN_PER_DEVICE_VULN_PER_YEAR / 60.0
    vuln_dollars = vuln_hours * hourly_rate
    
    # Report Generation & Distribution (simplified calculation for demo)
    report_hours = reports_per_year * 312  # Based on your G12 formula
    report_dollars = report_hours * hourly_rate
    
    # Totals
    total_hours = warranty_hours + asset_hours + change_hours + vuln_hours + report_hours
    total_dollars = warranty_dollars + licence_spend_savings + asset_dollars + change_dollars + vuln_dollars + report_dollars
    
    roi_percentage = ((total_dollars - sub_cost) / sub_cost * 100) if sub_cost > 0 else 0
    
    # Display Results
    st.header("💡 Your ROI Results")
    
    # Top-level metrics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Annual Savings", f"${total_dollars:,.0f}", delta="vs current state")
    with col2:
        st.metric("Annual Investment", f"${sub_cost:,.0f}")
    with col3:
        st.metric("ROI", f"{roi_percentage:.0f}%", delta=f"${total_dollars - sub_cost:,.0f} net savings")
    
    st.divider()
    
    # Detailed breakdown
    st.subheader("📋 Detailed Savings Breakdown")
    
    # Create DataFrame for results
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
    
    # Display as formatted table
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )
    
    # Summary box
    st.markdown(f"""
    <div class="savings-total">
        <h3>Total Annual Savings: ${total_dollars:,.0f}</h3>
        <p><strong>Total Hours Saved:</strong> {total_hours:,.0f} hours/year</p>
        <p><strong>Annual Investment:</strong> ${sub_cost:,.0f}</p>
        <p><strong>Net Savings:</strong> ${total_dollars - sub_cost:,.0f}</p>
        <p><strong>Return on Investment:</strong> {roi_percentage:.0f}%</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    # Download section
    st.subheader("📥 Export Your Results")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # CSV Download
        csv = df.to_csv(index=False)
        st.download_button(
            label="Download Results (CSV)",
            data=csv,
            file_name="openaudit_roi_results.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with col2:
        # Generate PDF Report
        def create_pdf_report():
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter)
            story = []
            styles = getSampleStyleSheet()
            
            # Title
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=24,
                textColor=colors.HexColor('#1f4788'),
                spaceAfter=30
            )
            story.append(Paragraph("Open-AudIT ROI Analysis", title_style))
            story.append(Spacer(1, 0.2*inch))
            
            # Summary
            story.append(Paragraph(f"<b>Total Annual Savings:</b> ${total_dollars:,.0f}", styles['Normal']))
            story.append(Paragraph(f"<b>Annual Investment:</b> ${sub_cost:,.0f}", styles['Normal']))
            story.append(Paragraph(f"<b>Net Savings:</b> ${total_dollars - sub_cost:,.0f}", styles['Normal']))
            story.append(Paragraph(f"<b>ROI:</b> {roi_percentage:.0f}%", styles['Normal']))
            story.append(Spacer(1, 0.3*inch))
            
            # Table
            table_data = [['Automation Item', 'Hours Saved', '$ Saved']]
            for _, row in df.iterrows():
                table_data.append([row['Automation Item'], row['Hours Saved'], row['$ Saved']])
            
            t = Table(table_data, colWidths=[3.5*inch, 1.5*inch, 1.5*inch])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            story.append(t)
            
            doc.build(story)
            buffer.seek(0)
            return buffer
        
        pdf_buffer = create_pdf_report()
        st.download_button(
            label="Download Report (PDF)",
            data=pdf_buffer,
            file_name="openaudit_roi_report.pdf",
            mime="application/pdf",
            use_container_width=True
        )

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
        
        5. **Download** your results as CSV or a formatted PDF report
        
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
        <p>Questions about Open-AudIT? Contact us at sales@openaudit.com</p>
        <p style='font-size: 0.8rem;'>© 2025 Open-AudIT. All calculations are estimates based on industry averages.</p>
    </div>
""", unsafe_allow_html=True)