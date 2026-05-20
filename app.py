import streamlit as st
import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io

# 1. Page Configuration
st.set_page_config(page_title="Mithila Internal Audit", page_icon="🔐", layout="centered")

# 2. Session State Initialization
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "failed_attempts" not in st.session_state:
    st.session_state.failed_attempts = 0
if "blocked" not in st.session_state:
    st.session_state.blocked = False
if "current_step" not in st.session_state:
    st.session_state.current_step = 0

# App Name Constant
APP_NAME = "mithila_internal_audit"

# 3. Security PIN Calculation (Day + Month + "98")
now = datetime.datetime.now()
correct_pin = f"{now.day + now.month}98"

# --- LOGIN SCREEN ---
if not st.session_state.authenticated:
    st.title("🔐 Mithila Laghubitta Audit Security")
    st.subheader("Mithila Internal Audit System")

    if st.session_state.blocked:
        st.error("❌ ACCESS DENIED FOR TODAY. Maximum attempts reached! Please try again tomorrow.")
    else:
        pin_input = st.text_input("ENTER PIN", type="password", help="Contact Mithila Laghubitta for PIN")
        
        if st.button("Verify & Access", type="primary"):
            if pin_input == correct_pin:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.session_state.failed_attempts += 1
                if st.session_state.failed_attempts >= 3:
                    st.session_state.blocked = True
                    st.error("❌ Maximum attempts reached! App blocked until tomorrow.")
                    st.rerun()
                else:
                    st.error(f"❌ Wrong Pin! (Attempts left: {3 - st.session_state.failed_attempts}). Please contact with Mithila Laghubitta.")

    st.caption("Copyright Roshan Gurung 9845118748")

# --- MAIN AUDIT APP SCREEN ---
else:
    st.title("📋 MLBSL Checklist & Audit Form")
    
    # Step Selection Sidebar or Top Progress
    step = st.radio("Audit Navigation Steps", ["Step 1: General Info", "Step 2: Financial Metrics", "Step 3: Cash Count", "Step 4: Checklist Registers"], index=st.session_state.current_step, horizontal=True)

    # --- STEP 1: GENERAL INFO ---
    if step == "Step 1: General Info":
        st.header("🏢 Branch & Audit Information")
        branch_name = st.text_input("Branch Name", value=st.session_state.get('branch_name', ''))
        branch_address = st.text_input("Branch Address", value=st.session_state.get('branch_address', ''))
        
        col1, col2 = st.columns(2)
        with col1:
            audit_p_from = st.date_input("Audit Period From", value=st.session_state.get('audit_p_from', datetime.date.today()))
            audit_d_from = st.date_input("Audit Date From", value=st.session_state.get('audit_d_from', datetime.date.today()))
        with col2:
            audit_p_to = st.date_input("Audit Period To", value=st.session_state.get('audit_p_to', datetime.date.today()))
            audit_d_to = st.date_input("Audit Date To", value=st.session_state.get('audit_d_to', datetime.date.today()))
            
        # Save to state
        st.session_state.branch_name = branch_name
        st.session_state.branch_address = branch_address
        st.session_state.audit_p_from = audit_p_from
        st.session_state.audit_p_to = audit_p_to
        st.session_state.audit_d_from = audit_d_from
        st.session_state.audit_d_to = audit_d_to

    # --- STEP 2: FINANCIAL METRICS ---
    elif step == "Step 2: Financial Metrics":
        st.header("📊 Branch Portfolio & Data Metrics")
        
        col1, col2 = st.columns(2)
        with col1:
            no_of_staffs = st.number_input("No of Staffs", min_value=0, step=1, value=st.session_state.get('no_of_staffs', 0))
            no_of_centers = st.number_input("No of Centers", min_value=0, step=1, value=st.session_state.get('no_of_centers', 0))
            no_of_members = st.number_input("No of Members", min_value=0, step=1, value=st.session_state.get('no_of_members', 0))
            no_of_borrowers = st.number_input("No of Borrowers", min_value=0, step=1, value=st.session_state.get('no_of_borrowers', 0))
        with col2:
            loan_portfolio = st.number_input("Loan Portfolio (Rs.)", min_value=0.0, step=100.0, value=st.session_state.get('loan_portfolio', 0.0))
            saving_balance = st.number_input("Saving Balance (Rs.)", min_value=0.0, step=100.0, value=st.session_state.get('saving_balance', 0.0))
            overdue_loan = st.number_input("Overdue Loan (Rs.)", min_value=0.0, step=100.0, value=st.session_state.get('overdue_loan', 0.0))
            watchlist = st.number_input("Watchlist (Rs.)", min_value=0.0, step=100.0, value=st.session_state.get('watchlist', 0.0))

        # Dynamic Calculations
        npa_val = (overdue_loan / loan_portfolio * 100) if loan_portfolio > 0 else 0.0
        prod_val = (loan_portfolio / no_of_staffs) if no_of_staffs > 0 else 0.0
        spc_val = (no_of_centers / no_of_staffs) if no_of_staffs > 0 else 0.0

        # NPA Metrics Check logic
        if npa_val <= 1.0:
            npa_status, npa_color = "Excellent", "green"
        elif npa_val <= 3.0:
            npa_status, npa_color = "Good", "blue"
        else:
            npa_status, npa_color = "Needs Improvement / High Risk", "red"

        st.markdown(f"**NPA Percentage:** :{npa_color}[{npa_val:.2f} % ({npa_status})]")
        st.metric("Productivity (Portfolio/Staff)", f"Rs. {prod_val:,.2f}")
        st.metric("Single Point Contact (Center/Staff)", f"{spc_val:.2f}")

        # Save to state
        st.session_state.no_of_staffs = no_of_staffs
        st.session_state.no_of_centers = no_of_centers
        st.session_state.no_of_members = no_of_members
        st.session_state.no_of_borrowers = no_of_borrowers
        st.session_state.loan_portfolio = loan_portfolio
        st.session_state.saving_balance = saving_balance
        st.session_state.overdue_loan = overdue_loan
        st.session_state.watchlist = watchlist
        st.session_state.npa_str = f"{npa_val:.2f} % ({npa_status})"

    # --- STEP 3: CASH COUNT ---
    elif step == "Step 3: Cash Count":
        st.header("💵 Cash Verification / Note Count")
        notes = [1000, 500, 100, 50, 20, 10, 5, 2, 1]
        
        if 'note_counts' not in st.session_state:
            st.session_state.note_counts = {n: 0 for n in notes}

        total_cash = 0.0
        st.markdown("Enter quantity for each denomination:")
        
        grid_cols = st.columns(3)
        for idx, n in enumerate(notes):
            col_target = grid_cols[idx % 3]
            with col_target:
                count = st.number_input(f"Rs. {n} Note", min_value=0, step=1, value=st.session_state.note_counts[n], key=f"note_{n}")
                st.session_state.note_counts[n] = count
                total_cash += count * n

        st.subheader(f"Total Cash Calculated: Rs. {total_cash:,.2f}")
        st.session_state.total_cash = total_cash

    # --- STEP 4: CHECKLIST REGISTERS & PDF EXPORT ---
    elif step == "Step 4: Checklist Registers":
        st.header("📑 Control Registers & Books Checklist")
        
        items = [
            "1. Attendance Register", "2. Field Register", "3. Leave Record Register",
            "4. Staff Leave Reg/File", "5. Day Book", "6. Cash Register",
            "7. Cash In Transit Register", "8. Check Book Register", "9. Vault Key Register",
            "10. Loan Deed Register", "11. Loan Sub-committee Reg.", "12. Loan Utilization Register",
            "13. Delay Follow-up Register", "14. Collateral Loan Register", "15. Darta/Chalani Register",
            "16. Darta/Chalani File", "17. Circular File", "18. Staff Meeting Register",
            "19. Staff Target/Progress File", "20. Stationery Stock Register", "21. Fixed Assets Register",
            "22. Low-Cost Assets Register", "23. Guest Register", "24. Area/HO Inspection Report/File",
            "25. Center Meeting File", "26. Nagadi Rasid Bharpai", "27. Complaint Box/Register",
            "28. Passbook Distr. Register", "29. Passbook Verification Reg.", "30. Member Registration Reg.",
            "31. Dropout Register", "32. Insurance Death Claim Reg."
        ]

        if 'checklist_data' not in st.session_state:
            st.session_state.checklist_data = {i: {"raised": True, "updated": True, "remarks": ""} for i in items}

        for item in items:
            with st.expander(item, expanded=False):
                c1, c2 = st.columns(2)
                with c1:
                    raised = st.checkbox("Raised / Maintained", value=st.session_state.checklist_data[item]["raised"], key=f"r_{item}")
                    updated = st.checkbox("Up to Date", value=st.session_state.checklist_data[item]["updated"], key=f"u_{item}")
                with c2:
                    remark = st.text_input("Remarks", value=st.session_state.checklist_data[item]["remarks"], key=f"rem_{item}")
                
                st.session_state.checklist_data[item] = {"raised": raised, "updated": updated, "remarks": remark}

        st.markdown("---")
        st.subheader("🏁 Finish and Generate Audit Report")

        # ReportLAB PDF Generation Mechanism
        def generate_pdf():
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            story = []
            styles = getSampleStyleSheet()
            
            # Custom Styles
            title_style = ParagraphStyle(name='TitleStyle', fontSize=18, leading=22, alignment=1, spaceAfter=15, textColor=colors.HexColor("#1A365D"))
            body_style = getSampleStyleSheet()['BodyText']
            
            story.append(Paragraph("<b>MITHILA INTERNAL AUDIT REPORT</b>", title_style))
            story.append(Paragraph(f"<b>Branch Name:</b> {st.session_state.get('branch_name', 'N/A')} | <b>Address:</b> {st.session_state.get('branch_address', 'N/A')}", body_style))
            story.append(Paragraph(f"<b>Audit Period:</b> {st.session_state.get('audit_p_from')} to {st.session_state.get('audit_p_to')} | <b>Audit Date:</b> {st.session_state.get('audit_d_from')} to {st.session_state.get('audit_d_to')}", body_style))
            story.append(Spacer(1, 15))
            
            # Portfolio Data Table
            story.append(Paragraph("<b>Financial Metrics & Indicators:</b>", styles['Heading3']))
            fin_data = [
                ["Metric", "Value", "Metric", "Value"],
                ["Total Staffs", str(st.session_state.get('no_of_staffs', 0)), "Loan Portfolio", f"Rs. {st.session_state.get('loan_portfolio', 0.0):,}"],
                ["Total Centers", str(st.session_state.get('no_of_centers', 0)), "Saving Balance", f"Rs. {st.session_state.get('saving_balance', 0.0):,}"],
                ["Overdue Loan", f"Rs. {st.session_state.get('overdue_loan', 0.0):,}", "NPA % Status", st.session_state.get('npa_str', '0.00 %')],
                ["Total Verified Cash", f"Rs. {st.session_state.get('total_cash', 0.0):,}", "-", "-"]
            ]
            t_fin = Table(fin_data, colWidths=[130, 130, 130, 130])
            t_fin.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('PADDING', (0,0), (-1,-1), 6),
            ]))
            story.append(t_fin)
            story.append(Spacer(1, 15))

            # Checklist Table
            story.append(Paragraph("<b>Registers & System Checklist Details:</b>", styles['Heading3']))
            chk_table_data = [["Register / Book Item", "Maintained", "Up to Date", "Remarks"]]
            
            for item in items:
                v = st.session_state.checklist_data[item]
                chk_table_data.append([
                    item, 
                    "Yes" if v["raised"] else "No", 
                    "Yes" if v["updated"] else "No", 
                    v["remarks"]
                ])
                
            t_chk = Table(chk_table_data, colWidths=[200, 70, 70, 180])
            t_chk.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A365D")),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('PADDING', (0,0), (-1,-1), 4),
                ('FONTSIZE', (0,0), (-1,-1), 9),
            ]))
            story.append(t_chk)
            
            doc.build(story)
            buffer.seek(0)
            return buffer

        # Download Button
        pdf_file = generate_pdf()
        st.download_button(
            label="📥 Download Audit PDF Report",
            data=pdf_file,
            file_name=f"Audit_Report_{st.session_state.get('branch_name','Branch')}.pdf",
            mime="application/pdf"
        )
        
        if st.button("Logout / Reset System", type="secondary"):
            st.session_state.authenticated = False
            st.session_state.current_step = 0
            st.rerun()

    # Footer Copyright text
    st.markdown("---")
    st.caption("Mithila Internal Audit Web Application • Developed with Streamlit")
