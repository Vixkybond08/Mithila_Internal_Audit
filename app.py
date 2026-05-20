import streamlit as st
import pandas as pd
import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

st.set_page_config(page_title="Mithila Audit System - Page 1 & 2", layout="wide")

# ==============================================================================
# १. प्रयोगकर्ता इनपुट संकलन खण्ड (SIDEBAR & CONTROLS FOR USER ENTRY)
# ==============================================================================
st.sidebar.title("अडिट प्रतिवेदन सेटिङहरू")

st.sidebar.header("पाना १: कभर पेज विवरण")
shakha_name = st.sidebar.text_input("शाखा कार्यालयको नाम:", "चोहर्वा, सिरहा")
pesh_miti = st.sidebar.text_input("पेस गरिएको मिति:", "३१ बैशाख २०Check")
auditor_1 = st.sidebar.text_input("आन्तरिक लेखापरीक्षक १:", "रोशन गुरुङ्ग (ब.अधिकृत)")
auditor_2 = st.sidebar.text_input("आन्तरिक लेखापरीक्षक २:", "सन्तोष पटेल (सह-अधिकृत)")

st.sidebar.header("पाना २: संक्षिप्त जानकारी मितिहरू")
audit_start_date = st.sidebar.text_input("लेखापरिक्षण सुरु मिति (देखि):", "२०Check/०४/०१")
audit_end_date = st.sidebar.text_input("लेखापरिक्षण अन्तिम मिति (सम्म):", "२०Check/११/३०")
audit_duration_days = st.sidebar.text_input("लागेको समय (दिन):", "३ दिन")
audit_period_str = st.sidebar.text_input("लेखापरिक्षण कार्य अवधि:", "२०Check/०१/०७ देखी २०Check/०१/०९ सम्म")

# ==============================================================================
# मुख्य स्क्रिन लेआउट (MAIN INTERFACE DISPLAYING THE PDF LOOK-ALIKE)
# ==============================================================================
tabs = st.tabs(["📄 पाना १ (Cover Page)", "📊 पाना २ (Summary & Cash Audit)"])

# ------------------------------------------------------------------------------
# TAB 1: COVER PAGE DESIGN
# ------------------------------------------------------------------------------
with tabs[0]:
    st.markdown("<p style='text-align: right; font-style: italic; color: gray;'>Mithila Laghubitta Bittiya Sanstha Ltd / Choharwa Branch</p>", unsafe_allow_html=True)
    st.write("---")
    
    st.markdown(f"<h1 style='text-align: center; font-size: 42px; margin-top: 50px;'>आन्तरिक लेखापरीक्षण प्रतिबेदन</h1>", unsafe_allow_html=True)
    st.markdown(f"<h3 style='text-align: center; font-size: 26px;'>शाखा कार्यालय {shakha_name}</h3>", unsafe_allow_html=True)
    
    st.markdown("<h1 style='text-align: center; margin-top: 40px; color: gray;'>↕</h1>", unsafe_allow_html=True)
    
    st.markdown("<h4 style='text-align: center; font-size: 20px; margin-top: 40px;'><b>पेश गरिएको :</b></h4>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px;'>मिथिला लघुवित्त वित्तीय संस्था लिमिटेड<br>केन्द्रीय कार्यालय, ढल्केवर, धनुषा ।</p>", unsafe_allow_html=True)
    
    st.markdown("<h1 style='text-align: center; margin-top: 30px; color: gray;'>↕</h1>", unsafe_allow_html=True)
    
    st.markdown("<h4 style='text-align: center; font-size: 20px; margin-top: 30px;'><b>प्रतिबेदन पेश गर्ने :</b></h4>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px;'>आन्तरिक लेखापरीक्षण बिभाग<br>केन्द्रीय कार्यालय, ढल्केवर, धनुषा ।</p>", unsafe_allow_html=True)
    
    st.markdown(f"<p style='text-align: center; font-size: 16px;'><b>आ. लेखापरीक्षकहरु :</b> {auditor_1} / {auditor_2}</p>", unsafe_allow_html=True)
    
    st.write("---")
    st.markdown("<table style='width:100%'><tr><td><small>Mithila Laghubitta Bittiya Sanstha Ltd / Internal Audit Department</small></td><td style='text-align:right'><small>Page 1</small></td></tr></table>", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 2: PAGE 2 DESIGN WITH MANUAL INDICATORS & DYNAMIC CASH DENOMINATION
# ------------------------------------------------------------------------------
with tabs[1]:
    st.markdown("<p style='text-align: right; font-style: italic; color: gray;'>Mithila Laghubitta Bittiya Sanstha Ltd / Choharwa Branch</p>", unsafe_allow_html=True)
    st.markdown(f"<h3 style='text-align: right;'>{pesh_miti}</h3>", unsafe_allow_html=True)
    
    st.header("१. शाखाको संक्षिप्त जानकारी")
    
    # Dynamic Data Entry Fields for Indicators Table
    col_ind1, col_ind2, col_ind3 = st.columns(3)
    with col_ind1:
        staff_count = st.text_input("शाखामा कार्यरत कर्मचारी:", "४ जना, (१ जना म्यासेन्जर सहित)")
        center_count = st.text_input("केन्द्र संख्या (#):", "९० वटा")
        member_count = st.text_input("सदस्य संख्या (#):", "१२१२")
    with col_ind2:
        borrower_count = st.text_input("ऋणी सदस्य संख्या (#):", "८८४")
        total_loan = st.text_input("लगानीमा रहीरहेको रकम (रु.):", "१२,०१,८१,९२०.९२/-")
        total_saving = st.text_input("बचत परिचालन (रु.):", "३,०४,१४,४८७.५४/-")
    with col_ind3:
        npa_loan = st.text_input("भाखा नाघेको कर्जा (रु.):", "२,११,००,१२३/- (१७५ जना)")
        npa_percent = st.text_input("भाखा नाघेको कर्जाको प्रतिशत:", "१७.५५%")
        watch_list = st.text_input("सुक्ष्म निगरानीमा रहेको कर्जा (*):", "३,७५,२१,९९६/- (३०८ जना)")
        productivity = st.text_input("कर्मचारी उत्पादकत्व / प्रति कर्मचारी केन्द्र:", "६०६ जना / ४५ वटा")

    # Interactive Display Table
    indicator_data = {
        "क्र.स": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],
        "सूचकहरु": [
            "शाखा कार्यालयको नाम :", "कार्यालयको ठेगाना :", "लेखापरिक्षण गरिएको अवधि :", 
            "लेखापरिक्षणमा लागेको समयः", "शाखामा कार्यरत कर्मचारी :", "केन्द्र संख्या :", 
            "सदस्य संख्या :", "ऋणी सदस्य संख्या :", "लगानीमा रहीरहेको रकम :", 
            "बचत परिचालन :", "भाखा नाघेको कर्जा :", "भाखा नाघेको कर्जाको प्रतिशत :", 
            "सुक्ष्म निगरानीमा रहेको कर्जा :", "कर्मचारी उत्पादकत्व / प्रति कर्मचारी केन्द्र :"
        ],
        "विवरण": [
            shakha_name.split(",")[0], shakha_name, f"{audit_start_date} देखी {audit_end_date} सम्म",
            f"{audit_period_str} सम्म {audit_duration_days}", staff_count, center_count,
            member_count, borrower_count, total_loan, total_saving, npa_loan, npa_percent, watch_list, productivity
        ]
    }
    st.table(pd.DataFrame(indicator_data).set_index("क्र.स"))
    
    # Editable Footnotes
    note_star = st.text_input("फुटनोट (*):", "* NPA मा नआएको तर भाखा नाघेर सुक्षम निगरानीमा (Watch List) रहेको कर्जा")
    note_hash = st.text_input("फुटनोट (#):", f"# मिति {pesh_miti} सम्मको")
    
    st.header("२. आन्तरीक लेखापरिक्षण गर्दा अपनाईएका विधिहरु तथा आधारहरु")
    st.write(f"क) कार्यालयको लेखा सम्बन्धका विषयबस्तुहरुको पुर्ण लेखा परिक्षण तथा कार्याक्रम तथा अन्य विषयबस्तुहरुको नमून परिक्षण विधिबाट गरिएको छ ।")
    st.write(f"ख) शाखा कार्यालयले लेखा परिक्षण अवधिभरमा उपलब्ध गराएको तथ्याङ्कको आधारमा यो प्रतिवेदन तयार पारिएको छ ।")
    st.write(f"ग) मिति {audit_end_date.split('/')[0]} {audit_end_date.split('/')[1]} मसान्तको सन्तुलन परिक्षणको आधारमा लेखा सम्बन्धी कारोबारहरुको परिक्षण गरिएको छ ।")

    # --------------------------------------------------------------------------
    # ३. नगद तथा ढुकुटीको निरिक्षण (DENOMINATION CASHTABLE RECONCILIATION)
    # --------------------------------------------------------------------------
    st.header("३. नगद तथा ढुकुटीको निरिक्षण")
    st.write("ढुकुटीमा फेला परेका भौतिक नोट तथा सिक्काहरूको परिमाण (Quantity) प्रविष्ट गर्नुहोस्:")
    
    denominations = [1000, 500, 100, 50, 20, 10, 5, "सिक्का (Coins)"]
    qty_inputs = {}
    amt_calculated = {}
    
    col_c1, col_c2, col_c3 = st.columns(3)
    
    for i, denom in enumerate(denominations):
        # Multi-column grid interface for notes entry
        if i % 3 == 0: target_col = col_c1
        elif i % 3 == 1: target_col = col_c2
        else: target_col = col_c3
        
        with target_col:
            if denom == "सिक्का (Coins)":
                qty_inputs[denom] = st.number_input("सिक्का कुल मूल्य (रु.):", min_value=0, value=0, step=1, key="coin_qty")
                amt_calculated[denom] = qty_inputs[denom]
            else:
                qty_inputs[denom] = st.number_input(f"रु. {denom} को नोट संख्या:", min_value=0, value=0, step=1, key=f"note_{denom}")
                amt_calculated[denom] = denom * qty_inputs[denom]
                st.write(f"↳ उप-योग: रु. {amt_calculated[denom]:,}")

    total_physical_cash = sum(amt_calculated.values())
    
    st.subheader("लेखा प्रणालीसँग भिडान (Reconciliation)")
    col_sys1, col_sys2 = st.columns(2)
    with col_sys1:
        system_cash = st.number_input("सफ्टवेयर (CBS) मा देखिएको नगद मौज्दात (रु.):", min_value=0.0, value=0.0, step=100.0)
    with col_sys2:
        inspection_date = st.text_input("भौतिक निरीक्षण गरिएको मिति/समय:", "मिति २०८३/०१/०६")

    # Math Logic for Reconciliation
    farak_amount = total_physical_cash - system_cash
    
    if farak_amount == 0:
        remarks_str = "मिलेको (भौतिक निरीक्षण र सफ्टवेयर ब्यालेन्स दुरुस्त रहेको)"
        status_color = "green"
    elif farak_amount < 0:
        remarks_str = f"रु. {abs(farak_amount):,} ले ढुकुटीमा कमी (नगद मौज्दात कम फेला परेको)"
        status_color = "red"
    else:
        remarks_str = f"रु. {farak_amount:,} ले ढुकुटीमा बढी (भौतिक नगद बढी फेला परेको)"
        status_color = "orange"

    # Display Calculated Output Table on Screen
    st.write("### ढुकुटी वर्गीकरण तथा भिडान विवरण")
    cash_table_display = {
        "विवरण": ["भौतिक निरीक्षणमा पाइएको नगद (Total Physical Cash)", "सफ्टवेयरमा देखिएको नगद (System CBS Cash)", "फरक रकम (Reconciled Difference)", "कैफियत (Remarks)"],
        "रकम / विवरण": [f"रु. {total_physical_cash:,}/-", f"रु. {system_cash:,}/-", f"रु. {farak_amount:,}/-", remarks_str]
    }
    st.table(pd.DataFrame(cash_table_display))
    
    st.write("---")
    st.markdown("<table style='width:100%'><tr><td><small>Mithila Laghubitta Bittiya Sanstha Ltd / Internal Audit Department</small></td><td style='text-align:right'><small>Page 2</small></td></tr></table>", unsafe_allow_html=True)

# ==============================================================================
# WORD (.DOCX) GENERATOR FUNCTION FOR EXACT FORMAT MATCHING
# ==============================================================================
def build_word_document():
    doc = Document()
    
    # Adjust Margins to fit templates nicely
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # PAGE 1: COVER PAGE
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_top = p_top.add_run("Mithila Laghubitta Bittiya Sanstha Ltd / Choharwa Branch")
    run_top.font.name = 'Kalimati'
    run_top.font.size = Pt(10)
    run_top.italic = True
    
    doc.add_paragraph("\n\n")
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("आन्तरिक लेखापरीक्षण प्रतिबेदन\n")
    run_title.font.name = 'Kalimati'
    run_title.font.size = Pt(28)
    run_title.bold = True
    
    run_subtitle = p_title.add_run(f"शाखा कार्यालय {shakha_name}")
    run_subtitle.font.name = 'Kalimati'
    run_subtitle.font.size = Pt(18)
    run_subtitle.bold = True
    
    doc.add_paragraph("\n\n\n")
    p_pesh = doc.add_paragraph()
    p_pesh.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_p1 = p_pesh.add_run("पेश गरिएको :\n")
    r_p1.bold = True
    r_p1.font.size = Pt(14)
    r_p2 = p_pesh.add_run("मिथिला लघुवित्त वित्तीय संस्था लिमिटेड\nकेन्द्रीय कार्यालय, ढल्केवर, धनुषा ।\n")
    r_p2.font.size = Pt(12)
    
    doc.add_paragraph("\n\n")
    p_done = doc.add_paragraph()
    p_done.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_d1 = p_done.add_run("प्रतिबेदन पेश गर्ने :\n")
    r_d1.bold = True
    r_d1.font.size = Pt(14)
    r_d2 = p_done.add_run("आन्तरिक लेखापरीक्षण बिभाग\nकेन्द्रीय कार्यालय, ढल्केवर, धनुषा ।\n")
    r_d2.font.size = Pt(12)
    
    r_aud = p_done.add_run(f"\nआ. लेखापरीक्षकहरु : {auditor_1} / {auditor_2}")
    r_aud.bold = True
    r_aud.font.size = Pt(11)
    
    # Page Break for Page 2
    doc.add_page_break()
    
    # PAGE 2
    p2_top = doc.add_paragraph()
    p2_top.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2_top.add_run(f"{pesh_miti}\n").bold = True
    
    h1 = doc.add_heading(level=1)
    h1.add_run("१. शाखाको संक्षिप्त जानकारी").font.name = "Kalimati"
    
    # Build Table 1 (Indicators)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'क्र.स'
    hdr_cells[1].text = 'सूचकहरु'
    hdr_cells[2].text = 'विवरण'
    
    # Shading header row
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="D3D3D3"/>')
    hdr_cells[1]._tc.get_or_add_tcPr().append(shading_elm)
    
    for i in range(len(indicator_data["क्र.स"])):
        row_cells = table.add_row().cells
        row_cells[0].text = str(indicator_data["क्र.स"][i])
        row_cells[1].text = indicator_data["सूचकहरु"][i]
        row_cells[2].text = str(indicator_data["विवरण"][i])
        
    doc.add_paragraph(note_star)
    doc.add_paragraph(note_hash)
    
    doc.add_heading("२. आन्तरीक लेखापरीक्षण गर्दा अपनाईएका विधिहरु तथा आधारहरु", level=1)
    doc.add_paragraph(f"क) कार्यालयको लेखा सम्बन्धका विषयबस्तुहरुको पुर्ण लेखा परिक्षण तथा कार्याक्रम तथा अन्य विषयबस्तुहरुको नमून परिक्षण विधिबाट गरिएको छ ।")
    doc.add_paragraph(f"ख) शाखा कार्यालयले लेखा परिक्षण अवधिभरमा उपलब्ध गराएको तथ्याङ्कको आधारमा यो प्रतिवेदन तयार पारिएको छ ।")
    doc.add_paragraph(f"ग) मिति {audit_end_date.split('/')[0]} फाल्गुन मसान्तको सन्तुलन परिक्षणको आधारमा लेखा सम्बन्धी कारोबारहरुको परिक्षण गरिएको छ ।")
    
    doc.add_heading("३. नगद तथा ढुकुटीको निरिक्षण", level=1)
    
    # Cash Table
    table2 = doc.add_table(rows=1, cols=3)
    table2.style = 'Table Grid'
    t2_hdr = table2.rows[0].cells
    t2_hdr[0].text = 'विवरण'
    t2_hdr[1].text = 'मौज्दात रकम'
    t2_hdr[2].text = 'कैफियत'
    
    r1 = table2.add_row().cells
    r1[0].text = "ढुकुटी/नगद रजिष्टरमा रहेको भौतिक नगद"
    r1[1].text = f"रु. {total_physical_cash:,}/-"
    r1[2].text = inspection_date
    
    r2 = table2.add_row().cells
    r2[0].text = "सफट्वेयरमा देखिएको नगद"
    r2[1].text = f"रु. {system_cash:,}/-"
    r2[2].text = ""
    
    r3 = table2.add_row().cells
    r3[0].text = "फरक रकम (Reconciled Difference)"
    r3[1].text = f"रु. {farak_amount:,}/-"
    r3[2].text = remarks_str
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==============================================================================
# FILE DOWNLOAD ACTION BUTTON
# ==============================================================================
st.sidebar.markdown("---")
st.sidebar.subheader("प्रतिवेदन डाउनलोड गर्नुहोस्")
st.sidebar.download_button(
    label="📝 Word (.docx) फाइल डाउनलोड गर्नुहोस्",
    data=build_word_document(),
    file_name=f"Audit_Report_Page_1_2_{shakha_name.split(',')[0]}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
