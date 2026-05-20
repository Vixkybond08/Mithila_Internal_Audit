import streamlit as st
import pandas as pd
import io
import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="Mithila Audit System - Grid", layout="wide")

# ==============================================================================
# ०. रोमनलाई नेपाली युनिकोडमा बदल्ने फङ्ग्सन
# ==============================================================================
def convert_to_nepali(text_input):
    if not text_input or text_input.strip() == "":
        return text_input
    try:
        url = f"https://google.com{text_input}&itc=ne-t-i0-und&num=1"
        response = requests.get(url, timeout=3)
        if response.status_code == 200:
            data = response.json()
            if data[0] == "SUCCESS":
                return data[1][0][1][0]
    except:
        pass
    return text_input

# ==============================================================================
# १. प्रतिवेदन सेटिङहरू (SIDEBAR)
# ==============================================================================
st.sidebar.title("अडिट प्रतिवेदन सेटिङहरू")
shakha_input = st.sidebar.text_input("शाखा कार्यालय (रोमन):", "choharwa, siraha")
shakha_name = convert_to_nepali(shakha_input)

pesh_input = st.sidebar.text_input("पेस गरिएको मिति:", "31 baishakh 2083")
pesh_miti = convert_to_nepali(pesh_input)

aud1_input = st.sidebar.text_input("लेखापरीक्षक १:", "roshan gurung")
auditor_1 = convert_to_nepali(aud1_input)

aud2_input = st.sidebar.text_input("लेखापरीक्षक २:", "suresh patel")
auditor_2 = convert_to_nepali(aud2_input)

# ==============================================================================
# मुख्य स्क्रिन लेआउट (TABS)
# ==============================================================================
tab1, tab2 = st.tabs(["📄 पाना १ (Cover Page)", "📊 पाना २ (Summary & Cash Audit)"])

with tab1:
    st.markdown("<p style='text-align: right; font-style: italic; color: gray;'>Mithila Laghubitta Bittiya Sanstha Ltd / Choharwa Branch</p>", unsafe_allow_html=True)
    st.write("---")
    st.markdown("<h1 style='text-align: center; font-size: 42px; margin-top: 50px;'>आन्तरिक लेखापरीक्षण प्रतिबेदन</h1>", unsafe_allow_html=True)
    st.markdown(f"<h3 style='text-align: center; font-size: 26px;'>शाखा कार्यालय {shakha_name}</h3>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align: center; margin-top: 30px; color: gray;'>↕</h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; font-size: 20px;'><b>पेश गरिएको :</b></h4>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px;'>मिथिला लघुवित्त वित्तीय संस्था लिमिटेड<br>केन्द्रीय कार्यालय, ढल्केवर, धनुषा ।</p>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align: center; margin-top: 20px; color: gray;'>↕</h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; font-size: 20px;'><b>प्रतिबेदन पेश गर्ने :</b></h4>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px;'>आन्तरिक लेखापरीक्षण विभाग<br>केन्द्रीय कार्यालय, ढल्केवर, धनुषा ।</p>", unsafe_allow_html=True)
    st.markdown(f"<p style='text-align: center; font-size: 16px;'><b>आ. लेखापरीक्षकहरु :</b> {auditor_1} / {auditor_2}</p>", unsafe_allow_html=True)
    st.write("---")

# ------------------------------------------------------------------------------
# TAB 2: EXCEL-STYLE GRID INTERFACE
# ------------------------------------------------------------------------------
with tab2:
    st.markdown(f"<p style='text-align: right; font-weight: bold;'>मिति: {pesh_miti}</p>", unsafe_allow_html=True)
    
    st.header("१. शाखाको संक्षिप्त जानकारी")
    st.write("💡 तालिका भित्र सिधै डबल क्लिक गरेर विवरण परिमार्जन गर्न सक्नुहुन्छ:")
    
    init_indicators = {
        "सूचकहरु (Indicators)": [
            "लेखापरिक्षण अवधि (देखि - सम्म):", "शाखामा कार्यरत कर्मचारी:", "केन्द्र संख्या (#):", 
            "सदस्य संख्या (#):", "ऋणी सदस्य संख्या (#):", "लगानीमा रहीरहेको रकम (रु.):", 
            "बचत परिचालन (रु.):", "भाखा नाघेको कर्जा (रु.):", "भाखा नाघेको कर्जा प्रतिशत:", 
            "सुक्ष्म निगरानीमा रहेको कर्जा (*):", "कर्मचारी उत्पादकत्व / केन्द्र:"
        ],
        "विवरण / तथ्याङ्क (Data)": [
            "२०८२/०४/०१ देखि २०८२/११/३० सम्म", "४ जना (१ म्यासेन्जर सहित)", "९० वटा", 
            "१२१२ जना", "८८४ जना", "१२,०१,८१,९२०.९२/-", 
            "३,०४,१४,४८७.५४/-", "२,११,००,१२३/- (१७५ जना)", "१७.५५%", 
            "३,७५,२१,९९६/-", "६०६ जना / ४५ वटा"
        ]
    }
    df_ind_edit = st.data_editor(pd.DataFrame(init_indicators), use_container_width=True, key="editor_ind")
    
    st.markdown("<small>* NPA मा नआएको तर भाखा नाघेर सुक्षम निगरानीमा रहेको कर्जा</small>", unsafe_allow_html=True)

    st.header("२. आन्तरीक लेखापरिक्षण गर्दा अपनाईएका विधिहरु तथा आधारहरु")
    st.write("क) कार्यालयको लेखा सम्बन्धका विषयबस्तुहरुको पुर्ण लेखा परिक्षण नमून परिक्षण विधिबाट गरिएको छ ।")
    st.write("ख) शाखा कार्यालयले उपलब्ध गराएको तथ्याङ्कको आधारमा यो प्रतिवेदन तयार पारिएको छ ।")

    # --------------------------------------------------------------------------
    # ३. नगद तथा ढुकुटीको निरिक्षण
    # --------------------------------------------------------------------------
    st.header("३. नगद तथा ढुकुटीको निरिक्षण")
    st.write("📋 **एक्सेल ढाँचा तालिका:** तलको **'परिमाण (Quantity)'** कोलममा सिधै दर अनुसारको संख्या टाइप गर्नुहोस्:")

    notes_base = [1000, 500, 100, 50, 20, 10, 5, 2, 1]
    
    # Initialize dictionary structure for clean caching
    if "cash_df" not in st.session_state:
        st.session_state.cash_df = pd.DataFrame({
            "विवरण (cash dino)": [f"रु. {x}" for x in notes_base],
            "परिमाण (Quantity)": [0, 0, 0, 0, 0, 0, 0, 0, 0],
            "कुल रकम (Amount)": [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
        })

    # Render data editor grid
    edited_cash_df = st.data_editor(
        st.session_state.cash_df,
        column_config={
            "विवरण (cash dino)": st.column_config.TextColumn("विवरण (cash dino)", disabled=True),
            "परिमाण (Quantity)": st.column_config.NumberColumn("परिमाण (Quantity)", min_value=0, default=0),
            "कुल रकम (Amount)": st.column_config.NumberColumn("कुल रकम (Amount)", disabled=True)
        },
        use_container_width=True,
        key="cash_matrix_editor_v2"
    )

    # Recalculate row amounts instantly
    total_physical_cash = 0.0
    for idx, note_val in enumerate(notes_base):
        qty = edited_cash_df.at[idx, "परिमाण (Quantity)"]
        row_amt = note_val * qty
        edited_cash_df.at[idx, "कुल रकम (Amount)"] = row_amt
        total_physical_cash += row_amt

    st.markdown(f"### 💰 ढुकुटीमा फेला परेको कुल नगद: **रु. {total_physical_cash:,}/-**")
    
    col_cbs1, col_cbs2 = st.columns(2)
    with col_cbs1:
        system_cash = st.number_input("सफ्टवेयर (CBS) मा देखिएको नगद मौज्दात (रु.):", min_value=0.0, value=0.0, step=1.0)
    with col_cbs2:
        inspection_date = st.text_input("भौतिक निरीक्षण विवरण:", "मिति २०८३/०१/०६ (१०:१५ बजे)")

    farak_amount = total_physical_cash - system_cash
    if farak_amount == 0:
        remarks_str = "मिलेको (दुरुस्त रहेको)"
    elif farak_amount < 0:
        remarks_str = f"रु. {abs(farak_amount):,}/- ले ढुकुटीमा कमी फेला परेको"
    else:
        remarks_str = f"रु. {farak_amount:,}/- ले ढुकुटीमा बढी फेला परेको"

    recon_rows = {
        "विवरण (Reconciliation)": ["भौतिक निरीक्षणमा पाइएको नगद", "सफट्वेयर (CBS) मा देखिएको नगद", "फरक रकम (Difference)", "कैफियत (Remarks)"],
        "रकम / विवरण": [f"रु. {total_physical_cash:,}/-", f"रु. {system_cash:,}/-", f"रु. {farak_amount:,}/-", remarks_str]
    }
    st.table(pd.DataFrame(recon_rows).set_index("विवरण (Reconciliation)"))

    # --------------------------------------------------------------------------
    # ३.१ Day End तथा दैनिक कारोबार
    # --------------------------------------------------------------------------
    st.header("३.१ Day End तथा दैनिक कारोबार")
    st.write("📅 तलको तालिकामा दिन संख्या सिधै परिवर्तन गर्नुहोस:")
    
    init_dayend = {
        "Day End अवस्था विवरण": [
            "समयमै अर्थात् कारोबार भएकै दिन Day End गरेको जम्मा दिन:",
            "१ दिन देखि ४ दिन सम्म ढिला Day End गरेको जम्मा दिन:",
            "५ दिन देखि ९ दिन सम्म ढिला Day End गरेको जम्मा दिन:",
            "१० दिन देखि १४ दिन सम्म ढिला Day End गरेको जम्मा दिन:"
        ],
        "जम्मा दिन संख्या": ["९ दिन", "१०६ दिन", "८० दिन", "४६ दिन"]
    }
    df_de_edit = st.data_editor(pd.DataFrame(init_dayend), use_container_width=True, key="editor_de")

# ==============================================================================
# WORD (.DOCX) COMPILE ENGINE (FIXED INDEX BUG)
# ==============================================================================
def build_word_document():
    doc = Document()
    
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(1), Inches(1)
        section.left_margin, section.right_margin = Inches(1), Inches(1)

    # Page 1 Cover
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_top.add_run("Mithila Laghubitta / Internal Audit Department").italic = True
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run("\n\nआन्तरिक लेखापरीक्षण प्रतिबेदन\n").font.size = Pt(28)
    p_title.add_run(f"शाखा कार्यालय {shakha_name}\n").font.size = Pt(18)
    
    p_p = doc.add_paragraph()
    p_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p.add_run("\n पेश गरिएको : \n").bold = True
    p_p.add_run("मिथिला लघुवित्त वित्तीय संस्था लिमिटेड\nकेन्द्रीय कार्यालय, ढल्केवर, धनुषा ।\n")
    p_p.add_run(f"\nआ. लेखापरीक्षकहरु : {auditor_1} / {auditor_2}").bold = True
    
    doc.add_page_break()
    
    # Page 2 Summary
    doc.add_heading("१. शाखाको संक्षिप्त जानकारी", level=1)
    t1 = doc.add_table(rows=1, cols=2)
    t1.style = 'Table Grid'
    hdr_cells1 = t1.rows[0].cells
    hdr_cells1[0].text, hdr_cells1[1].text = 'सूचकहरु', 'विवरण'
    
    for _, row in df_ind_edit.iterrows():
        rc = t1.add_row().cells
        rc[0].text = str(row["सूचकहरु (Indicators)"])
        rc[1].text = str(row["विवरण / तथ्याङ्क (Data)"])

    doc.add_heading("२. विधि तथा आधारहरु", level=1)
    doc.add_paragraph("क) पूर्ण तथा नमूना परिक्षण विधिबाट लेखा परीक्षण गरिएको छ।")
    
    # Page 2 Cash
    doc.add_heading("३. नगद तथा ढुकुटीको निरिक्षण", level=1)
    t2 = doc.add_table(rows=1, cols=3)
    t2.style = 'Table Grid'
    hdr_cells2 = t2.rows[0].cells
    hdr_cells2[0].text, hdr_cells2[1].text, hdr_cells2[2].text = 'cash dino', 'Quantity', 'Amount'
    
    for idx, row in edited_cash_df.iterrows():
        rc = t2.add_row().cells
        rc[0].text = str(row["विवरण (cash dino)"])
        rc[1].text = str(row["परिमाण (Quantity)"])
        rc[2].text = f"रु. {row['कुल रकम (Amount)']:,}/-"
        
    doc.add_paragraph("\n")
    t3 = doc.add_table(rows=1, cols=3)
    t3.style = 'Table Grid'
    hdr_cells3 = t3.rows[0].cells
    hdr_cells3[0].text, hdr_cells3[1].text, hdr_cells3[2].text = 'विवरण', 'मौज्दात रकम', 'कैफियत'
    
    r1 = t3.add_row().cells
    r1[0].text, r1[1].text, r1[2].text = "भौतिक नगद मौज्दात", f"रु. {total_physical_cash:,}/-", inspection_date
    r2 = t3.add_row().cells
    r2[0].text, r2[1].text, r2[2].text = "सफट्वेयर (CBS) नगद", f"रु. {system_cash:,}/-", ""
    r3 = t3.add_row().cells
    r3[0].text, r3[1].text, r3[2].text = "फरक रकम", f"रु. {farak_amount:,}/-", remarks_str

    # Page 2 Dayend
    doc.add_heading("३.१ Day End तथा दैनिक कारोबार", level=1)
    t4 = doc.add_table(rows=1, cols=2)
    t4.style = 'Table Grid'
    hdr_cells4 = t4.rows[0].cells
    hdr_cells4[0].text, hdr_cells4[1].text = 'Day End अवस्था विवरण', 'जम्मा दिन संख्या'
    
    for _, row in df_de_edit.iterrows():
        rc = t4.add_row().cells
        rc[0].text = str(row["Day End अवस्था विवरण"])
        rc[1].text = str(row["जम्मा दिन संख्या"])

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# Export button
st.sidebar.markdown("---")
st.sidebar.download_button(
    label="📝 Word (.docx) डाउनलोड गर्नुहोस्",
    data=build_word_document(),
    file_name=f"Mithila_Audit_Report_{shakha_name}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
