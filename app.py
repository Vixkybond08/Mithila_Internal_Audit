import streamlit as st
import pandas as pd
import io
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# १. नेपाली फन्ट रजिस्टर गर्ने (PDF को लागि)
# तपाईंको फोल्डरमा भएको Kalimati.ttf वा Mangal.ttf को नाम यहाँ ठ्याक्कै मिल्नुपर्छ
try:
    pdfmetrics.registerFont(TTFont('NepaliFont', 'Kalimati.ttf'))
    font_available = True
except:
    font_available = False

st.set_page_config(page_title="Mithila Audit System", layout="wide")
st.title("Mithila Laghubitta - Automated Audit Report System")
st.write("Excel फाइलहरू अपलोड गर्नुहोस् र भौचर त्रुटिहरू टाइप गर्नुहोस्। प्रणालीले स्वतः एउटै प्रतिवेदन तयार गर्नेछ।")

# डकुमेन्टमा थपिने विवरण संकलन गर्ने लिस्ट
compiled_findings = []

# ==========================================
# खण्ड १: एक्सेल अपलोड (सदस्य र कर्जा विवरण)
# ==========================================
st.header("१. एक्सेल फाइल अपलोड खण्ड (Excel Upload Section)")
col1, col2 = st.columns(2)

with col1:
    st.subheader("क) सदस्य विवरण एक्सेल")
    member_file = st.file_uploader("सदस्यहरूको एक्सेल फाइल छान्नुहोस् (.xlsx)", type=["xlsx"])
    if member_file:
        df_member = pd.read_excel(member_file)
        st.write("अपलोड भएको डाटाको केही अंश:")
        st.dataframe(df_member.head(3))
        
        # ६५ वर्ष नाघेका ओभर-एज सदस्य स्वतः खोज्ने लोजिक
        # एक्सेलमा 'उमेर' वा 'Age' भन्ने कोलम हुनुपर्छ
        age_col = [c for c in df_member.columns if 'उमेर' in str(c) or 'Age' in str(c)]
        name_col = [c for c in df_member.columns if 'नाम' in str(c) or 'Name' in str(c)]
        code_col = [c for c in df_member.columns if 'कोड' in str(c) or 'Code' in str(c)]
        
        if age_col and name_col:
            over_aged = df_member[df_member[age_col[0]] > 65]
            if not over_aged.empty:
                st.warning(f"त्रुटि भेटियो: ६५ वर्ष नाघेका {len(over_aged)} जना सदस्य भेटिए!")
                for idx, row in over_aged.iterrows():
                    m_code = row[code_col[0]] if code_col else idx
                    finding = f"सदस्य कोड {m_code}, नाम {row[name_col[0]]} को उमेर {row[age_col[0]]} वर्ष पाइयो। संस्थाको नीति बमोजिम १८ देखि ६५ वर्ष भित्र हुनुपर्नेमा सो विपरित सदस्यता कायम रहेको।"
                    compiled_findings.append({"शीर्षक": "सदस्य उमेर सम्बन्धमा", "विवरण": finding})

with col2:
    st.subheader("ख) कर्जा विवरण एक्सेल")
    loan_file = st.file_uploader("कर्जाको एक्सेल फाइल छान्नुहोस् (.xlsx)", type=["xlsx"])
    if loan_file:
        df_loan = pd.read_excel(loan_file)
        st.write("अपलोड भएको डाटाको केही अंश:")
        st.dataframe(df_loan.head(3))
        
        # एक्सेलमा 'कैफियत' वा 'Remarks' कोलममा भएका गल्ती संकलन गर्ने
        rem_col = [c for c in df_loan.columns if 'कैफियत' in str(c) or 'Remarks' in str(c)]
        l_name_col = [c for c in df_loan.columns if 'नाम' in str(c) or 'Name' in str(c)]
        
        if rem_col and l_name_col:
            for idx, row in df_loan.dropna(subset=[rem_col[0]]).iterrows():
                finding = f"ऋणी सदस्य {row[l_name_col[0]]} को कर्जामा कैफियत: {row[rem_col[0]]}।"
                compiled_findings.append({"शीर्षक": "कर्जा लगानी तथा स्वीकृती सम्बन्धमा", "विवरण": finding})

# ==========================================
# खण्ड २: म्यानुअल टाइप (भौचर त्रुटिहरू)
# ==========================================
st.header("२. भौचर त्रुटि प्रविष्टि (Voucher Errors Entry)")
st.write("एक्सेल बाहेकका अन्य फुटकर वा भौचरका गल्तीहरू यहाँ टाइप गर्नुहोस्:")

if 'vouchers' not in st.session_state:
    st.session_state.vouchers = []

# फारम मार्फत भौचर गल्ती थप्ने
with st.form("voucher_form", clear_on_submit=True):
    v_date = st.text_input("कारोबार मिति (जस्तै: २०८२/१०/२९):")
    v_error = st.text_area("भेटिएको त्रुटि/कैफियत (जस्तै: TDS कट्टी नगरेको, बिल नभएको):")
    v_suggest = st.text_area("सुझाव/निर्देशन:")
    submit_v = st.form_submit_button("यो भौचर त्रुटि थप्नुहोस्")
    
    if submit_v and v_error:
        st.session_state.vouchers.append({"मिति": v_date, "त्रुटि": v_error, "सुझाव": v_suggest})

# थपिएका भौचर गल्तीहरू स्क्रिनमा देखाउने
if st.session_state.vouchers:
    st.write("तपाईंले थप्नुभएका भौचर गल्तीहरू:")
    for v in st.session_state.vouchers:
        st.info(f"📅 मिति: {v['मिति']} | ❌ कैफियत: {v['त्रुटि']}")

# ==========================================
# खण्ड ३: प्रतिवेदन कम्पाइल र डाउनलोड
# ==========================================
st.header("३. प्रतिवेदन कम्पाइल र डाउनलोड खण्ड")

# सबै डाटालाई एउटै टेक्स्ट ढाँचामा ल्याउने फंक्शन
def generate_report_text():
    lines = ["आन्तरिक लेखापरीक्षण प्रतिबेदन (Compiled Audit Report)", "="*50, "\n"]
    
    if compiled_findings:
        lines.append("क) एक्सेल अपलोडबाट फेला परेका कैफियतहरू:\n")
        for f in compiled_findings:
            lines.append(f"📌 {f['शीर्षक']}: {f['विवरण']}")
        lines.append("\n" + "-"*40 + "\n")
        
    if st.session_state.vouchers:
        lines.append("ख) भौचर जाँच गर्दा फेला परेका कैफियत तथा सुझावहरू:\n")
        for v in st.session_state.vouchers:
            lines.append(f"📅 मिति: {v['मिति']}")
            lines.append(f"❌ कैफियत: {v['त्रुटि']}")
            lines.append(f"💡 सुझाव: {v['सुझाव']}\n")
            
    if not compiled_findings and not st.session_state.vouchers:
        lines.append("कुनै पनि कैफियतहरू प्रविष्टि गरिएको छैन।")
        
    return "\n".join(lines)

report_content = generate_report_text()
st.text_area("तयार भएको प्रतिवेदनको मस्यौदा (Preview):", report_content, height=250)

# १. Word (.docx) फाइल बनाउने लोजिक
def make_docx():
    doc = Document()
    doc.add_heading('आन्तरिक लेखापरीक्षण प्रतिबेदन', 0)
    doc.add_paragraph(report_content)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# २. PDF फाइल बनाउने लोजिक
def make_pdf():
    bio = io.BytesIO()
    p = canvas.Canvas(bio, pagesize=letter)
    
    # यदि नेपाली फन्ट उपलब्ध छ भने प्रयोग गर्ने, नत्र डिफल्ट
    if font_available:
        p.setFont('NepaliFont', 12)
    else:
        p.setFont('Helvetica', 12)
        
    y = 750
    p.drawString(50, y, "Mithila Laghubitta Bittiya Sanstha Ltd - Compiled Audit Report")
    y -= 30
    
    for line in report_content.split('\n'):
        if y < 50:  # नयाँ पाना थप्ने लोजिक
            p.showPage()
            if font_available: p.setFont('NepaliFont', 12)
            y = 750
        p.drawString(50, y, line)
        y -= 18
        
    p.showPage()
    p.save()
    bio.seek(0)
    return bio

# डाउनलोड बटनहरू
col_d1, col_d2 = st.columns(2)
with col_d1:
    st.download_button(
        label="📝 Word (.docx) फाइल डाउनलोड गर्नुहोस्",
        data=make_docx(),
        file_name="Audit_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

with col_d2:
    st.download_button(
        label="📕 PDF (.pdf) फाइल डाउनलोड गर्नुहोस्",
        data=make_pdf(),
        file_name="Audit_Report.pdf",
        mime="application/pdf"
    )
