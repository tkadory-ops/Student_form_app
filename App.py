import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from PIL import Image
import io
import os
from datetime import datetime

# ===== إعداد الصفحة =====
st.set_page_config(
    page_title="نموذج تسجيل طالب",
    page_icon="🎓",
    layout="centered"
)

# ===== CSS للتنسيق العربي =====
st.markdown("""
<style>
    .stApp { direction: rtl; }
    h1, h2, h3, p, label, .stMarkdown { text-align: right; direction: rtl; }
    .stTextInput input { direction: rtl; text-align: right; }
    div[data-testid="stFileUploader"] { direction: rtl; }
    
    /* تنسيق عبارة الإعداد أسفل التطبيق */
    .footer-credit {
        text-align: center;
        padding: 20px 0 10px 0;
        margin-top: 40px;
        border-top: 2px solid #1e3a8a;
        color: #1e3a8a;
        font-size: 16px;
        font-weight: bold;
        direction: rtl;
    }
</style>
""", unsafe_allow_html=True)

st.title("🎓 نموذج تسجيل طالب جامعي")
st.markdown("---")

# ===== دوال مساعدة =====
def save_uploaded_image(uploaded_file, filename):
    if uploaded_file is None:
        return None
    img = Image.open(uploaded_file)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    img.thumbnail((1500, 1500))
    path = f"/tmp/{filename}"
    img.save(path, "JPEG", quality=85)
    return path

def add_image_to_doc(doc, img_path, caption, width=Inches(4.5)):
    if img_path and os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(img_path, width=width)
        
        cap = doc.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if cap.runs:
            cap.runs[0].font.size = Pt(10)
            cap.runs[0].font.italic = True
    else:
        p = doc.add_paragraph(f"⚠️ {caption}: غير مرفقة")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ===== النموذج =====
with st.form("student_form"):
    
    st.subheader("📝 البيانات الشخصية")
    
    full_name = st.text_input(
        "الاسم الثلاثي *",
        placeholder="مثال: أحمد محمد علي"
    )
    
    mother_full = st.text_input(
        "اسم الأم وكنيتها *",
        placeholder="مثال: فاطمة الحلبي"
    )
    
    col1, col2 = st.columns(2)
    with col1:
        current_year = st.selectbox(
            "السنة الحالية للطالب *",
            ["السنة الأولى", "السنة الثانية", "السنة الثالثة",
             "السنة الرابعة", "السنة الخامسة", "السنة السادسة"]
        )
    with col2:
        university_id = st.text_input(
            "الرقم الجامعي *",
            placeholder="مثال: 2024001"
        )
    
    col1, col2 = st.columns(2)
    with col1:
        phone = st.text_input(
            "رقم الهاتف *",
            placeholder="مثال: 0912345678"
        )
    with col2:
        national_id = st.text_input(
            "الرقم الوطني *",
            placeholder="مثال: 01234567890"
        )
    
    col1, col2 = st.columns(2)
    with col1:
        cert_number = st.text_input(
            "رقم اكتتاب الشهادة الثانوية *",
            placeholder="مثال: 123456"
        )
    with col2:
        cert_year = st.text_input(
            "دورة الشهادة الثانوية *",
            placeholder="مثال: 2023"
        )
    
    st.markdown("---")
    st.subheader("🖼️ الصور المطلوبة")
    
    st.markdown("**🪪 الهوية (وجهان)**")
    col1, col2 = st.columns(2)
    with col1:
        id_front = st.file_uploader("الوجه الأمامي", type=["jpg","jpeg","png"], key="id_front")
    with col2:
        id_back = st.file_uploader("الوجه الخلفي", type=["jpg","jpeg","png"], key="id_back")
    
    st.markdown("**🎓 البطاقة الجامعية (وجهان)**")
    col1, col2 = st.columns(2)
    with col1:
        card_front = st.file_uploader("الوجه الأمامي", type=["jpg","jpeg","png"], key="card_front")
    with col2:
        card_back = st.file_uploader("الوجه الخلفي", type=["jpg","jpeg","png"], key="card_back")
    
    st.markdown("**📜 الشهادة الثانوية (صورة واحدة)**")
    cert_image = st.file_uploader("الشهادة الثانوية", type=["jpg","jpeg","png"], key="cert")
    
    st.markdown("---")
    submitted = st.form_submit_button("📄 توليد ملف Word", use_container_width=True)

# ===== عند الإرسال =====
if submitted:
    errors = []
    if not full_name: errors.append("الاسم الثلاثي")
    if not mother_full: errors.append("اسم الأم وكنيتها")
    if not university_id: errors.append("الرقم الجامعي")
    if not phone: errors.append("رقم الهاتف")
    if not national_id: errors.append("الرقم الوطني")
    if not cert_number: errors.append("رقم اكتتاب الشهادة")
    if not cert_year: errors.append("دورة الشهادة الثانوية")
    
    if errors:
        st.error(f"❌ الرجاء تعبئة الحقول التالية: {', '.join(errors)}")
    else:
        with st.spinner("⏳ جاري إنشاء الملف..."):
            
            images = {
                "id_front": save_uploaded_image(id_front, "id_front.jpg"),
                "id_back": save_uploaded_image(id_back, "id_back.jpg"),
                "card_front": save_uploaded_image(card_front, "card_front.jpg"),
                "card_back": save_uploaded_image(card_back, "card_back.jpg"),
                "cert": save_uploaded_image(cert_image, "cert.jpg"),
            }
            
            doc = Document()
            
            style = doc.styles['Normal']
            style.font.name = 'Arial'
            style.font.size = Pt(12)
            
            # العنوان
            h = doc.add_heading('استمارة بيانات طالب جامعي', level=0)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            h.runs[0].font.color.rgb = RGBColor(0x1e, 0x3a, 0x8a)
            
            # التاريخ
            date_p = doc.add_paragraph(f"تاريخ التحرير: {datetime.now().strftime('%Y/%m/%d')}")
            date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_paragraph()
            
            # جدول البيانات (8 صفوف)
            table = doc.add_table(rows=8, cols=2)
            table.style = 'Light Grid Accent 1'
            
            rows_data = [
                ("الاسم الثلاثي", full_name),
                ("اسم الأم وكنيتها", mother_full),
                ("السنة الحالية للطالب", current_year),
                ("الرقم الجامعي", university_id),
                ("رقم الهاتف", phone),
                ("الرقم الوطني", national_id),
                ("رقم اكتتاب الشهادة الثانوية", cert_number),
                ("دورة الشهادة الثانوية", cert_year),
            ]
            
            for i, (label, value) in enumerate(rows_data):
                table.cell(i, 0).text = label
                table.cell(i, 1).text = value
                for cell in table.rows[i].cells:
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    for para in cell.paragraphs:
                        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            # صفحة الصور
            doc.add_page_break()
            doc.add_heading('📎 المستندات المرفقة', level=1)
            
            doc.add_heading('🪪 الهوية', level=2)
            add_image_to_doc(doc, images["id_front"], "الوجه الأمامي")
            add_image_to_doc(doc, images["id_back"], "الوجه الخلفي")
            
            doc.add_heading('🎓 البطاقة الجامعية', level=2)
            add_image_to_doc(doc, images["card_front"], "الوجه الأمامي")
            add_image_to_doc(doc, images["card_back"], "الوجه الخلفي")
            
            doc.add_heading('📜 الشهادة الثانوية', level=2)
            add_image_to_doc(doc, images["cert"], "الشهادة الثانوية")
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("✅ تم إنشاء الملف بنجاح!")
            
            safe_name = full_name.replace(" ", "_")
            file_name = f"{safe_name}_{university_id}_{datetime.now().strftime('%Y%m%d')}.docx"
            
            st.download_button(
                label="⬇️ تحميل ملف Word",
                data=buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

# ═══════════════════════════════════════════════
# ⭐ عبارة "إعداد الطالب: توفيق قدوري" - في التطبيق فقط
# ═══════════════════════════════════════════════
st.markdown(
    '<div class="footer-credit">✍️ إعداد الطالب: توفيق قدوري</div>',
    unsafe_allow_html=True
)
