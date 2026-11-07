import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from PIL import Image
import io
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
</style>
""", unsafe_allow_html=True)

st.title("🎓 نموذج تسجيل طالب جامعي")
st.markdown("---")

# ===== دوال مساعدة =====
def save_uploaded_image(uploaded_file, filename):
    """يحفظ الصورة المرفوعة ويعيد مسارها"""
    if uploaded_file is None:
        return None
    img = Image.open(uploaded_file)
    # ضغط الصورة لتقليل الحجم
    img.thumbnail((1500, 1500))
    path = f"/tmp/{filename}"
    img.save(path, "JPEG", quality=85)
    return path

def add_image_to_doc(doc, img_path, caption, width=Inches(4.5)):
    """يضيف صورة إلى مستند Word"""
    if img_path:
        st_img = doc.add_paragraph()
        st_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st_img.add_run().add_picture(img_path, width=width)
        
        cap = doc.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.runs[0].font.size = Pt(10)
        cap.runs[0].font.italic = True
    else:
        p = doc.add_paragraph(f"⚠️ {caption}: غير مرفقة")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ===== النموذج =====
with st.form("student_form"):
    
    st.subheader("📝 البيانات الشخصية")
    
    col1, col2 = st.columns(2)
    with col1:
        national_id = st.text_input("الرقم الوطني *", placeholder="مثال: 01234567890")
        university_id = st.text_input("الرقم الجامعي *", placeholder="مثال: 2024001")
        phone = st.text_input("رقم هاتف للتواصل *", placeholder="مثال: 0912345678")
    with col2:
        mother_name = st.text_input("كنية الأم *", placeholder="مثال: فاطمة")
        cert_number = st.text_input("رقم اكتتاب الشهادة الثانوية *", placeholder="مثال: 123456")
        year = st.selectbox("سنة الطالب الحالية *", 
                            ["السنة الأولى", "السنة الثانية", "السنة الثالثة",
                             "السنة الرابعة", "السنة الخامسة", "السنة السادسة"])
    
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
    # التحقق
    errors = []
    if not national_id: errors.append("الرقم الوطني")
    if not university_id: errors.append("الرقم الجامعي")
    if not phone: errors.append("رقم الهاتف")
    if not mother_name: errors.append("كنية الأم")
    if not cert_number: errors.append("رقم اكتتاب الشهادة")
    
    if errors:
        st.error(f"❌ الرجاء تعبئة الحقول التالية: {', '.join(errors)}")
    else:
        with st.spinner("⏳ جاري إنشاء الملف..."):
            
            # حفظ الصور مؤقتاً
            images = {
                "id_front": save_uploaded_image(id_front, "id_front.jpg"),
                "id_back": save_uploaded_image(id_back, "id_back.jpg"),
                "card_front": save_uploaded_image(card_front, "card_front.jpg"),
                "card_back": save_uploaded_image(card_back, "card_back.jpg"),
                "cert": save_uploaded_image(cert_image, "cert.jpg"),
            }
            
            # إنشاء المستند
            doc = Document()
            
            # تنسيق عام
            style = doc.styles['Normal']
            style.font.name = 'Arial'
            style.font.size = Pt(12)
            
            # العنوان
            h = doc.add_heading('استمارة بيانات طالب جامعي', level=0)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            h.runs[0].font.color.rgb = RGBColor(0x1e, 0x3a, 0x8a)
            
            date_p = doc.add_paragraph(f"تاريخ التحرير: {datetime.now().strftime('%Y/%m/%d')}")
            date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_paragraph()
            
            # جدول البيانات
            table = doc.add_table(rows=6, cols=2)
            table.style = 'Light Grid Accent 1'
            
            rows_data = [
                ("الرقم الوطني", national_id),
                ("كنية الأم", mother_name),
                ("الرقم الجامعي", university_id),
                ("رقم اكتتاب الشهادة الثانوية", cert_number),
                ("رقم هاتف للتواصل", phone),
                ("سنة الطالب الحالية", year),
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
            
            # الهوية
            doc.add_heading('🪪 الهوية', level=2)
            add_image_to_doc(doc, images["id_front"], "الوجه الأمامي")
            add_image_to_doc(doc, images["id_back"], "الوجه الخلفي")
            
            # البطاقة الجامعية
            doc.add_heading('🎓 البطاقة الجامعية', level=2)
            add_image_to_doc(doc, images["card_front"], "الوجه الأمامي")
            add_image_to_doc(doc, images["card_back"], "الوجه الخلفي")
            
            # الشهادة
            doc.add_heading('📜 الشهادة الثانوية', level=2)
            add_image_to_doc(doc, images["cert"], "الشهادة الثانوية")
            
            # حفظ في الذاكرة
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("✅ تم إنشاء الملف بنجاح!")
            
            # زر التحميل
            st.download_button(
                label="⬇️ تحميل ملف Word",
                data=buffer,
                file_name=f"طالب_{university_id}_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )