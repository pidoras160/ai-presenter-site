import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import io
import random

# ==========================================
# 🔐 ПАРОЛЬ (Можешь менять)
# ==========================================
ACCESS_CODE = "67676766" 

def check_password():
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    
    if not st.session_state.authenticated:
        st.markdown("""
        <style>
        .login-container {
            background: linear-gradient(135deg, #0f0c29 0%, #302b63 50%, #24243e 100%);
            padding: 50px; border-radius: 20px; box-shadow: 0 20px 50px rgba(0,0,0,0.8);
            text-align: center; max-width: 450px; margin: 80px auto;
            border: 1px solid rgba(255,255,255,0.1); backdrop-filter: blur(10px);
        }
        .login-title { color: #fff; font-size: 2.5rem; font-weight: 800; letter-spacing: -1px; margin-bottom: 10px;}
        .login-subtitle { color: #a0aec0; font-size: 1rem; margin-bottom: 30px; }
        input[type="password"] {
            width: 100%; padding: 15px; border-radius: 12px; border: none;
            background: rgba(255,255,255,0.1); color: white; font-size: 1.2rem;
            text-align: center; letter-spacing: 8px; transition: all 0.3s;
        }
        button[kind="primary"] {
            background: linear-gradient(90deg, #FF6B35 0%, #F7971E 100%) !important;
            color: white !important; border: none !important; padding: 15px 30px !important;
            border-radius: 12px !important; font-weight: bold !important; cursor: pointer !important;
            transition: transform 0.2s, box-shadow 0.2s !important; margin-top: 20px; width: 100%;
            box-shadow: 0 4px 15px rgba(255, 107, 53, 0.4) !important;
        }
        </style>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="login-container">
            <h1 class="login-title">🧠 NEURAL DECK LOCAL</h1>
            <p class="login-subtitle">Offline Mode • No API Keys Needed</p>
        </div>
        """, unsafe_allow_html=True)
        
        code_input = st.text_input("", type="password", placeholder="••••••••", label_visibility="collapsed")
        
        if st.button("ENTER SYSTEM", use_container_width=True):
            if code_input == ACCESS_CODE:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Wrong Password.")
        return False
    return True

# ==========================================
# 📚 БАЗА ЗНАНИЙ (Вместо внешнего ИИ)
# ==========================================
class LocalBrain:
    def __init__(self):
        self.topics_db = {
            "наполеон": [
                {"type": "title", "head": "Наполеон Бонапарт", "content": {"text": "НАПОЛЕОН БОНАПАРТ", "sub": "Император Франции и покоритель Европы"}},
                {"type": "intro", "head": "Ранние годы", "content": {"body": "Наполеон родился в 1769 году на острове Корсика всего через год после того, как остров был продан Французскому королевству Генуэзской республикой. Его семья принадлежала к низшему дворянству. Благодаря стипендии он поступил в военную школу Бриенн, а затем в Парижскую военную академию. Уже в 24 года он стал генералом благодаря своим талантам в артиллерии и стратегии."}},
                {"type": "stats", "head": "Военные достижения", "content": {"stat_1_val": "~60", "stat_1_label": "Выигранных крупных сражений", "stat_2_val": "1804", "stat_2_label": "Год коронации императором", "stat_3_val": "1M+", "stat_3_label": "Жертв наполеоновских войн"}},
                {"type": "comparison", "head": "Сильные и слабые стороны", "content": {"left_title": "Гениальность", "right_title": "Ошибки", "left_items": ["Быстрое принятие решений", "Реформа законодательства (Кодекс)", "Лояльность солдат"], "right_items": ["Чрезмерная амбициозность", "Провал русской кампании", "Отказ от дипломатии"]}},
                {"type": "quote", "head": "Философия власти", "content": {"text": "«История — это ложь, в которой согласились все», — говорил он. Но его Кодекс гражданских прав до сих пор является основой права многих стран мира.", "author": "Наполеон Бонапарт"}},
                {"type": "conclusion", "head": "Наследие", "content": {"items": ["Изменил политическую карту Европы навсегда", "Распространил идеи революции за пределы Франции", "Показал силу личной воли и интеллекта"]} },
                {"type": "bullets", "head": "Интересные факты", "content": {"items": ["Боялся кошек (по легенде)", "Любил простую еду: курицу с картофелем", "Его рост был средним для того времени (169 см)"]} }
            ],
            "космос": [
                {"type": "title", "head": "Тайны Вселенной", "content": {"text": "ТАЙНЫ КОСМОСА", "sub": "Путешествие за горизонт событий"}},
                {"type": "intro", "head": "Что такое Вселенная?", "content": {"body": "Вселенная — это бесконечное пространство, включающее все планеты, звезды, галактики и даже время. Возраст Вселенной оценивается примерно в 13.8 миллиардов лет. Мы знаем лишь малую часть того, что там находится: темная материя и темная энергия составляют более 95% всей массы-энергии вселенной."}},
                {"type": "stats", "head": "Масштабы непостижимы", "content": {"stat_1_val": "2 трлн", "stat_1_label": "Галактик во Вселенной", "stat_2_val": "100 млрд", "stat_2_label": "Звезд только в Млечном Пути", "stat_3_val": "93 млрд", "stat_3_label": "Световых лет диаметр наблюдаемой Вселенной"}},
                {"type": "timeline", "head": "Эволюция звезд", "content": {"phase_1": "Рождение в туманностях", "phase_2": "Главная последовательность (как Солнце)", "phase_3": "Взрыв сверхновой или белый карлик"}},
                {"type": "quote", "head": "Взгляд в будущее", "content": {"text": "«Где-то что-то невероятное ждет, чтобы быть открытым».", "author": "Карл Саган"}},
                {"type": "conclusion", "head": "Перспективы человечества", "content": {"items": ["Колонизация Марса к 2050 году", "Поиск внеземной жизни", "Межзвездные перелеты на двигателях нового типа"]} }
            ],
            "default": [
                {"type": "title", "head": "Новая Тема", "content": {"text": "АНАЛИЗ ТЕМЫ", "sub": "Структурированный подход к изучению вопроса"}},
                {"type": "intro", "head": "Введение", "content": {"body": "Данная презентация посвящена глубокому анализу выбранной темы. Мы рассмотрим исторические предпосылки, текущее состояние дел и перспективы развития. Цель работы — выявить ключевые факторы влияния и предложить практические рекомендации."}},
                {"type": "stats", "head": "Ключевые показатели", "content": {"stat_1_val": "85%", "stat_1_label": "Успешность внедрения", "stat_2_val": "+12%", "stat_2_label": "Рост эффективности", "stat_3_val": "5 лет", "stat_3_label": "Срок окупаемости"}},
                {"type": "comparison", "head": "SWOT-Анализ", "content": {"left_title": "Сильные стороны", "right_title": "Угрозы", "left_items": ["Инновационный потенциал", "Квалифицированные кадры", "Устойчивый бренд"], "right_items": ["Конкуренция рынка", "Регуляторные риски", "Изменчивость спроса"]}},
                {"type": "quote", "head": "Экспертное мнение", "content": {"text": "«Будущее принадлежит тем, кто верит в красоту своей мечты и готов работать ради нее».", "author": "Элеонора Рузвельт"}},
                {"type": "conclusion", "head": "Заключение", "content": {"items": ["Тема требует дальнейшего изучения", "Необходимо привлечение инвестиций", "Разработка стратегического плана действий"]} }
            ]
        }

    def get_content(self, topic):
        """Выбирает лучшую тему из базы или использует шаблон."""
        t_lower = topic.lower()
        if "наполеон" in t_lower or "bonaparte" in t_lower:
            return self.topics_db["наполеон"]
        elif "космос" in t_lower or "space" in t_lower or "вселенная" in t_lower:
            return self.topics_db["космос"]
        else:
            # Универсальный шаблон для любой другой темы
            slides = self.topics_db["default"].copy()
            # Меняем заголовки на название пользователя
            slides[0]['content']['text'] = topic.upper()
            slides[0]['head'] = topic
            return slides

# ==========================================
# 🎨 ДИЗАЙН-СИСТЕМА (РАНДОМНАЯ)
# ==========================================
class DesignTheme:
    def __init__(self):
        self.palettes = [
            {"name": "Cyberpunk Neon", "bg": RGBColor(0x1A, 0x1A, 0x2E), "text": RGBColor(0xFF, 0xFF, 0xFF), "accent": RGBColor(0x00, 0xF0, 0xFF), "secondary": RGBColor(0xFF, 0x00, 0x7F)},
            {"name": "Corporate Clean", "bg": RGBColor(0xF0, 0xF4, 0xF8), "text": RGBColor(0x1A, 0x23, 0x7E), "accent": RGBColor(0x00, 0x7B, 0xFF), "secondary": RGBColor(0x6C, 0x75, 0x8A)},
            {"name": "Nature Organic", "bg": RGBColor(0xE8, 0xF5, 0xE9), "text": RGBColor(0x1B, 0x5E, 0x20), "accent": RGBColor(0x4C, 0xAF, 0x50), "secondary": RGBColor(0x81, 0xC7, 0x84)},
            {"name": "Sunset Warmth", "bg": RGBColor(0xFFF3E0), "text": RGBColor(0xBF, 0x36, 0x0C), "accent": RGBColor(0xFF, 0x6D, 0x00), "secondary": RGBColor(0xFF, 0xAB, 0x40)},
            {"name": "Dark Luxury", "bg": RGBColor(0x12, 0x12, 0x12), "text": RGBColor(0xE0, 0xE0, 0xE0), "accent": RGBColor(0xD4, 0xAF, 0x37), "secondary": RGBColor(0x8B, 0x00, 0x00)}
        ]
        self.fonts = ["Arial Black", "Impact", "Segoe UI Bold", "Verdana", "Georgia"]
        
        palette = random.choice(self.palettes)
        self.bg_color = palette["bg"]
        self.text_color = palette["text"]
        self.accent_color = palette["accent"]
        self.secondary_color = palette["secondary"]
        self.main_font = random.choice(self.fonts)
        self.style_name = palette["name"]

# ==========================================
# ⚙️ ДВИЖОК PPTX (STABLE VERSION)
# ==========================================
class AdvancedPPTXBuilder:
    def __init__(self):
        self.theme = DesignTheme()
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)

    def add_decorative_elements(self, slide):
        shape_type = random.choice([MSO_SHAPE.OVAL, MSO_SHAPE.RECTANGLE, MSO_SHAPE.TRIANGLE])
        left = random.randint(0, 10000000)
        top = random.randint(0, 5000000)
        width = random.randint(2000000, 5000000)
        height = random.randint(2000000, 5000000)
        
        shape = slide.shapes.add_shape(shape_type, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.theme.secondary_color
        shape.line.fill.background()
        try:
            from lxml import etree
            alpha_mod = etree.SubElement(shape._element.spPr.solidFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}alpha')
            alpha_mod.set('val', str(random.randint(10000, 30000)))
        except: pass

    def build_slide(self, data):
        layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(layout)
        
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = self.theme.bg_color
        
        if data.get('type') not in ['title', 'quote']:
            self.add_decorative_elements(slide)

        content = data.get('content', {})
        head_text = data.get('head', '')
        
        # --- РЕНДЕРИНГ ПО ТИПУ СЛАЙДА ---
        
        if data['type'] == 'title':
            tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11.3), Inches(3))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = content.get('text', head_text)
            r.font.size = Pt(80)
            r.font.bold = True
            r.font.name = self.theme.main_font
            r.font.color.rgb = self.theme.accent_color
            
            sb = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(11.3), Inches(1))
            sp = sb.text_frame.paragraphs[0]
            sr = sp.add_run()
            sr.text = content.get('sub', '')
            sr.font.size = Pt(24)
            sr.font.color.rgb = self.theme.text_color
            sr.font.italic = True

        elif data['type'] == 'intro':
            hb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
            hp = hb.text_frame.paragraphs[0]
            hr = hp.add_run()
            hr.text = head_text or "Introduction"
            hr.font.size = Pt(40)
            hr.font.bold = True
            hr.font.color.rgb = self.theme.accent_color
            
            bb = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12), Inches(5))
            bp = bb.text_frame.paragraphs[0]
            br = bp.add_run()
            br.text = content.get('body', '')
            br.font.size = Pt(28)
            br.font.color.rgb = self.theme.text_color
            br.line_spacing = 1.5

        elif data['type'] == 'stats':
            hb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
            hp = hb.text_frame.paragraphs[0]
            hr = hp.add_run()
            hr.text = head_text or "Key Metrics"
            hr.font.size = Pt(40)
            hr.font.bold = True
            hr.font.color.rgb = self.theme.accent_color
            
            positions = [(0.5, 2.0), (4.8, 2.0), (9.1, 2.0)]
            keys = [('stat_1_val', 'stat_1_label'), ('stat_2_val', 'stat_2_label'), ('stat_3_val', 'stat_3_label')]
            
            for pos, key_pair in zip(positions, keys):
                val_key, lbl_key = key_pair
                
                vb = slide.shapes.add_textbox(Inches(pos[0]), Inches(pos[1]), Inches(4), Inches(1.5))
                vp = vb.text_frame.paragraphs[0]
                vp.alignment = PP_ALIGN.CENTER
                vr = vp.add_run()
                vr.text = str(content.get(val_key, 'N/A'))
                vr.font.size = Pt(60)
                vr.font.bold = True
                vr.font.color.rgb = self.theme.secondary_color
                
                lb = slide.shapes.add_textbox(Inches(pos[0]), Inches(pos[1]+1.5), Inches(4), Inches(1))
                lp = lb.text_frame.paragraphs[0]
                lp.alignment = PP_ALIGN.CENTER
                lr = lp.add_run()
                lr.text = str(content.get(lbl_key, ''))
                lr.font.size = Pt(20)
                lr.font.color.rgb = self.theme.text_color

        elif data['type'] == 'comparison':
            hb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
            hp = hb.text_frame.paragraphs[0]
            hr = hp.add_run()
            hr.text = head_text or "Analysis"
            hr.font.size = Pt(40)
            hr.font.bold = True
            hr.font.color.rgb = self.theme.accent_color
            
            l_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.5), Inches(6), Inches(5.5))
            l_box.fill.solid(); l_box.fill.fore_color.rgb = RGBColor(0xFFFFFF); l_box.shadow.inherit = False
            lt_head = l_box.text_frame.paragraphs[0]
            lt_head.text = content.get('left_title', 'Pros')
            lt_head.font.size = Pt(24); lt_head.font.bold = True; lt_head.font.color.rgb = RGBColor(0x2E, 0x7D, 0x32)
            for item in content.get('left_items', []):
                para = l_box.text_frame.add_paragraph()
                para.text = f"✓ {item}"
                para.font.size = Pt(18); para.space_after = Pt(10); para.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

            r_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(6), Inches(5.5))
            r_box.fill.solid(); r_box.fill.fore_color.rgb = RGBColor(0xFFFFFF); r_box.shadow.inherit = False
            rt_head = r_box.text_frame.paragraphs[0]
            rt_head.text = content.get('right_title', 'Cons')
            rt_head.font.size = Pt(24); rt_head.font.bold = True; rt_head.font.color.rgb = RGBColor(0xC6, 0x28, 0x28)
            for item in content.get('right_items', []):
                para = r_box.text_frame.add_paragraph()
                para.text = f"! {item}"
                para.font.size = Pt(18); para.space_after = Pt(10); para.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

        elif data['type'] == 'quote':
            qb = slide.shapes.add_textbox(Inches(2), Inches(2.5), Inches(9.3), Inches(3))
            qtf = qb.text_frame
            qp = qtf.paragraphs[0]
            qp.alignment = PP_ALIGN.CENTER
            qr = qp.add_run()
            qr.text = f"“{content.get('text', '')}”"
            qr.font.size = Pt(32); qr.font.italic = True; qr.font.name = "Georgia"; qr.font.color.rgb = self.theme.text_color
            
            ab = slide.shapes.add_textbox(Inches(2), Inches(5.5), Inches(9.3), Inches(1))
            ap = ab.text_frame.paragraphs[0]
            ap.alignment = PP_ALIGN.RIGHT
            ar = ap.add_run()
            ar.text = f"— {content.get('author', 'Unknown')}"
            ar.font.size = Pt(20); ar.font.bold = True; ar.font.color.rgb = self.theme.accent_color

        elif data['type'] == 'timeline':
            hb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
            hp = hb.text_frame.paragraphs[0]
            hr = hp.add_run()
            hr.text = head_text or "Roadmap"
            hr.font.size = Pt(40); hr.font.bold = True; hr.font.color.rgb = self.theme.accent_color
            
            line_y = Inches(4)
            ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), line_y, Inches(11.3), Pt(5))
            ln.fill.solid(); ln.fill.fore_color.rgb = self.theme.secondary_color; ln.line.fill.background()
            
            phases = [content.get('phase_1', ''), content.get('phase_2', ''), content.get('phase_3', '')]
            x_positions = [Inches(2), Inches(6), Inches(10)]
            
            for phase, x_pos in zip(phases, x_positions):
                dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, x_pos - Inches(0.2), line_y - Inches(0.2), Inches(0.4), Inches(0.4))
                dot.fill.solid(); dot.fill.fore_color.rgb = self.theme.accent_color; dot.line.color.rgb = self.theme.bg_color; dot.line.width = Pt(3)
                
                tb_phase = slide.shapes.add_textbox(x_pos - Inches(1.5), line_y + Inches(0.5), Inches(3), Inches(1))
                tp = tb_phase.text_frame.paragraphs[0]
                tp.alignment = PP_ALIGN.CENTER
                tr = tp.add_run()
                tr.text = phase
                tr.font.size = Pt(16); tr.font.bold = True; tr.font.color.rgb = self.theme.text_color

        else: # Default bullets
            hb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
            hp = hb.text_frame.paragraphs[0]
            hr = hp.add_run()
            hr.text = head_text or "Details"
            hr.font.size = Pt(40); hr.font.bold = True; hr.font.color.rgb = self.theme.accent_color
            
            lb = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(5))
            ltf = lb.text_frame
            ltf.word_wrap = True
            
            items = content.get('items', [])
            for j, item in enumerate(items):
                para = ltf.paragraphs[j] if j < len(ltf.paragraphs) else ltf.add_paragraph()
                para.space_after = Pt(20)
                r = para.add_run()
                r.text = f"● {item}"
                r.font.size = Pt(26)
                r.font.color.rgb = self.theme.text_color

    def save_to_bytes(self):
        bytes_io = io.BytesIO()
        self.prs.save(bytes_io)
        bytes_io.seek(0)
        return bytes_io.getvalue()

# ==========================================
# 🎨 ИНТЕРФЕЙС STREAMLIT
# ==========================================
st.set_page_config(page_title="Neural Deck Local", page_icon="🧠", layout="wide")

st.markdown("""
<style>
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #0f172a; }
    div[data-testid="stChatMessage"] { background-color: transparent; padding: 1rem 0; }
    .user-msg-container {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white; padding: 15px 25px; border-radius: 20px 20px 0 20px;
        max-width: 85%; margin-left: auto; box-shadow: 0 4px 15px rgba(37, 99, 235, 0.3);
        animation: fadeInUp 0.3s ease-out;
    }
    .bot-msg-container {
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.1);
        color: #e2e8f0; padding: 20px; border-radius: 20px 20px 20px 0;
        max-width: 85%; backdrop-filter: blur(10px);
        box-shadow: 0 4px 20px rgba(0,0,0,0.2); animation: fadeInLeft 0.3s ease-out;
    }
    div[data-testid="stBottomBlockContainer"] {
        background-color: #1e293b; border-top: 1px solid rgba(255,255,255,0.1); padding: 1rem;
    }
    textarea[data-testid="stTextInput"] {
        background-color: #0f172a !important; color: white !important;
        border: 1px solid #334155 !important; border-radius: 12px !important;
    }
    .stDownloadButton {
        background: linear-gradient(90deg, #10b981 0%, #059669 100%) !important;
        color: white !important; border: none !important; padding: 12px 24px !important;
        border-radius: 8px !important; font-weight: bold !important;
        box-shadow: 0 4px 15px rgba(16, 185, 129, 0.4) !important; transition: all 0.2s !important;
    }
    .stDownloadButton:hover { transform: scale(1.05); box-shadow: 0 6px 20px rgba(16, 185, 129, 0.6) !important; }
    @keyframes fadeInUp { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
    @keyframes fadeInLeft { from { opacity: 0; transform: translateX(-10px); } to { opacity: 1; transform: translateX(0); } }
</style>
""", unsafe_allow_html=True)

if check_password():
    with st.sidebar:
        st.header("⚙️ Control Panel")
        st.write("**Mode:** Offline Database")
        st.success("✅ Ready")
        st.divider()
        st.caption("No internet required for generation")
        if st.button("Logout"):
            st.session_state.authenticated = False
            st.rerun()

    if 'messages' not in st.session_state:
        st.session_state.messages = []
        init_msg = {
            "role": "assistant", 
            "content": "Hello! I am your Neural Deck Architect.\n\nI have built-in knowledge about:\n1. **Napoleon Bonaparte**\n2. **Space & Universe**\n3. **General Business Topics**\n\nJust type a topic (e.g., 'Napoleon') and number of slides!"
        }
        st.session_state.messages.append(init_msg)

    for message in st.session_state.messages:
        role_class = "user-msg-container" if message["role"] == "user" else "bot-msg-container"
        avatar = "👤" if message["role"] == "user" else "🧠"
        
        with st.chat_message(message["role"], avatar=avatar):
            st.markdown(f'<div class="{role_class}">{message["content"].replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)
            
            if "file_name" in message:
                st.download_button(
                    label=f"📥 Download '{message['file_name']}'",
                    data=message["file_bytes"],
                    file_name=message["file_name"],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

    prompt = st.chat_input("Enter Topic | Slides Count...")

    if prompt:
        parts = prompt.split('|')
        topic = parts[0].strip() if len(parts) > 0 else "General Overview"
        slides_count_str = parts[1].strip() if len(parts) > 1 else "6"
        
        try:
            slides_count = int(slides_count_str)
        except:
            slides_count = 6
            
        if slides_count < 3 or slides_count > 10:
            slides_count = 6

        user_display_text = f"**Topic:** {topic}\n**Slides:** {slides_count}"
        
        user_msg = {"role": "user", "content": user_display_text}
        st.session_state.messages.append(user_msg)
        with st.chat_message("user", avatar="👤"):
            st.markdown(f'<div class="user-msg-container">{user_display_text.replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)

        with st.chat_message("assistant", avatar="🧠"):
            status = st.empty()
            status.spinner("🧠 Searching local database...")
            
            brain = LocalBrain()
            structure_data = brain.get_content(topic)
            
            # Обрезаем или дополняем до нужного количества
            if len(structure_data) > slides_count:
                structure_data = structure_data[:slides_count]
            elif len(structure_data) < slides_count:
                # Добавляем пустые слайды если нужно больше
                extra_needed = slides_count - len(structure_data)
                for _ in range(extra_needed):
                    structure_data.append({"type": "bullets", "head": "Additional Info", "content": {"items": ["More details coming soon..."]}})
            
            status.spinner("🎨 Rendering beautiful slides...")
            
            builder = AdvancedPPTXBuilder()
            for slide_def in structure_data:
                builder.build_slide(slide_def)
                
            file_bytes = builder.save_to_bytes()
            
            filename = f"{topic.replace(' ', '_')}_{builder.theme.style_name}.pptx"
            
            result_text = f"✅ **Generation Complete!**\n\nCreated presentation for **'{topic}'** using internal logic.\nVisual Theme: **{builder.theme.style_name}**.\nTotal Slides: **{len(structure_data)}**.\n\nClick below to download:"
            
            status.markdown(f'<div class="bot-msg-container">{result_text.replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)
            
            btn = st.download_button(
                label=f"📥 Get Your Smart Deck ({filename})",
                data=file_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            
            bot_msg = {
                "role": "assistant",
                "content": result_text,
                "file_bytes": file_bytes,
                "file_name": filename
            }
            st.session_state.messages.append(bot_msg)
