import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import io

# ==========================================
# 🔐 НАСТРОЙКИ ДОСТУПА
# ==========================================
ACCESS_CODE = "67676767" 

def check_password():
    """Экран входа с красивым дизайном."""
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    
    if not st.session_state.authenticated:
        # CSS для экрана логина
        st.markdown("""
        <style>
        .login-container {
            background: linear-gradient(135deg, #1a1a1a 0%, #2d2d2d 100%);
            padding: 50px;
            border-radius: 20px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.5);
            text-align: center;
            max-width: 400px;
            margin: 100px auto;
            border: 1px solid #333;
        }
        .login-title {
            color: #FF6B35;
            font-size: 2rem;
            font-weight: bold;
            margin-bottom: 10px;
        }
        .login-subtitle {
            color: #aaa;
            font-size: 1rem;
            margin-bottom: 30px;
        }
        input[type="password"] {
            width: 100%;
            padding: 15px;
            margin-top: 10px;
            border-radius: 10px;
            border: none;
            background: #333;
            color: white;
            font-size: 1.2rem;
            text-align: center;
            letter-spacing: 5px;
        }
        button[kind="primary"] {
            background-color: #FF6B35 !important;
            color: white !important;
            border: none !important;
            padding: 15px 30px !important;
            border-radius: 10px !important;
            font-weight: bold !important;
            cursor: pointer !important;
            transition: transform 0.2s ease !important;
            margin-top: 20px;
        }
        button[kind="primary"]:hover {
            transform: scale(1.05) !important;
            background-color: #e65c29 !important;
        }
        </style>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="login-container">
            <h1 class="login-title">🔒 AI Presenter Pro</h1>
            <p class="login-subtitle">Введите секретный код доступа</p>
        </div>
        """, unsafe_allow_html=True)
        
        code_input = st.text_input("", type="password", placeholder="••••••••", label_visibility="collapsed")
        
        if st.button("ВОЙТИ В СИСТЕМУ", use_container_width=True):
            if code_input == ACCESS_CODE:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Неверный код! Попробуйте снова.")
        return False
    return True

# ==========================================
# ⚙️ ЛОГИКА ГЕНЕРАЦИИ PPTX (PRO VERSION)
# ==========================================
class PPTXGenerator:
    def __init__(self):
        # Цветовая палитра "Modern Tech"
        self.colors = {
            "primary": RGBColor(0x1A, 0x23, 0x7E), # Deep Blue
            "accent": RGBColor(0xFF, 0x6B, 0x35),  # Vibrant Orange
            "bg_light": RGBColor(0xF8, 0xF9, 0xFA), # Off White
            "bg_dark": RGBColor(0x1E, 0x1E, 0x2E),   # Dark Gray for contrast
            "white": RGBColor(0xFF, 0xFF, 0xFF),
            "text_main": RGBColor(0x2D, 0x37, 0x48),
            "text_muted": RGBColor(0x6C, 0x75, 0x8A)
        }

    def generate_content(self, topic):
        """Генерирует структуру слайдов."""
        return [
            {"type": "title", "text": f"{topic.upper()}", "sub": "Интеллектуальная презентация • 2026"},
            {"type": "bullets", "head": "Введение", "items": ["Актуальность темы", "Цели исследования", "Методология"]},
            {"type": "two_col", "head": "Анализ ситуации", "left_title": "Сильные стороны", "right_title": "Зоны роста", "left_items": ["Технологичность", "Быстрая адаптация"], "right_items": ["Высокая стоимость", "Нехватка кадров"]},
            {"type": "quote", "head": "Мнение эксперта", "text": f"«{topic} — это не просто тренд, а фундамент будущего индустрии.»", "author": "Dr. Alex Smith, CEO FutureTech"},
            {"type": "conclusion", "head": "Заключение", "items": ["Резюме ключевых идей", "Практические рекомендации", "Перспективы развития"]}
        ]

    def create_pptx_bytes(self, slides_data):
        prs = Presentation()
        prs.slide_width = Inches(13.333) 
        prs.slide_height = Inches(7.5)

        for i, slide_def in enumerate(slides_data):
            layout = prs.slide_layouts[6] # Blank
            slide = prs.slides.add_slide(layout)
            
            bg = slide.background.fill
            bg.solid()
            
            is_title = (slide_def['type'] == 'title')
            is_quote = (slide_def['type'] == 'quote')
            is_two_col = (slide_def['type'] == 'two_col')
            
            # Выбор фона и цветов текста
            if is_title or is_quote:
                bg.fore_color.rgb = self.colors["primary"]
                txt_col = self.colors["white"]
                acc_col = self.colors["accent"]
            elif is_two_col:
                bg.fore_color.rgb = self.colors["bg_light"]
                txt_col = self.colors["text_main"]
                acc_col = self.colors["primary"]
            else:
                bg.fore_color.rgb = self.colors["bg_light"]
                txt_col = self.colors["text_main"]
                acc_col = self.colors["primary"]

            # --- РЕНДЕРИНГ СЛАЙДА ---
            
            if is_title:
                # Декоративная линия сверху
                line_top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.2))
                line_top.fill.solid()
                line_top.fill.fore_color.rgb = self.colors["accent"]
                line_top.line.fill.background()

                tb = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.3), Inches(2))
                tf = tb.text_frame
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                r = p.add_run()
                r.text = slide_def['text']
                r.font.size = Pt(60)
                r.font.bold = True
                r.font.name = "Arial Black"
                r.font.color.rgb = txt_col
                
                sb = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(11.3), Inches(1))
                sp = sb.text_frame.paragraphs[0]
                sp.alignment = PP_ALIGN.CENTER
                sr = sp.add_run()
                sr.text = slide_def.get('sub', '')
                sr.font.size = Pt(24)
                sr.font.color.rgb = acc_col

            elif is_quote:
                qb = slide.shapes.add_textbox(Inches(2), Inches(2.5), Inches(9.3), Inches(3))
                qtf = qb.text_frame
                qp = qtf.paragraphs[0]
                qp.alignment = PP_ALIGN.CENTER
                qr = qp.add_run()
                qr.text = slide_def['text']
                qr.font.size = Pt(36)
                qr.font.italic = True
                qr.font.name = "Georgia"
                qr.font.color.rgb = txt_col
                
                ab = slide.shapes.add_textbox(Inches(2), Inches(5.5), Inches(9.3), Inches(1))
                ap = ab.text_frame.paragraphs[0]
                ap.alignment = PP_ALIGN.RIGHT
                ar = ap.add_run()
                ar.text = f"— {slide_def.get('author', '')}"
                ar.font.size = Pt(20)
                ar.font.bold = True
                ar.font.color.rgb = acc_col

            elif is_two_col:
                # Заголовок
                hb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.3), Inches(1))
                htf = hb.text_frame
                hp = htf.paragraphs[0]
                hr = hp.add_run()
                hr.text = slide_def['head']
                hr.font.size = Pt(36)
                hr.font.bold = True
                hr.font.color.rgb = acc_col

                # Левая колонка
                lb_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(6), Inches(5))
                ltf = lb_box.text_frame
                lp_head = ltf.paragraphs[0]
                lr_head = lp_head.add_run()
                lr_head.text = slide_def['left_title']
                lr_head.font.size = Pt(24)
                lr_head.font.bold = True
                lr_head.font.color.rgb = self.colors["primary"]
                
                for item in slide_def['left_items']:
                    para = ltf.add_paragraph()
                    para.space_after = Pt(10)
                    r = para.add_run()
                    r.text = f"✓ {item}"
                    r.font.size = Pt(20)
                    r.font.color.rgb = txt_col

                # Правая колонка
                rb_box = slide.shapes.add_textbox(Inches(6.8), Inches(1.8), Inches(6), Inches(5))
                rtf = rb_box.text_frame
                rp_head = rtf.paragraphs[0]
                rr_head = rp_head.add_run()
                rr_head.text = slide_def['right_title']
                rr_head.font.size = Pt(24)
                rr_head.font.bold = True
                rr_head.font.color.rgb = self.colors["accent"]
                
                for item in slide_def['right_items']:
                    para = rtf.add_paragraph()
                    para.space_after = Pt(10)
                    r = para.add_run()
                    r.text = f"! {item}"
                    r.font.size = Pt(20)
                    r.font.color.rgb = txt_col

            else:
                # Стандартные списки
                hb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.3), Inches(1))
                htf = hb.text_frame
                hp = htf.paragraphs[0]
                hr = hp.add_run()
                hr.text = slide_def['head']
                hr.font.size = Pt(36)
                hr.font.bold = True
                hr.font.color.rgb = acc_col
                
                line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.2), Inches(12.3), Pt(3))
                line.fill.solid()
                line.fill.fore_color.rgb = self.colors["primary"]
                line.line.fill.background()

                lb = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(5))
                ltf = lb.text_frame
                ltf.word_wrap = True
                
                items = slide_def.get('items', [])
                for j, item in enumerate(items):
                    para = ltf.paragraphs[j] if j < len(ltf.paragraphs) else ltf.add_paragraph()
                    para.space_after = Pt(15)
                    r = para.add_run()
                    r.text = f"• {item}"
                    r.font.size = Pt(24)
                    r.font.color.rgb = txt_col

        bytes_io = io.BytesIO()
        prs.save(bytes_io)
        bytes_io.seek(0)
        return bytes_io.getvalue()

# ==========================================
# 🎨 ИНТЕРФЕЙС STREAMLIT (CHATGPT STYLE PRO)
# ==========================================
st.set_page_config(page_title="AI Presenter Pro", page_icon="✨", layout="wide")

# Глобальные стили
st.markdown("""
<style>
    /* Скрываем меню */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Фон страницы */
    .stApp {
        background-color: #f9fafb;
    }
    
    /* Чат сообщения */
    div[data-testid="stChatMessage"] {
        padding: 10px 0;
    }
    div[data-testid="stChatMessageContentAvatar"] {
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.5rem;
    }
    
    /* Пузыри чата */
    .user-bubble {
        background-color: #ffffff;
        color: #1e1b4b;
        padding: 15px 20px;
        border-radius: 20px 20px 0 20px;
        max-width: 80%;
        margin-left: auto;
        display: inline-block;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        border: 1px solid #e5e7eb;
    }
    .bot-bubble {
        background-color: #ffffff;
        color: #374151;
        padding: 15px 20px;
        border-radius: 20px 20px 20px 0;
        max-width: 80%;
        border: 1px solid #e5e7eb;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
    
    /* Кнопка скачивания */
    .stDownloadButton {
        background-color: #FF6B35 !important;
        color: white !important;
        border: none !important;
        padding: 10px 20px !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        cursor: pointer !important;
        margin-top: 10px;
    }
    .stDownloadButton:hover {
        background-color: #e65c29 !important;
    }
    
    /* Инпут */
    div[data-testid="stTextInput"] input {
        border-radius: 20px;
        border: 1px solid #ddd;
        padding: 10px 15px;
    }
</style>
""", unsafe_allow_html=True)

if check_password():
    # Боковое меню
    with st.sidebar:
        st.header("⚙️ Настройки")
        st.write("Статус системы:")
        st.success("✅ Онлайн")
        st.divider()
        st.caption("Версия 2.0 Pro Edition")
        if st.button("🚪 Выйти"):
            st.session_state.authenticated = False
            st.rerun()

    # Инициализация истории
    if 'messages' not in st.session_state:
        st.session_state.messages = []
        welcome_msg = {
            "role": "assistant", 
            "content": "Привет! Я твой персональный AI-ассистент по созданию презентаций.\n\nНапиши тему (например: **'Маркетинг'** или **'История Древнего Рима'**), и я разработаю структуру и создам готовый файл PowerPoint."
        }
        st.session_state.messages.append(welcome_msg)

    # Отображение истории
    for message in st.session_state.messages:
        role_class = "user-bubble" if message["role"] == "user" else "bot-bubble"
        avatar = "👤" if message["role"] == "user" else "🤖"
        
        with st.chat_message(message["role"], avatar=avatar):
            st.markdown(f'<div class="{role_class}">{message["content"].replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)
            
            if "file_name" in message:
                st.download_button(
                    label=f"📥 Скачать '{message['file_name']}'",
                    data=message["file_bytes"],
                    file_name=message["file_name"],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

    # Поле ввода
    prompt = st.chat_input("О чем сделать презентацию?")

    if prompt:
        user_msg = {"role": "user", "content": prompt}
        st.session_state.messages.append(user_msg)
        with st.chat_message("user", avatar="👤"):
            st.markdown(f'<div class="user-bubble">{prompt}</div>', unsafe_allow_html=True)

        with st.chat_message("assistant", avatar="🤖"):
            status = st.empty()
            status.spinner("🤖 Анализирую запрос...")
            
            gen = PPTXGenerator()
            slides_data = gen.generate_content(prompt)
            
            status.spinner("🎨 Генерирую дизайн слайдов...")
            file_bytes = gen.create_pptx_bytes(slides_data)
            
            filename = f"Presentation_{prompt.replace(' ', '_')[:15]}.pptx"
            
            summary_text = f"✅ Готово! Презентация по теме **«{prompt}»** создана.\n\n**Структура:**\n"
            for idx, s in enumerate(slides_data):
                title = s.get('text') or s.get('head')
                summary_text += f"- Слайд {idx+1}: {title}\n"
            
            status.markdown(f'<div class="bot-bubble">{summary_text.replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)
            
            btn = st.download_button(
                label=f"📥 Скачать '{filename}'",
                data=file_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            
            bot_msg = {
                "role": "assistant",
                "content": summary_text,
                "file_bytes": file_bytes,
                "file_name": filename
            }
            st.session_state.messages.append(bot_msg)
