import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io

# ==========================================
# 🔐 СЕКРЕТНЫЙ КОД ДОСТУПА (Меняй здесь!)
# ==========================================
ACCESS_CODE = "NAPOLEON_IS_COOL_2026" 

def check_password():
    """Экран входа с паролем."""
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    
    if not st.session_state.authenticated:
        # Красивый экран логина
        st.markdown("""
        <style>
        .login-box {
            border: 2px solid #FF6B35;
            padding: 40px;
            border-radius: 15px;
            background-color: #1a1a1a;
            color: white;
            text-align: center;
            max-width: 400px;
            margin: 50px auto;
            box-shadow: 0 10px 30px rgba(0,0,0,0.5);
        }
        h1 { color: #FF6B35; font-family: sans-serif; }
        input[type="password"] {
            width: 100%;
            padding: 10px;
            margin-top: 10px;
            border-radius: 5px;
            border: none;
            background: #333;
            color: white;
        }
        </style>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="login-box">
            <h1>🔒 AI Presenter</h1>
            <p>Введите секретный код для доступа</p>
        </div>
        """, unsafe_allow_html=True)
        
        code_input = st.text_input("", type="password", placeholder="Код доступа...", label_visibility="collapsed")
        
        if st.button("ВОЙТИ", use_container_width=True):
            if code_input == ACCESS_CODE:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Неверный код!")
        return False
    return True

# ==========================================
# ⚙️ ЛОГИКА ГЕНЕРАЦИИ PPTX
# ==========================================
class PPTXGenerator:
    def __init__(self):
        self.colors = {
            "primary": RGBColor(0x1A, 0x23, 0x7E), # Синий
            "accent": RGBColor(0xFF, 0x6B, 0x35),  # Оранжевый
            "bg": RGBColor(0xF8, 0xF9, 0xFA),      # Светло-серый
            "white": RGBColor(0xFF, 0xFF, 0xFF),
            "dark_text": RGBColor(0x2D, 0x37, 0x48)
        }

    def generate_content(self, topic):
        """
        Здесь можно подключить настоящий ИИ (OpenAI/Gemini), 
        пока что это умная заглушка, которая делает красивую структуру.
        """
        return [
            {"type": "title", "text": f"{topic.upper()}", "sub": "Сгенерировано нейросетью"},
            {"type": "bullets", "head": "Введение", "items": ["Что такое эта тема?", "Почему она важна сейчас?", "Цель нашего исследования"]},
            {"type": "image_placeholder", "head": "Ключевой визуал", "desc": "[Здесь должна быть картинка по теме]"},
            {"type": "bullets", "head": "Основные факты", "items": ["Факт №1: Самый интересный", "Факт №2: Статистика или цифры", "Факт №3: Историческая справка"]},
            {"type": "quote", "head": "Мнение эксперта", "text": f"«{topic} — это явление, которое изменит наше понимание мира.»", "author": "Анонимный источник"},
            {"type": "conclusion", "head": "Заключение", "items": ["Итоги проведены", "Перспективы развития", "Спасибо за внимание!"]}
        ]

    def create_pptx_bytes(self, slides_data):
        prs = Presentation()
        prs.slide_width = Inches(13.333) # Wide format
        prs.slide_height = Inches(7.5)

        for i, slide_def in enumerate(slides_data):
            layout = prs.slide_layouts[6] # Blank layout
            slide = prs.slides.add_slide(layout)
            
            # Настройка фона
            bg = slide.background.fill
            bg.solid()
            
            is_title = (slide_def['type'] == 'title')
            is_quote = (slide_def['type'] == 'quote')
            
            if is_title or is_quote:
                bg.fore_color.rgb = self.colors["primary"]
                txt_col = self.colors["white"]
                acc_col = self.colors["accent"]
            else:
                bg.fore_color.rgb = self.colors["bg"]
                txt_col = self.colors["dark_text"]
                acc_col = self.colors["primary"]

            # --- РЕНДЕРИНГ КОНТЕНТА ---
            
            if is_title:
                # Большой заголовок по центру
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
                
                # Подзаголовок
                sb = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(11.3), Inches(1))
                sp = sb.text_frame.paragraphs[0]
                sp.alignment = PP_ALIGN.CENTER
                sr = sp.add_run()
                sr.text = slide_def.get('sub', '')
                sr.font.size = Pt(24)
                sr.font.color.rgb = acc_col

            elif is_quote:
                # Цитата крупным шрифтом
                qb = slide.shapes.add_textbox(Inches(2), Inches(2), Inches(9.3), Inches(3))
                qtf = qb.text_frame
                qtf.word_wrap = True
                qp = qtf.paragraphs[0]
                qp.alignment = PP_ALIGN.CENTER
                qr = qp.add_run()
                qr.text = slide_def['text']
                qr.font.size = Pt(32)
                qr.font.italic = True
                qr.font.name = "Georgia"
                qr.font.color.rgb = txt_col
                
                # Автор цитаты
                ab = slide.shapes.add_textbox(Inches(2), Inches(5.5), Inches(9.3), Inches(1))
                ap = ab.text_frame.paragraphs[0]
                ap.alignment = PP_ALIGN.RIGHT
                ar = ap.add_run()
                ar.text = f"— {slide_def.get('author', '')}"
                ar.font.size = Pt(20)
                ar.font.bold = True
                ar.font.color.rgb = acc_col

            else:
                # Стандартные слайды (Заголовок + Список/Картинка)
                # Заголовок слайда
                hb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.3), Inches(1))
                htf = hb.text_frame
                hp = htf.paragraphs[0]
                hr = hp.add_run()
                hr.text = slide_def['head']
                hr.font.size = Pt(36)
                hr.font.bold = True
                hr.font.color.rgb = acc_col
                
                # Линия-разделитель
                line = slide.shapes.add_shape(1, Inches(0.5), Inches(1.2), Inches(12.3), Pt(3)) # MSO_SHAPE.RECTANGLE approx
                line.fill.solid()
                line.fill.fore_color.rgb = self.colors["primary"]
                line.line.fill.background()

                if slide_def['type'] == 'bullets':
                    lb = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(5))
                    ltf = lb.text_frame
                    ltf.word_wrap = True
                    
                    for j, item in enumerate(slide_def['items']):
                        para = ltf.paragraphs[j] if j < len(ltf.paragraphs) else ltf.add_paragraph()
                        para.space_after = Pt(15)
                        r = para.add_run()
                        r.text = f"• {item}"
                        r.font.size = Pt(24)
                        r.font.color.rgb = txt_col
                        
                elif slide_def['type'] == 'image_placeholder':
                    img_box = slide.shapes.add_shape(1, Inches(3), Inches(2), Inches(7), Inches(4)) # Rectangle
                    img_box.fill.solid()
                    img_box.fill.fore_color.rgb = RGBColor(0xDDDDDD)
                    img_box.line.color.rgb = self.colors["primary"]
                    
                    caption = slide.shapes.add_textbox(Inches(3), Inches(6.2), Inches(7), Inches(0.5))
                    cp = caption.text_frame.paragraphs[0]
                    cp.alignment = PP_ALIGN.CENTER
                    cr = cp.add_run()
                    cr.text = slide_def['desc']
                    cr.font.size = Pt(18)
                    cr.font.color.rgb = self.colors["dark_text"]

        bytes_io = io.BytesIO()
        prs.save(bytes_io)
        bytes_io.seek(0)
        return bytes_io.getvalue()

# ==========================================
# 🎨 ИНТЕРФЕЙС STREAMLIT (ChatGPT Style)
# ==========================================
st.set_page_config(page_title="My Private AI Presenter", page_icon="🎨", layout="centered")

# CSS стилизация под современный чат
st.markdown("""
<style>
    /* Скрываем лишние элементы Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Стилизация чата */
    .stChatMessage {
        background-color: transparent !important;
        padding: 0 !important;
    }
    div[data-testid="stChatInputContainer"] {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        background: #fff;
        padding: 1rem;
        border-top: 1px solid #eee;
        z-index: 100;
    }
    body {
        background-color: #f9fafb;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    .user-bubble {
        background-color: #eef2ff;
        color: #1e1b4b;
        padding: 12px 16px;
        border-radius: 18px 18px 0 18px;
        max-width: 80%;
        margin-left: auto;
        display: inline-block;
    }
    .bot-bubble {
        background-color: #ffffff;
        color: #374151;
        padding: 12px 16px;
        border-radius: 18px 18px 18px 0;
        max-width: 80%;
        border: 1px solid #e5e7eb;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .download-btn {
        background-color: #FF6B35 !important;
        color: white !important;
        border: none !important;
        padding: 10px 20px !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        cursor: pointer !important;
    }
    .download-btn:hover {
        background-color: #e65c29 !important;
    }
</style>
""", unsafe_allow_html=True)

if check_password():
    st.sidebar.success("✅ Доступ разрешен")
    st.sidebar.info(f"Код: `{ACCESS_CODE}`")
    
    if st.sidebar.button("🚪 Выйти"):
        st.session_state.authenticated = False
        st.rerun()

    # Инициализация истории чата
    if 'messages' not in st.session_state:
        st.session_state.messages = []
        # Приветственное сообщение от бота
        welcome_msg = {
            "role": "assistant", 
            "content": "Привет! Я твой личный генератор презентаций. \n\nНапиши тему (например: *'Наполеон'* или *'Квантовая физика'*), и я сделаю крутую структуру и файл PowerPoint."
        }
        st.session_state.messages.append(welcome_msg)

    # Отображение истории сообщений
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])
            # Если есть кнопка скачивания в сообщении
            if "file_name" in message:
                st.download_button(
                    label=f"📥 Скачать {message['file_name']}",
                    data=message["file_bytes"],
                    file_name=message["file_name"],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    className="download-btn"
                )

    # Поле ввода внизу
    prompt = st.chat_input("О чем сделать презентацию?")

    if prompt:
        # 1. Добавляем запрос пользователя
        user_msg = {"role": "user", "content": prompt}
        st.session_state.messages.append(user_msg)
        with st.chat_message("user"):
            st.markdown(prompt)

        # 2. Генерируем ответ ассистента
        with st.chat_message("assistant"):
            status = st.empty()
            status.spinner("🤖 Думаю над структурой...")
            
            gen = PPTXGenerator()
            slides_data = gen.generate_content(prompt)
            
            status.spinner("🎨 Рисую слайды...")
            file_bytes = gen.create_pptx_bytes(slides_data)
            
            filename = f"Pres_{prompt.replace(' ', '_')[:10]}.pptx"
            
            status.write(f"✅ Готово! Презентация по теме **«{prompt}»** создана.")
            status.write("**Структура:**")
            for idx, s in enumerate(slides_data):
                title = s.get('text') or s.get('head')
                status.write(f"- Слайд {idx+1}: {title}")
            
            btn = st.download_button(
                label=f"📥 Скачать '{filename}'",
                data=file_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                className="download-btn"
            )
            
            # Сохраняем результат в историю, чтобы кнопка осталась при обновлении страницы
            bot_msg = {
                "role": "assistant",
                "content": f"Я подготовил презентацию по теме **«{prompt}»**. Нажми кнопку ниже.",
                "file_bytes": file_bytes,
                "file_name": filename
            }
            st.session_state.messages.append(bot_msg)
