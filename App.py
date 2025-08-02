import streamlit as st
import pandas as pd
import re
import io
import requests
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from PIL import Image
import nltk
nltk.download('punkt')

from transformers import pipeline

# Load summarizer and expander only once (cache for Streamlit)
@st.cache_resource
def load_summarizer():
    return pipeline("summarization", model="facebook/bart-large-cnn")

@st.cache_resource
def load_expander():
    return pipeline("text2text-generation", model="t5-base")

summarizer = load_summarizer()
expander = load_expander()

def remove_unwanted_phrases(text):
    unwanted_phrases = [
        "For confidential support call the Samaritans on 08457 90 90 90, visit a local Samaritans branch or see www.samaritans.org.",
        "For confidential support call the Samaritans on 08457 90 90, visit a local Samaritans branch or see www.samaritans.org.",
        "For confidential support, call the Samaritans on 08457 90 90 90, visit a local Samaritans branch or see www.samaritans.org.",
        "For confidential support, call the Samaritans on 08457 90 90, visit a local Samaritans branch or see www.samaritans.org.",
        "For confidential support call the Samaritans in the UK on 08457 90 90 90, visit a local Samaritans branch or see www.samaritans.org for details.",
        "For confidential support, call the Samaritans in the UK on 08457 90 90 90, visit a local Samaritans branch or see www.samaritans.org for details.",
        "For confidential support call the Samaritans",
        "For confidential support, call the Samaritans",
        "visit a local Samaritans branch or see www.samaritans.org",
        "visit a local Samaritans branch or see www.samaritans.org for details"
    ]
    for phrase in unwanted_phrases:
        text = text.replace(phrase, "")
    # Catch other possible forms
    text = re.sub(r'\bFor confidential support.*?samaritans\.org\.?', '', text, flags=re.IGNORECASE)
    text = re.sub(r'contact (us|this) .*?samaritans\.org\.?', '', text, flags=re.IGNORECASE)
    return text.strip()

def summarize_text(text):
    try:
        summary = summarizer(
            text, max_length=300, min_length=100, length_penalty=0.9,
            no_repeat_ngram_size=3, do_sample=False
        )
        result = summary[0]['summary_text'].strip()
        return remove_unwanted_phrases(result)
    except Exception as e:
        st.warning(f"Summarization failed: {e}")
        return remove_unwanted_phrases(text)

def expand_text(text):
    prompt = f"Provide a more detailed explanation: {text}"
    try:
        expanded = expander(prompt, max_new_tokens=120, do_sample=False)
        result = expanded[0]['generated_text'].strip()
        if result.lower().startswith(prompt.lower()) or len(result) <= len(text) + 8:
            return remove_unwanted_phrases(text)
        return remove_unwanted_phrases(result)
    except Exception as e:
        st.warning(f"Expansion failed: {e}")
        return remove_unwanted_phrases(text)

def parse_content_to_bullets(text):
    bullets = []
    for line in text.split('\n'):
        stripped = line.strip()
        if not stripped:
            continue
        if re.match(r'^[-*]\s', stripped):
            bullets.append(('sub', stripped[1:].strip()))
        else:
            bullets.append(('main', stripped))
    return bullets

def add_image_autofit(slide, image_path, left_inches, top_inches, width_inches, height_inches, min_size=1.0):
    try:
        if pd.isna(image_path) or not str(image_path).strip():
            return
        if str(image_path).startswith('http'):
            resp = requests.get(image_path)
            img = Image.open(io.BytesIO(resp.content))
        else:
            img = Image.open(image_path)
        img_w, img_h = img.size
        dpi = img.info.get('dpi', (96, 96))[0] or 96
        img_width_in = img_w / dpi
        img_height_in = img_h / dpi
        box_aspect = width_inches / height_inches
        img_aspect = img_width_in / img_height_in

        if img_aspect > box_aspect:
            final_width = max(min(width_inches, img_width_in), min_size)
            final_height = final_width / img_aspect
            if final_height < min_size:
                final_height = min_size
                final_width = final_height * img_aspect
        else:
            final_height = max(min(height_inches, img_height_in), min_size)
            final_width = final_height * img_aspect
            if final_width < min_size:
                final_width = min_size
                final_height = final_width / img_aspect

        pos_left = left_inches + max((width_inches - final_width) / 2, 0)
        pos_top = top_inches + max((height_inches - final_height) / 2, 0)

        img_obj = io.BytesIO(resp.content) if str(image_path).startswith('http') else image_path

        slide.shapes.add_picture(
            img_obj,
            Inches(pos_left), Inches(pos_top),
            width=Inches(final_width), height=Inches(final_height)
        )
    except Exception as e:
        st.warning(f"Could not add image {image_path}: {e}")

def create_slide(prs, title, content_bullets, image_path=None):
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = title
    para = title_shape.text_frame.paragraphs[0]
    para.font.size = Pt(34)
    para.font.bold = True
    para.font.color.rgb = RGBColor(10, 60, 120)
    slide_width_in = prs.slide_width / 914400
    slide_height_in = prs.slide_height / 914400
    margin = 0.5
    gap = 0.3
    MAX_BULLETS = 10

    if image_path and str(image_path).strip():
        img_area_width = max(slide_width_in * 0.38, 2.5)
        text_area_width = slide_width_in - img_area_width - gap - 2 * margin
        text_area_left = margin
        text_area_top = 1.5
        text_area_height = slide_height_in - text_area_top - margin
        img_left = text_area_left + text_area_width + gap
        img_top = text_area_top
        img_width = img_area_width
        img_height = text_area_height
        for shape in slide.shapes:
            if shape.is_placeholder and shape.placeholder_format.idx == 1:
                slide.shapes._spTree.remove(shape._element)
                break
        text_box = slide.shapes.add_textbox(
            Inches(text_area_left), Inches(text_area_top),
            Inches(text_area_width), Inches(text_area_height)
        )
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        text_frame.clear()
        for i, (level, text) in enumerate(content_bullets):
            if i >= MAX_BULLETS:
                paragraph = text_frame.add_paragraph()
                paragraph.text = "... (content truncated)"
                paragraph.font.size = Pt(18)
                paragraph.font.color.rgb = RGBColor(150, 150, 150)
                paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
                break
            paragraph = text_frame.add_paragraph()
            paragraph.text = text
            paragraph.level = 0 if level == 'main' else 1
            paragraph.font.size = Pt(20) if level == 'main' else Pt(18)
            paragraph.font.color.rgb = RGBColor(35, 35, 35)
            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
        add_image_autofit(slide, image_path, img_left, img_top, img_width, img_height)
    else:
        text_area_left = margin
        text_area_top = 1.5
        text_area_width = slide_width_in - 2 * margin
        text_area_height = slide_height_in - text_area_top - margin
        text_box = slide.shapes.add_textbox(
            Inches(text_area_left), Inches(text_area_top),
            Inches(text_area_width), Inches(text_area_height)
        )
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        text_frame.clear()
        for i, (level, text) in enumerate(content_bullets):
            if i >= MAX_BULLETS:
                paragraph = text_frame.add_paragraph()
                paragraph.text = "... (content truncated)"
                paragraph.font.size = Pt(18)
                paragraph.font.color.rgb = RGBColor(150, 150, 150)
                paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
                break
            paragraph = text_frame.add_paragraph()
            paragraph.text = text
            paragraph.level = 0 if level == 'main' else 1
            paragraph.font.size = Pt(20) if level == 'main' else Pt(18)
            paragraph.font.color.rgb = RGBColor(35, 35, 35)
            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
    return slide

def decide_enrichment(title, content_raw):
    if not content_raw or str(content_raw).lower() == "nan":
        return "Content not available."
    cleaned_text = remove_unwanted_phrases(content_raw)
    length = len(cleaned_text)
    if length < 80:
        st.info(f"[{title}] Expanding short content ({length} chars).")
        return expand_text(cleaned_text)
    if length > 100:
        st.info(f"[{title}] Summarizing long content ({length} chars).")
        return summarize_text(cleaned_text)
    st.info(f"[{title}] Using as-is ({length} chars).")
    return cleaned_text

# --- Streamlit APP UI ---

st.set_page_config(page_title="Powerpoint Generator", layout="wide")
st.title("Powerpoint Generator – AI-Powered PPTX from Tabular Data")

uploaded_file = st.file_uploader("Upload CSV or Excel for your slides", type=["csv", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    df.columns = df.columns.str.strip()
    st.write("Table Preview:")
    st.dataframe(df)

    if st.button("Generate Powerpoint"):
        prs = Presentation()
        for idx, row in df.iterrows():
            title = str(row.get('Title', 'Untitled Slide'))
            content_raw = str(row.get('Content', ''))
            content_used = decide_enrichment(title, content_raw)
            bullets = parse_content_to_bullets(content_used)
            image_path = row.get('Image', None)
            create_slide(prs, title, bullets, image_path)
        pptx_file = "Powerpoint Generator.pptx"
        prs.save(pptx_file)
        with open(pptx_file, "rb") as f:
            st.download_button("Download Powerpoint", f, file_name=pptx_file, mime='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        st.success("Presentation generated and ready for download!")
    st.caption("Expand, summarize, and generate *beautiful presentations* from structured data – fully within your browser.")
