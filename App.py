import streamlit as st
import pandas as pd
from pptx import Presentation
# import or paste your parse_content_to_bullets, create_slide, add_image_autofit,
# remove_unwanted_phrases, summarize_text, expand_text, decide_enrichment here

st.set_page_config(page_title="Powerpoint Generator", layout="wide")
st.title("Powerpoint Generator")

uploaded_file = st.file_uploader("Upload CSV or Excel", type=['csv', 'xlsx'])

if uploaded_file:
    if uploaded_file.name.endswith("csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    st.write("Slide Data Preview:")
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
            st.download_button("Download Powerpoint", f, file_name=pptx_file)
    st.info("Upload your slide data, review, then generate and download your PowerPoint.")
