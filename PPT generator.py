# PPT generator.py

# (Paste the **full same set** of helper functions from app.py above:
# remove_unwanted_phrases, summarize_text, expand_text, parse_content_to_bullets,
# add_image_autofit, create_slide, decide_enrichment)

def main():
    file_path = ask_for_file()
    if not file_path:
        print("No file selected. Exiting.")
        return

    if file_path.endswith('.csv'):
        try:
            df = pd.read_csv(file_path, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(file_path, encoding='cp1252')
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, encoding='latin1')
    else:
        df = pd.read_excel(file_path)

    df.columns = df.columns.str.strip()
    prs = Presentation()

    for idx, row in df.iterrows():
        title = str(row.get('Title', 'Untitled Slide'))
        content_raw = str(row.get('Content', ''))
        content_used = decide_enrichment(title, content_raw)
        bullets = parse_content_to_bullets(content_used)
        image_path = row.get('Image', None)
        create_slide(prs, title, bullets, image_path)

    output_file = 'Powerpoint Generator.pptx'
    prs.save(output_file)
    print(f'Presentation generated: {output_file}')

if __name__ == '__main__':
    main()
