import streamlit as st
import json
import io
import re
import zipfile
import docx
from docx import Document
from docx.shared import Inches

# --- HJÄLPFUNKTIONER ---

def add_hyperlink(paragraph, text, url):
    """Skapar en klickbar länk i Word."""
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id)
    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')
    c = docx.oxml.shared.OxmlElement('w:color')
    c.set(docx.oxml.shared.qn('w:val'), "0000FF")
    rPr.append(c)
    u = docx.oxml.shared.OxmlElement('w:u')
    u.set(docx.oxml.shared.qn('w:val'), 'single')
    rPr.append(u)
    new_run.append(rPr)
    t = docx.oxml.shared.OxmlElement('w:t')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def normalize(name):
    """Tvättar filnamn för matchning."""
    clean = name.replace('\xa0', ' ').replace('□', ' ')
    return re.sub(r'[^a-zA-Z0-9]', '', clean).lower()

def find_image_in_uploads(target_name, uploaded_images):
    """Letar bild bland uppladdningar."""
    target_norm = normalize(target_name)
    for original_path, file_obj in uploaded_images.items():
        if target_norm in normalize(original_path):
            return file_obj
    return None

def process_node_list(nodes, doc_obj, paragraph_obj, uploaded_images, warnings, note_name):
    """Huvudparser för den moderna BoxNote-strukturen."""
    for node in nodes:
        n_type = node.get('type')
        
        if n_type == 'text':
            text = node.get('text', '')
            marks = node.get('marks', [])
            link = next((m['attrs']['href'] for m in marks if m['type'] == 'link'), None)
            if link:
                add_hyperlink(paragraph_obj, text, link)
            else:
                run = paragraph_obj.add_run(text)
                if any(m['type'] == 'strong' for m in marks): run.bold = True
                    
        elif n_type == 'image':
            img_fn = node.get('attrs', {}).get('fileName')
            if img_fn:
                img_data = find_image_in_uploads(img_fn, uploaded_images)
                if img_data:
                    # Vi skapar en kopia av datan för att inte stänga strömmen
                    img_bytes = io.BytesIO(img_data.getvalue())
                    doc_obj.add_picture(img_bytes, width=Inches(5.5))
                else:
                    warnings.append(f"Saknas: {img_fn} i {note_name}")

        elif n_type in ['paragraph', 'heading']:
            level = node.get('attrs', {}).get('level', 0) if n_type == 'heading' else 0
            p = doc_obj.add_heading('', level=level) if level > 0 else doc_obj.add_paragraph()
            if 'content' in node:
                process_content_recursive(node['content'], doc_obj, p, uploaded_images, warnings, note_name)

        elif n_type in ['bullet_list', 'ordered_list']:
            for item in node.get('content', []): # list_item
                for sub in item.get('content', []):
                    p = doc_obj.add_paragraph(style='List Bullet')
                    process_content_recursive(sub.get('content', []), doc_obj, p, uploaded_images, warnings, note_name)

def process_content_recursive(content, doc_obj, p_obj, images, warnings, name):
    """Hjälpfunktion för rekursion."""
    if content:
        process_node_list(content, doc_obj, p_obj, images, warnings, name)

# --- APP ---

st.set_page_config(page_title="BoxNote Pro", page_icon="📝")
st.title("📝 BoxNote Pro-konverterare")

files = st.file_uploader("Dra in .boxnote och bildmapp (eller bilder)", accept_multiple_files=True)

if files:
    notes = [f for f in files if f.name.endswith('.boxnote')]
    imgs = {f.name: f for f in files if f.name.lower().endswith(('.png', '.jpg', '.jpeg'))}
    
    if notes and st.button(f"Konvertera {len(notes)} filer"):
        zip_buf = io.BytesIO()
        all_warns = []

        with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
            for n_file in notes:
                try:
                    raw_data = json.loads(n_file.getvalue().decode("utf-8"))
                    doc = Document()
                    
                    # Logik för olika JSON-versioner
                    processed = False
                    
                    # Version A: Modern (doc -> content)
                    if 'doc' in raw_data:
                        process_node_list(raw_data['doc'].get('content', []), doc, None, imgs, all_warns, n_file.name)
                        processed = True
                    
                    # Version B: Gammal (ate -> text)
                    elif 'ate' in raw_data and 'text' in raw_data['ate']:
                        doc.add_heading(n_file.name, 0)
                        for line in raw_data['ate']['text'].split('\n'):
                            doc.add_paragraph(line)
                        processed = True

                    if not processed:
                        st.error(f"Kunde inte tolka formatet i: {n_file.name}")
                        st.write(f"Hittade nycklar: {list(raw_data.keys())}")
                    else:
                        d_buf = io.BytesIO()
                        doc.save(d_buf)
                        zip_f.writestr(n_file.name.replace(".boxnote", ".docx"), d_buf.getvalue())

                except Exception as e:
                    st.error(f"Fel vid {n_file.name}: {e}")

        if all_warns:
            with st.expander("Saknade bilder"):
                for w in all_warns: st.warning(w)

        st.success("Konvertering klar!")
        st.download_button("📥 Ladda ner Word-filer (ZIP)", zip_buf.getvalue(), "boxnotes_export.zip")