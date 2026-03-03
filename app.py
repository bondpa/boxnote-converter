import streamlit as st
import json
import io
import re
import zipfile
import base64
import urllib.parse
import unicodedata
import docx
from docx import Document
from docx.shared import Inches

# --- TERMINAL-LOGGNING ---
def log(msg):
    print(f"[LOGG] {msg}")

# --- HJÄLPFUNKTIONER ---

def add_hyperlink(paragraph, text, url):
    """Skapar en klickbar blå länk i Word-dokumentet."""
    try:
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
    except Exception:
        paragraph.add_run(f"{text} ({url})")

def normalize(name):
    """Normaliserar filnamn för matchning."""
    if not name: return ""
    n = unicodedata.normalize('NFD', name)
    n = "".join([c for c in n if unicodedata.category(c) != 'Mn']).lower()
    n = n.replace('\xa0', ' ').replace('□', ' ')
    return re.sub(r'[^a-z0-9]', '', n)

def find_image_in_uploads(target_name, target_id, uploaded_images):
    """Letar efter bild via namn eller Box-ID."""
    t_name_norm = normalize(target_name)
    t_id_str = str(target_id) if target_id else None

    for path, file_obj in uploaded_images.items():
        if t_id_str and t_id_str in path:
            return file_obj
        if t_name_norm and t_name_norm in normalize(path):
            return file_obj
    return None

def extract_unique_legacy_images(pool):
    """Extraherar unika bilder och deras Box-länkar från poolen."""
    unique_images = {}
    if not pool or 'numToAttrib' not in pool:
        return []
    
    for attr in pool['numToAttrib'].values():
        val = attr[0] if isinstance(attr, list) else str(attr)
        if val.startswith('image-'):
            try:
                parts = val.split('-')
                if len(parts) >= 3:
                    encoded_part = parts[-1]
                    padding = len(encoded_part) % 4
                    if padding: encoded_part += '=' * (4 - padding)
                    
                    decoded = base64.b64decode(urllib.parse.unquote(encoded_part)).decode('utf-8')
                    info = json.loads(urllib.parse.unquote(decoded))
                    
                    fname = info.get('fileName', 'image.png')
                    fid = info.get('boxFileId') or info.get('fileId')
                    flink = info.get('boxSharedLink') # Länken vi vill ha!
                    
                    name_key = normalize(fname)
                    if name_key not in unique_images:
                        unique_images[name_key] = {'name': fname, 'id': fid, 'link': flink}
            except Exception: pass
    return list(unique_images.values())

def process_node_list(nodes, doc_obj, paragraph_obj, uploaded_images, note_name):
    """Parser för det moderna formatet. Returnerar True om bilder saknas."""
    missing_any = False
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
            attrs = node.get('attrs', {})
            img = find_image_in_uploads(attrs.get('fileName'), attrs.get('boxFileId'), uploaded_images)
            if img:
                doc_obj.add_picture(io.BytesIO(img.getvalue()), width=Inches(5.5))
            else:
                p = doc_obj.add_paragraph("[BILD SAKNAS LOKALT] ")
                # Moderna noter har sällan direktlänkar inbäddade på samma sätt,
                # men vi skriver ut filnamnet så man vet vad man letar efter.
                p.add_run(f"Filnamn: {attrs.get('fileName')}").italic = True
                missing_any = True
        elif n_type in ['paragraph', 'heading']:
            level = node.get('attrs', {}).get('level', 0) if n_type == 'heading' else 0
            p = doc_obj.add_heading('', level=level) if level > 0 else doc_obj.add_paragraph()
            if 'content' in node:
                if process_node_list(node['content'], doc_obj, p, uploaded_images, note_name):
                    missing_any = True
    return missing_any

# --- APP ---
st.set_page_config(page_title="BoxNote Pro", layout="centered")
st.title("📝 BoxNote Pro-konverterare")

uploaded_files = st.file_uploader("Dra in filer här", accept_multiple_files=True)

if uploaded_files:
    notes = [f for f in uploaded_files if f.name.endswith('.boxnote')]
    imgs = {f.name: f for f in uploaded_files if f.name.lower().endswith(('.png', '.jpg', '.jpeg'))}
    
    if notes and st.button(f"Konvertera {len(notes)} noter"):
        zip_buf = io.BytesIO()

        with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_file:
            for n_file in notes:
                log(f"Bearbetar: {n_file.name}")
                try:
                    data = json.loads(n_file.getvalue().decode("utf-8"))
                    doc = Document()
                    missing_images = False
                    
                    # 1. MODERN
                    if 'doc' in data:
                        missing_images = process_node_list(data['doc'].get('content', []), doc, None, imgs, n_file.name)
                    
                    # 2. LEGACY
                    elif 'atext' in data:
                        doc.add_heading(n_file.name.replace(".boxnote", ""), 0)
                        for line in data['atext'].get('text', '').split('\n'):
                            doc.add_paragraph(line)
                        
                        legacy_imgs = extract_unique_legacy_images(data.get('pool', {}))
                        if legacy_imgs:
                            doc.add_heading("Bifogade bilder", level=2)
                            for info in legacy_imgs:
                                img_data = find_image_in_uploads(info['name'], info['id'], imgs)
                                if img_data:
                                    doc.add_picture(io.BytesIO(img_data.getvalue()), width=Inches(5.5))
                                else:
                                    missing_images = True
                                    p = doc.add_paragraph(f"⚠️ Bild saknas: {info['name']} - ")
                                    if info['link']:
                                        add_hyperlink(p, "KLICKA HÄR FÖR ATT SE BILDEN PÅ BOX", info['link'])
                                    else:
                                        p.add_run("(Ingen direktlänk hittades)")

                    # Spara filnamn med markering om något saknas
                    base_name = n_file.name.replace(".boxnote", ".docx")
                    final_name = f"[FIXA] {base_name}" if missing_images else base_name
                    
                    d_io = io.BytesIO()
                    doc.save(d_io)
                    zip_file.writestr(final_name, d_io.getvalue())
                    
                except Exception as e:
                    st.error(f"Fel vid {n_file.name}: {e}")

        st.success("Klar! Filer som saknar bilder är markerade med [FIXA].")
        st.download_button("📥 Ladda ner ZIP", zip_buf.getvalue(), "box_export.zip")