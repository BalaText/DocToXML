from zipfile import ZipFile
from lxml import etree
import re
import configparser

# --- Setup ---
docx_file = "/Users/balajimohandoss/Desktop/1_Work_Folder/210525/03 - RL - Siani_Updated.docx"
output_file = "output.xml"
ini_path = "/Users/balajimohandoss/Desktop/1_Work_Folder/210525/style_map.ini"

# Namespaces for parsing DOCX XML
ns = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}

# Load style mappings from .ini file
config = configparser.ConfigParser()
config.optionxform = str
config.read(ini_path)

style_map = {}

for style_name, style_def in config.items("style_mapping"):
    parts = style_def.split("|")
    tag = parts[0].strip()
    attributes = []
    child = None
    wrap = None
    include_text = None

    for part in parts[1:]:
        part = part.strip()
        if part.startswith("child="):
            child = part.split("=", 1)[1].strip()
        elif part.startswith("wrap="):
            wrap = part.split("=", 1)[1].strip()
        elif part.startswith("include_text="):
            include_text = part.split("=", 1)[1].strip().strip('"')
        else:
            attributes.append(part)

    style_map[style_name] = {
        "tag": tag,
        "attributes": " " + " ".join(attributes) if attributes else "",
        "child": child,
        "wrap": wrap,
        "include_text": include_text
    }

print("Loaded styles:", style_map)

# Step 1: Extract external hyperlinks (optional, keep your own code if needed)
rels = {}
with ZipFile(docx_file) as docx:
    with docx.open("word/_rels/document.xml.rels") as rels_file:
        rels_tree = etree.parse(rels_file)
        for rel in rels_tree.xpath("//rels:Relationship", namespaces={
            "rels": "http://schemas.openxmlformats.org/package/2006/relationships"}):
            if rel.get("Type").endswith("hyperlink"):
                rels[rel.get("Id")] = rel.get("Target")

# Step 2: Read document content
with ZipFile(docx_file) as docx:
    xml = docx.read("word/document.xml")
tree = etree.fromstring(xml)

# Step 3: Find figure bookmarks (optional, keep your own code if needed)
bookmark_map = {}
for bookmark in tree.xpath(".//w:bookmarkStart", namespaces=ns):
    name = bookmark.get(f"{{{ns['w']}}}name")
    if re.match(r'HueD_Fig\d+', name):
        fig_num = re.search(r'\d+', name).group()
        fig_id = f"F{fig_num}"
        parent = bookmark.getparent()
        bookmark_map[fig_id] = parent

# Step 4: Parse paragraphs and output XML
output = '<?xml version="1.0" encoding="UTF-8"?>\n<document>\n'
wrap_buffer = {}    # collect content for wraps
wrap_started = {}   # mark if wrap started for each wrap tag
unmapped_styles = set()

for para in tree.xpath("//w:p", namespaces=ns):
    pStyle = para.xpath(".//w:pPr/w:pStyle/@w:val", namespaces=ns)
    style_name = pStyle[0] if pStyle else "Paragraph"
    style_info = style_map.get(style_name)

    if not style_info:
        tag_name = re.sub(r'[^a-zA-Z0-9]', '', style_name)
        attributes = ""
        child = None
        wrap = None
        include_text = None
        unmapped_styles.add(style_name)
    else:
        tag_name = style_info["tag"]
        attributes = style_info["attributes"]
        child = style_info.get("child")
        wrap = style_info.get("wrap")
        include_text = style_info.get("include_text")

    # Collect text inside paragraph (including hyperlinks and runs)
    para_text = ""
    for node in para:
        # Hyperlinks may contain runs
        if node.tag == f"{{{ns['w']}}}hyperlink":
            for r in node.xpath(".//w:r", namespaces=ns):
                texts = r.xpath(".//w:t/text()", namespaces=ns)
                if texts:
                    para_text += "".join(texts)
        elif node.tag == f"{{{ns['w']}}}r":
            texts = node.xpath(".//w:t/text()", namespaces=ns)
            if texts:
                para_text += "".join(texts)

    raw_text = para_text.strip()

    # Debug print to track processing and trigger_wrap
    print(f"Style: {style_name}, Tag: {tag_name}, Wrap: {wrap}, IncludeText: {include_text}, Text: '{raw_text}'")

    # Determine if wrap trigger hit
    trigger_wrap = wrap and include_text and (include_text in raw_text)

    from xml.sax.saxutils import escape

    content = f"<{tag_name}{attributes}>"
    if child:
        content += f"<{child}>{escape(para_text)}</{child}>"
    else:
        content += escape(para_text)
    content += f"</{tag_name}>"

    # Handle wrapping logic
    if wrap:
        if wrap not in wrap_buffer:
            wrap_buffer[wrap] = []

        if trigger_wrap and not wrap_started.get(wrap, False):
            # Insert heading at the start of buffer to ensure correct order
            wrap_buffer[wrap].insert(0, content)
            wrap_started[wrap] = True
        elif wrap_started.get(wrap, False):
            wrap_buffer[wrap].append(content)
        else:
            # Accumulate content before heading trigger
            wrap_buffer[wrap].append(content)

    else:
        fig_id = next((fid for fid, node in bookmark_map.items() if para == node), None)
        if fig_id:
            output += f"  <fig id=\"{fig_id}\">"
        else:
            output += "  "
        output += content
        if fig_id:
            output += f"</fig>\n"
        else:
            output += "\n"

# After processing all paragraphs, output wraps if any
for wrap_tag, elements in wrap_buffer.items():
    if wrap_started.get(wrap_tag, False):
        output += f"<{wrap_tag}>\n"
        for el in elements:
            output += f"  {el}\n"
        output += f"</{wrap_tag}>\n"
    else:
        # Wrap never started - output all elements directly
        for el in elements:
            output += f"  {el}\n"

output += "</document>"

from lxml import etree

# Parse the generated XML string back into an lxml tree
doc_tree = etree.fromstring(output.encode("utf-8"))

# Find the <title style='EH'>REFERENCES</title> element
refs_title = None
for el in doc_tree.iter():
    if el.tag == 'title' and el.get('style') == 'EH' and el.text and el.text.strip().upper() == 'REFERENCES':
        refs_title = el
        break

if refs_title is not None:
    parent = refs_title.getparent()
    siblings = list(parent)
    idx = siblings.index(refs_title)

    # Collect consecutive <ref style='REF'> elements immediately after REFERENCES title
    refs_to_wrap = []
    for sibling in siblings[idx+1:]:
        if sibling.tag == 'ref' and sibling.get('style') == 'REF':
            refs_to_wrap.append(sibling)
        else:
            break

    if refs_to_wrap:
        # Create <ref-list> element
        ref_list = etree.Element('ref-list')

        # Insert <ref-list> at the position of the first ref element
        first_ref_idx = siblings.index(refs_to_wrap[0])
        parent.insert(first_ref_idx, ref_list)

        # Move refs inside <ref-list>
        for ref in refs_to_wrap:
            parent.remove(ref)
            ref_list.append(ref)

# --- Insert reorder block here to move Author bio after References ---

def reorder_author_bio_after_references(root):
    refs_title = root.xpath("//title[text()='REFERENCES']")
    refs_list = root.xpath("//ref-list")
    author_title = root.xpath("//title[text()='Author bio']")
    bio_block = root.xpath("//BIO")

    if refs_title and refs_list and author_title and bio_block:
        refs_title = refs_title[0]
        refs_list = refs_list[0]
        author_title = author_title[0]
        bio_block = bio_block[0]

        # Remove Author bio elements from their current position
        author_title.getparent().remove(author_title)
        bio_block.getparent().remove(bio_block)

        # Insert Author bio after ref-list
        refs_list.addnext(author_title)
        author_title.addnext(bio_block)
    else:
        print("⚠️ Could not find all elements to reorder Author bio")

# Call reorder function
reorder_author_bio_after_references(doc_tree)

# Serialize back to string
output = etree.tostring(doc_tree, pretty_print=True, encoding='unicode')

# Save output XML
with open(output_file, "w", encoding="utf-8") as f:
    f.write(output)

if unmapped_styles:
    print("⚠️ Unmapped styles:", ", ".join(sorted(unmapped_styles)))

print("✅ Done. Output saved to:", output_file)
