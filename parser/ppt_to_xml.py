import os
import zipfile
from lxml import etree

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
}

def load_xml(file_path):
    """Load an XML file and return its root element, or None if it fails."""
    if not os.path.exists(file_path):
        print(f"Warning: {file_path} not found, skipping.")
        return None
    parser = etree.XMLParser(remove_blank_text=True)
    try:
        return etree.parse(file_path, parser).getroot()
    except Exception as e:
        print(f"Error loading {file_path}: {e}")
        return None

def pptx_to_xml(pptx_path, output_xml_file):
    # Extract base name from pptx_path (e.g., "test diyea 3")
    base_name = os.path.splitext(os.path.basename(pptx_path))[0]
    # Define extraction directory as /parser/temp_pptx/{base_name}
    extract_dir = os.path.join(os.path.dirname(__file__), "temp_pptx", base_name)
    
    # Ensure extraction directory exists
    os.makedirs(extract_dir, exist_ok=True)
    
    # Extract PPTX contents to extract_dir
    try:
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        print(f"Extracted PPTX contents to {extract_dir}")
    except Exception as e:
        print(f"Error unzipping {pptx_path}: {e}")
        return None  # Return None on failure

    # Create the root element for the custom XML
    root = etree.Element("pptx")

    # 1. Load presentation.xml
    pres_path = os.path.join(extract_dir, "ppt", "presentation.xml")
    pres_xml = load_xml(pres_path)
    if pres_xml is not None:
        presentation_elem = etree.SubElement(root, "presentation")
        presentation_elem.append(pres_xml)
        slide_ids = [(sld.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'), 
                      sld.get('id')) 
                     for sld in pres_xml.findall('.//p:sldIdLst/p:sldId', NS)]
    else:
        slide_ids = []
        print("Warning: No presentation.xml found.")

    # 2. Load all slide masters
    master_dir = os.path.join(extract_dir, "ppt", "slideMasters")
    if os.path.exists(master_dir):
        master_files = sorted([f for f in os.listdir(master_dir) if f.endswith(".xml")])
        masters_elem = etree.SubElement(root, "slideMasters")
        for master_file in master_files:
            master_path = os.path.join(master_dir, master_file)
            master_xml = load_xml(master_path)
            if master_xml is not None:
                master_elem = etree.SubElement(masters_elem, "slideMaster", file=master_file)
                master_elem.append(master_xml)

    # 3. Load all themes
    theme_dir = os.path.join(extract_dir, "ppt", "theme")
    if os.path.exists(theme_dir):
        theme_files = sorted([f for f in os.listdir(theme_dir) if f.endswith(".xml")])
        themes_elem = etree.SubElement(root, "themes")
        for theme_file in theme_files:
            theme_path = os.path.join(theme_dir, theme_file)
            theme_xml = load_xml(theme_path)
            if theme_xml is not None:
                theme_elem = etree.SubElement(themes_elem, "theme", file=theme_file)
                theme_elem.append(theme_xml)

    # 4. Load all slide layouts
    layout_dir = os.path.join(extract_dir, "ppt", "slideLayouts")
    if os.path.exists(layout_dir):
        layout_files = sorted([f for f in os.listdir(layout_dir) if f.endswith(".xml")])
        layouts_elem = etree.SubElement(root, "slideLayouts")
        for layout_file in layout_files:
            layout_path = os.path.join(layout_dir, layout_file)
            layout_xml = load_xml(layout_path)
            if layout_xml is not None:
                layout_elem = etree.SubElement(layouts_elem, "slideLayout", file=layout_file)
                layout_elem.append(layout_xml)

    # 5. Load all slides with relationships
    slide_dir = os.path.join(extract_dir, "ppt", "slides")
    if os.path.exists(slide_dir):
        slide_files = sorted([f for f in os.listdir(slide_dir) if f.endswith(".xml")])
        slides_elem = etree.SubElement(root, "slides")
        for slide_file in slide_files:
            slide_path = os.path.join(slide_dir, slide_file)
            slide_xml = load_xml(slide_path)
            if slide_xml is not None:
                slide_elem = etree.SubElement(slides_elem, "slide", file=slide_file)
                slide_idx = slide_files.index(slide_file)
                if slide_idx < len(slide_ids):
                    r_id, sld_id = slide_ids[slide_idx]
                    slide_elem.set("rId", r_id)
                    slide_elem.set("sldId", sld_id or f"unknown_{r_id}")
                slide_elem.append(slide_xml)

    # 6. Include slide-specific relationships
    slide_rels_dir = os.path.join(extract_dir, "ppt", "slides", "_rels")
    if os.path.exists(slide_rels_dir):
        slide_rels_files = sorted([f for f in os.listdir(slide_rels_dir) if f.endswith(".rels")])
        rels_elem = etree.SubElement(root, "relationships")
        for rels_file in slide_rels_files:
            rels_path = os.path.join(slide_rels_dir, rels_file)
            rels_xml = load_xml(rels_path)
            if rels_xml is not None:
                rel_elem = etree.SubElement(rels_elem, "relationship", file=rels_file)
                rel_elem.append(rels_xml)

    # Save the combined XML
    tree = etree.ElementTree(root)
    with open(output_xml_file, 'wb') as f:
        tree.write(f, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    print(f"Combined PPTX XML saved to {output_xml_file}")
    print(f"Components included: "
          f"Slides={len(slide_files if 'slide_files' in locals() else [])}, "
          f"Layouts={len(layout_files if 'layout_files' in locals() else [])}, "
          f"Masters={len(master_files if 'master_files' in locals() else [])}, "
          f"Themes={len(theme_files if 'theme_files' in locals() else [])}")
    
    return extract_dir  # Return the extraction directory for reuse