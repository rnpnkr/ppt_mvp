import os
import zipfile
import json
from lxml import etree
import shutil
from ppt_to_xml import pptx_to_xml

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
}

def extract_background(slide_elem, rels_elem, extract_dir, slide_index):
    bg = slide_elem.find('.//p:bg', NS)
    if bg is None:
        return {"type": "color", "value": "FFFFFF"}
    bg_pr = bg.find('.//p:bgPr', NS)
    if bg_pr is not None:
        solid_fill = bg_pr.find('.//a:solidFill/a:srgbClr', NS)
        if solid_fill is not None:
            return {"type": "color", "value": solid_fill.get('val', 'FFFFFF')}
        scheme_fill = bg_pr.find('.//a:solidFill/a:schemeClr', NS)
        if scheme_fill is not None:
            return {"type": "color", "value": scheme_fill.get('val', 'FFFFFF')}
        blip_fill = bg_pr.find('.//a:blipFill/a:blip', NS)
        if blip_fill is not None:
            r_id = blip_fill.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            rel = rels_elem.find(f'.//r:Relationship[@Id="{r_id}"]', NS) if rels_elem is not None else None
            if rel:
                img_path = rel.get('Target', '').replace('../media/', '')
                full_path = os.path.join(extract_dir, 'ppt', 'media', img_path)
                if os.path.exists(full_path):
                    output_path = f"background_slide_{slide_index}.{img_path.split('.')[-1]}"
                    shutil.copy(full_path, output_path)
                    return {"type": "image", "file": output_path}
    return {"type": "color", "value": "FFFFFF"}

def extract_text_attributes(run):
    attrs = {
        "size": 18.0,
        "bold": False,
        "italic": False,
        "underline": False,
        "font": "Arial",
        "color": "000000"
    }
    rpr = run.find('.//a:rPr', NS)
    if rpr is not None:
        attrs["size"] = int(rpr.get('sz', 1800)) / 100
        attrs["bold"] = rpr.get('b') == '1'
        attrs["italic"] = rpr.get('i') == '1'
        attrs["underline"] = rpr.get('u') is not None and rpr.get('u') != 'none'
        latin = rpr.find('.//a:latin', NS)
        if latin is not None:
            attrs["font"] = latin.get('typeface', 'Arial')
        color = rpr.find('.//a:srgbClr', NS)
        if color is not None:
            attrs["color"] = color.get('val', '000000')
        scheme_color = rpr.find('.//a:schemeClr', NS)
        if scheme_color is not None:
            attrs["color"] = scheme_color.get('val', '000000')
    return attrs

def extract_position(element):
    off = element.find('.//a:off', NS)
    ext = element.find('.//a:ext', NS)
    return {
        "x": int(off.get('x', 0)) if off is not None else 0,
        "y": int(off.get('y', 0)) if off is not None else 0,
        "width": int(ext.get('cx', 0)) if ext is not None else 0,
        "height": int(ext.get('cy', 0)) if ext is not None else 0
    }

def extract_shape_style(shape):
    style = {}
    solid_fill = shape.find('.//a:solidFill/a:srgbClr', NS)
    if solid_fill is not None:
        style["fill_color"] = solid_fill.get('val')
    scheme_fill = shape.find('.//a:solidFill/a:schemeClr', NS)
    if scheme_fill is not None:
        style["fill_color"] = scheme_fill.get('val')
    ln = shape.find('.//a:ln', NS)
    if ln is not None:
        ln_fill = ln.find('.//a:solidFill/a:srgbClr', NS)
        if ln_fill is not None:
            style["border_color"] = ln_fill.get('val')
            style["border_width"] = int(ln.get('w', 0)) / 9525 if ln.get('w') else 0
        ln_scheme = ln.find('.//a:solidFill/a:schemeClr', NS)
        if ln_scheme is not None:
            style["border_color"] = ln_scheme.get('val')
            style["border_width"] = int(ln.get('w', 0)) / 9525 if ln.get('w') else 0
        if "border_color" not in style:
            style["border_color"] = None
            style["border_width"] = 0
    xfrm = shape.find('.//a:xfrm', NS)
    if xfrm is not None:
        style["rotation"] = int(xfrm.get('rot', 0)) / 60000
    return style

def group_text_content(paragraphs):
    content = []
    for p in paragraphs:
        runs = p.findall('.//a:r', NS)
        p_text = ""
        current_attrs = None
        for r in runs:
            text = r.find('.//a:t', NS)
            if text is not None and text.text:
                p_text += text.text
                current_attrs = extract_text_attributes(r)
        if p_text:
            content.append({"text": p_text, "attributes": current_attrs or extract_text_attributes(p)})
    return content

def extract_text_shape(shape, is_header=False, is_background=False, is_page_number=False):
    text_data = {
        "type": "text",
        "content": [],
        "position": extract_position(shape),
        "shape_background": None,
        "z_order": int(shape.get('order', 0)),
        "is_header": is_header,
        "is_background": is_background,
        "is_page_number": is_page_number
    }
    tx_body = shape.find('.//p:txBody', NS)
    if tx_body is not None:
        text_data["content"] = group_text_content(tx_body.findall('.//a:p', NS))
        if not text_data["content"]:
            return None
    fill = shape.find('.//a:solidFill/a:srgbClr', NS)
    if fill is not None:
        text_data["shape_background"] = fill.get('val')
    scheme_fill = shape.find('.//a:solidFill/a:schemeClr', NS)
    if scheme_fill is not None:
        text_data["shape_background"] = scheme_fill.get('val')
    return text_data

def extract_shape(shape):
    geom = shape.find('.//a:prstGeom', NS)
    if geom is None:
        return None
    return {
        "type": "shape",
        "shape_type": geom.get('prst', 'unknown'),
        "position": extract_position(shape),
        "style": extract_shape_style(shape),
        "z_order": int(shape.get('order', 0))
    }

def extract_image(pic, rels_elem, extract_dir, slide_index):
    blip = pic.find('.//a:blip', NS)
    if blip is None:
        return None
    r_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
    rel = rels_elem.find(f'.//r:Relationship[@Id="{r_id}"]', NS) if rels_elem is not None else None
    if rel:
        img_path = rel.get('Target', '').replace('../media/', '')
        full_path = os.path.join(extract_dir, 'ppt', 'media', img_path)
        if os.path.exists(full_path):
            output_path = f"image_slide_{slide_index}_{os.path.basename(img_path)}"
            shutil.copy(full_path, output_path)
            return {
                "type": "image",
                "file": output_path,
                "position": extract_position(pic),
                "z_order": int(pic.get('order', 0))
            }
    return None

def extract_table(graphic_frame):
    table = graphic_frame.find('.//a:tbl', NS)
    if table is None:
        return None
    table_data = {
        "type": "table",
        "position": extract_position(graphic_frame),
        "rows": [],
        "z_order": int(graphic_frame.get('order', 0))
    }
    for tr in table.findall('.//a:tr', NS):
        row = []
        for tc in tr.findall('.//a:tc', NS):
            text = tc.find('.//a:t', NS)
            content = text.text or "" if text is not None else ""
            row.append({"content": content, "attributes": extract_text_attributes(tc)})
        table_data["rows"].append(row)
    return table_data

def extract_chart(graphic_frame, rels_elem, extract_dir, slide_index):
    chart = graphic_frame.find('.//c:chart', NS)
    if chart is None:
        return None
    r_id = chart.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    rel = rels_elem.find(f'.//r:Relationship[@Id="{r_id}"]', NS) if rels_elem is not None else None
    if rel:
        chart_path = rel.get('Target', '').replace('../charts/', '')
        full_path = os.path.join(extract_dir, 'ppt', 'charts', chart_path)
        if os.path.exists(full_path):
            output_path = f"chart_slide_{slide_index}_{os.path.basename(chart_path)}"
            shutil.copy(full_path, output_path)
            return {
                "type": "chart",
                "chart_file": output_path,
                "position": extract_position(graphic_frame),
                "z_order": int(graphic_frame.get('order', 0))
            }
    return None

def parse_master(master_elem):
    background_elements = []
    for sp in master_elem.findall('.//p:sp', NS):
        ph = sp.find('.//p:nvSpPr/p:ph', NS)
        is_background = ph is not None and ph.get('type') in ['title', 'sldNum', 'ftr', 'hdr']
        is_page_number = ph is not None and ph.get('type') == 'sldNum'
        text_shape = extract_text_shape(sp, is_header=ph is not None and ph.get('type') in ['title', 'hdr'], 
                                        is_background=is_background, is_page_number=is_page_number)
        if text_shape and text_shape["content"]:
            background_elements.append(text_shape)
        else:
            shape = extract_shape(sp)
            if shape:
                background_elements.append(shape)
    return background_elements

def parse_slide(slide_elem, rels_elem, extract_dir, slide_index, background_elements):
    slide_data = {
        "background": extract_background(slide_elem, rels_elem, extract_dir, slide_index),
        "background_elements": background_elements.copy(),
        "elements": []
    }
    
    shapes = slide_elem.findall('.//p:sp', NS)
    for i, sp in enumerate(shapes):
        ph = sp.find('.//p:nvSpPr/p:ph', NS)
        is_header = ph is not None and ph.get('type') in ['title', 'ctrTitle']
        is_background = ph is not None and ph.get('type') in ['sldNum', 'ftr', 'hdr']
        if not is_background:
            text_shape = extract_text_shape(sp, is_header=is_header)
            if text_shape:
                if ph and ph.get('type') == 'sldNum':
                    text_shape["is_page_number"] = True
                slide_data["elements"].append(text_shape)
            else:
                shape = extract_shape(sp)
                if shape:
                    slide_data["elements"].append(shape)
    
    for pic in slide_elem.findall('.//p:pic', NS):
        image = extract_image(pic, rels_elem, extract_dir, slide_index)
        if image:
            slide_data["elements"].append(image)
    
    for gf in slide_elem.findall('.//p:graphicFrame', NS):
        table = extract_table(gf)
        if table:
            slide_data["elements"].append(table)
        chart = extract_chart(gf, rels_elem, extract_dir, slide_index)
        if chart:
            slide_data["elements"].append(chart)
    
    slide_data["elements"].sort(key=lambda x: (x["z_order"], x["position"]["y"]))
    return slide_data

def parse_layout(layout_elem):
    c_sld = layout_elem.find('.//p:cSld', NS)
    name = c_sld.get('name', 'unknown') if c_sld is not None else 'unknown'
    placeholders = [
        {
            "type": ph.get('type', 'body'),
            "position": extract_position(sp)
        }
        for sp in layout_elem.findall('.//p:sp', NS)
        if (ph := sp.find('.//p:nvSpPr/p:ph', NS)) is not None
    ]
    return {"name": name, "placeholders": placeholders}

def parse_theme(theme_elem):
    clr_scheme = theme_elem.find('.//a:clrScheme', NS)
    theme_data = {"colors": {}}
    if clr_scheme is not None:
        for clr in clr_scheme:
            srgb_clr = clr.find('.//a:srgbClr', NS)
            if srgb_clr is not None:
                theme_data["colors"][clr.tag.split('}')[-1]] = srgb_clr.get('val')
    return theme_data

def main(pptx_file):
    xml_output_dir = "/Users/aryan98/ppt_mvp/parser/extracted_xml"
    json_output_dir = "/Users/aryan98/ppt_mvp/parser/extracted_json"

    # Ensure directories exist
    os.makedirs(xml_output_dir, exist_ok=True)
    os.makedirs(json_output_dir, exist_ok=True)

    # Extract the base name from pptx_file (e.g., "test diyea 3")
    base_name = os.path.splitext(os.path.basename(pptx_file))[0]
    xml_file = os.path.join(xml_output_dir, f"{base_name}.xml")
    output_file = os.path.join(json_output_dir, f"{base_name}.json")

    # Call pptx_to_xml and get the extraction directory
    extract_dir = pptx_to_xml(pptx_file, output_xml_file=xml_file)
    if extract_dir is None:
        print("Failed to extract PPTX in ppt_to_xml, aborting.")
        return

    # Load the generated XML
    parser = etree.XMLParser(remove_blank_text=True)
    try:
        tree = etree.parse(xml_file, parser)
    except Exception as e:
        print(f"Error loading {xml_file}: {e}")
        return
    
    root = tree.getroot()
    output_data = {"slides": [], "template": {"layouts": [], "theme": {}}}
    
    # Get slide order from presentation.xml
    pres_elem = root.find('.//presentation')
    slide_order = []
    if pres_elem is not None:
        slide_ids = [(sld.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'), 
                      sld.get('id')) 
                     for sld in pres_elem.findall('.//p:sldIdLst/p:sldId', NS)]
        slide_order = [f"slide{i+1}.xml" for i in range(len(slide_ids))]
        print(f"Slide order from presentation.xml: {slide_order}")
    else:
        print("No presentation element found in XML")
    
    # Parse slide masters
    master_elem = root.find('.//slideMasters/slideMaster')
    background_elements = parse_master(master_elem) if master_elem is not None else []
    print(f"Background elements parsed: {len(background_elements)}")

    # Parse slides
    slides_elem = root.find('.//slides')
    if slides_elem is not None:
        for i, slide_elem in enumerate(slides_elem.findall('.//slide')):
            slide_file = slide_elem.get('file')
            print(f"Processing slide: {slide_file}")
            if slide_file in slide_order:
                slide_idx = slide_order.index(slide_file)
                rels_file = f"{slide_file}.rels"
                rels_elem = root.find(f'.//relationships/relationship[@file="{rels_file}"]')
                if rels_elem is None:
                    print(f"Warning: No relationships found for {rels_file}, using None")
                slide_data = parse_slide(slide_elem, rels_elem, extract_dir, slide_idx, background_elements)
                output_data["slides"].append(slide_data)
            else:
                print(f"Slide {slide_file} not in expected order, skipping")
        print(f"Total slides parsed: {len(output_data['slides'])}")
    else:
        print("No slides element found in XML")

    # Parse layouts
    layouts_elem = root.find('.//slideLayouts')
    if layouts_elem is not None:
        for layout_elem in layouts_elem.findall('.//slideLayout'):
            layout_data = parse_layout(layout_elem)
            output_data["template"]["layouts"].append(layout_data)
    
    # Parse theme
    theme_elem = root.find('.//themes/theme')
    if theme_elem is not None:
        output_data["template"]["theme"] = parse_theme(theme_elem)
    
    # Save JSON
    with open(output_file, 'w') as f:
        json.dump(output_data, f, indent=2)
        print(f"Parsed data saved to {output_file}")

if __name__ == "__main__":
    pptx_file = "/Users/aryan98/ppt_mvp/ppt_samples/test diyea 3.pptx"
    main(pptx_file)