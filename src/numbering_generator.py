# src/numbering_generator.py

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# WordprocessingML namespace
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def create_numbering_definitions(numbering_part, numbering_definitions):
    """
    In numbering.xml, creates custom multi-level list definitions (<w:abstractNum>)
    and instances (<w:num>) for the document to use.

    Args:
        numbering_part (lxml.etree._Element): The root element of numbering.xml.
        numbering_definitions (list): The numbering definitions parsed from the AI JSON.

    Returns:
        dict: A map from the definition name to its instance numId.
    """
    if not numbering_definitions:
        return {}

    num_id_map = {}

    # Find the highest existing IDs to avoid conflicts
    existing_abstract_ids = [int(an.get(qn('w:abstractNumId'))) for an in numbering_part.findall(qn('w:abstractNum'))]
    existing_num_ids = [int(n.get(qn('w:numId'))) for n in numbering_part.findall(qn('w:num'))]

    current_abstract_id = (max(existing_abstract_ids) if existing_abstract_ids else 0) + 1
    current_num_id = (max(existing_num_ids) if existing_num_ids else 0) + 1

    for definition in numbering_definitions:
        abstract_num_id = current_abstract_id

        abstract_num = OxmlElement('w:abstractNum')
        abstract_num.set(qn('w:abstractNumId'), str(abstract_num_id))

        multi_level_type = OxmlElement('w:multiLevelType')
        multi_level_type.set(qn('w:val'), 'multilevel')
        abstract_num.append(multi_level_type)

        for level_def in definition['levels']:
            lvl = OxmlElement('w:lvl')
            lvl.set(qn('w:ilvl'), str(level_def['level']))

            start = OxmlElement('w:start')
            start.set(qn('w:val'), '1')
            lvl.append(start)

            num_fmt = OxmlElement('w:numFmt')
            num_fmt.set(qn('w:val'), level_def['number_format'])
            lvl.append(num_fmt)

            lvl_text = OxmlElement('w:lvlText')
            lvl_text.set(qn('w:val'), level_def['text_format'])
            lvl.append(lvl_text)

            # Add some default indentation
            pPr = OxmlElement('w:pPr')
            ind = OxmlElement('w:ind')
            indent_val = 720 * level_def['level']  # 0.5 inch indent per level
            ind.set(qn('w:left'), str(indent_val))
            ind.set(qn('w:hanging'), '360')
            pPr.append(ind)
            lvl.append(pPr)

            abstract_num.append(lvl)

        numbering_part.append(abstract_num)

        # Create the instance that links to the abstract definition
        num = OxmlElement('w:num')
        num_id = current_num_id
        num.set(qn('w:numId'), str(num_id))

        abstract_num_id_val = OxmlElement('w:abstractNumId')
        abstract_num_id_val.set(qn('w:val'), str(abstract_num_id))
        num.append(abstract_num_id_val)

        numbering_part.append(num)

        num_id_map[definition['name']] = str(num_id)

        current_abstract_id += 1
        current_num_id += 1

    return num_id_map


def link_styles_to_numbering(styles_part, numbering_definitions, num_id_map):
    """
    In styles.xml, modifies the style definitions to link them to the specified numbering instance.

    Args:
        styles_part (lxml.etree._Element): The root element of styles.xml.
        numbering_definitions (list): The numbering definitions from the AI JSON.
        num_id_map (dict): The map from definition name to numId.
    """
    if not numbering_definitions:
        return

    for definition in numbering_definitions:
        num_id = num_id_map.get(definition['name'])
        if not num_id:
            continue

        for style_name, level in definition['style_links'].items():
            # Word style IDs often remove spaces.
            style_id_to_find = style_name.replace(" ", "")

            # Use XPath to find the style element
            style_element = styles_part.find(f".//w:style[@w:styleId='{style_id_to_find}']", namespaces={'w': W_NS})

            if style_element is not None:
                print(f"  > Successfully found style '{style_id_to_find}', linking numbering...")

                pPr = style_element.find('w:pPr', namespaces={'w': W_NS})
                if pPr is None:
                    pPr = OxmlElement('w:pPr')
                    style_element.append(pPr)

                # Find or create the numbering properties element
                numPr = pPr.find('w:numPr', namespaces={'w': W_NS})
                if numPr is None:
                    numPr = OxmlElement('w:numPr')
                    pPr.append(numPr)
                else:
                    numPr.clear()

                ilvl = OxmlElement('w:ilvl')
                ilvl.set(qn('w:val'), str(level))
                numPr.append(ilvl)

                numId_el = OxmlElement('w:numId')
                numId_el.set(qn('w:val'), num_id)
                numPr.append(numId_el)
            else:
                print(f"  > WARNING: Could not find style with ID '{style_id_to_find}' in styles.xml. Skipping link.")