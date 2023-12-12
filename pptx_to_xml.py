# -*- coding: utf-8 -*-
"""
Created on Mon Dec 11 14:11:18 2023

@author: rbuck
"""

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import xml.etree.ElementTree as ET

def ungroup_shapes(slide):
    while any(shape.shape_type == MSO_SHAPE_TYPE.GROUP for shape in slide.shapes):
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                group = shape.element
                parent = group.getparent()
                index = parent.index(group)
                for member in group:
                    parent.insert(index,member)
                    index+=1
                parent.remove(group)

def process_text(shape, slide_element):
    """
    Function to process pptx shape's text and add to xml tree 

    Parameters
    ----------
    shape : pptx.slide.Slide
        python-pptx shape object with text_frame.
    slide_element : etree.ElementTree.Element
        parent slide.

    Returns
    -------
    None.

    """
    if "Title" in shape.name:
        text_type = 'title'
    elif "Subtitle" in shape.name:
        text_type = 'subtitle'
    elif "Footnote" in shape.name:
        text_type = 'footnote'
    elif "Placeholder" in shape.name:
        text_type = 'placeholder'
    else:
        text_type = 'body'
    for paragraph in shape.text_frame.paragraphs:
        level = paragraph.level
        text_element = ET.SubElement(slide_element, "text", attrib={'text_type':text_type,'level':str(level)})
        text_element.text = paragraph.text
        
def process_table(shape, slide_element):
    table = shape.table
    n_rows = len(table.rows)
    n_cols = len(table.columns)
    table_element = ET.SubElement(slide_element, "table", attrib={'rows':n_rows,'columns':n_cols})
    for r in range(0, n_rows):
        for c in range(0,n_cols):
            cell = table.cell(r,c)
            cell_element = ET.SubElement(table_element, "cell", attrib={'row_id':r,'c_id':c})
            if cell.has_text_frame:
                for paragraph in cell.text_frame.paragraphs:
                    level = paragraph.level
                    text_element = ET.SubElement(cell_element, "text", attrib={'text_type':'table_text','level':str(level)})
                    text_element.text = paragraph.text
            

def pptx_to_xml(pptx_path, xml_path = None, encoding = 'utf-8', pptx_attributes = None):
    """
    Extracts structured text from Microsoft pptx in the form of an xml

    Parameters
    ----------
    pptx_path : str
        filepath of pptx.
    xml_path : str, optional
        filepath to save xml. The default is None.
    encoding : str, optional
        text encoding for saving xml. The default is 'utf-8'.
    pptx_attributes : dict, optional
        attribute information (metadata) for presentation. The default is None.

    Returns
    -------
    xml if no path to save is provided.

    """
    powerpoint = Presentation(pptx_path)
    
    root = ET.Element('presentation')
    
    for i,slide in enumerate(powerpoint.slides):
        slide_element = ET.SubElement(root, "slide", attrib={'number':str(i)})
                
        ungroup_shapes(slide)
        
        for shape in slide.shapes:
            if shape.has_text_frame:
                if len(shape.text)>0:
                    process_text(shape, slide_element)
            if shape.has_table:
                process_table(shape, slide_element)
                    
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text
            if len(notes) > 0:
                text_element = ET.SubElement(slide_element, "text", attrib={'text_type':'note'})
                text_element.text = notes
            
    tree = ET.ElementTree(root)
    
    if pptx_attributes is not None:
        for key,value in pptx_attributes.items():
            root.set(key,value)
    
    if xml_path is None:
        return(tree)
    else:
        tree.write(xml_path, encoding=encoding, xml_declaration= True)
