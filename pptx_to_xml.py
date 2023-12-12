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

def pptx_to_xml(pptx_path, xml_path = None, encoding = 'utf-8', pptx_attributes = None):
    powerpoint = Presentation(pptx_path)
    
    root = ET.Element('presentation')
    
    for i,slide in enumerate(powerpoint.slides):
        slide_element = ET.SubElement(root, "slide", attrib={'number':str(i)})
                
        ungroup_shapes(slide)
        
        for shape in slide.shapes:
            if shape.has_text_frame:
                if len(shape.text)>0:
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
    
        