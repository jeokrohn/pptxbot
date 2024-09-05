#!/usr/bin/env python

import re
from typing import Dict, List, Tuple
from zipfile import ZipFile, ZIP_DEFLATED
import logging
from lxml import etree
from collections import defaultdict

log = logging.getLogger(__name__)


class PresentationRel:
    FILE = 'ppt/_rels/presentation.xml.rels'

    def __init__(self, zip_file: ZipFile):
        with zip_file.open(name=self.FILE) as rel_file:
            tree = etree.parse(rel_file)
        root = tree.getroot()
        data = defaultdict(list)
        for child in root:
            target = child.attrib['Target']
            var = child.attrib['Type'].split('/')[-1]
            data[var].append(target)
        for k, v in sorted((k, v) for k, v in data.items()):
            if len(v) == 1:
                v = v[0]
            else:
                v.sort()
            self.__dict__[k] = v
            log.debug(f'presentation rel: {k}={v}')


def read_color_map(zip_file: ZipFile, file_name: str) -> Tuple[str, Dict[str, str]]:
    """
    read color map from given theme file
    :param zip_file:
    :param file_name:
    :return: theme name and dict to map from color name to RGB value
    """
    with zip_file.open(name=file_name) as theme_file:
        tree = etree.parse(theme_file)
    root = tree.getroot()
    nsmap = root.nsmap
    clr_scheme = tree.xpath('a:themeElements/a:clrScheme', namespaces=nsmap)[0]
    name = clr_scheme.get('name')
    colors = list(clr_scheme)
    rgb_map = {}
    for color in colors:
        tag = color.tag.split('}')[-1]
        srgb = color[0].get('val')
        rgb_map[tag] = srgb
    for tag, rgb in rgb_map.items():
        log.debug(f'color map {file_name}, {name}: {tag}->{rgb}')
    return name, rgb_map


def color_schemes(zip_file: ZipFile) -> List[Tuple[str, Dict[str, str]]]:
    """
    Get color schemes of all themes of the PPT
    :param zip_file:
    :return:
    """
    file_list = zip_file.namelist()
    theme_files = [f_name for f_name in file_list
                   if f_name.startswith('ppt/theme')]
    themes = [read_color_map(zip_file, name) for name in theme_files]
    return themes


PRE_MAP = {
    'bg1': "lt1",
    'tx1': "dk1",
    'bg2': "lt2",
    'tx2': "dk2"
}


def convert(content, color_map):
    """
    Convert a single XML. Replace all references to theme colors with actual RGB colors

    :param content: XML file contents
    :param color_map: color map from a PPT theme
    :return:
    """

    root = etree.fromstring(content)
    nsmap = root.nsmap
    if nsmap.get('a') is None:
        return content

    # find all rpr tags and make sure that there is a solidFill child
    # if it's missing, then create one with "schemeClr" "tx1"
    # this is to address text fields where no specific color has been applied and thus
    # the default tx1 applies
    rprs = root.xpath('//a:rPr', namespaces=nsmap)
    target_rprs = [rpr for rpr in rprs
                   if not rpr.xpath('a:solidFill', namespaces=nsmap)]
    # prepare a namespace for the attributes to be added
    ns = '{' + nsmap['a'] + '}'
    for rpr in target_rprs:
        # we are only interested in rpr tags of where the parent actually has text
        if t := rpr.xpath('../a:t', namespaces=nsmap):
            # solidFill as new child of rpr tag
            solid_fill = etree.SubElement(rpr, f'{ns}solidFill')
            # .. and the color is schemeClr tx1
            etree.SubElement(solid_fill, f'{ns}schemeClr', val='tx1')

    # look for schemeClr tags anywhere in the XML
    targets = root.xpath('//a:schemeClr', namespaces=nsmap)
    if not targets:
        return content

    for target in targets:
        value = target.get('val')
        pm_value = PRE_MAP.get(value, value)

        if pm_value != value:
            # log.debug(f'pre map {value}->{pm_value}')
            pass
        rgb = color_map.get(pm_value)
        # log.debug(f'color map {pm_value}->{rgb}')
        if rgb is None:
            # unknown theme color name?
            continue
        target.tag = target.tag.replace('schemeClr', 'srgbClr')
        target.attrib['val'] = rgb
    # create new XML
    new_content = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    new_content = new_content.replace(b'\n', b'\r\n')
    return new_content


def convert_pptx_to_rgb(input_file, output_file):
    """
    Read one PPT file and write conversion result to new file
    :param input_file:
    :param output_file:
    :return:
    """
    with ZipFile(input_file) as input_pptx_file:
        # theme is defined in the presentation rels
        rel = PresentationRel(input_pptx_file)
        theme_name = f'ppt/{rel.theme}'
        theme = read_color_map(zip_file=input_pptx_file, file_name=theme_name)
        color_map = theme[1]

        new_file_data = []

        file_list = input_pptx_file.namelist()

        for file in file_list:
            with input_pptx_file.open(name=file) as current_file:
                file_contents = current_file.read()
                # only convert XML files of slides
                if re.match(r'ppt/slides/slide(\d+).xml', file):
                    new_contents = convert(content=file_contents, color_map=color_map)
                else:
                    new_contents = file_contents
            new_file_data.append(
                {
                    'file': file,
                    'data': new_contents
                }
            )

    # create/write the new ZIP
    with ZipFile(output_file, 'w') as output_pptx_file:
        for data in new_file_data:
            output_pptx_file.writestr(data['file'], data['data'], ZIP_DEFLATED)


def print_usage():
    print(f'''PPT theme color to RGB color converter 0.01 Usage:
    
    Specify the .pptx filename as an input: 
    {__file__} inputfile.pptx
    ''')


if __name__ == '__main__':

    logging.basicConfig(level=logging.DEBUG)
    from sys import argv

    if len(argv) == 2:
        filename = argv[1]
        if filename.endswith('.pptx'):
            newfile = re.sub('.pptx', '_rgb.pptx', filename)
            convert_pptx_to_rgb(filename, newfile)
            print(f'Conversion Completed. New file is {newfile}')
        else:
            print_usage()
            print('Filename must end in .pptx')
    else:
        print_usage()
