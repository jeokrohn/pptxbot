import pptxutil
import xml.etree.ElementTree as ET
import re
import bs4
import itertools

PPTX = '/Users/jkrohn/OneDrive - Cisco/Documents/proj/202106 Webex Rebrand/Template/Webex_PPT_LIGHT-05.21.pptx'
PPTX = 'test.pptx'


def main():
    pptx = pptxutil.PPTXHelper(file=PPTX)
    themes_and_color_specs = pptx.themes_color_specs()
    # get a sorted list of all color spec tags from all themes
    color_tags = sorted(
        set(itertools.chain.from_iterable((c_specs.keys() for c_specs in themes_and_color_specs.values()))))

    # print head line with all theme names
    print('{:8}  {}'.format('',
                            ' '.join(('{:7}'.format(n.split('/')[-1].split('.')[0]) for n in themes_and_color_specs))))

    # now for each tag print a row with the value from each theme
    for color_tag in color_tags:
        print('{tag:8}: {spec}'.format(tag=color_tag, spec=' '.join(
            ('{:7}'.format(s.get(color_tag, '')) for s in themes_and_color_specs.values()))))

    slides = pptx.slide_file_info_list()
    slides.sort(key=lambda x: int(re.search(r'slide(\d+)\.xml', x.filename).group(1)))
    slide_info = []
    for slide in slides:
        with pptx.open(slide) as slide_file:
            slide_data = slide_file.read()
        soup = bs4.BeautifulSoup(slide_data, 'xml')
        clrMapOvr = soup.find('p:clrMapOvr')
        srgb_clr_elements = soup.find_all('a:srgbClr')
        scheme_clr_elements = soup.find_all('a:schemeClr')
        slide_info.append((slide.filename.split('/')[-1].split('.')[0],
                           {'clr_map_ovr': clrMapOvr,
                            'srgb_clr_elements': srgb_clr_elements,
                            'scheme_clr_elements': scheme_clr_elements}))

    slides_with_override = [(slide, slide_data) for slide, slide_data in slide_info
                            if slide_data['clr_map_ovr'].find('a:overrideClrMapping')]
    slides_wo_override = [(slide, slide_data) for slide, slide_data in slide_info
                          if slide_data['clr_map_ovr'].find('a:masterClrMapping')]
    srgb_parents = sorted(
        set(itertools.chain.from_iterable(((e.parent.name for e in sd['srgb_clr_elements']) for _, sd in slide_info))))
    scheme_clr_parents = sorted(set(itertools.chain.from_iterable((('{}:{}:{}'.format(e.parent.parent.parent.name,
                                                                                      e.parent.parent.name,
                                                                                      e.parent.name) for e in
                                                                    sd['scheme_clr_elements']) for _, sd in
                                                                   slide_info))))

    files = pptx.namelist()
    file_info = pptx.infolist()
    presentation_xml_rels = next((f for f in file_info if f.filename.endswith('presentation.xml.rels')))
    with pptx.open(presentation_xml_rels) as presentation_xml_rels_file:
        rel_data = presentation_xml_rels_file.read()
    soup = bs4.BeautifulSoup(rel_data, 'xml')
    relationships = [r for r in soup.Relationships.children]
    relationship_types = sorted(set((r['Type'].split('/')[-1] for r in relationships)))
    theme_relationships = [r for r in relationships if r['Type'].endswith('theme')]

    paths = set(('/'.join(f.filename.split('/')[:-1]) for f in file_info))
    themes = [f for f in file_info if f.filename.startswith('ppt/theme')]
    for theme in themes:
        with pptx.open(theme) as theme_file:
            theme_data = theme_file.read()
        tree = ET.fromstring(theme_data)
        m = re.search(r'\{(.+)\}theme', tree.tag)
        ns = {'ns': m.group(1)}
        clr_scheme = next(tree.iterfind('.//ns:clrScheme', ns))
        color_tags = [c.tag.split('}')[-1] for c in clr_scheme]

        soup = bs4.BeautifulSoup(theme_data, 'xml')

        s_clr_scheme = soup.find('a:clrScheme')
        colors = [c for c in s_clr_scheme]
        for color in soup.find('a:clrScheme'):
            c_name = color.name
            c_spec = next(color.children)
            spec_name = c_spec.name
            spec_attr, spec_val = next(((k, v) for k, v in c_spec.attrs.items()))

        colors_with_srgb_spec = (color for color in soup.find('a:clrScheme'))
        colors_with_srgb_spec = [c for c in soup.find('a:clrScheme') if next(c.children).name == 'srgbClr']
        color_specs = {c.name: next(c.children).attrs['val'] for c in colors_with_srgb_spec}

        print(soup)

    return


if __name__ == '__main__':
    main()
