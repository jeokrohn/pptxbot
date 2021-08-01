import logging
import re
import zipfile
from collections import Mapping
from collections import defaultdict
from dataclasses import dataclass, field, fields
from typing import Dict, Tuple, List, get_origin

import bs4
import inflection
import lxml.etree
from lxml import etree

log = logging.getLogger(__name__)


@dataclass
class ColorSpec:
    name: str
    color_map: Dict[str, str]

    @classmethod
    def from_file(cls, zip_file: zipfile.ZipFile, file_name: str):
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
        result = cls(name=name, color_map=rgb_map)
        return result


@dataclass
class Rel:

    @classmethod
    def from_tree(cls, tree: lxml.etree.ElementTree) -> 'Rel':
        root = tree.getroot()
        data = defaultdict(list)
        for child in root:
            target: str = child.attrib['Target']
            if target.startswith('../'):
                target = target[3:]
            var = child.attrib['Type'].split('/')[-1]
            data[var].append(target)
        init = {}
        cls_fields = {f.name: f for f in fields(cls)}

        undefined = []
        for k, v in sorted((k, v) for k, v in data.items()):
            k = inflection.underscore(k)
            cls_fld = cls_fields.get(k)
            if cls_fld is None:
                undefined.append(k)
            else:
                # if the class field is not a list then convert from list to element
                if get_origin(cls_fld.type) is None:
                    v = v[0]
                else:
                    v.sort()
                init[k] = v
        result = cls(**init)
        if undefined:
            log.warning(f'!!! {cls.__name__}: rel file has unexpected rel type(s): {", ".join(undefined)} !!!')
        return result

    @classmethod
    def from_file(cls, zip_file: zipfile.ZipFile, file_name: str) -> 'Rel':
        with zip_file.open(name=file_name) as rel_file:
            tree = etree.parse(rel_file)
        return cls.from_tree(tree=tree)

    @classmethod
    def from_zip(cls, zip_file: zipfile.ZipFile) -> 'Rel':
        return cls.from_file(zip_file=zip_file, file_name=cls.FILE)


@dataclass
class PresentationRel(Rel):
    FILE = 'ppt/_rels/presentation.xml.rels'

    notes_master: str
    pres_props: str
    slide: List[str]
    slide_master: List[str]
    table_styles: str
    theme: str
    view_props: str

    def __post_init__(self):
        slide_re = re.compile(r'.+slide(\d+).xml')
        self.slide.sort(key=lambda s: int(slide_re.match(s).group(1)))

    @classmethod
    def from_zip(cls, zip_file: zipfile.ZipFile) -> 'PresentationRel':
        return super().from_zip(zip_file=zip_file)


@dataclass
class SlideRel(Rel):
    slide_layout: str
    notes_slide: str = None
    image: List[str] = field(default_factory=list)
    hyperlink: List[str] = field(default_factory=list)

    @classmethod
    def from_file(cls, zip_file: zipfile.ZipFile, file_name: str) -> 'SlideRel':
        return super().from_file(zip_file=zip_file, file_name=file_name)


@dataclass
class SlideLayoutRel(Rel):
    slide_master: str
    image: List[str] = field(default_factory=list)

    @classmethod
    def from_file(cls, zip_file: zipfile.ZipFile, file_name: str) -> 'SlideLayoutRel':
        return super().from_file(zip_file=zip_file, file_name=file_name)


@dataclass
class TreeAndRel:
    tree: lxml.etree.ElementTree
    _rel: Rel

    REL_TYPE = None

    @classmethod
    def from_path(cls, zip_file: zipfile.ZipFile, path: str) -> 'TreeAndRel':
        with zip_file.open(name=path) as file:
            tree = etree.parse(file)
        s_path = path.split('/')
        reL_path = '/'.join(s_path[:-1] + ['_rels', f'{s_path[-1]}.rels'])
        rel = cls.REL_TYPE.from_file(zip_file=zip_file, file_name=reL_path)

        return cls(tree=tree, _rel=rel)


class Slide(TreeAndRel):
    REL_TYPE = SlideRel

    @classmethod
    def from_path(cls, zip_file: zipfile.ZipFile, path: str) -> 'Slide':
        return super().from_path(zip_file, path)

    @property
    def rel(self) -> SlideRel:
        return self._rel


class SlideLayout(TreeAndRel):
    REL_TYPE = SlideLayoutRel

    @classmethod
    def from_path(cls, zip_file: zipfile.ZipFile, path: str) -> 'SlideLayout':
        return super().from_path(zip_file, path)

    @property
    def rel(self) -> SlideLayoutRel:
        return self._rel


class PPTProxy(Mapping[str, TreeAndRel]):
    PREFIX = None
    TYPE = None

    def __init__(self, zip_file: zipfile.ZipFile):
        self._zip_file = zip_file
        name_re = re.compile(f'(?:ppt/)?({self.PREFIX}s/{self.PREFIX}(\d+)\.xml)')
        files = [m.group(1) for f in self._zip_file.namelist()
                 if (m := name_re.match(f))]
        files.sort(key=lambda f: int(name_re.match(f).group(2)))
        self._files = files

    def __iter__(self):
        return iter(self._files)

    def __getitem__(self, item):
        item_match = re.match(f'.*({self.PREFIX}\d+\.xml)', item)
        if item_match is None:
            raise KeyError(f'invalid key "{item}"')
        name = item_match.group(1)
        path = f'ppt/{self.PREFIX}s/{name}'
        result = self.TYPE.from_path(zip_file=self._zip_file, path=path)
        return result

    def __len__(self):
        return len(self._files)


class SlideProxy(PPTProxy, Mapping[str, Slide]):
    PREFIX = 'slide'
    TYPE = Slide

    def __getitem__(self, item) -> Slide:
        return super().__getitem__(item=item)


class SlideLayoutProxy(PPTProxy, Mapping[str, SlideLayout]):
    PREFIX = 'slideLayout'
    TYPE = SlideLayout

    def __getitem__(self, item) -> SlideLayout:
        return super().__getitem__(item=item)


class PPTXHelper(zipfile.ZipFile):
    def __init__(self, file):
        zipfile.ZipFile.__init__(self, file=file)
        self.slides = SlideProxy(zip_file=self)
        self.slide_layouts = SlideLayoutProxy(zip_file=self)

    def slide_file_info_list(self):
        """
        get list of file info objects of all slides
        :return:
        """
        return [f for f in self.infolist() if f.filename.startswith('ppt/slides') and f.filename.endswith('.xml')]

    def theme_file_info_list(self):
        """
        get list of file info objects for all theme files included in the PPTX
        :return:
        """
        return [f for f in self.infolist() if f.filename.startswith('ppt/theme')]

    def theme_color_specs(self, theme) -> Tuple[Dict[str, str]]:
        with self.open(theme) as theme_file:
            theme_data = theme_file.read()

        soup = bs4.BeautifulSoup(theme_data, 'xml')
        r = {}
        scheme = soup.find('a:clrScheme')
        scheme_name = scheme.attrs['name']
        for color in scheme:
            c_name = color.name
            c_spec = next(color.children)
            spec_name = c_spec.name
            if spec_name == 'srgbClr':
                r[c_name] = c_spec.attrs['val']
            elif spec_name == 'sysClr':
                r[c_name] = '{}s'.format(c_spec.attrs['lastClr'])
            else:
                raise
            # if
        # for
        return scheme_name, r

    def themes_color_specs(self) -> Dict[str, Tuple[Dict[str, str]]]:
        return {t.filename: self.theme_color_specs(t) for t in self.theme_file_info_list()}

    @property
    def color_spec(self) -> ColorSpec:
        theme_file = f'ppt/{self.presentation_rel.theme}'
        return ColorSpec.from_file(zip_file=self, file_name=theme_file)

    @property
    def presentation_rel(self) -> PresentationRel:
        return PresentationRel.from_zip(self)

    def slide_rel(self, slide_name: str) -> SlideRel:
        slide_name = re.match(r'.*(slide\d+.xml)', slide_name).group(1)
        rel_name = f'ppt/slides/_rels/{slide_name}.rels'
        return SlideRel.from_file(zip_file=self, file_name=rel_name)
