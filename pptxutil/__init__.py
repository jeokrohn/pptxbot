import zipfile
import bs4
from typing import Dict, Tuple


class PPTXHelper(zipfile.ZipFile):
    def __init__(self, file):
        zipfile.ZipFile.__init__(self, file=file)

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
