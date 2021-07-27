# -*- coding: utf-8 -*-
'''
Created : 2021-07-27

@author: Stan Knutson, stan@agmonitor.com
'''

from docxtpl import DocxTemplate, InlineImage

# for height and width you have to use millimeters (Mm), inches or points(Pt) class :
from docx.shared import Mm, Inches
import jinja2



DETAIL_TEMPLATE = 'templates/subdoc_detail_page.docx'
OUTER_TEMPLATE = 'templates/subdoc_multipage_tpl.docx'

IMAGE_1 = 'templates/django.png'
IMAGE_2 = 'templates/zope.png'

# Note: the actual report generation uses matplotlib to make the images
specs = [
    {"case_name": "django",
     "image_file": IMAGE_1},
    {"pump_name": "Zope",
     "image_file": IMAGE_2},
]

OUTFILE = "output/multipage_report.docx"

def gen_report():
    detail_pages = []

    doc = DocxTemplate(OUTER_TEMPLATE)

    # build all of the inner subpages
    for ix, spec in enumerate(specs):
        detail_doc = DocxTemplate(DETAIL_TEMPLATE)

        # BUG:  Older version wanted the image to be part of OUTER template
        #  but that results in error (not rendering)
        #  /opt/agmonitor/lib/python3.9/site-packages/docxtpl/__init__.py in _insert_image(self)
        #     835
        #     836     def _insert_image(self):
        # --> 837         pic = self.tpl.current_rendering_part.new_pic_inline(
        #     838             self.image_descriptor,
        #     839             self.width,
        #
        # AttributeError: 'NoneType' object has no attribute 'new_pic_inline'

        # ALTERNATE: add to subdocument, but then the image is not actually included

        spec["case_image"] = InlineImage(detail_doc, spec["image_file"], width=Inches(6.5))

        detail_doc.render(context=spec)
        detail_subdoc = doc.new_subdoc()
        detail_subdoc.subdocx = detail_doc.docx
        detail_pages.append(detail_subdoc)

    # outer page context
    context = {"customer": "sample",
               "detail_pages": detail_pages}
    doc.render(context)
    doc.save(OUTFILE)
    print("Wrote", OUTFILE)

if __name__ == "__main__":
    gen_report()
