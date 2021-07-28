# -*- coding: utf-8 -*-
'''
Created : 2021-07-27

This file shows the WORKAROUND for generating a multi-page report

It uses the extra library docxcomposer, which should not be required

There is one restriction: the template for "first page" must not want to have any content AFTER the detail pages.
(or modify the composer call to generate and include a "trailer page")

@author: Stan Knutson, stan@agmonitor.com
'''

from docxtpl import DocxTemplate, InlineImage

# for height and width you have to use millimeters (Mm), inches or points(Pt) class :
from docx.shared import Mm, Inches
import jinja2



DETAIL_TEMPLATE = 'templates/subdoc_detail_page.docx'
OUTER_TEMPLATE = 'templates/subdoc_multipage_workaround.docx'

IMAGE_1 = 'templates/django.png'
IMAGE_2 = 'templates/zope.png'

# Note: the actual report generation uses matplotlib to make the images
specs = [
    {"case_name": "django",
     "location": "Fresno CA",
     "image_file": IMAGE_1},
    {"case_name": "Zope",
     "location": "Madera CA",
     "image_file": IMAGE_2},
]

TEMPDIR = "output/"

OUTFILE = "output/multipage_report_workaround.docx"

def gen_report():
    detail_pages = []

    # our report has one "front page"
    # and many "detail pages

    doc = DocxTemplate(OUTER_TEMPLATE)
    # generate the front page
    context = {"customer_name": "sample",
               #"detail_pages": detail_pages
               "case_specs": specs
               }
    doc.render(context)
    header_file = TEMPDIR + "/temp_multipage_header.docx"
    doc.save(header_file)

    files_to_merge = []

    # build all of the detail pages
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

        #spec["case_image"] = "file " + spec["image_file"]
        ctx = dict(spec)
        ctx["case_image"] = InlineImage(detail_doc, spec["image_file"], width=Inches(6.5))

        detail_doc.render(context=ctx)
        temp_file = TEMPDIR + f"temp_doc_{ix+1}.docx"
        detail_doc.save(temp_file)

        files_to_merge.append(temp_file)

    from docxcompose.composer import Composer
    from docx import Document
    master = Document(header_file)
    composer = Composer(master)
    for fn in files_to_merge:
        doc1 = Document(fn)
        composer.append(doc1)
    composer.save(OUTFILE)
    print("Wrote", OUTFILE)

if __name__ == "__main__":
    gen_report()
