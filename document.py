import glob
from PIL import ExifTags
from reportlab.lib import colors, utils
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from datetime import date
from spreadsheet import Equipment
import os
import sys
import PIL.Image


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


# setup variables
WIDTH, HEIGHT = letter
styles = getSampleStyleSheet()
seattle_logo = resource_path("resources/web-st-logo-horizontal-blue-rgb.png")
wsp_logo = resource_path("resources/img-png-wsp-red.png")
styleN = styles["Normal"]
styleT = ParagraphStyle(
    'newTitle',
    parent=styles['Title'],
    fontSize=14
)
styleS = ParagraphStyle(
    'newSubTitle',
    parent=styles['Normal'],
    fontSize=11,
    alignment=1
)
document_name = "DSTT NTIS Inspection Report.pdf"
# cs_colors = [colors.lightgreen, colors.yellow, colors.orange, colors.pink]

inspection_date = ""

for orientation in ExifTags.TAGS.keys():
    if ExifTags.TAGS[orientation] == 'Orientation':
        break


# for rotating portrait photos
class RotatedImage(Image):
    def draw(self):
        self.canv.rotate(-90)
        self.canv.translate(- self.drawWidth/2 - self.drawHeight/2, self.drawWidth/2-self.drawHeight/2)
        Image.draw(self)


deficiency_ctr = 1
def create_equipment_table(equip):
    global deficiency_ctr
    path = equip.image_path
    temp_path = path[:path.rfind("\\", 0, -1)] + "/temp/" + path[path.rfind("\\", 0, -1):]

    # fix rotated images
    with PIL.Image.open(path) as img:
        exif = img._getexif()
    if exif[orientation] == 6:
        image = RotatedImage(temp_path, width=2.5*inch, height=3*inch, kind="proportional")
    else:
        image = Image(temp_path, width=2.25*inch, height=2.5*inch, kind="proportional")

    # information paragraphs
    notes = Paragraph(equip.notes)
    notes.wrap(4.75 * inch, HEIGHT)
    # sol_text_p = Paragraph(equip.sol_text)
    # sol_text_p.wrap(4.75 * inch, HEIGHT)
    data = [
        [deficiency_ctr, Paragraph('<b>Location:</b> %s' % equip.room), Paragraph('<b>Station:</b> %s' % equip.station),
         Paragraph('<b>Component:</b> %s' % equip.component)],
        [Paragraph('<b>Deficiency/Notes: </b>'), "", "", image],
        [notes],
        ]
    t = Table(data, style=[('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # align top row centered
                           ('ALIGN', (3, 1), (-1, 1), 'CENTER'),  # align title and photo
                           # ('ALIGN', (0, 3), (0, 3), 'CENTER'),  # align solution title centered
                           ('VALIGN', (0, 2), (0, 2), 'TOP'),
                           # ('VALIGN', (0, 4), (0, 4), 'TOP'),
                           ('VALIGN', (-1, 1), (-1, -1), 'CENTER'),
                           ('BOX', (0, 0), (-1, -1), 1, colors.black),
                           ('BOX', (0, 0), (-1, 0), 1, colors.black),
                           ('GRID', (0, 0), (2, -1), 1, colors.black),
                           ('SPAN', (0, 1), (-2, 1)),  # span title
                           ('SPAN', (0, 2), (-2, 2)),  # span descr
                           # ('SPAN', (0, 3), (3, 3)),  # span sol title
                           # ('SPAN', (0, 4), (3, 4)),  # span sol descr
                           ('SPAN', (-1, 1), (-1, -1)),  # span picture box
                           # ('BACKGROUND', (-1, 0), (-1, 0), cs_colors[equip.cs-1])
                           ],
              colWidths=[.4 * inch, 1.25 * inch, 2.75 * inch, 2.75 * inch],
              rowHeights=[.25 * inch, .25 * inch, 2.5 * inch])

    deficiency_ctr += 1
    return t


# creates report table for top of document
def create_report_table(loc):
    data = [
        [Paragraph('<b>LOCATION:</b> %s' % loc.loc),
         Paragraph('<b>INSPECTION DATE:</b> %s' % loc.insp_date)],
        [Paragraph('<b>STATION:</b> %s' % loc.station)],
    ]
    t = Table(data)
    return t


# formatting for the first page
def first_page_format(canvas, doc):
    global inspection_date
    canvas.saveState()

    # setup header
    canvas.drawImage(seattle_logo, .45 * inch, HEIGHT - .05 * inch, width=100, height=15)
    canvas.drawImage(wsp_logo, WIDTH - 1.65 * inch, HEIGHT - .2 * inch, width=70, height=33)
    p = Paragraph("DSTT NTIS INSPECTION REPORT", styleT)
    w, h = p.wrap(inch * 4, HEIGHT)
    p.drawOn(canvas, (WIDTH - 16) / 2 - w / 2, HEIGHT - .25 - h / 2)

    # setup subheader
    p2 = Paragraph('<b>INSPECTION DATE:</b> %s' % inspection_date, styleS)
    w2, h2 = p2.wrap(inch * 4, HEIGHT)
    p2.drawOn(canvas, (WIDTH - 16) / 2 - w2 / 2, HEIGHT - 20)

    # setup footer
    canvas.setFont('Times-Roman', 9)
    canvas.drawString(.75 * inch, .75 * inch, "August 27, 2021")
    canvas.drawCentredString((WIDTH - 16) / 2, .75 * inch, "Page %d" % doc.page)
    canvas.drawRightString(WIDTH - inch, .90 * inch, "Contract No. 160415P-01")
    canvas.drawRightString(WIDTH - inch, .75 * inch, "Task Order: 17")
    canvas.restoreState()


# assembles the document and saves it to the sheet's location
def build_document(sheet):
    global inspection_date
    doc = SimpleDocTemplate(os.path.join(sheet.folder, document_name), pageSize=letter,
                            title="DSTT NTIS Inspection Report", author="WSP")  # start document template
    Story = []

    inspection_date = sheet.location.insp_date
    # t = create_report_table(sheet.location)
    # Story.append(t)  # add location information
    # Story.append(Spacer(1, 0.2 * inch))

    # compress images into temp folder
    sheet.compress_pictures()

    for row in sheet.fp.parse(1).itertuples():  # for row in spreadsheet
        e = Equipment.generate_equip(sheet.folder, row)
        t = create_equipment_table(e)

        Story.append(KeepTogether(t))
        Story.append(Spacer(1, 0.1 * inch))

    doc.build(Story, onFirstPage=first_page_format, onLaterPages=first_page_format)

    # delete compressed image temp folder
    sheet.delete_compressed_pictures()


# checks if document already exists in sheet's folder
def check_doc_exists(sheet):
    matches = glob.glob(os.path.join(sheet.folder, document_name))
    return len(matches) != 0
