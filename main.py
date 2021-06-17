from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Frame, PageTemplate, Table
from reportlab.lib.styles import getSampleStyleSheet
from datetime import date

WIDTH, HEIGHT = letter
styles = getSampleStyleSheet()
mbta_logo = "images/mbtalogo.png"
wsp_logo = "images/wsplogo.png"
styleN = styles["Normal"]
styleT = styles["Title"]


class Equipment():
    def __init__(self, discipline, num, room, equipment_id, cs, title, descr, sol_title, sol_text, image_path):
        self.discipline = discipline
        self.num = num
        self.id = discipline + str(num)
        self.room = room
        self.equipment_id = equipment_id
        self.cs = cs
        self.title = title
        self.descr = descr
        self.sol_title = sol_title
        self.sol_text = sol_text
        self.image_path = image_path  # auto calculate?


def create_table(equip):
    image = Image(equip.image_path)
    image._restrictSize(2*inch, 3*inch)
    descr_p = Paragraph(equip.descr)
    descr_p.wrap(4.75*inch, HEIGHT) # HEIGHT?
    sol_text_p = Paragraph(equip.sol_text)
    sol_text_p.wrap(4.75 * inch, HEIGHT)  # HEIGHT?
    data =[[equip.id, "Room:", equip.room, "Equipment ID:", equip.equipment_id, "CS:", equip.cs],
           [equip.title, "", "", "", image],
           [descr_p],
           [equip.sol_title],
           [sol_text_p],
           ]
    t = Table(data, style=[('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # align top row centered
                           ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # align title and photo
                           ('ALIGN', (0, 3), (0, 3), 'CENTER'),  # align solution title centered
                           ('VALIGN', (0, 2), (0, 2), 'TOP'),
                           ('VALIGN', (0, 4), (0, 4), 'TOP'),
                           ('VALIGN', (4, 1), (-1, -1), 'CENTER'),
                           ('BOX', (0, 0), (-1, -1), 1, colors.black),
                           ('BOX', (0, 0), (-1, 0), 1, colors.black),
                           ('GRID', (0, 1), (3, -1), 1, colors.black),
                           ('SPAN', (0, 1), (3, 1)),
                           ('SPAN', (0, 2), (3, 2)),
                           ('SPAN', (0, 3), (3, 3)),
                           ('SPAN', (0, 4), (3, 4)),
                           ('SPAN', (-3, 1), (-1, -1)),
                           ],
              colWidths=[.25*inch, .5*inch, 3*inch, inch, 1.75*inch, .25*inch, .25*inch],
              rowHeights=[.25*inch, .25*inch, 1.25*inch, .25*inch, 1.25*inch])
    return t

# formatting for the first page
def first_page_format(canvas, doc):
    canvas.saveState()

    # setup header
    canvas.drawImage(mbta_logo, .25*inch, HEIGHT-.25*inch, width=100, height=50)
    canvas.drawImage(wsp_logo, WIDTH-2*inch, HEIGHT-.25*inch, width=100, height=50)
    p = Paragraph("MBTA TUNNEL VENTILATION FACILITY AND SYSTEM ASSESSMENT", styleT)
    w, h = p.wrap(inch*5, HEIGHT)
    p.drawOn(canvas, WIDTH/2-w/2, HEIGHT-.25-h/2)

    # setup footer
    canvas.setFont('Times-Roman',9)
    canvas.drawString(inch, .75*inch, "%s" % date.today().strftime("%m/%d/%Y"))
    canvas.drawCentredString(WIDTH/2, .75*inch, "Page %d" % doc.page)
    canvas.drawRightString(WIDTH - inch, .90*inch, "Contract No. Z94PS10")
    canvas.drawRightString(WIDTH - inch, .75*inch, "Task Order: 03")
    canvas.restoreState()


# # formatting for all other pages
# def later_page_format(canvas, doc):
#     canvas.saveState()
#     canvas.setFont('Times-Roman',9)
#     canvas.drawString(inch, 0.75 * inch, "Page %d" % doc.page)
#     canvas.restoreState()


def build_document():
    doc = SimpleDocTemplate("MBTA Tunnel Vent and System Assessment.pdf", pageSize=letter,
                            title="MBTA Tunnel Vent and System Assessment", author="WSP")  # start document template
    # frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='normal')  # frame for doc margins
    # template = PageTemplate(id="test", frames=frame)
    # doc.addPageTemplates([template])

    Story = []

    for i in range(10):
        e = Equipment("E", 1, "EF-2 EXHAUST PLENUM", "PUSHBUTTON", 3, "CORROSION",
                      "PUSHBUTTON IS CORRODED, DEVICE IS STILL OPERATIONAL, BUT SHOULD BE MONITORED.",
                      "MONITOR", "MONITOR DEVICE OVER COMING YEARS", "images/pushbutton.png")
        t = create_table(e)
        Story.append(t)
        # Story.append(Image("images/10366.png", width=20, height=20))
        Story.append(Spacer(1,0.2*inch))
    doc.build(Story, onFirstPage=first_page_format, onLaterPages=first_page_format)


if __name__ == '__main__':
    build_document()
