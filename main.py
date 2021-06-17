from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

WIDTH, HEIGHT = letter

#want a form for each page

# if __name__ == '__main__':
#     c = canvas.Canvas("test.pdf", pagesize=letter)
#     c.drawString(100, 100, "Hello World")
#     c.showPage()  # ends page
#     c.save()

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Frame, PageTemplate
from reportlab.lib.styles import getSampleStyleSheet
styles = getSampleStyleSheet()
mbta_logo = "images/mbtalogo.png"
wsp_logo = "images/wsplogo.png"
styleN = styles["Normal"]
styleT = styles["Title"]

# formatting for the first page
def first_page_format(canvas, doc):
    canvas.saveState()

    # setup header
    canvas.drawImage(mbta_logo, .5*inch, HEIGHT-.25*inch, width=100, height=45)
    canvas.drawImage(wsp_logo, WIDTH-2*inch, HEIGHT-.25*inch, width=100, height=45)
    p = Paragraph("MBTA TUNNEL VENTILATION FACILITY AND SYSTEM ASSESSMENT", styleT)
    w, h = p.wrap(inch*5, HEIGHT)
    p.drawOn(canvas, WIDTH/2-w/2, HEIGHT-.25-h/2)

    # setup footer
    canvas.setFont('Times-Roman',9)
    canvas.drawString(inch, 0.75 * inch, "Page %d" % doc.page)
    canvas.restoreState()


# # formatting for all other pages
# def later_page_format(canvas, doc):
#     canvas.saveState()
#     canvas.setFont('Times-Roman',9)
#     canvas.drawString(inch, 0.75 * inch, "Page %d" % doc.page)
#     canvas.restoreState()


def build_document():
    doc = SimpleDocTemplate("test.pdf", pageSize=letter, title="Test", author="WSP")  # start document template
    # frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='normal')  # frame for doc margins
    # template = PageTemplate(id="test", frames=frame)
    # doc.addPageTemplates([template])

    Story = []

    for i in range(10):
        bogustext = ("This is Paragraph number %s. " % i) * 20
        p = Paragraph(bogustext, styleN)
        Story.append(p)
        Story.append(Image("images/10366.png", width=20, height=20))
        Story.append(Spacer(1,0.2*inch))
    doc.build(Story, onFirstPage=first_page_format, onLaterPages=first_page_format)


if __name__ == '__main__':
    build_document()
