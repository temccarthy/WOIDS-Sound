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

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
styles = getSampleStyleSheet()
mbta_logo = Image("images/mbtalogo.png")
wsp_logo = Image("images/wsplogo.png")


# formatting for the first page
def first_page_format(canvas, doc):
    canvas.saveState()
    canvas.setFont('Times-Roman',9)
    canvas.drawString(inch, 0.75 * inch, "Page %d" % doc.page)
    canvas.restoreState()


# formatting for all other pages
def later_page_format(canvas, doc):
    canvas.saveState()
    canvas.setFont('Times-Roman',9)
    canvas.drawString(inch, 0.75 * inch, "Page %d" % doc.page)
    canvas.restoreState()


def build_document():
    doc = SimpleDocTemplate("test.pdf", pageSize=letter)
    Story = []
    style = styles["Normal"]
    for i in range(100):
        bogustext = ("This is Paragraph number %s. " % i) * 20
        p = Paragraph(bogustext, style)
        Story.append(p)
        Story.append(Image("images/10366.png", width=20, height=20))
        Story.append(Spacer(1,0.2*inch))
    doc.build(Story, onFirstPage=first_page_format, onLaterPages=later_page_format)


if __name__ == '__main__':
    build_document()
