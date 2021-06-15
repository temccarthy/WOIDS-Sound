from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

WIDTH, HEIGHT = letter

if __name__ == '__main__':
    c = canvas.Canvas("test.pdf", pagesize=letter)
    c.drawString(100, 100, "Hello World")
    c.showPage()
    c.save()

