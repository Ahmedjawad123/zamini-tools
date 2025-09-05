from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# Create canvas and set page size
c = canvas.Canvas("image_example_reportlab.pdf", pagesize=A4)

# Set image path and position
image_path = "old_pdf_merger.png"
x = 50  # horizontal position
y = 750  # vertical position from bottom of page (A4 height is 842)

# Optional: set width and height manually (or let it scale automatically)
width = 200
height = 200

# Draw the image
c.drawImage(image_path, x, y, width, height)

# Finalize and save
c.save()


