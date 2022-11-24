from PIL import Image
from PyPDF2 import PdfFileMerger

image = Image.open('peginservice.jpg')
new_image = image.resize((100,100))
new_image.save('peginservice.jpg')
merger = PdfFileMerger()
print(merger)