from PIL import Image

image = Image.open('peginservice.jpg')
new_image = image.resize((100,100))
new_image.save('peginservice.jpg')