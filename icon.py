from PIL import Image
img = Image.open('icon.jpg')
img.save('icon.ico', format='ICO')