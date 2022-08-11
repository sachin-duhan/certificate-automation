from PIL import Image, ImageDraw, ImageFont

import pandas as pd

# specify the location of input excel file.
form = pd.read_excel(r"/home/duhan/Desktop/Python_Project/CetificateScript/Copy of WEBINARS (Responses).xlsx")
FONT_SIZE = 270
# Excel file should contain a column 'Name'
name_list = form['Name'].to_list()

def getXCord(name):
    OFFSET_FACTOR = 50
    BASE_X_VAL = 2800
    name = str(name)
    retVal = BASE_X_VAL - (OFFSET_FACTOR * len(name)-1)
    return retVal

current_index = 0
for i in name_list:
    # input image on which the text needs to be added.
    im = Image.open(r"/home/duhan/Desktop/Python_Project/CetificateScript/Webinar Internship.jpg")
    d = ImageDraw.Draw(im)
    # we define the X,Y position of the text on the image
    location = (getXCord(i), 1900)
    text_color = (8, 56, 96)
    # first argument is font family and second family is font size that we want
    font = ImageFont.truetype(r"./TIMES.ttf", FONT_SIZE)
    d.text(location, i, fill=text_color, font=font)
    current_index += 1
    # choose a name for output pdf file.
    im.save("certificate_"+str(current_index)+".pdf")