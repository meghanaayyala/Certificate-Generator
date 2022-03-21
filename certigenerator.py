#installing necessary libraries
pip install pil
pip install pandas

#importing necessary 
from PIL import Image, ImageDraw, ImageFont
import pandas as pd

#creating a to read the excel file data into a DataFrame object.
form = pd.read_excel("FeedbackForm.xlsx")

#to_list() which returns the contents of Series object as list. Use that to convert series names into a list
name_list = form['Name'].to_list()

#for each name in the list
for i in name_list:
    #opening certificate template
    im = Image.open("CertTemplate.png")
    d = ImageDraw.Draw(im)
    
    #choosing the location to start printing the name (In the form of pixels)
    location = (185,940)
    
    #choosing dark blue in RBG format
    text_color = (0, 120, 212)
    
    #selecting the font to print the names in
    font = ImageFont.truetype("Segoe UI.ttf", 150)
    d.text(location, i, fill=text_color,font=font)
    
    #saving each file with name as filename
    im.save("certificate_"+i+".png")
