import pytesseract as tess
import pandas as pd
from PIL import Image
import os

tess.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

path1, path2 = './Images', './Data'
listing = os.listdir(path1)

print("👉" * 10, "WELCOME", "👈" * 10)
print("🔃 Processing large no of images Might take more Time \n")

final = list()
FILE = list()

for file in listing:
    img = Image.open(path1 + "/" + file)
    section2 = img.crop((0, 0, 854, 310))
    section3 = img.crop((840, 200, 1180, 400))
    section4 = img.crop((1060, 65, 1600, 275))

    text = str(tess.image_to_string(section2))
    text2 = str(tess.image_to_string(section4))

    text = text.replace("\n", "")
    text2 = text2.replace("\n", "")

    if "Z" in text: text = text.replace("Z", "")
    if "=" in text2: text2 = text2.replace("=", "")
    if "-" not in text2: text2 = "-" + text2
    if "O" in text2: text2 = text2.replace("O", "")
    if "R" in text2: text2 = text2.replace("R", "8")
    if "e" in text2: text2 = text2.replace("e", "")

    t = text + "E" + text2
    final.append(t)
    FILE.append(file)

newDataframe = pd.DataFrame()

newDataframe["File Name"] = FILE
newDataframe["Scanned"] = final


newDataframe.to_excel('output.xlsx', index=False)
print(pd.read_excel('output.xlsx'))
print("DONE 😊")
