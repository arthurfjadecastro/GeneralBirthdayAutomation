import win32com.client as client
from PIL import Image, ImageFont, ImageDraw
import textwrap
import pandas as pd
import datetime

outlook = client.Dispatch("Outlook.Application")

df = pd.read_excel(
    r"C:\Users\Arthur\Desktop\analisar t\workspace\send_leyman_email_coletivo\sisrh\Empregados.xlsx",
    sheet_name='DataBase')

df = df[['Nome', 'Data', 'Matrícula']]

#
data_list = []
for i in range(len(df)):
    data_list.append({"Matrícula": df['Matrícula'][i]})

# print(data_list)
# index = 0
# my_image = Image.open(
#     "images/coletivo.jpg")
#
# title_text = []
# font = ImageFont.truetype(
#     'BebasNeue-Regular.ttf', 16)
#
# fontWeekDay = ImageFont.truetype(
#     'BebasNeue-Regular.ttf', 22)
#
# sortedList = sorted(data_list, key=lambda x: x['dayOfTheWeek'])
#
# title_text = ExistingDayOfTheWeek(sortedList)

# fontAss = ImageFont.truetype(
#     'BebasNeue-Regular.ttf', 18)
#
# image_editable = ImageDraw.Draw(my_image)
#
# i = 0
# y = 175
# width2, height2 = font.getsize(dayWeekText[0])


# image_editable.text(((800 - width2) / 3, y-30),
#                     dayWeekText[0], font=fontWeekDay, fill="white", stroke_width=1, stroke_fill="white", align="baseline")
# while i < len(title_text):
#
#     lines = textwrap.wrap(title_text[i], width=400)
#     for line in lines:
#         width, height = font.getsize(line)
#         print(len(title_text))
#         image_editable.text(((800 - width) / 3, y),
#                             line.title(), font=font, fill="white", stroke_width=0, stroke_fill="white",
#                             align="baseline")
#         print(((height * 1.5) / (len(title_text))))
#         y += height * \
#              2 if len(title_text) < 6 else (
#             ((height * 19) / (len(title_text))))
#     i = i + 1

# my_image.save("result.png", optimize=True, quality=100)
j = 0
message = outlook.CreateItem(0)
mats = []
while j < len(data_list):
    mats.append(data_list[j]["Matrícula"])
    j += 1

    # t = 0
    # while t < len(data_list):

    # t += 1
#     message.BCC = ass['Assinatura'][2]
#     message.Subject = "Feliz Aniversário!"

def received(mats):
    text = ""
    for x in mats:
        text += x
    return text

a = received(mats)

message.To = a
message.Display()

# html_body = """
#     <div>
#         <img src="result.png">
#     </div>
#     """

html_body = """
    <div>
        OI
    </div>
    """
message.HTMLBody = html_body
