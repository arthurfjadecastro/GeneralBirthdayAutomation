import win32com.client as client
from PIL import Image, ImageFont, ImageDraw
import textwrap
import pandas as pd
import datetime


def received(mats):
    text = ""
    for x in mats:
        text += x
    return text


outlook = client.Dispatch("Outlook.Application")

df = pd.read_excel(
    r"C:\Users\Arthur\Desktop\analisar t\workspace\send_leyman_email_coletivo\sisrh\Empregados.xlsx",
    sheet_name='DataBase')

df = df[['Nome', 'Data', 'Matrícula']]

#
data_list = []
for i in range(len(df)):
    data_list.append({"Matrícula": df['Matrícula'][i], "Nome": df["Nome"][i]})

j = 0

mats = []
while j < len(data_list):
    mats.append(data_list[j]["Matrícula"])
    j += 1

    # t = 0
    # while t < len(data_list):

    # t += 1
#     message.BCC = ass['Assinatura'][2]
#     message.Subject = "Feliz Aniversário!"


n = 0
nam = []
while n < len(data_list):
    nam.append(data_list[n]["Nome"] + "\n")
    n += 1

matriculas = received(mats)
names = received(nam)
message = outlook.CreateItem(0)
message.To = matriculas
# message.Display()

# index = 0

# Get original image
my_image = Image.open(
    "C:\\Users\\Arthur\\Desktop\\analisar t\\workspace\\send_leyman_email_coletivo\\images\\coletivo.jpg")

# Specific title fonts


# Specific body fonts
# fontWeekDay = ImageFont.truetype(
#     'BebasNeue-Regular.ttf', 22)


# Get image to design


# Constants to alignment text in image
i = 0
y = 175
# Define dimension fonts
# width2, height2 = font.getsize(names)


box = ((10, 10, 490, 190))
image_editable = ImageDraw.Draw(my_image)
image_editable.rectangle(box, outline="#000")

# names = "This is some\nexample text"
font_size = 100
size = None
while (size is None or size[0] > box[2] - box[0] or size[1] > box[3] - box[1]) and font_size > 0:
    font = ImageFont.truetype(
        'C:\\Users\\Arthur\\Desktop\\analisar t\\workspace\\send_leyman_email_coletivo\\fonts\\BebasNeue-Regular.ttf',
        16)
    size = font.getsize_multiline(names)
    font_size -= 1
    image_editable.multiline_text((box[0], box[1]), names, "#000", font)
#
# lines = ""
# # Break lines text`s on image
# while i < len(names):
#     lines = textwrap.wrap(names[i], width=400)
#     y += 100
#     i = i + 1
# # for line in lines:
# #     width, height = font.getsize(line)
# image_editable.text(((200) / 3, y - 30), names, font=font, fill="white", stroke_width=1, stroke_fill="white",
#                     align="baseline")

# Save image final result
my_image.save("C:\\Users\\Arthur\\Desktop\\analisar t\\workspace\\send_leyman_email_coletivo\\images\\result.png",
              optimize=True, quality=100)

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
