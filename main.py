import win32com.client as client
from PIL import Image, ImageFont, ImageDraw
import pandas as pd
import os

# Create Absolute Path
file_path = os.path.abspath(os.path.dirname(__file__))
absolutPath = "\"" + \
              file_path.replace("\\", "\\\\") + "\\\\result.png" + "\""


def received(mats):
    text = ""
    for x in mats:
        text += x
    return text


outlook = client.Dispatch("Outlook.Application")

df = pd.read_excel(
    r"C:\Users\Arthur\Desktop\analisar t\workspace\send_leyman_email_coletivo\sisrh\Empregados.xlsx",
    sheet_name='DataBase')

df = df[['Nome', 'Data', 'Matrícula', "Unidade"]]

#
data_list = []
for i in range(len(df)):
    data_list.append({"Matrícula": df['Matrícula'][i], "Nome": df["Nome"][i], "Unidade": df["Unidade"][i]})

j = 0

mats = []
while j < len(data_list):
    mats.append(data_list[j]["Matrícula"])
    j += 1

    # t = 0
    # while t < len(data_list):

    # t += 1
    #     message.BCC = ass['Assinatura'][2]

n = 0
nam = []
while n < len(data_list):
    nam.append(data_list[n]["Nome"] + "  -  " + data_list[n]["Unidade"] + "\n\n")
    n += 1

matriculas = received(mats)
names = received(nam)
message = outlook.CreateItem(0)
message.To = matriculas
message.Subject = "Feliz Aniversário - Parabenize seus colegas! 🎉🎈🎁"
message.Display()

# Get original image
my_image = Image.open(
    "C:\\Users\\Arthur\\Desktop\\analisar t\\workspace\\send_leyman_email_coletivo\\images\\coletivo.jpg")

box = ((100, 175, 490, 400))
image_editable = ImageDraw.Draw(my_image)
# image_editable.rectangle(box)

# names = "This is some\nexample text"
font_size = 100
size = None

print(names)
while (size is None or size[0] > box[2] - box[0] or size[1] > box[3] - box[1]) and font_size > 0:
    font = ImageFont.truetype(
        'C:\\Users\\Arthur\\Desktop\\analisar t\\workspace\\send_leyman_email_coletivo\\fonts\\BebasNeue-Regular.ttf',
        14)
    size = font.getsize_multiline(names)
    font_size -= 1
    image_editable.multiline_text((box[0], box[1]), names, font=font, align="center", fill="white", stroke_fill="white",
                                  spacing=2)

# Save image final result
my_image.save("C:\\Users\\Arthur\\Desktop\\analisar t\\workspace\\send_leyman_email_coletivo\\result.png",
              optimize=True, quality=100)

html_body = f"""
           <div>
               <img src={absolutPath}>
           </div> 
           """
message.HTMLBody = html_body

# html_body = """
#     <div>
#         OI
#     </div>
#     """

