import win32com.client as client
from PIL import Image, ImageFont, ImageDraw
import pandas as pd
import os
import math
from datetime import date


def convertText(text):
    # print(text[0][1])
    DIAS = [
        'Segunda-feira',
        'Terça-feira',
        'Quarta-feira',
        'Quinta-Feira',
        'Sexta-feira',
        'Sábado',
        'Domingo'
    ]
    concat1 = text[0] + text[1]
    concat2 = text[3] + text[4]

    year = date.today().strftime('%Y')
    data = date(year=int(year), month=int(concat2), day=int(concat1))
    indice_da_semana = data.weekday()
    dia_da_semana = DIAS[indice_da_semana]
    return dia_da_semana


# Create Absolute Path
file_path = os.path.abspath(os.path.dirname(__file__))

htmlAbsoluteImagePath = "\"" + file_path.replace("\\", "\\\\") + "\\\\images\\\\result.png\""

sisrhPath = file_path + "\\sisrh\\Busca_SISRH_SR2637.xlsm"

coletivoJPGPath = file_path + "\\images\\coletivo.jpg"

fontBebasNeueTTFPath = file_path + "\\fonts\\BebasNeue-Regular.ttf"

finalImagePath = file_path + "\\images\\result.png"


def received(matriculas):
    text = ""
    for x in matriculas:
        text += x
    return text


outlook = client.Dispatch("Outlook.Application")


df = pd.read_excel(sisrhPath, sheet_name='Dados')

df = df[['Name', 'Data', 'Matrícula', "Unidade"]]

#
data_list = []
for i in range(len(df)):
    data_list.append(
        {"Matrícula": df['Matrícula'][i], "Name": df["Name"][i], "Unidade": df["Unidade"][i], "Data": df['Data'][i]})

j = 0

matriculas = []

while j < len(data_list):
    matriculas.append(data_list[j]["Matrícula"])
    j += 1

half_length = math.ceil(len(matriculas) / 2)
first_half = matriculas[:half_length]
sec_half = matriculas[half_length:]

result = {}
foo = []
for n, g in df.groupby("Data"):
    foo.append(convertText(str(n)) + "\n")
    for x in g.values:
        if (n == x[1]):
            foo.append(x[0] + "  -  " + x[3] + "\n\n")

    if n in result:
        result[n] += g.values.tolist()  # ...se sim, concatena a lista em result com a lista obtida do grupo.
    else:
        result[
            n] = g.values.tolist()  # ...se não, cria a chave em result e adiciona a lista obtida do grupo como valor.

foo2 = received(foo)

matriculas1 = received(first_half)
matriculas2 = received(sec_half)

matriculas = []
matriculas.append(matriculas1)
matriculas.append(matriculas2)

i = 0
while i < 2:
    message = outlook.CreateItem(0)
    message.BCC = matriculas[i]
    message.Subject = "Parabenize seus colegas - Feliz Aniversário 🎉🎈🎁"
    message.Display()
    # Get original image
    my_image = Image.open(coletivoJPGPath)
    box = ((100, 175, 490, 400))
    image_editable = ImageDraw.Draw(my_image)

    font_size = 100
    size = None
    while (size is None or size[0] > box[2] - box[0] or size[1] > box[3] - box[1]) and font_size > 0:
        font = ImageFont.truetype(fontBebasNeueTTFPath, 14)
        size = font.getsize_multiline(str(foo2))
        font_size -= 1
        image_editable.multiline_text((box[0], box[1]), str(foo2), font=font, align="center", fill="white",
                                      stroke_fill="white")

    # Save image final result
    my_image.save(finalImagePath, optimize=True, quality=100)

    html_body = f"""
               <div>
                   <img src={htmlAbsoluteImagePath}>
               </div> 
               """
    message.HTMLBody = html_body
    i += 1
