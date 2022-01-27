from random import randint
import win32com.client as client
from PIL import Image, ImageFont, ImageDraw
import textwrap
import pandas as pd
import datetime


codeAndName = {3: 'AG AEROPORTO PRESIDENTE JK, DF',
               4: 'AG BERNARDO SAYAO, DF',
               6: 'AG MEXICO, DF',
               8: 'AG TAGUATINGA, DF',
               10:	'AG 105 SUDOESTE, DF',
               138:	'AG PARACATU, MG',
               609:	'AG PIRAPORA, MG',
               643:	'AG GUARA, DF',
               647:	'AG CAPITAL, DF',
               655: 'AG GAMA, DF',
               674:	'AG LAGO SUL, DF',
               688:	'AG NUCLEO BANDEIRANTE, DF',
               804:	'AG LUZIANIA, GO',
               816:	'AG 210 SUL, DF',
               942:	'AG UNAI, MG',
               1040: 'AG BRAZLANDIA, DF',
               1041:	'AG COMERCIAL SUL, DF',
               1057:	'AG 515 SUL, DF',
               1088:	'AG AGUAS LINDAS SHOPPING, GO',
               1502:	'AG SUDOESTE, DF',
               1511:	'AG SIG, DF',
               1556:	'AG CNB 12, DF',
               1803:	'AG ASA SUL, DF',
               1818:	'AG JOAO PINHEIRO, MG',
               1899:	'AG JARDIM INGA, GO',
               1985:	'AG AV RECANTO, DF',
               1990:	'AG FELICITTA SHOPPING, DF',
               2272:	'AG CEILANDIA, DF',
               2304:	'AG CASAPARK SHOPPING, DF',
               2399:	'AG TAGUASUL, DF',
               2407:	'AG SIA, DF',
               2437:	'AG VALPARAISO, GO',
               2889:	'AG VARZEA DA PALMA, MG',
               2893:	'AG LUCIO COSTA, DF',
               3001:	'AG CIDADE DE SANTA MARIA, DF',
               3002:	'AG GUARA II, DF',
               3035:	'AG RIACHO FUNDO, DF',
               3052:	'AG AGUAS LINDAS, GO',
               3129:	'PA CONAB, DF',
               3189:	'AG NOVO GAMA, GO',
               3247:	'PA SPO, DF',
               3369:	'AG CRISTALINA, GO',
               3494:	'AG AGUAS CLARAS, DF',
               3625:	'AG 102 SUDOESTE, DF',
               3697:	'AG VAZANTE, MG',
               3872:	'AG AV HELIO PRATES, DF',
               4166:	'AG CEILANDIA NORTE, DF',
               4167:	'AG SAMAMBAIA, DF',
               4221:	'AG PADRE BERNARDO, GO',
               4222:	'AG CIDADE OCIDENTAL, GO',
               4223:	'AG STO ATO DO DESCOBERTO, GO',
               4331:	'AG RECANTO DAS EMAS, DF',
               4461:	'AG GAMA LESTE, GO',
               4462:	'AG COMERCIAL NORTE, DF',
               4463:	'AG PISTAO SUL, DF',
               4482:	'AG 310 SUL, DF',
               4483:	'AG SAMAMBAIA SUL, DF',
               4501:	'AG VALPARAISO CENTRO, GO',
               4760:	'AG PARQUE CIDADE, DF',
               4979:	'AG BURITIS DE MINAS, MG',
               5725: 'SEV GUARÁ',
               6600: 'SEV CEILÂNDIA',
               5079: 'SEV TAGUATINGA',
               5038: 'SEV PLANO PILOTO',
               5295: 'SEV PARACATU',
               5731: 'SEV GAMA',
               2637: 'SR Brasília Sul',
               }


da = datetime.date.today().strftime("%d/%m/%Y")
outlook = client.Dispatch("Outlook.Application")


phrase = ['A Felicidade merece ser compartilhada']

df = pd.read_excel(
    r"C:\Users\Ravin\Desktop\send_email_coletivo\Empregados.xlsx", sheet_name='DataBase')

ass = pd.read_excel(
    r"C:\Users\Ravin\Desktop\send_email_coletivo\Empregados.xlsx", sheet_name='Assinatura')


srName = ass['Assinatura'][0]
srEntity = ass['Assinatura'][1]
office = ass['Assinatura'][3]

textAss = srName + '\n' + office + '\n' + srEntity

# df["dte_Nascimento_Empregado"] = pd.to_datetime(df["dte_Nascimento_Empregado"])

df = df[['Str_Mat_Outlook', 'str_Nome_Empregado',
         'int_CodLotacao_Empregado', 'dte_Nascimento_Empregado']]

dataFullList = [{}]
data_list = [{}]
for i in range(len(df)):
    if df['dte_Nascimento_Empregado'][i].strftime("%m-%d") == datetime.date.today().strftime('%m-%d'):
        data_list.append({"birthDate": df['dte_Nascimento_Empregado'][i].strftime(
            "%m-%d"), "name": df['str_Nome_Empregado'][i], "unity": df['int_CodLotacao_Empregado'][i], "mat": df['Str_Mat_Outlook'][i]})
    if df['Str_Mat_Outlook'][i] == 'c150713;':
        dataFullList.append({"mat": df['Str_Mat_Outlook'][i]})

print(dataFullList)
index = 1

my_image = Image.open(
    "C:\\Users\\Ravin\\Desktop\\send_email_coletivo\\coletivo.jpg")

title_text = []
font = ImageFont.truetype(
    'C:\\Users\\Ravin\\Desktop\\send_email_coletivo\\BebasNeue-Regular.ttf', 16)
# image_editable = ImageDraw.Draw(my_image)
while index < len(data_list):
    title_text.append(data_list[index]["name"] + '                               '
                      + codeAndName[data_list[index]['unity']])
    index = index + 1


fontAss = ImageFont.truetype(
    'C:\\Users\\Ravin\\Desktop\\send_email_coletivo\\BebasNeue-Regular.ttf', 18)


image_editable = ImageDraw.Draw(my_image)

i = 0
y = 200
while i < len(title_text):
    lines = textwrap.wrap(title_text[i], width=250)
    for line in lines:
        width, height = font.getsize(line)
        image_editable.text(((600 - width) / 3, y),
                            line, font=font, fill="white", stroke_width=0, stroke_fill="white", align="left")
        y += height * 1.5
    i = i + 1


my_image.save("result.png", optimize=True, quality=100)
# while j < len(data_list):
#     message = outlook.CreateItem(0)
#     message.Display()
#     message.To = data_list[j]["mat"]
#     message.BCC = ass['Assinatura'][2]
#     message.Subject = "Feliz Aniversário!"

# html_body = """
#     <div>
#         <img src="C:\\Users\\Ravin\\Desktop\\send_email_coletivo\\result.png">
#     </div>
#     """
# message.HTMLBody = html_body
