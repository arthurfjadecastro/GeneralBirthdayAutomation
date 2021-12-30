import win32com.client as client
from PIL import Image, ImageFont, ImageDraw
import textwrap
import pandas as pd
import datetime


da = datetime.date.today().strftime("%d/%m/%Y")
outlook = client.Dispatch("Outlook.Application")

df = pd.read_excel(
    r"C:\Users\Ravin\Desktop\send_email_python\Empregados.xlsx")

df["dte_Nascimento_Empregado"] = pd.to_datetime(df["dte_Nascimento_Empregado"])

df = df[['Str_Mat_Outlook', 'str_Nome_Empregado', 'dte_Nascimento_Empregado']]


data_list = [{}]
for i in range(len(df)):
    if df['dte_Nascimento_Empregado'][i].strftime("%m-%d") == datetime.date.today().strftime('%m-%d'):
        data_list.append({"birthDate": df['dte_Nascimento_Empregado'][i].strftime(
            "%m-%d"), "name": df['str_Nome_Empregado'][i], "mat": df['Str_Mat_Outlook'][i]})


index = 1
while index < len(data_list):
    print(data_list[index]["name"])
    print(data_list[index]["birthDate"])
    print(data_list[index]["mat"])
    message = outlook.CreateItem(0)
    message.Display()
    message.To = data_list[index]["mat"]
    message.BCC = "marciano.matos@caixa.gov.br"
    message.Subject = "Feliz Aniversário!"
    # Transformar primeiro nome
    # # Usar Matrícula em um while como destinatário
    # # Converter data de nascimento em Brasil e verificar se há necessidade de enviar o e-mail de aniversário
    my_image = Image.open(
        "C:\\Users\\Ravin\\Desktop\\send_email_python\\parabensind.jpg")
    title_text = data_list[index]["name"] + \
        ", este é seu dia, e por isso deve festejar com alegria. Espero que receba muito carinho, homenagens e surpresas boas. Parabéns e muitas felicidades!"
    lines = textwrap.wrap(title_text, width=44)
    y_text = 100
    font = ImageFont.truetype(
        'C:\\Users\\Ravin\\Desktop\\send_email_python\\BebasNeue-Regular.ttf', 32)
    image_editable = ImageDraw.Draw(my_image)
    for line in lines:
        width, height = font.getsize(line)
        image_editable.text(((550 - width) / 3, y_text),
                            line, font=font, fill="white")
        y_text += height

    my_image.save("result.jpg")

    html_body = """
        <div>
            <img src="C:\\Users\\Ravin\\Desktop\\send_email_python\\result.jpg" width=100%>
        </div>
        """
    message.HTMLBody = html_body
    index = index + 1
