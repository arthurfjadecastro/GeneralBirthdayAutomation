import win32com.client as client
from PIL import Image, ImageFont, ImageDraw


outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()
message.To = "arthur.ravin@gmail.com"
message.BCC = "arthur_ravin@hotmail.com"
message.Subject = "Feliz Aniversário!"

my_image = Image.open(
    "C:\\Users\\Ravin\\Desktop\\send_email_python\\parabensind.jpg")
title_font = ImageFont.truetype(
    'C:\\Users\\Ravin\\Desktop\\send_email_python\\BebasNeue-Regular.ttf', 32)
title_text = "Arthur de Castro, este é seu dia, e por isso deve festejar com alegria. Espero que receba muito carinho, homenagens e surpresas boas. Parabéns e muitas felicidades! "
image_editable = ImageDraw.Draw(my_image)
image_editable.text((15, 15), title_text, (237, 230, 211), font=title_font)
my_image.save("result.jpg")


html_body = """
    <div>
        <img src="C:\\Users\\Ravin\\Desktop\\send_email_python\\result.jpg" width=100%>
    </div>
    """

message.HTMLBody = html_body
