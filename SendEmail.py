import win32com.client as client
from PIL import Image, ImageFont, ImageDraw
import textwrap


outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()
message.To = "arthur.ravin@gmail.com"
message.BCC = "arthur_ravin@hotmail.com"
message.Subject = "Feliz Aniversário!"


my_image = Image.open(
    "C:\\Users\\Ravin\\Desktop\\send_email_python\\parabensind.jpg")

title_text = "Raphael, este é seu dia, e por isso deve festejar com alegria. Espero que receba muito carinho, homenagens e surpresas boas. Parabéns e muitas felicidades!"
lines = textwrap.wrap(title_text, width=40)
y_text = 100

font = ImageFont.truetype(
    'C:\\Users\\Ravin\\Desktop\\send_email_python\\BebasNeue-Regular.ttf', 32)
image_editable = ImageDraw.Draw(my_image)


for line in lines:
    width, height = font.getsize(line)
    image_editable.text(((450 - width) / 2, y_text),
                        line, font=font, fill="white")
    y_text += height


my_image.save("result.jpg")


html_body = """
    <div>
        <img src="C:\\Users\\Ravin\\Desktop\\send_email_python\\result.jpg" width=100%>
    </div>
    """

message.HTMLBody = html_body
