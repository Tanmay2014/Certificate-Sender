from PIL import Image, ImageDraw, ImageFont
import textwrap
import pandas as pd
import yagmail

sheet = pd.read_excel('btor.xlsx')
names = sheet["Names"].tolist()
mails = sheet["Mails"].tolist()

sender = str(input("Please Enter your Email Address : "))#"cdosbdce@gmail.com"

pwd = str(input("Please Enter your Password : "))#"BDCDOS#1"

mail = yagmail.SMTP(user=sender, password=pwd)

subj_ect = "Certificate"


for i in range(len(names)):
    astr = names[i].title()
    para = textwrap.wrap(astr, width=150)

    MAX_W, MAX_H = 3000, 2000
    cert = Image.open("btor1.jpg")
    draw = ImageDraw.Draw(cert)
    font = ImageFont.truetype('Madina.ttf', 200)

    current_h, pad = 50, 10
    for line in para:
        w, h = draw.textsize(line, font=font)
        draw.text(((MAX_W - w) / 2, 900), line, font=font, fill=(0,0,0))
        current_h += h + pad
    topdf=cert.convert('RGB')
    # topdf.save(f"try\{names[i].title}.png")
    savename = f"{names[i].title}"
    cert.save(f"btorcerts\{names[i].title()}.pdf")
    # cert.show('test.png')
    cont_ent = [
            f"Dear {names[i].title()}",
            "\n",
            "Thank you for participating in online workshop of <b>Basic Training On Robotics</b> by <b>Ms. Shreya Sarode.</b>",
            "We've attached your certificate with this email.",
            "You'll have to click on the print icon in the Adobe PDF viewer to export the pdf after verifying",
            "\n",
            "\n",
            "Regards,",
            "Team TFL",
            "BDCE Sewagram.",
        ]
    mail.send(to=mails[i], subject=subj_ect, contents=cont_ent, attachments=f"btorcerts\{names[i].title()}.pdf")
    print(names[i].title())
    print(mails[i])