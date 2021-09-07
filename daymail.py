import json
import os
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import schedule
import time


def load_json():
    # fname = r"C:\Users\Samit\PycharmProjects\HR\Files\Json\emaildates.json"
    fname=os.path.join(os.getcwd(),"Files","Json","emaildates.json")
    today = (datetime.now())
    today =datetime.strptime(F"{today.month}/{today.day}/{today.year}","%m/%d/%Y")
    with open(os.path.join(fname)) as json_file:
        data = json.loads(json.load(json_file))
    for x in data:
        # print(data[x][0])
        tmpdate = datetime.strptime(data[x][0],'%m/%d/%Y')
        print(tmpdate)
        if (tmpdate-today).days == 3:
            print("sending mail for ",x)
            tmp = []
            tmp.append(data[x][1])
            send_mail(tmp,data[x][2],F"{tmpdate.day}/{tmpdate.month}/{tmpdate.year}")


def send_mail(mail,name,date):
    fromaddr = "mail@example.com"
    toaddr = mail
    msg = MIMEMultipart()
    # date = f"{date[3:5]}/{date[:2]}/{date[6:]}"
    msg['From'] = fromaddr

    msg['To'] = ", ".join(toaddr)

    msg['Subject'] = "Joining Date"

    body = F"""Hi {name},
    This is a remainder mail to inform you that you have to join at mahindra office on {date}.
    If you have any queries please feel free to contact."""

    msg.attach(MIMEText(body, 'plain'))
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(fromaddr, "PASSWORD_HERE")
    text = msg.as_string()
    s.sendmail(fromaddr, toaddr, text)
    s.quit()


if __name__ == "__main__":
    schedule.every().day.at("01:49").do(load_json)
    # schedule.every().day.at("11:30").do(load_json)
    while True:
        schedule.run_pending()
        time.sleep(30) # wait one minute
