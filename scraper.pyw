import requests
from bs4 import BeautifulSoup
import smtplib
import time
from win10toast import ToastNotifier
import json
import xlsxwriter

# Add URL of site which allows html extraction
URL = "https://www.google.com/search?q=pound+to+rand&rlz=1C1CHBF_en-GBGB822GB822&oq=pund+to+rand&aqs=chrome.1.69i57j0l5.4103j1j4&sourceid=chrome&ie=UTF-8"
# Define your User Agent. Google Search: "My user agent"
headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36"}

toaster = ToastNotifier()
# Set target for rand to drop to
rand_target = 17.00
# Enter respective sender and receiver email addresses
sender = 'SENDER_EMAIL'
receiver = 'RECEIVER_EMAIL'

# Input password for sender email only!
password = int('PASSWORD')
rand_time = {}
def check_price():
    page = requests.get(URL,headers=headers)

    soup = BeautifulSoup(page.content,'html.parser')
    
    rand = soup.find(class_="DFlfde SwHCTb").get_text()
    date = time.strftime("%Y-%m-%d %H:%M")
    converted_rand = float(rand)
    #print("Rand: " + rand)
    toaster.show_toast(f"Price Check {loop}",f"Exchange script running...\nRand: {rand}",None,4,True)
    add_file(date,rand)
    excel(date,rand)
    if(converted_rand <= rand_target):
        send_mail(rand)
        send_toast(rand)


def add_file(date,rand):
    rand_time[date] = f"{rand}"
    with open ('full/path/rand_times.json') as my_dict:
        data = json.load(my_dict)
    data.update(rand_time)

    with open("full/path/rand_times.json",'w') as my_dict:
        json.dump(data,my_dict)

def excel(date, rand):
    workbook = xlsxwriter.Workbook('rand.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write('1','0',date)
    worksheet.write('1','1',rand)

def send_toast(rand):
    title = "Rand Exchange!"
    message = f"Rand: R{rand}"
    toaster.show_toast(title,message,None,4,False)

def send_mail(rand):
#     print("Establishing Connection...")
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.ehlo()
#     print("Connection to server successful")
    server.login(sender,password)
#     print("Login successful")

    subject = 'Rand is stronger'
    body = f'Rand: R{rand}\nCheck the link {URL}'

    msg = f'subject: {subject}\n\n{body}'

    server.sendmail(sender,receiver,msg)

#     print('CHECK MAIL')

    server.quit()
#     print("Server successfully quit")

#Allow less secure apps
loop = 1
check_price()
