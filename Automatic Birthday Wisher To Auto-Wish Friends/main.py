import pandas as pd
import datetime
import openpyxl
import  smtplib
import os
#from twilio.rest import Client
os.chdir(r"F:\Python\Practice\P5\Automatic Birthday Wisher To Auto-Wish Friends")
# os.mkdir("testing")
# Enter Your Details
GMAIL_ID = 'faisal25marcg@gmail.com'
GMAIL_PWD = 'messifacebook'



def sendEmail(to,sub,msg):
    print(f" Email to: {to}, Sent with subject: {sub}, Message: {msg} ")
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PWD)
    s.sendmail(GMAIL_ID,to,f"Subject: {sub}\n\n {msg}")
    s.quit()
# sendEmail(GMAIL_ID,"Subject","test Message")
# exit()

df = pd.read_excel("data.xlsx")
# print(df)
today = datetime.datetime.now().strftime("%d-%m")
yearNow = datetime.datetime.now().strftime("%Y")
# print(today)
writeInd = []
for index,item in df.iterrows():
    # print(index,item["Birthday"])
    bday = item["Birthday"].strftime("%d-%m")
    # print(bday)
    if today == bday and yearNow not in str(item["Year"]):
        sendEmail(item["Email"],"Happy Bithday",item["Wish"])
        writeInd.append(index)
# print(writeInd)
for i in writeInd:
    yr = df.loc[i,"Year"]
    # print(yr)
    df.loc[i,"Year"] = f"{yr},{yearNow}"
    # print(df.loc[i,"Year"])
# print(df)
df.to_excel('data.xlsx',index=False)