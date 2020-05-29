import pandas as pd 
import datetime
import smtplib


def sendemail(to,sub,content):
    new_file=open(r"other\ps.txt")
    server=smtplib.SMTP("smtp.gmail.com",587)
    server.ehlo()
    server.starttls()
    server.login("abc@gmail.com",new_file.read())
    new_file.close()
    server.sendmail("abc@gmail.com",to,f"subject: {sub}\n\n{content}")
    server.quit()

if __name__ == "__main__":
    data_frame=pd.read_excel("bday_data.xlsx")
    today=datetime.datetime.now().strftime("%d-%m")
    current_year=datetime.datetime.now().strftime("%Y")
    c=[]
    for index,value in data_frame.iterrows():
        birthday_date=value["Birthday"].strftime("%d-%m")
        wishing_year=value["Year"]
        if (today==birthday_date and current_year not in str(wishing_year)):
            sendemail(value["Email"],"Happy Birthday",value["Wish"])
            c.append(index)
    if (len(c)!=0):
        for i in c:
            lock_year=data_frame.loc[i,"Year"]
            data_frame.loc[i,"Year"]=str(lock_year)+","+current_year
    else:
        print("There is no one's birthday today")

    data_frame.to_excel("bday_data.xlsx",index=False)
