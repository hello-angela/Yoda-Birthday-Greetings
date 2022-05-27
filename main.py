from asyncore import write
import email
import pandas as pd
import datetime
import smtplib
import random
import mimetypes
import getpass
from email.utils import make_msgid
from email.message import EmailMessage

def makeCSV():
    """
    take an excel doc with name, email, birthday (Y-M-D format) saved as "birthdays"
    """
    birthdays = pd.DataFrame(pd.read_excel("birthdays.xlsx"))
    birthdays.to_csv ("birthdayCSV.csv", 
                  index = None,
                  header=True)
    return birthdays
    
    
def getBirthdays(birthdays):
    """
    retrieves the names of birthday celebrants (birthdayPeople),
    birthday emails to send (birthdayMessages),
    number of birthdays today (numbOfbirthdays)
    """
    dateToday = datetime.date.today()
    year = datetime.date.today().year

    todaysBirthdays = []
    birthdayEmails = []

    for index, row in birthdays.iterrows():
        person  = birthdays.iloc[index].iat[2]
        parsed = datetime.datetime.strptime(str(person), '%Y-%m-%d %H:%M:%S').date().replace(year)
        if parsed == dateToday:
                birthdayToday = birthdays.iloc[index].iat[0]
                todaysBirthdays.append(birthdayToday)
                emailAddresses = birthdays.iloc[index].iat[1]
                birthdayEmails.append(emailAddresses)
                
    birthdayPeople = todaysBirthdays
    birthdayMessages = str(birthdayEmails)
    numbOfbirthdays = len(todaysBirthdays)            
                
    return birthdayPeople, birthdayMessages, numbOfbirthdays


def emailAccount():
    """
    works for gmail
    NB: security details of gmail account needs to be changed to allow 3rd party apps to send emails
    """
    username = input("Please login to your account: ")
    password = getpass.getpass("Please enter your password: ")
    print("Sending birthday greetings...")
    return username, password


def writeBirthdayMessage(birthdayPeople, birthdayMessages, numbOfbirthdays, username):
    birthdayGreetings = ['fork.jpg', 'baby-yoda.jpg', 'original-yoda.jpg', 'baby-yoda2.jpg']
    birthday = random.choice(birthdayGreetings)

    for i in range(numbOfbirthdays):
        name = birthdayPeople[i]
        email = birthdayMessages[i]
        msg = EmailMessage()
        msg['From']= username
        msg['To']= email 
        msg['Subject']= "A special message for " + name

        body = "Happy birthday {name}! {birthday} Hope you have an awesome day!"

        image_cid = make_msgid()

        msg.add_alternative("""
            <html>
            <div> 
            <style>
            div {{
            text-align:center
                }}
            </style>
                <h1 style="font-family: 'Lucida Handwriting'; font-size: 56; font-weight: bold; color: #d74323;">
                    Happy Birthday {name}!!
                    <p> <img src="cid:{image_cid}"></p>
                <span style="font-family: 'Lucida Sans'; font-size: 28; color: #5927dc;">
                    Hope you have an awesome day!
                    <p></p>
            </html>
                """.format(name=name,image_cid=image_cid[1:-1]), subtype='html')

    with open(birthday, 'rb') as img:
        maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')
        msg.get_payload()[0].add_related(img.read(), 
                                maintype=maintype, 
                                subtype=subtype, 
                                cid=image_cid)
    
    return msg
        

def sendEmail(username, password, msg):                         
    connection = smtplib.SMTP(host='smtp.gmail.com', port=587)
    connection.starttls()
    connection.login(username,password)
    connection.send_message(msg)
    print("Message sent!")
    connection.quit()
    
    
def main():
    birthdays = makeCSV()
    birthdayPeople, birthdayMessages, numbOfbirthdays = getBirthdays(birthdays)
    username,password = emailAccount()
    msg = writeBirthdayMessage(birthdayPeople, birthdayMessages, numbOfbirthdays, username)
    sendEmail(username, password, msg)
    

if __name__ == "__main__":
    main()

    





