import pandas as pd
import random as rd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

path = r'path\\Secret Santa.xlsx'
secret_santa = pd.read_excel(path)


#Matches people and ensures people do not get pair with "Do not pair with:"
list_of_users = list(secret_santa['Name'])
for user in secret_santa.index:
    match = False
    while not match:
        choice = rd.choice(list_of_users)
        if secret_santa['Do not pair with:'][user] != choice and secret_santa['Name'][user] != choice:
            secret_santa['Match with:'][user] = choice
            match = True
    
    list_of_users.remove(choice)

#Saves a copy of file with a column showing who each person got paired with 
secret_santa.to_excel("Secret Santa Output.xlsx", index=False) 

secret_santa = secret_santa.set_index('Name').transpose().to_dict()

sender_address = 'yourmail@mail.com'
sender_pass = 'password'

for name, details in secret_santa.items():
    print(f'Sending email to: {name}')
    receiver_address = details['Email']
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = 'Your Secret Santa is.....'   #The subject line
    
    #Email message
    mail_content = f"""
    Hey {name.split(' ')[0]},
    
    Your secret santa is {details['Match with:']}. This person likes {secret_santa[details["Match with:"]]["Hobbies / Likes"]}. 
    We are going with a $45(plus tax) gift limit. See you on December 23th!

    Thanks,
    This email was auto generated using Python SMTP library.
    """

    #The body and the attachments for the mail
    message.attach(MIMEText(mail_content, 'plain'))

    #Create SMTP session for sending the mail through GMAIL
    session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
    session.starttls() #enable security
    session.login(sender_address, sender_pass) #login with mail_id and password
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)

session.quit()

