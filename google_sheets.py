#google sheet classes
import gspread
from oauth2client.service_account import ServiceAccountCredentials
#email classes
import smtplib
from email.message import EmailMessage


scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("sheetsconnection.json",scope)

client = gspread.authorize(creds)
sheet = client.open("EXOMUN DELEGATES ").sheet1



#reading row number from file
input_file = open("row number.txt","r")
number = input_file.readline()
#number = number.rstrip('\n')
input_file.close()
number = (int(number))



#access data
row_count = len(sheet.get_all_values())
counter = number
#col = (int)sheet.col_values(4)

print("Mails sent to the following addresses: ")

#loop to pick all the emails one by one
while counter <= row_count:
    email = sheet.cell(counter,4).value
    print(email)
    # send EmailMessage
    body= "Greetings! \n\nThank you for registering with EXOMUN. Your registration was successful.\n\nYou shall get your country and committee allotments by 16th August. \n\nThanks and Regards \nTeam EXOMUN"

    msg = EmailMessage()
    msg.set_content(body)

    msg["Subject"] = "Registration confirmation"
    msg["From"]= "confrimationexomun@gmail.com"
    msg["To"]= email

    server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    server.login("confrimationexomun@gmail.com", "exomunlak12")
    #server.sendmail("lakshaykarnwal@gmail.com","{}".format(email),"{}".format(body))
    server.send_message(msg)

    server.quit()
    counter = counter+1

print("\nTotal emails sent: " + str(counter-number))
print("Current row number: " + str(counter))

#output of the row number to the file
output_file = open("row number.txt","w")
output_file.write(str(counter))
output_file.close()


print("Program completed")
