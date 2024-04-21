import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import sys
import time

programConfig = {}
mailTitle = ""
mailBody = ""
mailRtf = ""
mailAttaches = []
mailingList = []

def loadConfig():
    global programConfig
    if not os.path.isfile("Config.ini") :
        print("Cannot find Config.ini, quitting...")
        exit()

    configFile = open("Config.ini", "r", encoding="utf-8")
    for configLine in configFile:
        configLine = configLine.replace('\n', '')
        configSplits = configLine.split('=')
        if len(configSplits) >= 2 : # Only operates when splits results are correct.
            configTitle = configSplits[0]
            configValue = configSplits[1]
            if (len(configSplits) > 2):
                print(configLine + " contains split character \'=\'.")
                print("Please check the config value.")
                for i in range(2, len(configSplits)):
                    configValue += '='
                    configValue += configSplits[i]
            programConfig[configTitle] = configValue
        else:
            print(configLine + " is not a correct config discription.")
    configFile.close()

def loadTitle():
    global mailTitle
    if not os.path.isfile("MailTitle.txt") :
        print("Cannot find MailTitle.txt, your Mail will be titled \"Untitled\"")
        mailTitle = "Untitled"
        return
    
    titleFile = open("MailTitle.txt", "r", encoding="utf-8")
    for titleLine in titleFile:
        titleLine = titleLine.replace('\n', ' ')
        mailTitle += titleLine
    titleFile.close()

def loadBody():
    global mailBody
    if not os.path.isfile("MailBody.txt") :
        print("Cannot find MailBody.txt, your Mail will be embedded with no plain text")
        mailBody = ""
        return

    bodyFile = open("MailBody.txt", "r", encoding="utf-8")
    for bodyLine in bodyFile:
        mailBody += bodyLine
    bodyFile.close()

def loadRtfBody():
    global mailRtf
    if not os.path.isfile("MailBody.html") :
        print("Cannot find MailBody.html, your Mail will not have RTF mail body")
        mailRtf = ""
        return
    
    rtfFile = open("MailBody.html", "r", encoding="utf-8")
    for rtfLine in rtfFile:
        rtfLine = rtfLine.replace('\n', '')
        mailRtf += rtfLine
    rtfFile.close()

def loadAttach():
    global mailAttaches
    if not os.path.isfile("MailAttach.txt") :
        print("This mail does not have any attachments due to not able to find MailAttach.txt")
        return

    attachFile = open("MailAttach.txt", "r", encoding="utf-8")
    for attachLine in attachFile:
        attachLine = attachLine.replace('\n', '')
        mailAttaches.append(attachLine)
    
    attachFile.close()

def loadList():
    global mailingList
    if not os.path.isfile("MailingList.txt") :
        print("You didn't provide any addresses to send this Email to.")
        exit()
    
    listFile = open("MailingList.txt", "r", encoding="utf-8")
    for listLine in listFile:
        listLine = listLine.replace('\n', '')
        mailingList.append(listLine)
    listFile.close()

def mailer():
    global programConfig
    global mailTitle
    global mailBody
    global mailRtf
    global mailAttaches
    global mailingList
    # Check encryption status
    if (stringToBoolean(programConfig.get('smtp_is_ssl'))):
        try:
            mailServer = smtplib.SMTP_SSL(programConfig.get('smtp_server'), int(programConfig.get('smtp_port')))
        except:
            print("SMTP Server Config is not valid.")
            print(sys.exc_info())
            exit()
        # If SMTP is using real SSL, it needs a different class to init
    elif (stringToBoolean(programConfig.get('smtp_is_starttls'))):
        try:
            mailServer = smtplib.SMTP(programConfig.get('smtp_server'), int(programConfig.get('smtp_port')))
            mailServer.ehlo()
            mailServer.starttls()
        except:
            print("SMTP Server Config is not valid.")
            print(sys.exc_info())
            exit()
    else:
        try:
            mailServer = smtplib.SMTP(programConfig.get('smtp_server'), int(programConfig.get('smtp_port')))
        except:
            print("SMTP Server Config is not valid.")
            print(sys.exc_info())
            exit()


    # Check if require logins
    if (stringToBoolean(programConfig.get('smtp_is_auth'))):
        try:
            mailServer.login(programConfig.get('smtp_login'), programConfig.get('smtp_password'))
        except:
            print("Login to SMTP is incorrect")
            print(sys.exc_info())
            exit()

    # Forming Email and Send out
    for toMail in mailingList :
        newMail = MIMEMultipart('alternative')
        newMail["To"] = toMail
        newMail["From"] = programConfig.get('mail_account')
        newMail["Subject"] = mailTitle

        # Check whether the Mail has pure text part, if so, embed the pure text
        if (len(mailBody) > 0):
            mailText = MIMEText(mailBody, "plain", "utf-8")
            newMail.attach(mailText)

        # Check whether the Mail has RTF part, if so, embed the RTF part
        if (len(mailRtf) > 0):
            rtfText = MIMEText(mailRtf, "html", "utf-8")
            newMail.attach(rtfText)

        # Check whether the Mail has attachment, if so, embed the file part
        if (len(mailAttaches) > 0):
            for mailAttach in mailAttaches:
                attachInfo = mailAttach.split(':')
                if (len(attachInfo) != 2) :
                    print(mailAttach + ' is not in correct format')
                else:
                    if not os.path.isfile(attachInfo[1]) :
                        print(attachInfo[1] + ' does not exist.')
                    else:
                        with open(attachInfo[1], 'rb') as toAttachFile:
                            attachPart = MIMEApplication(
                                toAttachFile.read(), name=attachInfo[1], _subtype=attachInfo[0]
                            )
                            attachPart['Content-Disposition'] = 'attachment; filename="%s"' % attachInfo[1]
                            newMail.attach(attachPart)
        
        try:
            mailServer.sendmail(programConfig.get('mail_account'), toMail, newMail.as_string())
            print('Email to ' + toMail + ' with content: ')
            print(newMail)
            print('has been sent')
        except:
            print('Cannot send this Email')
            print(newMail)
            print('It is because of')
            print(sys.exc_info())

        if (int(programConfig.get('send_interval')) > 0):
            print("Wait " + programConfig.get('send_interval') + " seconds")
            time.sleep(int(programConfig.get('send_interval')))
        else :
            print("Setting interval as 0 may result in being banned from server.")
        
    mailServer.quit()
    
def stringToBoolean(boolString):
    if boolString.lower() in ['true', '1', 't', 'y', 'yes', 'yeah', 'yup', 'certain', 'certainly'] :
        return True
    else:
        return False

def main():
    loadConfig()
    loadTitle()
    loadBody()
    loadRtfBody()
    loadAttach()
    loadList()
    mailer()

if __name__ == "__main__":
    main()
