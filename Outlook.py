
import win32com.client
import os
from datetime import datetime, timedelta

def outlook():
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")

    for account in mapi.Accounts:
        print(account.DeliveryStore.DisplayName)


    inbox = mapi.GetDefaultFolder(6)
    print(inbox)
    messages = inbox.Items
   # print(messages)
    #received_dt = datetime.now() - timedelta(days=1)
   # received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
   # messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
   #messages = messages.Restrict("[SenderEmailAddress] = 'razirahmanvelocity@gmail.com'")
   # messages = messages.Restrict("[Subject] = 'Test'")
    outputDir = r"C:\attachment"
    try:
        for message in list(messages):
          print(message)
          try:
             s = message.sender
             print(message.Subject)
             print(s)
             if message.Unread==True:
                print(message.subject)
             for attachment in message.Attachments:
                attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
                print(f"attachment {attachment.FileName} from {s} saved")
          except Exception as e:
            print("error when saving the attachment:" + str(e))
    except Exception as e:
            print("error when processing emails messages:" + str(e))