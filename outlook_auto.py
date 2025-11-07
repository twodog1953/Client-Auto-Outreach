from time import sleep
import pandas as pd
import win32com.client
import os

f_path = "C:/Users/Oliver/Desktop/Klaus's Code/Private_Code/email_lst.xlsx"
# f_path = "C:/Users/Oliver/Desktop/SPAM EMAILS/client_susan.xlsx"

attach_path = "C:/Users/Oliver/Desktop/Klaus's Code/Private_Code/SPAM EMAILS/hol.jpg"
body_path = "C:/Users/Oliver/Desktop/Klaus's Code/Private_Code/SPAM EMAILS/hol_2024.txt"
subject_txt = 'Holidays Greeting'

df = pd.read_excel(f_path)[["to", "cc"]]

f = open(body_path, "r", encoding='utf-8')
body_txt = f.read()
f.close()

# extract only to and cc column for mass spamming
to_n_cc = (df.dropna(subset=["to"]).drop_duplicates(["to", "cc"]).fillna(" ").reset_index(drop=True))
print(to_n_cc)

for i in range(len(to_n_cc["to"])):
    i_to = to_n_cc.loc[i, "to"]
    i_cc = to_n_cc.loc[i, "cc"]

    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0

    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = subject_txt
    newmail.To = i_to
    newmail.CC = i_cc

    newmail.BodyFormat = 2
    attachment = newmail.Attachments.Add(attach_path)  # Attach the image
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "myimage")  # Set Content ID for inline use
    newmail.HTMLBody = f'''
    <html>
        <body>
            <img src="cid:myimage">
            <pre style="font-family: Arial, sans-serif; font-size: 14px;">{body_txt}</pre>
        </body>
    </html>
    '''

    # newmail.Display()
    newmail.Send()
    # sleep(100)