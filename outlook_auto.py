import tkinter
from time import sleep
import pandas as pd
import win32com.client
from tkinter import *
import os, glob
import calendar

def contact_filter():
    # todo: for filtering out contact info directly from customer list

    return


def new_email():
    for i in range(len(to_n_cc["to"])):
        i_to = to_n_cc.loc[i, "to"]
        i_cc = to_n_cc.loc[i, "cc"]

        ol = win32com.client.Dispatch('Outlook.Application')
        olmailitem = 0x0

        newmail = ol.CreateItem(olmailitem)
        newmail.Subject = title_box.get()
        newmail.To = i_to
        newmail.CC = i_cc

        newmail.BodyFormat = 2
        attachment = newmail.Attachments.Add(os.path.abspath(attach_path))  # Attach the image
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                                                "myimage")  # Set Content ID for inline use
        newmail.HTMLBody = f'''
        <html>
            <body>
                <img src="cid:myimage">
                <pre style="font-family: Arial, sans-serif; font-size: 14px;">{body_txt}</pre>
            </body>
        </html>
        '''

        if if_direct_send.get() == 1:
            print(f"---Auto Send is ON, auto sending emails to {i_to} ---")
            newmail.Send()
        else:
            print(f"---Auto Send is OFF, displaying emails to {i_to} ---")
            newmail.Display()
    return

f_path = "outlook_contact.xlsx"
# f_path = "C:/Users/Oliver/Desktop/SPAM EMAILS/client_susan.xlsx"

attach_path = "hol.jpg"
body_path = "sample_body.txt"
# subject_txt = 'Holidays Greeting'

df = pd.read_excel(f_path)[["to", "cc"]]

f = open(body_path, "r", encoding='utf-8')
body_txt = f.read()
f.close()

# run GUI
root = Tk()
root.geometry("350x200")
root.title('Auto Email Outreach - By Klaus')

# for entering title for email
title_label = Label(root, text='Title of Email', font=("Comic Sans MS", 14))
title_label.grid(row=1, column=1)
title_box = Entry(root, width=20, font=("Comic Sans MS", 12))
title_box.grid(row=2, column=1)

# if direct send
# ds_label = Label(root, text='If direct send?', font=("Comic Sans MS", 8))
# ds_label.grid(row=1, column=2)
if_direct_send = tkinter.IntVar()
ds_box = Checkbutton(root, text='If direct send?', onvalue=1, offvalue=0, variable=if_direct_send)
ds_box.grid(row=1, column=2)

# label for showing current path
new_email_button = Button(root, text='New Email', font=("Comic Sans MS", 14),
                          command=new_email)
new_email_button.grid(row=3, column=1)

# extract only to and cc column for mass spamming
to_n_cc = (df.dropna(subset=["to"]).drop_duplicates(["to", "cc"]).fillna(" ").reset_index(drop=True))
print(to_n_cc)

root.mainloop()