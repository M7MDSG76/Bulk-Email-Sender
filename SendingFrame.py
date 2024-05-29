"""
This Project started on: 4/19/2022  10:00 PM
I start this project after a request from a client and as I have already made the SMTP email sender program.

in this project I will:
- create GUI for the program.
- Restructur a new instance of the program for the client needs.
"""


from mimetypes import init
import tkinter as tk
from tkinter import filedialog as fd
from email_sender import send_SMTP_email_to_multiple_addresses 
import sys as os




class Email():
    def __init__(self):
        self.user_email = 'user_email'
        self.user_password = 'user_password'
        self.email_subject = 'email_subject'
        self.email_message = 'email_message'
        self.addresses_file_path = 'addresses_file_path'
    
    def set_user_email(self, email):
        self.user_email =  email
        
       
    def get_user_email(self):
        return  self.user_email
    
    
    def set_user_password(self, password):
        self.user_password = password
    
    def get_user_password(self):
        return self.user_password
        
    
    def set_email_subject(self, subject):
        self.email_subject =  subject
       
    def get_email_subject(self):
        return  self.email_subject
    
    
    def set_email_message(self, message):
        self.email_message =  message
       
    def get_email_message(self):
        return  self.email_message
    
    
    def set_addresses_file_path(self, file_path):
        self.addresses_file_path =  file_path
       
    def get_addresses_file_path(self):
        return  self.addresses_file_path
    
    
    
# GUI 

# Tkinter window
window = tk.Tk()
window.title('Bulker')


email_obj = Email()


#User info
user_email_lable = tk.Label(window, text='User Email:')
user_password_lable = tk.Label(window, text='User password:')

user_email_entry = tk.Entry(window, width=40)
user_password_entry = tk.Entry(window, width=40)

# Set values to the email object.



# Email Info
email_subject_lable = tk.Label(window, text='Subject:')
email_message_lable = tk.Label(window, text='message:')

email_subject_entry = tk.Entry(window, width=40)
#email_message_entry = tk.Text(window, width=40)
email_message_entry = tk.Entry(window, width=40)
print('Text: ',email_message_entry.get())
print('Type: ',type(email_message_entry.get()))


# Target addresses source
def select_file(email_obj):
    file_address = fd.askopenfilename()
    email_obj.set_addresses_file_path(file_address)
    print('-------type: ',type(email_obj.get_addresses_file_path()))
    

addresses_file_path_lable = tk.Label(window, text='Addresses file path: ')
addresses_file_path_entry = tk.Button(window, text='Select Addresses file', command=lambda:select_file(email_obj))


def send_email(email_obj):
    # Set values
    email_obj.set_user_email(user_email_entry.get())
    email_obj.set_user_password(user_password_entry.get())
    
    email_obj.set_email_subject(email_subject_entry.get())
    email_obj.set_email_message(email_message_entry.get()) 
    
    
    user_email = email_obj.get_user_email()
    user_password = email_obj.get_user_password()
    email_subject = email_obj.get_email_subject()
    email_message = email_obj.get_email_message()
    addresses_file_path = email_obj.get_addresses_file_path()
    
    
    # Send 
    send_SMTP_email_to_multiple_addresses(user_email, user_password, addresses_file_path, email_subject, email_message)
    
    
def exite_view():
    exite_popup = tk.Tk()
    exite_popup.title('Exite Pop up.')
    exite_message = tk.Label(exite_popup, text='Are you ure you want to close the application?')
    exite_confirmation_button = tk.Button(exite_popup, text="I'm sure", command=os.exit)
    exite_cancellation_button = tk.Button(exite_popup, text="Cancel", command=exite_popup.destroy)
    exite_message.pack()
    exite_confirmation_button.pack()
    exite_cancellation_button.pack()
    exite_popup.mainloop()
    
   
# Sending
send_button = tk.Button(window, text='Send', command=lambda:send_email(email_obj))
exite_button = tk.Button(window, text='Exite', command=exite_view)


user_email_lable.pack()
user_email_entry.pack()

user_password_lable.pack()
user_password_entry.pack()

email_subject_lable.pack()
email_subject_entry.pack()

email_message_lable.pack()
email_message_entry.pack()

addresses_file_path_lable.pack()
addresses_file_path_entry.pack()

send_button.pack()
exite_button.pack()


window.mainloop()
