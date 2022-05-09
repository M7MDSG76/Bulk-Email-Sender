import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import sys, pandas as pd, os
from pathlib import Path

def start_smtp_gmail_session(sender_address, sender_pass):
   # Create SMTP session for sending the email
        session = smtplib.SMTP('smtp.gmail.com', 587)  # use gmail with port
        session.starttls()  # enable security
        session.login(sender_address, sender_pass)  # login 
        return session


def create_MIME_object(sender_address, receiver_address,
                   email_subject, email_message):
        # MIME settings
        MIME_object = MIMEMultipart()  # Create a MIME object.

        MIME_object['From'] = sender_address  # Set the sender address.

        MIME_object['To'] = receiver_address  # Set the Receiver address.

        MIME_object['Subject'] = email_subject # Set the The email Subject.

        body = email_message    # Set the mail body.

        MIME_object.attach(MIMEText(body, 'plain'))     # Attach the mail body to the message object.


        return MIME_object
                
                
def attach_file_to_MIME_object(MIME_object , *files):
     
        BASE_DIR = os.path.join(Path(__file__).resolve().parent.parent, '\\attachments\\')
        
        files_list = []
        
        for file in files:
            files_list.append(file)
        
        # attach the files to the MIME_obj.
        for i in range(len(files_list)): 
            print(files_list[i])
            attachment = open(os.path.join(BASE_DIR,files_list[i]), "rb")
            
            mb = MIMEBase("application", "octet-stream")    # Object of MIMEBase. 

            mb.set_payload((attachment).read())   # To change the payload into encoded form

            encoders.encode_base64(mb)      # Encode into base64 

            mb.add_header('Content-Disposition', f"attachment; fileName = {files_list[i]}")

            MIME_object.attach(mb)  # Attach Object 'mb' to object 'MIME_obj'.


def MIME_object_to_string(MIME_object):
        try:
            message = MIME_object.as_string()
            return message
        except:   
            print("Oops!", sys.exc_info()[0], "occurred.")
            print("Next entry.\n")
            print()   
   

def create_message(sender_address, receiver_address, email_subject, email_message, *files):
     # Create MIME object without attachments.
        MIME_object = create_MIME_object(sender_address, receiver_address,
                                            email_subject, email_message)
        
        # if attachmentes attached to the MIME object.
        if files:
            attach_file_to_MIME_object(MIME_object, *files)
        
        # Convert MIME object to string.
        message = MIME_object_to_string(MIME_object)   
        
        return message
        
              
def from_excel_to_list(xlsx_path, sheet_name,
                       col_name):
    # Declare Dataframe
    df = pd.read_excel(xlsx_path, sheet_name)
    
    # This list is acontainer of recivers addresses list
    targets_address_list = []
    
    # Iterate over the dataframe and move addresses to the receiver_address_list. 
    for email in df[col_name]:
         targets_address_list.append(email)
         
    return targets_address_list
 
 
def send_SMTP_email(sender_address, target_address, email_subject, email_message, *files, **kwargs):
    
    SMTP_session_status = False   
    
    # check if there is a session started or not.
    for key in  kwargs:
        if key == 'SMTP_session':
            SMTP_session_status = True
           
    if SMTP_session_status == False:
        
        try:
            # Start smtp session.
            SMTP_session = start_smtp_gmail_session(sender_address, kwargs['sender_password'])
        except:
            print('Make sure you enterd the sender password')
            print("Oops!", sys.exc_info()[0], "occurred.")
            
        try:
            # Send the message over the SMTP_session to the target address.
            SMTP_session.sendmail(sender_address, target_address, message)
            print(f'Email sent successfully to {target_address}')
            return f'sent successfully to {target_address}\n'
            
        except:
            print(f'Email: {target_address} got an error.')    
            print("Oops!", sys.exc_info()[0], "occurred.")
            print("Next entry.\n")
            return f'Error accurred with {target_address}\n'
    
    elif SMTP_session_status == True:
        
        SMTP_session = kwargs['SMTP_session']
        message = create_message(sender_address, target_address, email_subject, email_message, *files)
        try:
            # Send the message over the SMTP_session to the target address.
            SMTP_session.sendmail(sender_address, target_address, message)
            print(f'Email sent successfully to {target_address}')
            return f'sent successfully to {target_address}\n'
            
        except:
            print(f'Email: {target_address} got an error.')   
            print("Oops!", sys.exc_info()[0], "occurred.")
            print("Next entry.\n")
            return f'Error accurred with {target_address}\n'
  
            
def send_SMTP_email_to_multiple_addresses(sender_address, sender_password, addresses_source, email_subject, email_message, *files):
    
    try:
        # Start smtp session.
        SMTP_session = start_smtp_gmail_session(sender_address, sender_password)
    except:
        print("probably authentication error accured", sender_address)
        print("Oops!", sys.exc_info()[0], "occurred.")
    
    try:
        targets_address_list = from_excel_to_list(addresses_source, 'Emails', 'Email')
        target_status = []
        print(targets_address_list)
        for target in targets_address_list:
            target_status.append(send_SMTP_email(sender_address, target, email_subject, email_message, SMTP_session = SMTP_session))
            
        
    except:
        print(f'File path: {addresses_source}')
        print("Please Check The Addresses File Path!")
        print("Oops!", sys.exc_info()[0], "occurred.")
        
    print('Thank you for using Bulker!\n| Developed By Eng.Mohammed Saleh Alghanmi |')
    SMTP_session.quit()
    return target_status
    
    
        
           
    
    


# subject = '1/2 Teaspoon Burns Fat 728% Faster'
# message = """
# Before you go to bed tonight, eat 1/2 teaspoon of THIS (before 10pm) and boost your metabolism by over 728%!
# Here it is:
# ==> 1/2 Teaspoon Boosts Metabolism By 728% (Slow Metabolism Loophole)
# Skeptical?
# I was too, but then I saw the shocking proof for myself.
# Within weeks folks have dropped an average of 25.3 lbs, waists have shrunk by 7.2 inches.
# See their incredible results for yourself.
# To your amazing health,
# # """
# print('-----Start-----')    
# send_SMTP_email_to_multiple_addresses('mohammed.s.g76@gmail.com', 'nmaeqlgjjnxdzggd', 'sourceFile.xlsx',
#                                       subject, message, ['files\Mohammed_Saleh_Alghanmi_CV.pdf'])
# print('-----End-----')    
   






# def main():
#         # Sender address
#         sender_address = 'mohammed.s.g76@gmail.com' 
        
#         # Use google app pasword to login insted of using the orginal passowrd. 
#         sender_pass = 'nmaeqlgjjnxdzggd' 
        
#         # excel file address path.
#         Addresses_source = 'sourceFile.xlsx'
        
#         # Recruiters addresses list
#         targets_address_list = from_excel_to_list(Addresses_source, 'Emails', 'Email')
        
#         # Email Subject
#         email_subject = 'Computer Science Graduate Seeking for a Job'
        
#         # Email letter
#         email_message = '''
# Dear Hiring Manager, 


# I am Mohammed Saleh Alghanmi, and I am excited to apply for a job at your company.
# I am a recent graduate (April 2021) in Computer Science in Software Engineering track from King Abdul-Aziz University,
# and I am eager to enter the workforce. 

# My cover letter and resume are attached for your review.
# If you would like more information regarding my qualifications for the job, please do not hesitate to reach out. 

# Thank you so much for your time and consideration, and I looking forward to hearing from you soon.

# Sincerely,


# Mohammed Saleh AlGhanmi

# Jeddah | +966-500977323 | Linkedin: https://www.linkedin.com/in/mohammed-s-alghanmi/ | GitHub: https://github.com/M7MDSG76
#         '''
        
        
#         """
#     //// My Email attachments 
#          # Email Subject
#         email_subject = 'Computer Science Graduate Seeking for a Job'
        
#         # Email letter
#         email_message = '''
# Dear Hiring Manager, 


# I am Mohammed Saleh Alghanmi, and I am excited to apply for a job at your company.
# I am a recent graduate (April 2021) in Computer Science in Software Engineering track from King Abdul-Aziz University,
# and I am eager to enter the workforce. 

# My cover letter and resume are attached for your review.
# If you would like more information regarding my qualifications for the job, please do not hesitate to reach out. 

# Thank you so much for your time and consideration, and I looking forward to hearing from you soon.

# Sincerely,


# Mohammed Saleh AlGhanmi

# Jeddah | +966-500977323 | Linkedin: https://www.linkedin.com/in/mohammed-s-alghanmi/ | GitHub: https://github.com/M7MDSG76
#         '''
#         """
        
#         # The CV file name without the path (it should be in the same as the project file path)
#         CV_file_name = 'Mohammed_Saleh_Alghanmi_CV.pdf'
        
#         # The Cover letter file name without the path (it should be in the same as the project file path)
#         coverletter_file_name = 'Mohammed-S-Alghanmi_Coverletter.docx'
        
        
#         # Start smtp session.
#         session = start_smtp_gmail_session(sender_address, sender_pass)
        
#         # Iterate over the recruiters and send email for each.
#         for recruiter in range(len(targets_address_list)):
            
#             # declare a message object and attach every thing.
#             message = create_message(sender_address, targets_address_list[recruiter],
#                                      email_subject, email_message,
#                                      CV_file_name, coverletter_file_name)
#             try:
#                 # Send the message over the smtp session to the recruiter address.
#                 session.sendmail(sender_address, targets_address_list[recruiter], message)
#                 print(f'Email sent successfully to {targets_address_list[recruiter]}')
                
#             # If Error accure make an exception.
#             except:
#                 print(f'Email: {targets_address_list[recruiter]} got an error.')    
#                 print("Oops!", sys.exc_info()[0], "occurred.")
#                 print("Next entry.\n")
  
#         session.quit()    
        
# if __name__ == '__main__':

#         main()