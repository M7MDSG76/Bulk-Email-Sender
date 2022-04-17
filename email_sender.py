from email.mime import message
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


def create_message(sender_address, receiver_address,
                   email_subject, email_message,
                   CV_file_name, coverletter_file_name):
        # MIME settings
        message = MIMEMultipart()  # Create a MIME object.

        message['From'] = sender_address  # Set the sender address.

        message['To'] = receiver_address  # Set the Receiver address.

        message['Subject'] = email_subject # Set the The email Subject.

        body = email_message    # Set the mail body.

        message.attach(MIMEText(body, 'plain'))     # Attach the mail body to the message object.


        
        BASE_DIR = os.path.join(Path(__file__).resolve().parent.parent, 'email_sender\\files')

        files = [CV_file_name, coverletter_file_name]   #files path list.
        
        # attach the files to the message.
        for i in range(len(files)): 
            print(files[i])
            attachment = open(os.path.join(BASE_DIR,files[i]), "rb")
            
            mb = MIMEBase("application", "octet-stream")    # Object of MIMEBase. 

            mb.set_payload((attachment).read())   # To change the payload into encoded form

            encoders.encode_base64(mb)      # Encode into base64 

            mb.add_header('Content-Disposition', f"attachment; fileName = {files[i]}")

            message.attach(mb)  # Attach Object 'mb' to object 'message'.
            
        try:
            msg = message.as_string()
            return msg
        except:   
            
            print("Oops!", sys.exc_info()[0], "occurred.")
            print("Next entry.\n")
            print() 
                
   
def from_excel_to_list(xlsx_path, sheet_name,
                       col_name):
    # Declare Dataframe
    df = pd.read_excel(xlsx_path, sheet_name)
    
    # This list is acontainer of recivers addresses list
    receiver_address_list = []
    
    # Iterate over the dataframe and move addresses to the receiver_address_list. 
    for email in df[col_name]:
         receiver_address_list.append(email)
         
    return receiver_address_list



def main():
        # Sender address
        sender_address = 'mohammed.s.g76@gmail.com' 
        
        # Use google app pasword to login insted of using the orginal passowrd. 
        sender_pass = 'cmyvblrcqncrorux' 
        
        # excel file address path.
        Addresses_source = 'sourceFile.xlsx'
        
        # Recruiters addresses list
        recruiters_list = from_excel_to_list(Addresses_source, 'Emails', 'Email')
        
        # Email Subject
        email_subject = 'Computer Science Graduate Seeking for a Job'
        
        # Email letter
        email_message = '''
Dear Hiring Manager, 


I am Mohammed Saleh Alghanmi, and I am excited to apply for a job at your company.
I am a recent graduate (April 2021) in Computer Science in Software Engineering track from King Abdul-Aziz University,
and I am eager to enter the workforce. 

My cover letter and resume are attached for your review.
If you would like more information regarding my qualifications for the job, please do not hesitate to reach out. 

Thank you so much for your time and consideration, and I looking forward to hearing from you soon.

Sincerely,


Mohammed Saleh AlGhanmi

Jeddah | +966-500977323 | Linkedin: https://www.linkedin.com/in/mohammed-s-alghanmi/ | GitHub: https://github.com/M7MDSG76
        '''
        
        
        """
    //// My Email attachments 
         # Email Subject
        email_subject = 'Computer Science Graduate Seeking for a Job'
        
        # Email letter
        email_message = '''
Dear Hiring Manager, 


I am Mohammed Saleh Alghanmi, and I am excited to apply for a job at your company.
I am a recent graduate (April 2021) in Computer Science in Software Engineering track from King Abdul-Aziz University,
and I am eager to enter the workforce. 

My cover letter and resume are attached for your review.
If you would like more information regarding my qualifications for the job, please do not hesitate to reach out. 

Thank you so much for your time and consideration, and I looking forward to hearing from you soon.

Sincerely,


Mohammed Saleh AlGhanmi

Jeddah | +966-500977323 | Linkedin: https://www.linkedin.com/in/mohammed-s-alghanmi/ | GitHub: https://github.com/M7MDSG76
        '''
        """
        
        # The CV file name without the path (it should be in the same as the project file path)
        CV_file_name = 'Mohammed_Saleh_Alghanmi_CV.pdf'
        
        # The Cover letter file name without the path (it should be in the same as the project file path)
        coverletter_file_name = 'Mohammed-S-Alghanmi_Coverletter.docx'
        
        
        # Start smtp session.
        session = start_smtp_gmail_session(sender_address, sender_pass)
        
        # Iterate over the recruiters and send email for each.
        for recruiter in range(len(recruiters_list)):
            
            # declare a message object and attach every thing.
            message = create_message(sender_address, recruiters_list[recruiter],
                                     email_subject, email_message,
                                     CV_file_name, coverletter_file_name)
            try:
                # Send the message over the smtp session to the recruiter address.
                session.sendmail(sender_address, recruiters_list[recruiter], message)
                print(f'Email sent successfully to {recruiters_list[recruiter]}')
                
            # If Error accure make an exception.
            except:
                print(f'Email: {recruiters_list[recruiter]} is invalid.')    
                print("Oops!", sys.exc_info()[0], "occurred.")
                print("Next entry.\n")
                print()  
        session.quit()    
        
if __name__ == '__main__':

        main()