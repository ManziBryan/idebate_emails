from __future__ import print_function
import os.path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
import mimetypes
import base64
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import csv
import openpyxl
from pathlib import Path
import socket
import email.encoders

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://mail.google.com/']
# sender = 'jhabineza@debaterwanda.org'
sender = "brnmanzi130@gmail.com"
# to = 'manzi.k.bryan@gmail.com'
subject = 'Hold an online debate event with Rwanda'
message_text = """
<br>

I hope this email finds you well and keeping safe.<br><br>

<p>My name is Jean Michel Habineza; I’m the founder and director of iDebate Rwanda, a debate program that works with students from Rwanda and East Africa as a whole. </p>

<p>Since 2014, iDebate Rwanda has been doing a 3-months tour in the USA. A tour in which we speak about the story of Rwanda, <b>a tale of the rebirth</b> of a Nation after enduring years of ideology, propaganda and polarization, we talk about how dangerous all that is and why we need civil discourse. From 2014-2018 we have had the chance of speaking to more than <b>50,000 students</b> from more than 60 universities and high schools  in the United States.  <a href=https://youtu.be/05x7et3gusA> click here </a>to learn more about the 2018 US Tour.</p>

<p>In 2020, the world was hit by a global pandemic, and everything had to pause. After months of preparing yet another tour, we had to cancel to protect our students from the COVID19. However, last year, we were presented with an <b>opportunity</b> that allowed us to bring these discussions to the table with university and high school students from the USA.</p>

<p>In the fall of 2020, we launched the <b>Global Debate series</b> —a series of events where we host debates and Public presentations(zoom) with Universities and High schools.  Events aimed to raise awareness about the 1994 Genocide against Tutsi in Rwanda and the <b>importance of public discourse</b> and serve as a way to expose our students to the global debate world. These are events we created to bring our discussions online, and we thought this would be a great opportunity for us to host events with different <b>people from all over the world</b>.<p>

<p>Over the past year, we have had the honour to host events with 11 universities in the USA. We are currently looking into hosting this kind of event with high schools as well and we would <b>love to host an event with your school</b>. These events can be in the form of a debate, panel discussion, cultural presentations.  <a href=https://youtu.be/3HSvP0zE6rA> click here </a>to watch some of our global debate series.</p>

<p>We also have the opportunity to create more significant events with other African countries if you would find interest in that.</p>

<p>These events would also work to raise funds that allow us to continue doing the <b>work of civil discourse</b> around the world and in Rwanda.</p>

<p>We would love to do this with you. Would you be interested to do this, please find the event proposal here.</p>


Best Regards,<br><br>

Jean Michel Habineza<br>
Founder& Managing Director<br>
iDebate Rwanda"""


directory = '/Users/manzi/Downloads/Proposal.pdf'
def create_message(sender, to, subject, message_text):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  
  message = MIMEText(message_text, 'html')
  message['to'] = to
  # message['from'] = sender
  message['subject'] = subject
  return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}

def create_message_with_attachment(sender, recipient, subject, message_text, file, name):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file: The path to the file to be attached.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEMultipart()
  message['to'] = recipient
  # message['from'] = sender
  message['subject'] = subject
  with_name = 'Dear '
  with_name += str.title(name)
  with_name += ', <br>'
  with_name += message_text

  msg = MIMEText(with_name, 'html')
  message.attach(msg)

  content_type, encoding = mimetypes.guess_type(file)

  if content_type is None or encoding is not None:
    content_type = 'application/octet-stream'
  main_type, sub_type = content_type.split('/', 1)
  if main_type == 'text':
    fp = open(file, 'rb')
    msg = MIMEText(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'image':
    fp = open(file, 'rb')
    msg = MIMEImage(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'audio':
    fp = open(file, 'rb')
    msg = MIMEAudio(fp.read(), _subtype=sub_type)
    fp.close()
  else:
    fp = open(file, 'rb')
    msg = MIMEBase(main_type, sub_type)
    msg.set_payload(fp.read())
    fp.close()
  filename = os.path.basename(file)
  msg.add_header('Content-Disposition', 'attachment', filename=filename)
  email.encoders.encode_base64(msg)
  message.attach(msg)

  return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}



def main(name, recipient, row):
    print("sending to ", name, recipient)
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # msg = create_message(sender, to, subject, message_text)
    msg = create_message_with_attachment(sender, recipient, subject, message_text, directory, name)

    service = build('gmail', 'v1', credentials=creds)

    try:
      message = (service.users().messages().send(userId='me', body=msg)
                .execute())

      # Call the Gmail API
      results = service.users().labels().list(userId='me').execute()
      labels = results.get('labels', [])
    except:
      print(row)


# if __name__ == '__main__':
#   count = 1
#   maximum = 450
#   with open('emails.csv', 'r', encoding='utf-8-sig') as file:
#       reader = csv.reader(file)
#       for row in reader:
        
#         if len(row) > 0 and len(row[0]) > 6 and count < maximum:
#           if len(row) == 1:
#             recipient = row[0]
#             name = 'Sir / Madam'
#           else:
#             name = row[0]
#             recipient = row[1]
#           main(name, recipient, row)
#           count += 1

    
if __name__ == '__main__':

  loc = ("emails.xlsx")

  xlsx_file = Path('', 'emails.xlsx')
  sheet = openpyxl.load_workbook(xlsx_file).active
  name = ''
  destination = ''
  count = 1
  maximum = 450
  for row in sheet.iter_rows():
    if not row[1].value:
      continue
    destination = row[1].value
    if not (row[0].value): # We don't have a name
      name = 'Sir/Madam'
      print("no name", row)
    else:
      name = row[0].value
    main(name, 'brnmanzi130@gmail.com', name + " " + 'brnmanzi130@gmail.com')
    count += 1

      
  # count = 1
  # maximum = 450
  # with open('emails.csv', 'r', encoding='utf-8-sig') as file:
  #     reader = csv.reader(file)
  #     for row in reader:
        
  #       if len(row) > 0 and len(row[0]) > 6 and count < maximum:
  #         if len(row) == 1:
  #           recipient = row[0]
  #           name = 'Sir / Madam'
  #         else:
  #           name = row[0]
  #           recipient = row[1]
  #         main(name, recipient, row)
  #         count += 1

    