from email.mime.text import MIMEText
import base64
from googleapiclient.discovery import build
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = ['https://mail.google.com/']

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
  message = MIMEText(message_text)
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}




def getCreds():
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
    return creds


def send_message(service, user_id, message):
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
  try:
    message = (service.users().messages().send(userId=user_id, body='message')
               .execute())
    print('Message Id: %s' % message['id'])
    return message
  except:
   print (f'An error occurred:')
   return None

    
sender = 'brnmanzi130@gmail.com'
to = 'bryan.manzi.21@dartmouth.edu'
subject = 'Checking'
message_text = 'Hi, this is Manzi and I wanted to test out this emailing ability. To unsubscribe, click on this link https://idebate.com/unsubscribe'

msg = create_message(sender, to, subject, message_text)
creds = getCreds()
service = build('gmail', 'v1', credentials=creds)
send_message(service, 'me', msg)



