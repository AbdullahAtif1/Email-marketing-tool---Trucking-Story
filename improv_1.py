import smtplib, ssl
import pandas as pd
import re
import concurrent.futures
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


print_lock = threading.Lock()
port = 465
smtp_server = "smtp.hostinger.com"
context = ssl.create_default_context()

subject = "Exclusive Pre-Launch Offer - Save Big Before January 1st!"
msg = """

Dear {legal_name},

We're excited to announce the pre-launch of Trucking Story, your one-stop solution for starting a successful trucking business! Our website officially launches on January 1st, 2025, but we're offering an exclusive pre-launch deal you don't want to miss.

Here's what you'll get with our Startup Package:
✅ Professional Website tailored for your trucking business.
✅ Professional Email to build trust with clients.
✅ Custom Logo Design for your unique identity.
✅ Marketing Materials including flyers and business cards.
✅ Niche Selection Guidance to ensure a solid start.
✅ 1 Year Free Hosting for your website.
✅ 3 Months of Free Agent Support for personalized assistance.
✅ 1 Week of Free Dispatch Service to help you get loads.

Our team will guide you on:

How to use your website effectively.
How to find and connect with shippers.
Pre-Launch Pricing - Just $1500
Take advantage of our special pre-launch offer for $1500 (in 3 easy installments):

$500 to get started - Covers website creation, email setup, and marketing materials.
$500 after your website is live and ready.
$500 after one month of service.
Act Now and Save $3500!
After January 1st, 2025, the same package will be priced at $5000. Don't miss this chance to save big and launch your trucking business with confidence.

Money-Back Guarantee
We're confident in our services. If you're not satisfied within the first month, we'll give you a full refund.

If you're interested, simply reply to this email with “Interested”, and we'll get in touch with you right away.

Let's build your trucking business together!

Best regards,
The Trucking Story Team
"""

# Define the list of accounts and their account specific data
accounts = [
    {
        'sender_email': 'Jordan@truckingstory.com',
				# 'file_name': 'account1.xlsx',
				'file_name': '4_12_24.xlsx',
        'password': 'Ahmad@2134',
        'subject': subject,
        'msg': msg
    },
    # {
    #     'sender_email': 'sales@truckingstory.com',
		# 		# 'file_name': 'account2.xlsx',
		# 		'file_name': 'testing.xlsx',
    #     'password': 'Ahmad@2134',
    #     'subject': subject,
    #     'msg': msg
    # }
    # {
    #     'sender_email': 'account3@truckingstory.com',
		# 		'file_name': 'whatever_name',
    #     'password': 'Ahmad@2134',
    #     'subject': subject,
    #     'msg': msg
    # },
    # {
    #     'sender_email': 'account4@truckingstory.com',
		# 		'file_name': 'whatever_name',
    #     'password': 'Ahmad@2134',
    #     'subject': subject,
    #     'msg': msg
    # }
]

# count
# mc number
# time start end

def send_email_func(file_name, sender_email, password, smtp_server, port, subject_template, msg_template):
	
	file_path = f"C:\\Users\\Abdullah Atif\\Desktop\\Email marketing tool - Trucking Story\\{file_name}"
	df = pd.read_excel(file_path)

	with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
		server.login(sender_email, password)

		for index, row in df.iterrows():

			if row['PowerUnits'] < 5:
				continue

			mc_number = re.search(r'\d+', row['MCNumber']).group()

			# Format the subject and message with data from the Excel row
			subject = subject_template.format(mc_number=mc_number, legal_name=row.get('LegalName', 'Valued Customer'))
			msg = msg_template.format(mc_number=mc_number, legal_name=row.get('LegalName', 'Valued Customer'))

			
			email_message = MIMEMultipart()
			email_message['From'] = sender_email

			recipient_email = str(row['Email']).strip()

			email_message['To'] = recipient_email
			email_message['Subject'] = subject

      # Attach the message body
			email_message.attach(MIMEText(msg, 'plain', 'utf-8'))

			try:
				server.sendmail(sender_email, recipient_email, email_message.as_string()) # Message is rendered using as_string ()
			except Exception as e:
				print(f"Error sending email to {recipient_email}: {e}")
				continue

			with print_lock:
				print(f"Processing row {index + 1}: Email sent to {recipient_email}")

		server.sendmail(sender_email, "pyabdpy@gmail.com", email_message.as_string())

# Define a function to call send_email_func for each account
def send_email_for_account(account):
    send_email_func(
        account['file_name'], 
        account['sender_email'], 
        account['password'], 
        smtp_server, 
        port, 
        account['subject'], 
        account['msg']
    )

# Create a thread pool with 4 threads
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
    # Submit the send_email_for_account function for each account
    futures = [
				executor.submit(send_email_for_account, account)
				for account in accounts
		]

    # Wait for all threads to finish
    for future in concurrent.futures.as_completed(futures):
        future.result()

