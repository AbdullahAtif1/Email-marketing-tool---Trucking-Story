import re
import smtplib, ssl
import pandas as pd
import re
import concurrent.futures
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import time

# Record the start time
start_time = time.time()
start_time_formatted = datetime.now().strftime("%H:%M")


print_lock = threading.Lock()
port = 465
smtp_server = "smtp.hostinger.com"
context = ssl.create_default_context()

subject = "Exclusive Pre-Launch Offer - Save Big Before January 1st!"
msg = """

<!DOCTYPE html>
<html>
<body>
  <p><strong>Hi {legal_name},</strong></p>
  
  <p>We're giving away <strong>a free website</strong> to a few lucky trucking companies! 🚛</p>
  
  <p>A website helps you:</p>
  <ul>
    <li>Get direct freight & avoid expensive load boards</li>
    <li>Look professional & build trust with brokers</li>
    <li>Stand out from other carriers</li>
  </ul>
  
  <p><strong>How to enter?</strong> Just reply to this email with your <strong>truck type</strong> (Box Truck, Dry Van, Reefer, etc.), and you're in! 🎉</p>
  
  <p>Winners will be announced in <strong>7 days</strong>. Don't miss out! 🕒</p>
  
  <p><strong>Reply now with your truck type to enter the draw!</strong></p>
  
  <p>Best,<br>
  [Your Name]<br>
  Trucking Story</p>
</body>
</html>


"""

# Define the list of accounts and their account specific data
accounts = [
    {
        'sender_email': 'Jordan@truckingstory.com',
				# 'file_name': 'account1.xlsx',
				'file_name': 'testing.xlsx',
        'password': 'Ahmad@2134',
        'subject': subject,
        'msg': msg
    }
]


def send_email_func(file_name, sender_email, password, smtp_server, port, subject_template, msg_template):
	
	file_path = f"C:\\Users\\Abdullah Atif\\Desktop\\Email marketing tool - Trucking Story\\non-py files\\{file_name}"
	df = pd.read_excel(file_path)

	legal_name_col = 'LegalName' if 'LegalName' in df.columns else 'Legal Name'
	mc_number_col = 'MCNumber' if 'MCNumber' in df.columns else 'MC Number'

	with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
		server.login(sender_email, password)

		for index, row in df.iterrows():
			# mc_number = re.search(r'\d+', row[mc_number_col]).group()

			# # Format the subject and message with data from the Excel row
			# subject = subject_template.format(mc_number=mc_number, legal_name=row.get(legal_name_col, 'Valued Customer'))
			# msg = msg_template.format(mc_number=mc_number, legal_name=row.get(legal_name_col, 'Valued Customer'))

			mc_number_value = row.get(mc_number_col, '')
			match = re.search(r'\d+', str(mc_number_value))  # Ensure value is a string for regex
			mc_number = match.group() if match else 'Unknown'
			legal_name = row.get(legal_name_col, 'Valued Customer')

			# Format the subject and message
			subject = subject_template.format(mc_number=mc_number, legal_name=legal_name)
			msg = msg_template.format(mc_number=mc_number, legal_name=legal_name)

			email_message = MIMEMultipart()
			email_message['From'] = sender_email

			recipient_email = str(row['Email']).strip()

			email_message['To'] = recipient_email
			email_message['Subject'] = subject

      # Attach the message body
			email_message.attach(MIMEText(msg, 'plain', 'utf-8'))

			try:
					response = server.sendmail(sender_email, recipient_email, email_message.as_string())  # Attempt to send the email
					print(f"Email sent to {index + 1}/{len(df)}: {recipient_email} - MC Number {mc_number}")

			except Exception as e:
					print(f"Error sending email to {recipient_email}: {e}\n")
					continue

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


# Record the end time
end_time = time.time()
end_time_formatted = datetime.now().strftime("%H:%M") 

# Calculate the total time taken in hours
total_time_in_hours = (end_time - start_time) / 3600

# Print the final output
print(f"Start time: {start_time_formatted}")
print(f"End time: {end_time_formatted}")
print(f"Total time taken: {total_time_in_hours:.2f} hours")


# TNow can handle html messages
