import smtplib, ssl
import pandas as pd
import email.message
import re

port = 465
password = "Ahmad@2134"
sender_email = "Jordan@truckingstory.com"
smtp_server = "smtp.hostinger.com"

# Create a secure SSL context
context = ssl.create_default_context()


# Fetch Target emails
file_path = "C:\\Users\\Abdullah Atif\\Desktop\\Email marketing tool - Trucking Story\\15_16.xlsx"
df = pd.read_excel(file_path)


with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
		server.login(sender_email, password)

		for index, row in df.iterrows():

			mc_number = re.search(r'\d+', row['MC Number']).group()
			m = email.message.Message()
			m['From'] = sender_email
			m['To'] = row['Emails']
			m['Subject'] = f"Congratulations, your MC Number {mc_number} is active. One final step required"

			m.set_payload(
				f"Hello {row['Legal Name']}, we are very pleased to have you on board and are the first in line to cheer you and support you on your way to build a successful and ever-lasting business."
			)

			server.sendmail(sender_email, row['Emails'], m.as_string()) # Message is rendered using as_string ()
			print(f"Processing row {index + 1}: Email sent to {row['Emails']}")



