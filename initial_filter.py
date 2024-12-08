import re
from dns import resolver
import pandas as pd
import re

resolver.default_resolver = resolver.Resolver()
resolver.default_resolver.nameservers = ['1.1.1.1', '8.8.8.8']

def validate_email_syntax(email):
    """
    Validate the syntax of the email address.
    Returns True if valid, False otherwise.
    """
    regex = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    return re.match(regex, email) is not None

def validate_domain(email):
    try:
        domain = email.split('@')[-1]
        # First, try MX records
        mx_records = resolver.resolve(domain, 'MX')
        return True
    except (resolver.NoAnswer, resolver.NXDOMAIN):
        # Fallback to A record if no MX records exist
        try:
            a_records = resolver.resolve(domain, 'A')
            return True
        except resolver.NoAnswer:
            return False
    except Exception as e:
        print(f"DNS lookup failed for {domain}: {e}")
        return False

def validate_email(email):
    """
    Combine all validation steps.
    Returns a tuple: (is_valid, reason)
    """
    if not validate_email_syntax(email):
        return False, "Invalid syntax"
    if not validate_domain(email):
        return False, "Invalid domain"
    return True, "Valid email"

def process_email_validation(input_file, output_good_file, output_bounced_file):
    """
    Process the email validation for a given Excel file.
    Splits rows into 'good emails' and 'bounced emails' files.
    """
    # Read the input Excel file
    df = pd.read_excel(input_file)

    # Initialize lists for results
    good_rows = []
    bounced_rows = []

    for index, row in df.iterrows():
        try:
            email = str(row['Email']).strip()
            print(f"Processing row {index + 1}: {email}")
            
            # Validate the email
            is_valid, reason = validate_email(email)
            
            # Append the row to the appropriate list
            if is_valid:
                good_rows.append(row)
            else:
                row['Reason'] = reason  # Add a reason column to the bounced rows
                bounced_rows.append(row)

            print(f"Processed {index + 1}/{len(df)}: {email} - {reason}")
        except Exception as e:
            # Log the issue and skip the problematic row
            print(f"Error processing row {index + 1} ({email}): {e}")
            row['Reason'] = f"Error: {e}"
            bounced_rows.append(row)  # Optionally add problematic rows to bounced list

    # Convert the results to DataFrames
    good_df = pd.DataFrame(good_rows)
    bounced_df = pd.DataFrame(bounced_rows)

    # Write to new Excel files
    good_df.to_excel(output_good_file, index=False)
    bounced_df.to_excel(output_bounced_file, index=False)


# Input and output file paths
input_file = r"C:\\Users\\Abdullah Atif\\Desktop\\Email marketing tool - Trucking Story\\non-py files\\filtered_power_unit_5.xls"
output_good_file = r"C:\\Users\\Abdullah Atif\\Desktop\\Email marketing tool - Trucking Story\\non-py files\\good_emails.xlsx"
output_bounced_file = r"C:\\Users\\Abdullah Atif\\Desktop\\Email marketing tool - Trucking Story\\non-py files\\initial_filtered_out.xlsx"


# Run the email validation
process_email_validation(input_file, output_good_file, output_bounced_file)