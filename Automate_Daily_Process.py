import os
import win32com.client
import pandas as pd
from datetime import datetime

# Connects to the Outlook application and retrieves emails from the inbox
def connect_to_outlook():
    try:
        # Create an instance of the Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        # Get the inbox folder (6 refers to the inbox)
        inbox = namespace.GetDefaultFolder(6)
        messages = inbox.Items
        # Sort messages by received time, latest first
        messages.Sort("[ReceivedTime]", True)

        print("Connection to Outlook successful.")
        return messages
    except Exception as e:
        print(f"Failed to connect to Outlook: {e}")
        return None

# Downloads the latest Excel file from the email with the specified subject keyword
def download_latest_excel(subject_keyword, download_folder):
    messages = connect_to_outlook()
    if not messages:
        return None, None

    # Loop through each message in the inbox
    for message in messages:
        # Check if the subject of the message contains the specified keyword
        if subject_keyword in message.Subject:
            attachments = message.Attachments
            # Loop through each attachment in the message
            for attachment in attachments:
                if attachment.FileName.endswith(".xlsx"):
                    # Create the download folder if it doesn't exist
                    if not os.path.exists(download_folder):
                        os.makedirs(download_folder)
                    file_path = os.path.join(download_folder, attachment.FileName)
                    
                    try:
                        # Remove the existing file if it exists
                        if os.path.exists(file_path):
                            os.remove(file_path)
                        # Save the attachment to the specified file path
                        attachment.SaveAsFile(file_path)
                        print(f"Attachment {attachment.FileName} saved to {file_path}")
                        return file_path, message
                    except Exception as e:
                        print(f"Failed to save the attachment: {e}")
            break
    print("No email found with the specified subject or no Excel attachment found.")
    return None, None

# Validates the contents of the Excel file
def validate_excel(file_path):
    try:
        # Load the Excel file
        xls = pd.ExcelFile(file_path)
        df_events_parts = pd.read_excel(xls, 'Events_Parts')
        df_general_events = pd.read_excel(xls, 'General_Events')

        # Validation functions for various fields
        def validate_site_type(site_type):
            valid_values = [
                'Assembly Site', 'Fabrication Site', 'Material Location',
                'Final Test Site', '-', 'Shipping Site'
            ]
            return site_type in valid_values

        def validate_country(country):
            valid_values = [
                'Possible Impacted', 'Impacted',
                'Impact is being evaluated', 'Not Impacted'
            ]
            return country in valid_values

        def validate_url(url):
            return url.startswith('http://')

        def validate_date(date):
            if date == 'TBD':
                return True
            try:
                pd.to_datetime(date, format='%Y-%m-%d', errors='raise')
                return True
            except ValueError:
                return False

        def validate_threat_level(threat_level):
            valid_values = ['Minor', 'Moderate', 'Critical']
            return threat_level in valid_values
        
        def validate_event_type(event_type):
            valid_values = [
                'Market Insights', 'Earthquakes', 'Factory Fires', 'Business Withdrawal/Closure', 'Tariffs and Customs',
                'Typhoons', 'Floods', 'Droughts', 'Power Outages', 'Economic', 'Social', 'Political Situations',
                'Global Pandemic', 'Ports Disruptions', 'Business Expansions', 'Cyber Attacks', 'Civil Unrests',
                'Industrial Disputes', 'Wildfires', 'Product Shortages', 'Cyclones', 'Price Fluctuations',
                'Mergers and Acquisitions', 'Air Pollutions', 'Raw Material Shortages', 'Factory Explosions',
                'Mines Shut Down', 'Volcanoes', 'Health and Safety', 'Lead Time Variability', 'Storms',
                'Management Board updates', 'Chemical Accidents', 'Sanctions', 'Military Disputes', 'Extreme Weather',
                'Avalanches', 'Diseases', 'Shortages and Allocation', 'Stock Status/Forecasting', 'Lawsuits',
                'Dam burst', 'Supply/Demand Statement', 'Employees Layoff', 'Terrorist Acts', 'Business Relocation',
                'Partnerships', 'Patents and Copyrights', 'Explosions', 'Shipping Disruption', 'Landslides',
                'Bankruptcy', 'Spin Off', 'Product failure', 'Cyber Risks', 'Product Recall', 'Tsunamis',
                'Infrastructure Disruptions'
            ]
            return event_type in valid_values

        def validate_event_scope(event_scope):
            valid_values = [
                'City', 'Technical Feature', 'Technology node', 'Wafer Size', 'Process Technology', 'Supplier',
                'Region', 'State', 'Facility', 'Country', 'Raw Material', 'Series', 'Product line', 'Technical Features',
                'Site Type', 'Wafer Material', 'Company'
            ]
            return event_scope in valid_values

        # Check for issues in the 'Events_Parts' sheet
        def check_events_parts(df):
            issues = []
            for idx, row in df.iterrows():
                if not validate_site_type(row['Site Type']):
                    issues.append((idx, 'Site Type', row['Site Type']))
                if not validate_country(row['Impact Status']):
                    issues.append((idx, 'Impact Status', row['Impact Status']))
                if not validate_url(row['Event News URL']):
                    issues.append((idx, 'Event News URL', row['Event News URL']))
                if not validate_date(row['Event News Date']):
                    issues.append((idx, 'Event News Date', row['Event News Date']))
                if not validate_date(row['Event Start Date']):
                    issues.append((idx, 'Event Start Date', row['Event Start Date']))
                if not validate_date(row['Event End Date']):
                    issues.append((idx, 'Event End Date', row['Event End Date']))
                if not validate_threat_level(row['Event Threat Level']):
                    issues.append((idx, 'Event Threat Level', row['Event Threat Level']))
                if not validate_event_type(row['Event Type']):
                    issues.append((idx, 'Event Type', row['Event Type']))
                if not validate_event_scope(row['Event Scope']):
                    issues.append((idx, 'Event Scope', row['Event Scope']))
            return issues

        # Check for issues in the 'General_Events' sheet
        def check_general_events(df):
            issues = []
            for idx, row in df.iterrows():
                if not validate_site_type(row['Site Type']):
                    issues.append((idx, 'Site Type', row['Site Type']))
                if not validate_country(row['Impact Status']):
                    issues.append((idx, 'Impact Status', row['Impact Status']))
                if not validate_url(row['Event News URL']):
                    issues.append((idx, 'Event News URL', row['Event News URL']))
                if not validate_date(row['Event News Date']):
                    issues.append((idx, 'Event News Date', row['Event News Date']))
                if not validate_date(row['Event Start Date']):
                    issues.append((idx, 'Event Start Date', row['Event Start Date']))
                if not validate_date(row['Event End Date']):
                    issues.append((idx, 'Event End Date', row['Event End Date']))
                if not validate_threat_level(row['Event Threat Level']):
                    issues.append((idx, 'Event Threat Level', row['Event Threat Level']))
                if not validate_event_type(row['Event Type']):
                    issues.append((idx, 'Event Type', row['Event Type']))
                if not validate_event_scope(row['Event Scope']):
                    issues.append((idx, 'Event Scope', row['Event Scope']))
            return issues

        # Combine issues from both sheets
        issues_events_parts = check_events_parts(df_events_parts)
        issues_general_events = check_general_events(df_general_events)

        # Add the new condition to compare rows starting from column AL
        issues = issues_events_parts + issues_general_events
        for idx, row in df_events_parts.iterrows():
            matching_rows = df_general_events[df_general_events['AL'] == row['AL']]
            if matching_rows.empty:
                issues.append((idx, 'AL', row['AL']))  # Mark as issue if not found in General_Events

        return issues

    except Exception as e:
        print(f"validate the Excel file: {e}")
        return []

# Saves identified issues to a new Excel file
def save_issues_to_excel(issues, download_folder):
    today_str = datetime.today().strftime('%Y%m%d')
    issues_file_path = os.path.join(download_folder, f"suspected_issues_{today_str}.xlsx")

    # Create a DataFrame from the issues list
    issues_df = pd.DataFrame(issues, columns=['Row', 'Column', 'Value'])
    with pd.ExcelWriter(issues_file_path, engine='xlsxwriter') as writer:
        issues_df.to_excel(writer, sheet_name='Issues', index=False)
    
    print(f"Issues saved to {issues_file_path}")

# Sends an approval response email
def send_approval_response(email_message):
    outlook = win32com.client.Dispatch("Outlook.Application")
    response = email_message.Reply()
    response.Subject = "Re: " + email_message.Subject
    response.Body =

 """
Dear,

File approved with no issues.

Thanks,
Almohtadey Metwaly
Engineer II Quality Department

"""
    # Add CC recipients to the response email
    response.CC = ";".join([
        "eslam_ghanem@siliconexpert.com",
        "khaledi@siliconexpert.com",
        "mohammad_dawah@siliconexpert.com",
        "omar_seliem@siliconexpert.com",
        "mohammad_gouda@siliconexpert.com",
        "elsaid_hussein@siliconexpert.com",
        "hamed_abdalhameed@siliconexpert.com",
        "abdelmoemen_agha@siliconexpert.com",
        "elhussein_sobhy@siliconexpert.com",
        "sara_essam@siliconexpert.com",
        "hany_samy@siliconexpert.com",
        "mai_mahmoud@siliconexpert.com"
    ])
    response.Send()
    print("Approval response sent.")

if __name__ == "__main__":
    subject_keyword = "Nokia Event Delivery"
    download_folder = r"C:\Users\145989\Downloads\automated nokia"
    # Download the latest Excel file with the specified subject keyword
    downloaded_file_path, email_message = download_latest_excel(subject_keyword, download_folder)
    
    if downloaded_file_path:
        print(f"Downloaded file path: {downloaded_file_path}")
        # Validate the downloaded Excel file
        issues = validate_excel(downloaded_file_path)
        if not issues:
            print("File approved with no issues.")
            # Send approval response if no issues are found
            send_approval_response(email_message)
        else:
            print("Issues found in the Excel file. Saving to a new Excel file.")
            # Save the issues to a new Excel file
            save_issues_to_excel(issues, download_folder)
    else:
        print("No file downloaded.")
