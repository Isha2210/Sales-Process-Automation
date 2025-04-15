# Task 4: Automating the Email Campaign
# This script sends personalized emails to leads from the Excel file
# It includes tracking functionality and rate limiting to prevent being flagged as spam

import os
import pandas as pd
import smtplib
import time
import random
import uuid
import logging
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from jinja2 import Template
from urllib.parse import quote
from dotenv import load_dotenv

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_campaign.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Load environment variables from .env file (for email credentials)
load_dotenv()


class EmailCampaign:
    def __init__(self, excel_path=None):
        """
        Initialize the email campaign with configuration and templates

        Parameters:
        -----------
        excel_path : str, optional
            Path to the Excel file with lead data. If None, will find the most recent one.
        """
        # Campaign settings
        self.sender_email = os.getenv('EMAIL_ADDRESS', 'your_email@gmail.com')
        self.sender_password = os.getenv('EMAIL_PASSWORD', 'your_password')
        self.sender_name = os.getenv('SENDER_NAME', 'Your Name')
        self.sender_company = os.getenv('SENDER_COMPANY', 'Your Company')
        self.sender_position = os.getenv('SENDER_POSITION', 'Your Position')
        self.sender_phone = os.getenv('SENDER_PHONE', 'Your Phone')
        self.base_url = os.getenv('TRACKING_URL', 'https://yourwebsite.com/track')
        self.batch_size = int(os.getenv('BATCH_SIZE', 10))
        self.delay_min = int(os.getenv('DELAY_MIN', 60))  # Minimum delay in seconds
        self.delay_max = int(os.getenv('DELAY_MAX', 180))  # Maximum delay in seconds
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', 587))

        # Load lead data
        self.excel_path = self._find_excel_file() if excel_path is None else excel_path
        self.leads_df = self._load_leads()

        # Load email templates
        self.email_template = self._load_email_template()
        self.subject_template = self._load_subject_template()

        # Create/load campaign tracking data
        self.campaign_id = datetime.now().strftime('%Y%m%d%H%M%S')
        self.tracking_data = {}
        self.tracking_file = f'campaign_data_{self.campaign_id}.json'

        logger.info(f"Email campaign initialized with ID: {self.campaign_id}")
        logger.info(f"Found {len(self.leads_df)} leads to process")

    def _find_excel_file(self):
        """Find the most recent Excel file in the output directory"""
        output_dir = 'output'
        if os.path.exists(output_dir):
            excel_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx')]
            if excel_files:
                excel_files.sort(key=lambda x: os.path.getmtime(os.path.join(output_dir, x)), reverse=True)
                return os.path.join(output_dir, excel_files[0])

        # Check for last export path
        last_export_path_file = os.path.join('output', 'last_export_path.txt')
        if os.path.exists(last_export_path_file):
            with open(last_export_path_file, 'r') as f:
                excel_path = f.read().strip()
                if os.path.exists(excel_path):
                    return excel_path

        raise FileNotFoundError("No Excel file with lead data found. Run Task 2 first.")

    def _load_leads(self):
        """Load the leads from the Excel file"""
        try:
            df = pd.read_excel(self.excel_path)
            logger.info(f"Successfully loaded {len(df)} leads from {self.excel_path}")
            return df
        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            raise

    def _load_email_template(self):
        """Load the HTML email template"""
        template_path = os.path.join('templates', 'email_template.html')
        try:
            with open(template_path, 'r') as f:
                return Template(f.read())
        except FileNotFoundError:
            logger.warning("Email template not found. Will create a new one.")
            from task3_email_template import create_email_template
            return create_email_template()

    def _load_subject_template(self):
        """Load the subject line template"""
        template_path = os.path.join('templates', 'subject_template.txt')
        try:
            with open(template_path, 'r') as f:
                return Template(f.read())
        except FileNotFoundError:
            logger.warning("Subject template not found. Will create a new one.")
            from task3_email_template import create_subject_line_template
            return create_subject_line_template()

    def _generate_tracking_id(self, lead_id):
        """Generate a unique tracking ID for a particular lead"""
        return f"{self.campaign_id}_{lead_id}_{uuid.uuid4().hex[:8]}"

    def _add_tracking_to_html(self, html_content, tracking_id):
        """Add tracking pixel and unique IDs to links in the HTML"""
        # Add tracking pixel
        tracking_pixel = f'<img src="{self.base_url}/pixel/{tracking_id}" width="1" height="1" alt="" style="display:none;">'
        html_content = html_content.replace('</body>', f'{tracking_pixel}</body>')

        # Add tracking to CTA link
        if 'class="cta-button"' in html_content:
            # Extract the original href
            import re
            cta_match = re.search(r'<a href="([^"]+)" class="cta-button"', html_content)
            if cta_match:
                original_url = cta_match.group(1)
                # Create a tracking URL that redirects to the original
                tracking_url = f"{self.base_url}/click/{tracking_id}?url={quote(original_url)}"
                # Replace the original URL with the tracking URL
                html_content = html_content.replace(f'href="{original_url}"', f'href="{tracking_url}"')

        return html_content

    def _prepare_email(self, lead):
        """Prepare a personalized email for a single lead"""
        lead_id = lead.get('id', str(lead.name))
        tracking_id = self._generate_tracking_id(lead_id)

        # Extract the recipient's first name
        if 'Contact Person' in lead and lead['Contact Person']:
            full_name = lead['Contact Person']
            recipient_name = full_name.split()[0]  # Extract first name
        else:
            recipient_name = "there"  # Default if no name available

        # Prepare the context for template rendering
        context = {
            'recipient_name': recipient_name,
            'company_name': lead.get('Company Name', 'your company'),
            'industry': lead.get('Industry', 'your industry'),
            'location': lead.get('Location', 'your area')
        }

        # Create subject and HTML content
        subject = self.subject_template.render(context)
        html_content = self.email_template.render(context)

        # Add tracking elements to the HTML
        tracked_html = self._add_tracking_to_html(html_content, tracking_id)

        # Store tracking data
        self.tracking_data[tracking_id] = {
            'lead_id': lead_id,
            'company': lead.get('Company Name', 'Unknown'),
            'contact': lead.get('Contact Person', 'Unknown'),
            'email': lead.get('Email', 'Unknown'),
            'sent_time': None,
            'opened': False,
            'clicked': False,
            'responded': False,
            'last_activity': None
        }

        return subject, tracked_html, tracking_id

    def _send_email(self, recipient_email, subject, html_content):
        """Send an email to a single recipient"""
        try:
            # Create message container
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = f"{self.sender_name} <{self.sender_email}>"
            msg['To'] = recipient_email

            # Create HTML part
            html_part = MIMEText(html_content, 'html')
            msg.attach(html_part)

            # Connect to SMTP server and send email
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.send_message(msg)

            return True
        except Exception as e:
            logger.error(f"Error sending email to {recipient_email}: {e}")
            return False

    def _save_tracking_data(self):
        """Save the tracking data to a JSON file"""
        with open(self.tracking_file, 'w') as f:
            json.dump(self.tracking_data, f, indent=4)
        logger.info(f"Tracking data saved to {self.tracking_file}")

    def run_campaign(self):
        """Run the full email campaign, sending emails to all leads"""
        logger.info(f"Starting email campaign with {len(self.leads_df)} leads")

        # Create directory for sent emails if it doesn't exist
        sent_dir = 'sent_emails'
        if not os.path.exists(sent_dir):
            os.makedirs(sent_dir)

        # Process leads in batches
        total_leads = len(self.leads_df)
        successful_sends = 0
        failed_sends = 0

        for i, (_, lead) in enumerate(self.leads_df.iterrows()):
            # Skip leads without email
            if 'Email' not in lead or not lead['Email'] or pd.isna(lead['Email']):
                logger.warning(
                    f"Lead {i + 1}/{total_leads} ({lead.get('Company Name', 'Unknown')}) has no email address. Skipping.")
                continue

            # Prepare the email
            subject, html_content, tracking_id = self._prepare_email(lead)
            recipient_email = lead['Email']

            # Log the email being sent
            logger.info(
                f"Sending email {i + 1}/{total_leads} to {recipient_email} ({lead.get('Company Name', 'Unknown')})")

            # Send the email
            success = self._send_email(recipient_email, subject, html_content)

            if success:
                # Update tracking data with sent time
                current_time = datetime.now().isoformat()
                self.tracking_data[tracking_id]['sent_time'] = current_time
                self.tracking_data[tracking_id]['last_activity'] = current_time

                # Save a copy of the sent email
                company_name = lead.get('Company Name', f'Unknown_{i}')
                safe_name = ''.join(c if c.isalnum() else '_' for c in company_name)
                sent_file = os.path.join(sent_dir, f"sent_{safe_name}_{tracking_id[-8:]}.html")

                with open(sent_file, 'w') as f:
                    f.write(html_content)

                successful_sends += 1
                logger.info(f"Email successfully sent to {recipient_email}")
            else:
                failed_sends += 1
                logger.error(f"Failed to send email to {recipient_email}")

            # Save tracking data after each email
            self._save_tracking_data()

            # Add delay between emails, except for the last one
            if i < total_leads - 1:
                delay = random.randint(self.delay_min, self.delay_max)
                logger.info(f"Waiting {delay} seconds before sending next email...")
                time.sleep(delay)

            # If we've reached the batch size, take a longer break
            if (i + 1) % self.batch_size == 0 and i < total_leads - 1:
                batch_delay = random.randint(self.delay_max * 2, self.delay_max * 3)
                logger.info(f"Completed batch of {self.batch_size} emails. Taking a {batch_delay} second break...")
                time.sleep(batch_delay)

        # Final stats
        logger.info(f"Campaign completed: {successful_sends} emails sent successfully, {failed_sends} failed")
        return successful_sends, failed_sends

    def generate_campaign_report(self):
        """Generate a summary report of the campaign"""
        total_leads = len(self.tracking_data)
        sent_count = sum(1 for data in self.tracking_data.values() if data['sent_time'])
        opened_count = sum(1 for data in self.tracking_data.values() if data['opened'])
        clicked_count = sum(1 for data in self.tracking_data.values() if data['clicked'])

        # Calculate rates
        open_rate = (opened_count / sent_count * 100) if sent_count > 0 else 0
        click_rate = (clicked_count / sent_count * 100) if sent_count > 0 else 0
        click_to_open_rate = (clicked_count / opened_count * 100) if opened_count > 0 else 0

        report = {
            'campaign_id': self.campaign_id,
            'timestamp': datetime.now().isoformat(),
            'total_leads': total_leads,
            'emails_sent': sent_count,
            'emails_opened': opened_count,
            'links_clicked': clicked_count,
            'open_rate': open_rate,
            'click_rate': click_rate,
            'click_to_open_rate': click_to_open_rate,
            'hot_leads': [
                {
                    'company': data['company'],
                    'contact': data['contact'],
                    'email': data['email'],
                    'opened': data['opened'],
                    'clicked': data['clicked']
                }
                for tracking_id, data in self.tracking_data.items()
                if data['clicked']  # Consider clicked leads as "hot"
            ],
            'warm_leads': [
                {
                    'company': data['company'],
                    'contact': data['contact'],
                    'email': data['email'],
                    'opened': data['opened'],
                    'clicked': data['clicked']
                }
                for tracking_id, data in self.tracking_data.items()
                if data['opened'] and not data['clicked']  # Opened but not clicked
            ]
        }

        # Save report to file
        report_file = f'campaign_report_{self.campaign_id}.json'
        with open(report_file, 'w') as f:
            json.dump(report, f, indent=4)

        logger.info(f"Campaign report generated and saved to {report_file}")

        # Print summary to console
        print("\n" + "=" * 50)
        print(f"CAMPAIGN SUMMARY (ID: {self.campaign_id})")
        print("=" * 50)
        print(f"Total leads processed: {total_leads}")
        print(f"Emails sent: {sent_count}")
        print(f"Emails opened: {opened_count} ({open_rate:.1f}%)")
        print(f"Links clicked: {clicked_count} ({click_rate:.1f}%)")
        print(f"Click-to-open rate: {click_to_open_rate:.1f}%")
        print(f"Hot leads: {len(report['hot_leads'])}")
        print(f"Warm leads: {len(report['warm_leads'])}")
        print("=" * 50)

        return report


def create_env_file():
    """Create a template .env file if it doesn't exist"""
    env_file = '.env'
    if not os.path.exists(env_file):
        with open(env_file, 'w') as f:
            f.write("""# Email Campaign Configuration
# Email credentials - USE AN APP PASSWORD FOR GMAIL
EMAIL_ADDRESS=your_email@gmail.com
EMAIL_PASSWORD=your_app_password
SENDER_NAME=Your Name
SENDER_COMPANY=Your Company
SENDER_POSITION=Sales Development Representative
SENDER_PHONE=+1 (555) 123-4567

# Email sending settings
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
BATCH_SIZE=10
DELAY_MIN=60
DELAY_MAX=180

# Tracking URL (replace with your actual tracking server)
TRACKING_URL=https://yourwebsite.com/track
""")
        print(f"Created template .env file at {env_file}")
        print("Please edit this file with your actual email credentials and settings before running the campaign.")


def setup_tracking_server():
    """Create a simple Flask app for tracking email opens and clicks"""
    tracking_server_file = 'tracking_server.py'

    if not os.path.exists(tracking_server_file):
        with open(tracking_server_file, 'w') as f:
            f.write('''# Simple Flask server for tracking email opens and clicks
from flask import Flask, request, redirect, send_file
import json
import os
from datetime import datetime
import logging

app = Flask(__name__)

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('tracking_server.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

@app.route('/track/pixel/<tracking_id>')
def track_open(tracking_id):
    """Track when an email is opened"""
    logger.info(f"Email opened: {tracking_id}")

    # Update tracking data
    update_tracking_data(tracking_id, 'opened')

    # Return a 1x1 transparent pixel
    return send_file('pixel.png', mimetype='image/png')

@app.route('/track/click/<tracking_id>')
def track_click(tracking_id):
    """Track when a link is clicked and redirect"""
    url = request.args.get('url', 'https://www.example.com')
    logger.info(f"Link clicked: {tracking_id}, redirecting to {url}")

    # Update tracking data
    update_tracking_data(tracking_id, 'clicked')

    # Redirect to the original URL
    return redirect(url)

def update_tracking_data(tracking_id, action):
    """Update the tracking data file with the new event"""
    # Find the campaign data file based on the tracking ID
    campaign_id = tracking_id.split('_')[0]
    data_file = f'campaign_data_{campaign_id}.json'

    if not os.path.exists(data_file):
        logger.error(f"Campaign data file not found: {data_file}")
        return

    try:
        # Load the existing data
        with open(data_file, 'r') as f:
            tracking_data = json.load(f)

        # Update the record if it exists
        if tracking_id in tracking_data:
            current_time = datetime.now().isoformat()
            tracking_data[tracking_id][action] = True
            tracking_data[tracking_id]['last_activity'] = current_time

            # Save the updated data
            with open(data_file, 'w') as f:
                json.dump(tracking_data, f, indent=4)

            logger.info(f"Updated tracking data for {tracking_id}: {action}")
        else:
            logger.warning(f"Tracking ID not found in data: {tracking_id}")

    except Exception as e:
        logger.error(f"Error updating tracking data: {e}")

# Create a 1x1 transparent pixel if it doesn't exist
if not os.path.exists('pixel.png'):
    from PIL import Image
    img = Image.new('RGBA', (1, 1), (0, 0, 0, 0))
    img.save('pixel.png')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
''')
        print(f"Created tracking server script at {tracking_server_file}")
        print("Requirements for tracking server: Flask, Pillow")
        print("Run with: python tracking_server.py")


def main():
    """Main function to run Task 4"""
    print("\n" + "=" * 60)
    print("Task 4: Automating Email Campaign".center(60))
    print("=" * 60)

    # Create template .env file
    create_env_file()

    # Setup tracking server
    setup_tracking_server()

    # Ask user if they want to proceed with sending emails
    print("\nBefore proceeding, make sure you've:")
    print("1. Edited the .env file with your email credentials")
    print("2. Confirmed that your email template is appropriate")
    print("3. Ensured your tracking server is set up (optional)")

    proceed = input("\nDo you want to proceed with sending the email campaign? (yes/no): ").strip().lower()
    if proceed != 'yes':
        print("Campaign aborted. You can run this script again when ready.")
        return

    # Initialize and run the campaign
    try:
        campaign = EmailCampaign()
        successful, failed = campaign.run_campaign()
        campaign.generate_campaign_report()

        print(f"\nCampaign completed: {successful} emails sent successfully, {failed} failed")
        print("Check the campaign report for more details.")

    except Exception as e:
        logger.error(f"Error running campaign: {e}")
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    main()