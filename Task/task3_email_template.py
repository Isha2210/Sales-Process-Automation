# Task 3: Create an Email Template
# This script creates personalized email templates for the automated email campaign
# It can dynamically populate placeholders with data from the Excel file

import os
import pandas as pd
from jinja2 import Template
import re


def create_email_template():
    """
    Create an HTML email template with placeholders for personalization

    Returns:
    --------
    jinja2.Template
        A Jinja2 template object that can be used to generate personalized emails
    """
    # HTML email template with placeholders for dynamic content
    html_template = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Outreach Email</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                line-height: 1.6;
                color: #333333;
                margin: 0;
                padding: 0;
            }
            .container {
                max-width: 600px;
                margin: 0 auto;
                padding: 20px;
            }
            .header {
                margin-bottom: 20px;
            }
            .footer {
                margin-top: 30px;
                font-size: 12px;
                color: #777777;
                border-top: 1px solid #eeeeee;
                padding-top: 20px;
            }
            .cta-button {
                display: inline-block;
                background-color: #4CAF50;
                color: white !important;
                padding: 12px 20px;
                text-decoration: none;
                border-radius: 4px;
                font-weight: bold;
                margin: 20px 0;
            }
            .highlight {
                color: #0066cc;
                font-weight: bold;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <p>Hi {{ recipient_name }},</p>
            </div>

            <div class="content">
                <p>I hope this email finds you well. I came across {{ company_name }} while researching leading {{ industry }} companies in {{ location }} and was particularly impressed by your work in this space.</p>

                <p>My name is Isha, and I'm reaching out from [Your Company], where we specialize in helping {{ industry }} businesses like yours improve their [specific value proposition relevant to recipient's industry].</p>

                <p>Based on my research of {{ company_name }}'s online presence, I think we could help you:</p>

                <ul>
                    <li>Increase your benefit</li>
                    <li>Optimize your benefit </li>
                    <li>Streamline your benefit </li>
                </ul>

                <p>We've helped companies similar to yours achieve [specific result, ideally with a metric], and I'd love to share how we might be able to do the same for {{ company_name }}.</p>

                <a href="https://calendly.com/[your-link]" class="cta-button">Schedule a 15-minute call</a>

                <p>Would you be open to a brief conversation to explore how we might be able to help? If you're not the right person to speak with, could you kindly point me in the direction of whoever handles [relevant department/decision]?</p>

                <p>Thank you for your time, {{ recipient_name }}. I look forward to potentially working with {{ company_name }}.</p>
            </div>

            <div class="footer">
                <p>Best regards,<br>
                [Your Name]<br>
                [Your Position]<br>
                [Your Company]<br>
                [Your Contact Information]</p>

                <p>If you'd prefer not to receive further emails, please <a href="[unsubscribe-link]">click here to unsubscribe</a>.</p>
            </div>
        </div>
    </body>
    </html>
    """

    # Create a Jinja2 template object
    template = Template(html_template)

    # Save the template to a file for reference
    output_dir = 'templates'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with open(os.path.join(output_dir, 'email_template.html'), 'w') as f:
        f.write(html_template)

    print(f"Email template created and saved to {os.path.join(output_dir, 'email_template.html')}")

    return template


def create_subject_line_template():
    """
    Create a template for the email subject line

    Returns:
    --------
    jinja2.Template
        A template object for the subject line
    """
    subject_template = "{{ recipient_name }}, let's improve {{ company_name }}'s performance in {{ industry }}"

    # Create a Jinja2 template object
    template = Template(subject_template)

    # Save the template to a file for reference
    output_dir = 'templates'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with open(os.path.join(output_dir, 'subject_template.txt'), 'w') as f:
        f.write(subject_template)

    print(f"Subject line template created and saved to {os.path.join(output_dir, 'subject_template.txt')}")

    return template


def generate_personalized_email(lead_data, email_template, subject_template):
    """
    Generate a personalized email for a specific lead

    Parameters:
    -----------
    lead_data : dict
        Dictionary containing the lead's data
    email_template : jinja2.Template
        The HTML email template
    subject_template : jinja2.Template
        The subject line template

    Returns:
    --------
    tuple
        (subject, html_content) - The personalized subject and HTML content
    """
    # Extract the recipient's first name from the full name
    if 'Contact Person' in lead_data and lead_data['Contact Person']:
        full_name = lead_data['Contact Person']
        recipient_name = full_name.split()[0]  # Extract first name
    else:
        recipient_name = "there"  # Default if no name available

    # Prepare the context for template rendering
    context = {
        'recipient_name': recipient_name,
        'company_name': lead_data.get('Company Name', 'your company'),
        'industry': lead_data.get('Industry', 'your industry'),
        'location': lead_data.get('Location', 'your area')
    }

    # Render the templates with the context
    subject = subject_template.render(context)
    html_content = email_template.render(context)

    return subject, html_content


def preview_personalized_emails(excel_path=None):
    """
    Preview personalized emails for the first few leads in the Excel file

    Parameters:
    -----------
    excel_path : str, optional
        Path to the Excel file containing lead data. If None, it will look for the most recent export.
    """
    # Find the most recent Excel file if not specified
    if excel_path is None:
        output_dir = 'output'
        if os.path.exists(output_dir):
            excel_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx')]
            if excel_files:
                excel_files.sort(key=lambda x: os.path.getmtime(os.path.join(output_dir, x)), reverse=True)
                excel_path = os.path.join(output_dir, excel_files[0])

        # Check if path was found
        if excel_path is None or not os.path.exists(excel_path):
            print("No Excel file found. Please provide a valid path.")
            return

    # Load the Excel file
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return

    # Create email and subject templates
    email_template = create_email_template()
    subject_template = create_subject_line_template()

    # Preview for ALL leads
    num_previews = len(df)  # Modified to use all leads

    print(f"\nPreviewing personalized emails for {num_previews} leads:")

    # Create output directory for previews
    preview_dir = 'email_previews'
    if not os.path.exists(preview_dir):
        os.makedirs(preview_dir)

    for i in range(num_previews):
        lead = df.iloc[i].to_dict()
        subject, html_content = generate_personalized_email(lead, email_template, subject_template)

        print(f"\n--- Preview {i + 1}: Email for {lead.get('Company Name', 'Unknown Company')} ---")
        print(f"Subject: {subject}")
        print("HTML content saved to preview file.")

        # Save preview to file
        safe_filename = re.sub(r'[^\w\s-]', '', lead.get('Company Name', f'lead_{i + 1}')).strip().replace(' ', '_')
        preview_filename = os.path.join(preview_dir, f"preview_{safe_filename}_{i + 1}.html")

        with open(preview_filename, 'w') as f:
            f.write(html_content)

        print(f"Full preview saved to: {preview_filename}")


def main():
    """Main function to execute Task 3"""
    print("Task 3: Creating Email Template...")

    # Create the email and subject templates
    email_template = create_email_template()
    subject_template = create_subject_line_template()

    # Try to find the Excel file from Task 2
    excel_path = None
    last_export_path_file = os.path.join('output', 'last_export_path.txt')
    if os.path.exists(last_export_path_file):
        with open(last_export_path_file, 'r') as f:
            excel_path = f.read().strip()

    # Preview personalized emails
    if excel_path and os.path.exists(excel_path):
        preview_personalized_emails(excel_path)
    else:
        print("Excel file from Task 2 not found. Run Task 2 first or provide path manually.")
        print("Continuing without previews...")

    print("\nTask 3 completed! Email templates created and saved to the 'templates' directory.")
    print("These templates will be used in Task 4 (Email Campaign Automation).")


if __name__ == "__main__":
    main()