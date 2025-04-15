# Sales-Process-Automation
Sales Process Automation using Python and AI/ML Concepts
Project Overview:
This project automates key aspects of the sales process using Python programming and AI/ML techniques. The automation covers data scraping from online sources, data extraction to Excel, email template creation, email campaign automation, and analytics gathering.

Table of Contents:
1.	Data Scraping and Extraction
2.	Data Export to Excel
3.	Email Template Creation
4.	Email Campaign Automation
5.	Analytics Gathering
6.	Installation and Setup
7.	Usage Guide
8.	Technologies Used

(1) Data Scraping and Extraction:
Ideal Customer Profile (ICP)
For this project, we defined our ICP with the following criteria:
•	Industry: Software as a Service (SaaS)
•	Company Size: 50-500 employees
•	Location: United States, primarily tech hubs (Bay Area, New York, Austin)
•	Decision Makers: C-level executives, VPs of Sales/Marketing

The scraping process follows these steps:
1.	Define targeted search queries based on our ICP
2.	Use Selenium to scrape Google search results for matching companies
3.	Extract basic company information from search results
4.	Visit company websites to enrich data with contact information
5.	Filter and deduplicate the results

(2) Data Export to Excel:
After scraping and enriching the data, we export it to Excel for easy management and manipulation.
The Excel file includes the following columns:
•	Company Name
•	Website URL
•	Contact Email (when available)
•	Phone Number (when available)
•	Location
•	Company Description

(3) Email Template Creation:
We created a dynamic email template system using HTML and Python string formatting.
The email template includes:
•	Personalized greeting with the recipient's name
•	Dynamic content mentioning the company name and industry
•	A clear call-to-action button for scheduling a meeting
•	Professional signature with sender details
•	Hidden tracking pixel for open rate tracking

(4) Email Campaign Automation:
Key features of the email automation:
•	Unique tracking IDs for each recipient
•	Personalized content for each lead
•	Randomized delays between emails to avoid spam detection
•	Error handling and logging
•	Campaign data storage for analytics tracking

(5) Analytics Gathering:
For analytics gathering, we implemented a tracking server (Flask) to monitor email opens and link clicks.
The analytics gathering process:
1.	Tracks email opens using a transparent tracking pixel
2.	Monitors link clicks through redirected URLs
3.	Stores all events with timestamps, IP addresses, and user agents
4.	Categorizes leads as hot (clicked), warm (opened), or cold (no engagement)
5.	Generates comprehensive reports in Excel format

(6) Installation and Setup:
Prerequisites
•	Python 3.8+
•	Required Python packages: 
o	pandas
o	requests
o	beautifulsoup4
o	selenium
o	flask
o	openpyxl
o	pillow (optional, for tracking pixel creation)

(7) Usage Guide:

1. Data Scraping
To scrape data based on your ICP:
python web_scraping-_script.py
This will:
•	Search for companies matching your ICP
•	Extract and enrich company data
•	Save the results to an Excel file

2. Email Campaign Setup:
To set up and run an email campaign:
python email_campaign.py --leads-file sales_leads.xlsx
This will:
•	Load lead data from Excel
•	Generate personalized emails for each lead
•	Send emails with tracking pixels and unique links
•	Record campaign data for analytics

3. Tracking Server:
To start the tracking server for analytics:
python tracking_server.py
This will:
•	Start a Flask server on port 5000
•	Track email opens and link clicks
•	Store tracking data in JSON files

5. Analytics Generation:
To generate campaign analytics:
python task5_analytics.py --campaign-id 20250414145110
This will:
•	Process tracking data for the specified campaign
•	Categorize leads based on engagement
•	Generate an Excel report with

(8) Technologies used:
Python and AI/ML
Editor: PyCharm 
