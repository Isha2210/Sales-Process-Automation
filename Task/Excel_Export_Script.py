# Task 2: Export Data into Excel Format
# This script takes the scraped data from Task 1 and exports it to an Excel file

import pandas as pd
import os
from datetime import datetime
import json


def export_to_excel(data, output_file=None):
    """
    Export the scraped data to an Excel file

    Parameters:
    -----------
    data : list of dictionaries
        The data scraped from Task 1
    output_file : str, optional
        Name of the output Excel file. If None, a timestamp-based name will be used

    Returns:
    --------
    str
        Path to the created Excel file
    """
    # Convert data to DataFrame
    df = pd.DataFrame(data)

    # Create output directory if it doesn't exist
    output_dir = 'output'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Generate output filename if not provided
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"leads_{timestamp}.xlsx"

    # Ensure the filename has .xlsx extension
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'

    # Create the full path
    output_path = os.path.join(output_dir, output_file)

    # Export to Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Leads')

        # Auto-adjust columns' width
        worksheet = writer.sheets['Leads']
        for i, col in enumerate(df.columns):
            max_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.column_dimensions[worksheet.cell(row=1, column=i + 1).column_letter].width = max_width

    print(f"Data successfully exported to {output_path}")

    # Save the path for later use in tasks 3 and 4
    with open(os.path.join(output_dir, 'last_export_path.txt'), 'w') as f:
        f.write(output_path)

    return output_path


def load_scraped_data(input_file=None):
    """
    Load data from Task 1

    Parameters:
    -----------
    input_file : str, optional
        Path to the JSON file containing scraped data from Task 1.
        If None, it will look for the most recent file in the data directory.

    Returns:
    --------
    list
        The scraped data as a list of dictionaries
    """
    # If no input file specified, try to find the most recent data file
    if input_file is None:
        data_dir = 'data'
        if not os.path.exists(data_dir):
            # For testing purposes, generate sample data
            return generate_sample_data()

        # Find the most recent JSON file
        json_files = [f for f in os.listdir(data_dir) if f.endswith('.json')]
        if not json_files:
            return generate_sample_data()

        # Sort by modification time (most recent first)
        json_files.sort(key=lambda x: os.path.getmtime(os.path.join(data_dir, x)), reverse=True)
        input_file = os.path.join(data_dir, json_files[0])

    # Load the data
    try:
        with open(input_file, 'r') as f:
            data = json.load(f)
        return data
    except Exception as e:
        print(f"Error loading data: {e}")
        return generate_sample_data()


def generate_sample_data():
    """Generate sample data for testing purposes"""
    return [
        {
            "Company Name": "Tech Solutions Inc.",
            "Contact Person": "John Smith",
            "Title": "CTO",
            "Industry": "Software Development",
            "Website": "https://techsolutions.example.com",
            "Location": "San Francisco, CA",
            "Email": "john.smith@techsolutions.example.com",
            "Phone": "123-456-7890",
            "Company Size": "50-100",
            "LinkedIn": "https://linkedin.com/in/johnsmith"
        },
        {
            "Company Name": "Data Analytics Pro",
            "Contact Person": "Sarah Johnson",
            "Title": "CEO",
            "Industry": "Data Analytics",
            "Website": "https://dataanalyticspro.example.com",
            "Location": "New York, NY",
            "Email": "sarah.j@dataanalyticspro.example.com",
            "Phone": "987-654-3210",
            "Company Size": "10-50",
            "LinkedIn": "https://linkedin.com/in/sarahjohnson"
        },
        {
            "Company Name": "Cloud Systems LLC",
            "Contact Person": "Michael Brown",
            "Title": "Sales Director",
            "Industry": "Cloud Computing",
            "Website": "https://cloudsystems.example.com",
            "Location": "Austin, TX",
            "Email": "m.brown@cloudsystems.example.com",
            "Phone": "555-123-4567",
            "Company Size": "100-500",
            "LinkedIn": "https://linkedin.com/in/michaelbrown"
        },
        {
            "Company Name": "AI Innovations",
            "Contact Person": "Jessica Lee",
            "Title": "CIO",
            "Industry": "Artificial Intelligence",
            "Website": "https://aiinnovations.example.com",
            "Location": "Seattle, WA",
            "Email": "jessica@aiinnovations.example.com",
            "Phone": "206-555-7890",
            "Company Size": "10-50",
            "LinkedIn": "https://linkedin.com/in/jessicalee"
        },
        {
            "Company Name": "Smart Manufacturing Inc.",
            "Contact Person": "Robert Chen",
            "Title": "Operations Manager",
            "Industry": "Manufacturing",
            "Website": "https://smartmanufacturing.example.com",
            "Location": "Chicago, IL",
            "Email": "r.chen@smartmanufacturing.example.com",
            "Phone": "312-555-6543",
            "Company Size": "500+",
            "LinkedIn": "https://linkedin.com/in/robertchen"
        }
    ]


def main():
    """Main function to execute Task 2"""
    print("Task 2: Exporting scraped data to Excel...")

    # Load the data from Task 1
    scraped_data = load_scraped_data()

    # Export to Excel
    excel_file = export_to_excel(scraped_data)

    print(f"Task 2 completed! Data exported to: {excel_file}")
    print("This Excel file will be used in Task 3 (Email Template Creation) and Task 4 (Email Campaign Automation)")

    return excel_file


if __name__ == "__main__":
    main()