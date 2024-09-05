# PerFi
PerFi is a Python-based tool designed to help users track their investments. 
The goal is to make financial insights actionable and visualizations simple while reducing hassle and boredom.

# Motivation

Most brokerages offer features to ensure that your contributions remain within the IRS limits. However, this can become complicated if you change jobs during the year.

PerFi v1 lets you track HSA & 401K contibutions and comparison against the IRS limit for 2024 & 2025.

## Features

- Calculates remaining or exceeded contributions for HSA and 401(k).
- Generates and updates an Excel file with contribution data.
- Updates a Google Sheet with contribution data.
- Highlights exceeded contributions with a red fill.
- Automatically resizes columns for optimal viewing.

## Setup

### Prerequisites

- Ensure you have Python 3 installed.
- Works best if you have MS Excel installed
- Also works with Google sheets with extra work of Google Cloud project with access to the Google Sheets API and a service account for authentication.

### Installation

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/nishk/PerFi.git
   cd PerFi
   
Install the required Python packages using pip and the requirements.txt file:
   
   pip install -r requirements.txt

Google Sheets API Setup (Optional):
	Enable the Google Sheets API in your Google Cloud project.
	Create a service account and download the JSON credentials file.
	Share your Google Sheet with the service account email.

Update input.yaml file with the following structure:
  file_path: "/path/to/save/excel/files/"
  google_sheet_url: "https://docs.google.com/spreadsheets/d/your_google_sheet_id/edit"
  credentials_file: "/path/to/your/service_account_credentials.json"

**Usage**

Command-Line Interface

You can run the script with the following options:

python contribution_tracker.py --year {2024,2025} --hsa HSA_CONTRIBUTED --k401 K401_CONTRIBUTED [--family]

	•	--year: Year for which you want to calculate contributions (2024 or 2025).
	•	--hsa: Amount contributed to HSA so far in the selected year.
	•	--k401: Amount contributed to 401(k) so far in the selected year.
	•	--family: (Optional) Flag indicating if HSA is for family (if not set, assumed individual).

**Examples**

	1.	Calculate Contributions for 2024 as an Individual:
 python contribution_tracker.py --year 2024 --hsa 2000 --k401 15000

	2.	Calculate Contributions for 2025 as a Family:
 python contribution_tracker.py --year 2025 --hsa 5000 --k401 18000 --family

**Output**

The script will generate:

	•	Excel File: A file named contribution_summary.xlsx in the specified file_path.
	•	Google Sheet: Updates to the specified Google Sheet with the contribution data

**Acknowledgments**

This tool leverages several open-source Python libraries, including:

	•	gspread
	•	gspread-formatting
	•	openpyxl
	•	pandas
	•	oauth2client

**License**

This project is licensed under the MIT License. See the LICENSE file for more details.

**Contributing**

Feel free to fork this repository, submit pull requests, or create issues for any bugs or feature requests.


