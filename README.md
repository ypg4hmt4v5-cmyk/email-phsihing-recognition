# Email Phishing Checker

## Overview

The Email Phishing Checker is a Python-based cybersecurity project designed to analyse email files and identify potential phishing characteristics.

The programme allows users to select email files (`.eml`) from their computer, extracts important email information and checks for suspicious phrases and links. It then calculates a phishing risk score and stores the results in an Excel spreadsheet.

The project was developed to explore cybersecurity, email security and Python programming while creating a practical tool for identifying potentially suspicious emails.

## Features

- **Email Selection:** Select any saved `.eml` email file using a file-selection window.
- **Email Extraction:** Extracts the sender, subject, reply-to address, date and email body.
- **Link Detection:** Identifies URLs contained within the email.
- **Suspicious Phrase Detection:** Searches for common phishing phrases, including:
  - Urgent
  - Verify your account
  - Account suspended
  - Confirm your details
  - Immediate action
- **Link Security Checks:** Identifies potentially suspicious links, including unusually long URLs, IP address-based links and URL shorteners.
- **Risk Scoring:** Calculates a phishing risk score out of 100 based on detected characteristics.
- **Excel Reporting:** Saves analysis results into an Excel spreadsheet.
- **Multiple Email Analysis:** Allows multiple emails to be analysed and saves each analysis in a separate worksheet within the same spreadsheet.

## Technologies Used

| Technology | Purpose |
|---|---|
| Python | Main programming language |
| Email Parser | Extracting and analysing email content |
| Regular Expressions (re) | Identifying URLs and suspicious patterns |
| Tkinter | Selecting email files |
| OpenPyXL | Creating and updating Excel spreadsheets |

## How It Works

The programme follows several stages:

1. **Email Selection:** The user selects an `.eml` file from their computer.
2. **Email Parsing:** Python extracts the email headers and message body.
3. **Content Analysis:** The programme searches for suspicious phrases and extracts URLs.
4. **Link Analysis:** URLs are checked for potentially suspicious characteristics.
5. **Risk Calculation:** A risk score is calculated using the number of suspicious phrases and links identified.
6. **Results Storage:** The results are saved into an Excel spreadsheet, with a separate worksheet for each email.

## Risk Scoring

The programme uses a basic rule-based scoring system.

| Detection | Points |
|---|---:|
| Each suspicious phrase | +10 |
| Each suspicious link | +20 |

The maximum possible score is 100.

### Risk Categories

| Risk Score | Classification |
|---|---|
| 0–20 | Low Risk |
| 30–50 | Medium Risk |
| 60–100 | High Risk |

## Installation

### Requirements

- Python 3
- OpenPyXL

### Install Dependencies

Open your terminal and run:

```bash
python3 -m pip install openpyxl
```

### Run the Programme

1. Clone or download this repository.
2. Open the project folder.
3. Run the Python file:

```bash
python3 phishing.py
```

4. Select an `.eml` email file when the file-selection window appears.
5. View the analysis in Terminal.
6. Open the generated `Email_Analysis.xlsx` spreadsheet to view the saved results.

## Output

The programme automatically creates an Excel spreadsheet called:

`Email_Analysis.xlsx`

Each analysed email receives its own worksheet, named after the email subject.

All results are stored in the same spreadsheet, allowing multiple email analyses to be kept together.

## Limitations

This project uses rule-based detection rather than advanced machine learning or artificial intelligence.

- Some legitimate emails may contain suspicious phrases.
- Some phishing emails may not contain any of the currently recognised phrases.
- A suspicious-looking URL is not necessarily malicious.
- The risk score is an indication of potential risk, not a definitive security verdict.
- The programme currently analyses saved `.eml` files rather than connecting directly to an email inbox.

## Future Improvements

Potential improvements include:

- Checking sender addresses against trusted domains.
- Analysing email attachments.
- Detecting domain spoofing.
- Improving URL analysis.
- Introducing machine learning for more advanced phishing detection.
- Connecting directly to email services such as Gmail.
- Improving the accuracy of the phishing risk scoring system.

## Disclaimer

This project is intended for educational purposes and cybersecurity awareness. It is designed to help identify potentially suspicious emails but should not be relied upon as a complete email security solution.

## Author

Developed as a personal Python and cybersecurity project to develop practical programming skills and explore email security and phishing detection.
