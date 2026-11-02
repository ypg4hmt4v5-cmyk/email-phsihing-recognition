 
from email import policy
from email.parser import BytesParser
from tkinter import Tk, filedialog
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
import os
import re

# ----------------------------------
# SET UP SPREADSHEET
# ----------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
spreadsheet_path = os.path.join(BASE_DIR, "Email_Analysis.xlsx")

# Open existing spreadsheet or create a new one
if os.path.exists(spreadsheet_path):
    workbook = load_workbook(spreadsheet_path)
else:
    workbook = Workbook()
    workbook.remove(workbook.active)

# ----------------------------------
# SUSPICIOUS WORDS
# ----------------------------------

suspicious_words = [
    "urgent",
    "verify your account",
    "password",
    "click here",
    "account suspended",
    "confirm your details",
    "limited time",
    "immediate action"
]

# ----------------------------------
# CREATE FILE PICKER
# ----------------------------------

root = Tk()
root.withdraw()

# ----------------------------------
# EMAIL ANALYSIS FUNCTION
# ----------------------------------

def analyse_email(email_path):

    # Parse email
    with open(email_path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)

    # Extract email information
    sender = str(msg["From"] or "Not provided")
    subject = str(msg["Subject"] or "No subject")
    reply_to = str(msg["Reply-To"] or "Not provided")
    date = str(msg["Date"] or "Not provided")

    # Extract email body
    body = ""

    for part in msg.walk():

        if part.get_content_type() in ["text/plain", "text/html"]:

            if part.get_content_disposition() != "attachment":

                try:
                    content = part.get_content()

                    if isinstance(content, str):
                        body += content + "\n"

                except Exception:
                    pass

    # Find URLs in email
    links = re.findall(
        r"""https?://[^\s<>"']+""",
        body
    )

    # Check for suspicious phrases
    found_words = []

    for word in suspicious_words:

        if word in body.lower():
            found_words.append(word)

    # Check for suspicious links
    suspicious_links = []

    for link in links:

        if len(link) > 100:

            suspicious_links.append(
                (link, "Unusually long URL")
            )

        elif re.search(
            r"https?://\d{1,3}(\.\d{1,3}){3}",
            link
        ):

            suspicious_links.append(
                (link, "Uses an IP address")
            )

        elif any(shortener in link.lower() for shortener in [
            "bit.ly", "tinyurl.com", "t.co"
        ]):

            suspicious_links.append(
                (link, "Uses a URL shortener")
            )

    # Calculate phishing risk score
    risk_score = 0

    risk_score += len(found_words) * 10
    risk_score += len(suspicious_links) * 20

    risk_score = min(risk_score, 100)

    # Determine risk level
    if risk_score >= 60:
        result = "HIGH RISK"

    elif risk_score >= 30:
        result = "MEDIUM RISK"

    else:
        result = "LOW RISK"

    # Return analysis results
    return {
        "sender": sender,
        "subject": subject,
        "reply_to": reply_to,
        "date": date,
        "body": body,
        "links": links,
        "found_words": found_words,
        "suspicious_links": suspicious_links,
        "risk_score": risk_score,
        "result": result
    }


# ----------------------------------
# CREATE SPREADSHEET WORKSHEET
# ----------------------------------

def add_to_spreadsheet(data, filename):

    # Use email subject as worksheet name
    sheet_name = data["subject"]

    # Remove characters Excel does not allow in sheet names
    sheet_name = re.sub(r'[\\/*?:\[\]]', '', sheet_name)

    sheet_name = sheet_name.strip("'")[:31]

    if not sheet_name:
        sheet_name = os.path.splitext(filename)[0][:31]

    # Make sure worksheet name is unique
    original_name = sheet_name
    counter = 2

    while sheet_name in workbook.sheetnames:

        suffix = f" ({counter})"
        sheet_name = original_name[:31 - len(suffix)] + suffix
        counter += 1

    # Create new worksheet
    sheet = workbook.create_sheet(title=sheet_name)

    # Set column widths
    sheet.column_dimensions["A"].width = 28
    sheet.column_dimensions["B"].width = 85

    # Add heading
    sheet["A1"] = "EMAIL PHISHING ANALYSIS"
    sheet["A1"].font = Font(
        bold=True,
        size=16,
        color="FFFFFF"
    )
    sheet["A1"].fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )

    sheet.merge_cells("A1:B1")

    # Email information
    rows = [
        ("Sender", data["sender"]),
        ("Subject", data["subject"]),
        ("Reply-To", data["reply_to"]),
        ("Date", data["date"]),
        ("Original File", filename),
    ]

    current_row = 3

    for label, value in rows:

        sheet.cell(current_row, 1, label)
        sheet.cell(current_row, 2, value)

        current_row += 1

    # Email body
    current_row += 1

    sheet.cell(current_row, 1, "EMAIL BODY")
    sheet.cell(current_row, 1).font = Font(bold=True, size=12)

    current_row += 1

    sheet.cell(current_row, 2, data["body"][:32000])
    sheet.cell(current_row, 2).alignment = Alignment(
        wrap_text=True,
        vertical="top"
    )

    current_row += 2

    # Links found
    sheet.cell(current_row, 1, "LINKS FOUND")
    sheet.cell(current_row, 1).font = Font(bold=True, size=12)

    current_row += 1

    if data["links"]:

        for link in data["links"]:
            sheet.cell(current_row, 2, link)
            current_row += 1

    else:
        sheet.cell(current_row, 2, "No links found")
        current_row += 1

    current_row += 1

    # Suspicious phrases
    sheet.cell(current_row, 1, "SUSPICIOUS PHRASES")
    sheet.cell(current_row, 1).font = Font(bold=True, size=12)

    current_row += 1

    if data["found_words"]:

        for word in data["found_words"]:
            sheet.cell(current_row, 2, word)
            current_row += 1

    else:
        sheet.cell(current_row, 2, "No suspicious phrases detected")
        current_row += 1

    current_row += 1

    # Suspicious links
    sheet.cell(current_row, 1, "SUSPICIOUS LINKS")
    sheet.cell(current_row, 1).font = Font(bold=True, size=12)

    current_row += 1

    if data["suspicious_links"]:

        for link, reason in data["suspicious_links"]:

            sheet.cell(current_row, 1, reason)
            sheet.cell(current_row, 2, link)

            current_row += 1

    else:
        sheet.cell(current_row, 2, "No suspicious link characteristics detected")
        current_row += 1

    current_row += 1

    # Final phishing assessment
    sheet.cell(current_row, 1, "FINAL PHISHING ASSESSMENT")
    sheet.cell(current_row, 1).font = Font(bold=True, size=12)

    current_row += 1

    sheet.cell(current_row, 1, "Risk Score")
    sheet.cell(current_row, 2, f'{data["risk_score"]}/100')

    current_row += 1

    sheet.cell(current_row, 1, "Result")
    sheet.cell(current_row, 2, data["result"])

    # Format result
    if data["result"] == "HIGH RISK":
        sheet.cell(current_row, 2).fill = PatternFill(
            "solid", fgColor="FF9999"
        )

    elif data["result"] == "MEDIUM RISK":
        sheet.cell(current_row, 2).fill = PatternFill(
            "solid", fgColor="FFD966"
        )

    else:
        sheet.cell(current_row, 2).fill = PatternFill(
            "solid", fgColor="A9D18E"
        )

    sheet.freeze_panes = "A3"

    # Save spreadsheet
    workbook.save(spreadsheet_path)

    print("\nEmail successfully added to spreadsheet.")
    print("Worksheet name:", sheet_name)
    print("Spreadsheet:", spreadsheet_path)


# ----------------------------------
# MAIN PROGRAMME
# ----------------------------------

print("================================")
print("       PHISHING CHECKER")
print("================================")

while True:

    print("\nSelect an email to analyse.")

    email_path = filedialog.askopenfilename(
        title="Select an email",
        filetypes=[
            ("Email files", "*.eml"),
            ("All files", "*.*")
        ]
    )

    # Stop if no email was selected
    if not email_path:
        print("\nNo email selected.")
        break

    print("\nAnalysing email...")

    try:

        # Analyse selected email
        data = analyse_email(email_path)

        # Display results in terminal
        print("\nEMAIL ANALYSIS")
        print("------------------------")
        print("Sender:", data["sender"])
        print("Subject:", data["subject"])
        print("Reply-To:", data["reply_to"])
        print("Date:", data["date"])

        print("\nLINKS FOUND")
        print("------------------------")

        for link in data["links"]:
            print(link)

        if not data["links"]:
            print("No links found.")

        print("\nPHISHING CHECK")
        print("------------------------")

        for word in data["found_words"]:
            print("-", word)

        if not data["found_words"]:
            print("No suspicious phrases detected.")

        print("\nFINAL PHISHING ASSESSMENT")
        print("------------------------")
        print("Risk score:", data["risk_score"], "/100")
        print("Result:", data["result"])

        # Add results to Excel
        add_to_spreadsheet(
            data,
            os.path.basename(email_path)
        )

    except Exception as error:

        print("\nError analysing email:", error)

    # Ask whether to analyse another email
    another = input(
        "\nWould you like to analyse another email? (y/n): "
    ).strip().lower()

    if another != "y":
        break

root.destroy()

print("\n================================")
print("       PROGRAMME FINISHED")
print("================================")
print("Spreadsheet saved at:")
print(spreadsheet_path)
