# Task_Automation-Validation-Pandas-Outlook-
This script automatically connects to Outlook, downloads the latest Excel attachment from emails with a specified subject, validates the file’s contents for data quality issues, saves a report of any problems found, and sends an approval response if the file passes all checks.
Absolutely. Here’s a **detailed, professional README.md** for your script, following best practices (clear intro, usage, requirements, explanation of each part, limitations, and credits).
If you want it even more concise or more formal, let me know and I’ll push back or revise as needed.


# Automated Outlook Excel Validator & Approval Script

This script automates the process of downloading Excel files from Outlook emails, validating the contents for data quality issues, and handling approval or reporting via automated email replies.

---

## Features

- **Connects to Microsoft Outlook** to read inbox emails.
- **Downloads the latest Excel attachment** from emails matching a specific subject keyword.
- **Validates the Excel file’s content** for required formats, values, and fields across two sheets (`Events_Parts` and `General_Events`).
- **Generates a report** listing any detected data issues.
- **Sends automated approval emails** if the file is valid, or saves an issues report otherwise.
- **Customizable** subject keyword and download folder.

---

## How It Works

1. **Connect to Outlook** and scan inbox for emails with a specific subject.
2. **Download** the most recent Excel attachment from that email.
3. **Validate** the Excel file by:
    - Checking fields in both `Events_Parts` and `General_Events` sheets for format, values, dates, etc.
    - Reporting mismatches or missing/incorrect data.
    - Ensuring cross-sheet consistency (e.g., matching `AL` column).
4. **If no issues:** Sends an approval email (with optional CC) to the sender.
5. **If issues are found:** Saves a detailed Excel file listing the problems for review.

---

## Requirements

- Windows OS (requires COM automation)
- Microsoft Outlook (installed and configured)
- Python 3.7+
- Required Python libraries:
    - `pywin32`
    - `pandas`
    - `xlsxwriter`

Install dependencies via pip:
```bash
pip install pywin32 pandas xlsxwriter
````

---

## Usage

1. **Update Configuration:**

   * Change `subject_keyword` and `download_folder` in the `__main__` section as needed.

2. **Run the script:**

   ```bash
   python your_script.py
   ```

3. **Behavior:**

   * Downloads the latest Excel file matching your subject.
   * Validates its contents.
   * If all checks pass: Sends an approval email.
   * If issues exist: Saves an `suspected_issues_YYYYMMDD.xlsx` report in your download folder.

---

## Example Output

* **Approval email** (if no issues):
  The script automatically sends a standardized approval response to the sender, CC’ing a defined list.

* **Issues report** (if problems found):
  A new Excel file is created listing the row, column, and problematic value for each issue detected.

---

## Customization

* **Validation rules** can be extended by editing the functions inside `validate_excel()`.
* **CC recipients** for approval can be updated in the `send_approval_response()` function.
* To change the inbox folder, adjust the folder ID (currently `6` for Inbox) in `connect_to_outlook()`.

---

## Limitations

* Only works with Outlook on Windows (uses COM automation).
* Assumes Excel files follow the specified schema (`Events_Parts` and `General_Events`).
* Current validation logic is hardcoded—needs updates if your data schema changes.
* Requires Outlook to be installed and the user to be signed in.

---

## Credits

Developed by \[Almohtadey Metwaly].
For questions, [almohtadey.metwaly@gmail.com].

