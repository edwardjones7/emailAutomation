# emailAutomation

A simple Python script that reads a list of email addresses from an Excel `.xlsx` file and sends customized emails via Gmail’s SMTP server.

This is ideal for automating outreach or sending personalized email campaigns without manual sending. The script uses SMTP to securely connect to Gmail, loops through all recipients, and sends a formatted HTML email.

---

## 👨‍💻 Features

- Reads email addresses from an Excel file (`emails.xlsx`)
- Sends HTML-formatted emails to multiple recipients
- Uses Gmail SMTP with TLS encryption
- Simple structure — easy to customize

---

## 🚀 Requirements

- Python 3.8+
- `pandas` library for reading Excel files

Install dependencies:

```bash
pip install pandas openpyxl
