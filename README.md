# Automated Employee Report Generation System

## 📌 Project Overview

This project is an end-to-end automation solution developed using Python and Microsoft Power Automate. The system automatically generates employee reports, stores them in OneDrive, and sends them to stakeholders through Outlook email without any manual intervention.

The main objective of this project is to reduce repetitive manual reporting tasks, improve accuracy, and automate report delivery in real time.

---

## 🚀 Key Features

- Automated Excel report generation using Python
- Data extraction and processing from multiple sources
- Automatic OneDrive file monitoring
- Real-time workflow triggering using Power Automate
- Automated Outlook email delivery with attachments
- Conditional file validation before sending emails
- Zero manual effort after setup

---

## 🛠️ Technologies Used

| Technology | Purpose |
|---|---|
| Python | Data processing and report generation |
| Power Automate | Workflow automation |
| OneDrive for Business | File storage and trigger monitoring |
| Microsoft Outlook | Automated email delivery |
| Excel | Report format |

---

## ⚙️ System Workflow

```text
Python Script
      ↓
Generate Employee_Report.xlsx
      ↓
Save file to OneDrive folder
      ↓
Power Automate detects new file
      ↓
Validate report file name
      ↓
Get file content
      ↓
Send automated email with attachment
