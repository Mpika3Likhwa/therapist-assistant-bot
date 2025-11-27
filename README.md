# 🧠 Therapist Appointment Automation Bot

This project automates the process of handling therapist appointment emails using **Zapier**, **Google Sheets**, **Google Apps Script**, and an **AI API** for data extraction.

---

---

## ⚙️ How the System Works

### 1. Email Intake
- A client sends an appointment request email to the therapist’s inbox.
- The subject usually includes words like “appointment”, “booking”, or “request”.

### 2. Zapier Integration
- A Zapier automation monitors the therapist’s mailbox.
- When a new appointment-related email arrives, Zapier extracts the email’s subject and body.
- The extracted details are added as a new row to a connected **Google Sheet**.

### 3. Google Apps Script Processing
- A Google Apps Script is attached to that Google Sheet.
- Whenever a new row is added, the script:
  - Sends the email content to an **AI endpoint** (OpenAI API) for structured data extraction.
  - Creates a **Google Calendar event** using the extracted date, time, and client info.
  - Sends **confirmation emails** to both the client and the therapist.

### 4. AI Data Extraction
- The AI API formats messy, unstructured email text into structured JSON with fields such as:
  ```json
  {
    "client_name": "Jane Doe",
    "appointment_date": "2025-11-04",
    "appointment_time": "14:00",
    "email": "janedoe@example.com"
  }
