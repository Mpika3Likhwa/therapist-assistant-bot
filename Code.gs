/* 
Therapist Appointment Bot
- Monitors a Google Sheet for new appointment emails,
- Extracts details using AI API,
- Adds them to Google Calendar,
- Sends confirmation emails.
*/

const AI_API_KEY = "";    //<-- this is where the AI secret key goes 
const CALENDAR_NAME = "Therapist Appointments";
const SHEET_NAME = "Therapist Appointment Queue";
const SHEET_ID = "1xkKHgh7HaY2PRytQL9qMJuTtXbIRQGQ2Xy8lLbkZZMI";
const THERAPIST_EMAIL = "slmstroke0@gmail.com";

/* 
Main function:
- Entry point for the bot
- Calls processAppointmentsFromSheet to handle all pending appointments
*/
function main() {
  processAppointmentsFromSheet();
}

/* 
Function: callAIForDetails
- Sends email text to AI API
- Returns raw response text
*/
function callAIForDetails(emailBody, senderEmail) {
  const prompt = `
Extract the following appointment details from the email text below.

Fields to extract:
- date (format: YYYY-MM-DD)
- time (24-hour format: HH:MM)
- name
- email

If a value cannot be found, leave it as an empty string ("").

Return ONLY valid JSON, no text or explanation. Example:
{"date": "2025-11-03", "time": "14:00", "name": "John Doe", "email": "john@example.com", "purpose": "Therapy consultation"}

Email text:
${emailBody}
Sender email: ${senderEmail}
`;

  const response = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + AI_API_KEY,
      "HTTP-Referer": "https://script.google.com",
      "X-Title": "Therapist Appointment Bot"
    },
    payload: JSON.stringify({
      model: "",    // <-- this is where the model goes 
      messages: [
        { role: "system", content: "You are a helpful assistant that extracts appointment info and replies with only JSON." },
        { role: "user", content: prompt }
      ],
      temperature: 0.2
    }),
    muteHttpExceptions: true
  });

  return response.getContentText();
}

/* 
Function: parseAIResponse
- Parses AI response into usable details object
*/
function parseAIResponse(rawText) {
  let details;
  try {
    const parsed = JSON.parse(rawText);
    let content = parsed?.choices?.[0]?.message?.content;

    if (!content) {
      const fallbackMatch = rawText.match(/({[\s\S]*})/);
      if (fallbackMatch) content = fallbackMatch[1];
    }

    if (!content) throw new Error("No content found in AI response.");

    try {
      details = typeof content === "string" ? JSON.parse(content) : content;
    } catch (inner) {
      const innerMatch = String(content).match(/({[\s\S]*})/);
      if (!innerMatch) throw inner;
      details = JSON.parse(innerMatch[1]);
    }

    Logger.log("Extracted details object: " + JSON.stringify(details, null, 2));
    return details;
  } catch (e) {
    throw new Error("Error extracting appointment details: " + e + "\nRaw response:\n" + rawText);
  }
}

/* 
Function: createCalendarEvent
- Creates a calendar event if slot is available
*/
function createCalendarEvent(details, senderEmail) {
  if (!details.date || details.date.trim() === "") {
    throw new Error("AI returned no date; cannot schedule without date.");
  }
  if (!details.time || details.time.trim() === "") {
    throw new Error("AI returned no time; cannot schedule without time.");
  }

  const calendar = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0];
  if (!calendar) throw new Error("Calendar not found: " + CALENDAR_NAME);

  const startTime = new Date(`${details.date}T${details.time}:00`);
  if (isNaN(startTime.getTime())) {
    throw new Error("Invalid start time constructed: " + `${details.date}T${details.time}:00`);
  }
  const endTime = new Date(startTime.getTime() + 60 * 60 * 1000);

  const existingEvents = calendar.getEvents(startTime, endTime);
  if (existingEvents.length > 0) {
    throw new Error(`Time slot already booked for ${details.date} ${details.time}`);
  }

  calendar.createEvent(
    `Therapy Appointment - ${details.name || "Unknown"}`,
    startTime,
    endTime,
    {
      guests: details.email || senderEmail
    }
  );
}

/* 
Function: sendConfirmationEmails
- Sends confirmation to client and notification to therapist
*/
function sendConfirmationEmails(details, senderEmail) {
  if (details.email || senderEmail) {
    MailApp.sendEmail({
      to: details.email || senderEmail,
      subject: "Appointment Confirmation",
      body: `Hi ${details.name || "Client"},\n\nYour appointment is confirmed for ${details.date} at ${details.time}.\nLooking forward to what I hope will be a productive session`
    });
  }

  MailApp.sendEmail({
    to: THERAPIST_EMAIL,
    subject: "New Appointment Scheduled",
    body: `Appointment Details:\n-Client: ${details.name || "Unknown"}\n-Date: ${details.date}\n-Time: ${details.time}\n-Purpose: Therapy Consultation`
  });
}

/* 
Function: processAppointmentEmail
- Orchestrates AI call, parsing, event creation, and emails
*/
function processAppointmentEmail(emailBody, senderEmail) {
  try {
    const rawText = callAIForDetails(emailBody, senderEmail);
    const details = parseAIResponse(rawText);
    createCalendarEvent(details, senderEmail);
    sendConfirmationEmails(details, senderEmail);
    return "Success";
  } catch (error) {
    Logger.log("Error in processAppointmentEmail: " + error);
    return "Error: " + error;
  }
}

/* 
Function: processAppointmentsFromSheet
- Loops through Google Sheet
- Reads unprocessed emails and calls processAppointmentEmail
*/
function processAppointmentsFromSheet() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log("Sheet not found: " + SHEET_NAME);
      return;
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const processed = row[4]; // Column E (Processed)
      if (processed) continue;

      const senderEmail = row[1]; // Column B (Sender Email)
      const emailBody = row[3];   // Column D (Email Body)

      const result = processAppointmentEmail(emailBody, senderEmail);
      Logger.log(`Processed row ${i + 1}: ${result}`);

      sheet.getRange(i + 1, 5).setValue("YES");
    }
  } catch (error) {
    Logger.log("Error in processAppointmentsFromSheet: " + error);
  }
}
