/* 
Therapist Appointment Bot 
- Monitors a Google Sheet for new appointment emails,
- extracts details using OpenAI, adds them to Google Calendar,
- and sends confirmation emails.
*/

const OPENAI_API_KEY = "API KEY HERE";
const CALENDAR_NAME = "Therapist Appointments";
const SHEET_NAME = "Therapist Appointment Queue";
const SHEET_ID = "1xkKHgh7HaY2PRytQL9qMJuTtXbIRQGQ2Xy8lLbkZZMI";
const THERAPIST_EMAIL = "slmstroke0@gmail.com";

/*
processAppointmentEmail function:
- sends email over to AI API for data extraction
- receives formatted response
- converts it to a usable JavaScript object
- accesses the calendar and creates an appointment using the data returned by the AI
- sends a confirmation email to therapist and client
*/
function processAppointmentEmail(emailBody, senderEmail) {
  try {
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

    // Call the AI API
    const response = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + OPENAI_API_KEY,
        "HTTP-Referer": "https://script.google.com",
        "X-Title": "Therapist Appointment Bot"
      },
      payload: JSON.stringify({
        model: "minimax/minimax-m2:free", // the model you used in your logs
        messages: [
          { role: "system", content: "You are a helpful assistant that extracts appointment info and replies with only JSON." },
          { role: "user", content: prompt }
        ],
        temperature: 0.2
      }),
      muteHttpExceptions: true
    });

    const rawText = response.getContentText();
    Logger.log("Raw AI response: " + rawText);

    let details;

    try {
      // Parse the top-level API response (this is the structure you pasted)
      const parsed = JSON.parse(rawText);

      // The assistant's content (often a stringified JSON)
      let content = parsed?.choices?.[0]?.message?.content;

      if (!content) {
        // Fallback: try to find the first JSON object anywhere in the top-level text
        const fallbackMatch = rawText.match(/({[\s\S]*})/);
        if (fallbackMatch) content = fallbackMatch[1];
      }

      if (!content) throw new Error("No content found in AI response.");

      // content may be a JSON string already (example: "{\"date\":\"\",\"time\":\"14:00\"...}") OR a plain object string.
      // First try to parse directly; if that fails try to extract the JSON substring.
      try {
        details = typeof content === "string" ? JSON.parse(content) : content;
      } catch (inner) {
        // Try to extract a JSON object from the content string (handles appended reasoning)
        const innerMatch = String(content).match(/({[\s\S]*})/);
        if (!innerMatch) throw inner;
        details = JSON.parse(innerMatch[1]);
      }

      Logger.log("Extracted details object: " + JSON.stringify(details, null, 2));
    } catch (e) {
      throw new Error("Error extracting appointment details: " + e + "\nRaw response:\n" + rawText);
    }

    // Use GetDateFromText function to get the date from email body
    if (!details.date || details.date.trim() === "") {
      const dateFT = GetDateFromText(emailBody || "");    //dateFT - date from text
      if (dateFT) {
        details.date = dateFT;
        Logger.log("deduced date from email body: " + details.date);
      } else {
        // As a safe fallback use today's date (so event creation still proceeds — optional policy)
        details.date = formatDateISO(new Date());
        Logger.log("No date found; using fallback date (today): " + details.date);
      }
    }

    // Validate time — if missing, throw (we need a time to create the event)
    if (!details.time || details.time.trim() === "") {
      throw new Error("AI returned no time; cannot schedule without time: " + JSON.stringify(details));
    }

    // Access Google Calendar
    const calendar = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0];
    if (!calendar) throw new Error("Calendar not found: " + CALENDAR_NAME);

    // Build start and end time objects
    const startTime = new Date(`${details.date}T${details.time}:00`);
    if (isNaN(startTime.getTime())) {
      throw new Error("Invalid start time constructed: " + `${details.date}T${details.time}:00`);
    }
    const endTime = new Date(startTime.getTime() + 60 * 60 * 1000); // add 1 hour

    // Check if time is available
    const existingEvents = calendar.getEvents(startTime, endTime);
    if (existingEvents.length > 0) {
      throw new Error(`Time slot already booked for ${details.date} ${details.time}`);
    }

    // Create new calendar event if slot is available
    calendar.createEvent(
      `Therapy Appointment - ${details.name || "Unknown"}`,
      startTime,
      endTime,
      {
        guests: details.email || senderEmail // fallback
      }
    );

    // Send confirmation email to client
    if (details.email || senderEmail) {
      MailApp.sendEmail({
        to: details.email || senderEmail,
        subject: `Appointment Confirmation`,
        body: `Hi ${details.name || "Client"},\n\nYour appointment is confirmed for ${details.date} at ${details.time}.\nLooking forward to what I hope will be a productive session`
      });
    }

    // Notify the therapist
    MailApp.sendEmail({
      to: THERAPIST_EMAIL,
      subject: "New Appointment Scheduled",
      body: `\nAppointment Details:\n-Client: ${details.name || "Unknown"}\n-Date: ${details.date}\n-Time: ${details.time}\n-Purpose: Therapy Consultation`
    });

    return "Success";

  } catch (error) {
    Logger.log("Error in processAppointmentEmail: " + error);
    return "Error: " + error;
  }
}

/*
processAppointmentsFromSheet function:
- loops through the google sheet
- reads unprocessed emails and calls function processAppointmentEmail for each one
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

//-----------------------------------------------------------------------------------------------------------------

// Helpers:
/*
- Natural-language date extraction for 'today', 'tomorrow', 'next <weekday>', or '<weekday>'
Returns YYYY-MM-DD or null if nothing found.
*/
function GetDateFromText(text) {
  if (!text) return null;
  const s = text.toLowerCase();

  // today / tomorrow
  if (s.includes("today")) return formatDateISO(new Date());
  if (s.includes("tomorrow")) {
    const t = new Date();
    t.setDate(t.getDate() + 1);
    return formatDateISO(t);
  }

  // weekdays
  const weekdays = ["sunday","monday","tuesday","wednesday","thursday","friday","saturday"];
  for (let i = 0; i < weekdays.length; i++) {
    const w = weekdays[i];
    // look for "next <weekday>"
    if (s.includes("next " + w)) {
      const d = getNextWeekdayDate(i, true);
      return formatDateISO(d);
    }
    // look for "<weekday>" alone (e.g., "on tuesday" or "this tuesday" or "tuesday")
    const regex = new RegExp(`\\b(on |this |coming |)${w}\\b`);
    if (regex.test(s)) {
      const d = getNextWeekdayDate(i, false);
      return formatDateISO(d);
    }
  }

  // If user wrote "in 2 days" or "in 3 days" — basic support
  const inDaysMatch = s.match(/in\s+(\d{1,2})\s+days?/);
  if (inDaysMatch) {
    const offset = parseInt(inDaysMatch[1], 10);
    const d = new Date();
    d.setDate(d.getDate() + offset);
    return formatDateISO(d);
  }

  return null;
}

/*
Helper: returns next date for a given weekday index (0=Sunday..6=Saturday)
If addWeek is true, returns the next week's same weekday.
*/
function getNextWeekdayDate(weekdayIndex, addWeek) {
  const today = new Date();
  const todayIndex = today.getDay();
  let daysAhead = (weekdayIndex - todayIndex + 7) % 7;
  if (daysAhead === 0 && addWeek) daysAhead = 7;
  if (addWeek && daysAhead < 7) daysAhead += 7 - (daysAhead === 0 ? 0 : 0);
  const result = new Date(today.getFullYear(), today.getMonth(), today.getDate() + (daysAhead === 0 && !addWeek ? 0 : daysAhead));
  return result;
}

/*
Helper: format a Date object as "YYYY-MM-DD"
*/
function formatDateISO(d) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}
