function sendCompanyOutreach() {
  // Get the active spreadsheet and the first sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  // Get the header row to map column names to indices
  const headers = data[0];
  const getColIndex = (colName) => headers.indexOf(colName);
  
  // Map column indices
  const cols = {
    firstName: getColIndex('First Name'),
    lastName: getColIndex('Last Name'),
    email: getColIndex('Email Address'),
    attachment: getColIndex('Attachment')
  };

  // Add status columns if they don't exist
  let statusSentCol = getColIndex('Status Sent');
  if (statusSentCol === -1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Status Sent');
    headers.push('Status Sent');
    statusSentCol = headers.length - 1;
  }
  
  let emailOpenedCol = getColIndex('Email Opened');
  if (emailOpenedCol === -1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Email Opened');
    headers.push('Email Opened');
    emailOpenedCol = headers.length - 1;
  }

  // Process each row starting from row 2 (skipping headers)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = row[cols.email];
    
    // Skip if email is empty or invalid
    if (!email || email.trim() === '') continue;
    
    const firstName = row[cols.firstName];
    // Skip if already sent
    if (row[statusSentCol] === "Sent") continue;

    // Create HTML email body with tracking image
    const trackingId = Utilities.getUuid();
    const trackingUrl = ScriptApp.getService().getUrl() + "?trackingId=" + trackingId + "&email=" + encodeURIComponent(email);
    const attachmentLink = row[cols.attachment] || "https://edubridge-academy.com/resources"; // Default link if none provided
    
    // Full-width HTML template for better PC display
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body, table, td {
      font-family: Arial, Helvetica, sans-serif;
      color: #333333;
      line-height: 1.5;
      margin: 0;
      padding: 0;
    }
    .highlight-box {
      background-color: #f8f8f8;
      padding: 15px 20px;
      border-left: 4px solid #2b6cb0;
      margin-bottom: 25px;
    }
    .info-box {
      background-color: #edf2f7;
      padding: 15px 20px;
      margin-bottom: 25px;
      border-radius: 4px;
    }
    .action-box {
      background-color: #e6f7ff;
      padding: 15px 20px;
      border-radius: 4px;
      margin-bottom: 25px;
    }
    .action-button {
      display: inline-block;
      background-color: #2b6cb0;
      color: white !important;
      padding: 8px 16px;
      text-decoration: none;
      border-radius: 4px;
      font-weight: bold;
    }
    h2 {
      color: #2b6cb0;
      margin-top: 0;
      font-size: 18px;
    }
    ul {
      padding-left: 20px;
      margin-bottom: 0;
    }
    li {
      margin-bottom: 8px;
    }
  </style>
</head>
<body>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="#ffffff">
    <tr>
      <td align="left">
        <table border="0" cellpadding="0" cellspacing="0" width="100%">
          <tr>
            <td style="padding: 15px 20px;">
                <p style="font-size: 16px; margin-bottom: 15px;"><strong>Dear ${firstName},</strong></p>

                <p style="font-size: 16px; margin-bottom: 20px;">I hope you're doing well.</p>

                <p style="font-size: 16px; margin-bottom: 20px;">I'm reaching out from EduBridge Academy to share some of our specialized professional training programs designed to support your personal and career growth.</p>

                <p style="font-size: 16px; margin-bottom: 20px;">We currently offer two focused learning tracks through our <strong>Open Academy</strong>:</p>

                <div class="highlight-box" style="background-color: #f8f8f8; padding: 15px 20px; border-left: 4px solid #2b6cb0; margin-bottom: 25px;">
                  <h2 style="color: #2b6cb0; margin-top: 0; font-size: 18px;"><strong>School of Finance</strong></h2>
                  <p style="margin-bottom: 10px;">Courses include:</p>
                  <ul style="padding-left: 20px; margin-bottom: 0;">
                    <li style="margin-bottom: 8px;">Financial Modeling & Valuation</li>
                    <li style="margin-bottom: 8px;">Financial Instrument</li>
                    <li style="margin-bottom: 8px;">Financial Analysis</li>
                    <li style="margin-bottom: 8px;">Basic Accounting & Financial Statement Analysis</li>
                    <li style="margin-bottom: 8px;">Prompt Engineering</li>
                    <li style="margin-bottom: 8px;">Business Planning</li>
                    <li style="margin-bottom: 8px;">Excel for Finance</li>
                    <li style="margin-bottom: 8px;">Power Point/Pitch Deck</li>
                  </ul>
                </div>

                <div class="highlight-box" style="background-color: #f8f8f8; padding: 15px 20px; border-left: 4px solid #2b6cb0; margin-bottom: 25px;">
                  <h2 style="color: #2b6cb0; margin-top: 0; font-size: 18px;"><strong>📊 Data School</strong></h2>
                  <p style="margin-bottom: 10px;">Courses include:</p>
                  <ul style="padding-left: 20px; margin-bottom: 0;">
                    <li style="margin-bottom: 8px;">Excel Essentials</li>
                    <li style="margin-bottom: 8px;">Power BI for Analysis & Reporting</li>
                    <li style="margin-bottom: 8px;">SQL</li>
                    <li style="margin-bottom: 8px;">Python</li>
                    <li style="margin-bottom: 8px;">Business Communication</li>
                  </ul>
                </div>

                <div class="info-box" style="background-color: #edf2f7; padding: 15px 20px; margin-bottom: 25px; border-radius: 4px;">
                  <h2 style="color: #2b6cb0; margin-top: 0; font-size: 18px;"><strong>Why choose EduBridge?</strong></h2>
                  <ul style="padding-left: 20px; margin-bottom: 0;">
                    <li style="margin-bottom: 10px;">Flexible learning: in-person or virtual sessions</li>
                    <li style="margin-bottom: 10px;">Expert-led training</li>
                    <li style="margin-bottom: 10px;">Practical content tailored to real-world applications</li>
                  </ul>
                </div>

                <p style="font-size: 16px; margin-bottom: 20px;">If you're looking to sharpen your skills or explore new areas professionally, feel free to reply to this email or contact me directly at <strong>08120288047</strong> or <strong>08101786422</strong> to learn more.</p>

                <div class="action-box" style="background-color: #e6f7ff; padding: 15px 20px; border-radius: 4px; margin-bottom: 25px;">
                  <p style="margin-top: 15px;">
                    <a href="${attachmentLink}" class="action-button" style="display: inline-block; background-color: #2b6cb0; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; font-weight: bold;">DOWNLOAD OUR PROGRAM BROCHURE</a>
                  </p>
                </div>

                <p style="font-size: 16px; margin-bottom: 20px;">Looking forward to supporting your growth journey!</p>

                <p style="font-size: 16px; margin-bottom: 10px;">Warm regards,</p>

                <p style="font-size: 16px; margin-bottom: 5px;"><strong>Esther Delano</strong></p>
                <p style="font-size: 16px; margin-bottom: 5px; color: #666;">Business Development Manager</p>
                <p style="font-size: 16px; margin-bottom: 20px; color: #2b6cb0;"><strong>EduBridge Academy</strong></p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
    <img src="${trackingUrl}" width="1" height="1" alt="" style="display:block" />
  </body>
</html>`;

    // Create plain text version as fallback
    const plainBody = `
Dear ${firstName},

I hope you're doing well.

I'm reaching out from EduBridge Academy to share some of our specialized professional training programs designed to support your personal and career growth.

We currently offer two focused learning tracks through our Open Academy:

School of Finance
Courses include:
* Financial Modeling & Valuation
* Financial Instrument
* Financial Analysis
* Basic Accounting & Financial Statement Analysis
* Prompt Engineering
* Business Planning
* Excel for Finance
* Power Point/Pitch Deck

📊 Data School
Courses include:
* Excel Essentials
* Power BI for Analysis & Reporting
* SQL
* Python
* Business Communication

Why choose EduBridge?
* Flexible learning: in-person or virtual sessions
* Expert-led training
* Practical content tailored to real-world applications

If you're looking to sharpen your skills or explore new areas professionally, feel free to reply to this email or contact me directly at 08120288047 or 08101786422 to learn more.

DOWNLOAD OUR PROGRAM BROCHURE: ${attachmentLink}

Looking forward to supporting your growth journey!

Warm regards,
Esther Delano
Business Development Manager
EduBridge Academy`;

    // Send email
    try {
      MailApp.sendEmail({
        to: email,
        subject: "Upskill with Practical, Career-Boosting Training Opportunities",
        htmlBody: htmlBody,
        body: plainBody
      });
      
      // Update status in spreadsheet
      sheet.getRange(i + 1, statusSentCol + 1).setValue("Sent");
      sheet.getRange(i + 1, statusSentCol + 1).setNote(new Date().toString());
      
      // Store tracking ID in script properties
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty(trackingId, JSON.stringify({
        email: email,
        row: i + 1,
        sent: new Date().toString()
      }));
      
      // Log success
      Logger.log(`Email sent successfully to ${firstName} (${email})`);
    } catch (error) {
      // Update status in spreadsheet
      sheet.getRange(i + 1, statusSentCol + 1).setValue("Failed");
      sheet.getRange(i + 1, statusSentCol + 1).setNote(error.toString());
      
      // Log error
      Logger.log(`Failed to send email to ${firstName} (${email}): ${error.toString()}`);
    }
  }
}

// This function handles email open tracking
function doGet(e) {
  const trackingId = e.parameter.trackingId;
  const email = decodeURIComponent(e.parameter.email);
  
  if (trackingId && email) {
    try {
      // Get tracking data
      const scriptProperties = PropertiesService.getScriptProperties();
      const trackingData = JSON.parse(scriptProperties.getProperty(trackingId) || "{}");
      
      if (trackingData.row) {
        // Update spreadsheet
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const emailOpenedCol = headers.indexOf('Email Opened');
        
        if (emailOpenedCol !== -1) {
          sheet.getRange(trackingData.row, emailOpenedCol + 1).setValue("Opened");
          sheet.getRange(trackingData.row, emailOpenedCol + 1).setNote(new Date().toString());
        }
      }
    } catch (error) {
      Logger.log(`Error tracking email open: ${error.toString()}`);
    }
  }
  
  // Return a 1x1 transparent GIF
  return ContentService.createTextOutput(
    "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"
  )
  .setMimeType(ContentService.MimeType.IMAGE_GIF);
}