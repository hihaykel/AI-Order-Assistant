// ====== CONFIGURATION ======
const API_KEY = "Your_API_KEY";
const GEMINI_URL = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=" + API_KEY;

/**
 * Executed via Installable Trigger (On Edit)
 */
function processCustomerMessage(e) {
  if (!e || !e.range) return;
  
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Triggers only when modifying Column B (Customer Message) from Row 2 downwards
  if (range.getColumn() !== 2 || range.getRow() < 2) return;
  
  const row = range.getRow();
  const userMessage = range.getValue().toString().trim();
  
  // CLEAR COLUMNS C, D, E IF MESSAGE IS DELETED
  if (userMessage === "") {
    sheet.getRange(row, 3, 1, 3).clearContent();
    return;
  }

  const customerEmail = sheet.getRange(row, 1).getValue().toString().trim();
  
  // Status indicator
  sheet.getRange(row, 3).setValue("🤖 Analyzing...");
  SpreadsheetApp.flush();
  
  // SYSTEM PROMPT WITH STRICT PROTOCOL RULES
  const systemPrompt = `
  You are an automated support AI agent for an e-commerce store.
  You MUST strictly follow the COMPANY POLICIES below:

  1. EXCHANGE POLICY:
     - 7-day exchange window for unused items in original packaging.
     - Customer must contact support to initiate.
  2. DEFECTIVE ITEMS:
     - Must be reported within 48 hours of delivery.
     - Requires photos and order details.
     - Offer replacement or refund upon verification.
  3. SHIPPING STATUS:
     - Remind customer to use tracking link from confirmation email.

  CRITICAL RULE FOR HUMAN ESCALATION:
  - If the query is unclear, aggressive, or requests something outside the above 3 rules (e.g., custom requests, late returns, payment issues), set status to "Escalated to Human Support".

  Analyze this message: "${userMessage}"

  Return ONLY a raw JSON string (no markdown, no \`\`\`json) with these keys:
  - "status": Choice between ("Exchange Approved", "Defective Report - Photos Needed", "Shipping Tracking Info", or "Escalated to Human Support")
  - "reply": The exact email reply to send to the customer.
  `;
                 
  const payload = {
    "contents": [{ "parts": [{ "text": systemPrompt }] }]
  };
  
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    const response = UrlFetchApp.fetch(GEMINI_URL, options);
    const jsonResponse = JSON.parse(response.getContentText());
    let rawText = jsonResponse.candidates[0].content.parts[0].text.trim();
    
    // Clean JSON markdown formatting if present
    if (rawText.includes("```")) {
      rawText = rawText.replace(/```json/g, "").replace(/```/g, "").trim();
    }
    
    const result = JSON.parse(rawText);
    
    // Update AI Status and Reply Draft columns
    sheet.getRange(row, 3).setValue(result.status);
    sheet.getRange(row, 4).setValue(result.reply);
    
    // Handle action and notification logic
    if (result.status === "Escalated to Human Support") {
      sheet.getRange(row, 5).setValue("🔴 Pending Human Review");
    } else if (customerEmail && customerEmail.includes("@")) {
      GmailApp.sendEmail(customerEmail, "Support Request Update", result.reply);
      sheet.getRange(row, 5).setValue("✉️ Sent Automatically");
    } else {
      sheet.getRange(row, 5).setValue("⚠️ No Email Provided");
    }
    
  } catch (error) {
    sheet.getRange(row, 3).setValue("⚠️ Error");
    sheet.getRange(row, 4).setValue(error.toString());
  }
}
