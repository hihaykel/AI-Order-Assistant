# 🤖 AI-Powered Customer Support Assistant (Google Sheets + Gemini API)

An automated customer support agent built inside Google Sheets using **Google Apps Script** and the **Gemini API**. This system automatically analyzes incoming customer emails, references them against custom company policies, and decides whether to send an instant professional response or escalate the request to human review.

---

## 🚀 Features

* **Automated Ticket Classification:** Instantly detects the customer's intent (e.g., order tracking, refund requests).
* **Smart Policy Enforcement:** Cross-checks customer messages with predefined company rules (e.g., 7-day refund policy).
* **Auto-Reply & Delivery:** Automatically sends high-quality, professional email responses using `GmailApp`.
* **Human-in-the-Loop Escalation:** Flags complex or out-of-policy requests as `Pending Human Review` and creates a draft note without sending automated emails.
* **100% Autopilot:** Uses an `On edit` Trigger to process incoming messages hands-free.

---

## 🛠️ Setup Instructions

### 1. Google Sheets Setup
1. Create a new Google Sheet.
2. Define your columns (e.g., `Customer Email`, `Message`, `Status`, `AI Reply`).
3. Define your **Company Rules** (e.g., Order Tracking, Damage Claims, 7-day Exchanges) to train your AI model.

### 2. Apps Script Integration
1. Inside your Google Sheet, navigate to **Extensions > Apps Script**.
2. Open the default `Code.gs` file.
3. Delete any existing template code.
4. Copy the entire script from the `Code.gs` file in this repository and paste it into your editor.

### 3. Get a Gemini API Key
1. Go to [Google AI Studio](https://aistudio.google.com/).
2. Click on **Create API Key**.
3. Copy your free key and paste it at the top of your `Code.gs` file in Google Apps Script:
   ```javascript
   const GEMINI_API_KEY = "YOUR_API_KEY_HERE";
