Google Sheets AI Agent (Gemini-Powered)
This project transforms a standard Google Sheet into a "Smart Sheet" using Google Apps Script and the Gemini API. It allows users to generate complex formulas, create pivot dashboards, and summarize data using natural language commands.

✨ Features
Formula Assistant: Type what you need (e.g., "Find the salary for Riya") and the AI writes the XLOOKUP or SUMIF formula for you.

Automated Dashboards: Automatically generates Pivot Tables and Charts (Bar/Line) based on your column headers.

Header Awareness: The agent reads your first row to understand your data context (e.g., "Salary", "Department", "Location").

Custom Menu: Adds a "🤖 AI Agent" menu directly to the Google Sheets toolbar.

🚀 Setup Instructions
1. Get a Gemini API Key
Visit the Google AI Studio.

Generate a new API Key.

2. Install the Script
Open a Google Sheet.

Go to Extensions > Apps Script.

Copy the code from Code.gs in this repository and paste it into the editor.

Crucial: Replace the API_URL placeholder in the script with your actual API key or store it in Project Settings > Script Properties.

3. Usage
Refresh your Google Sheet.

Click the 🤖 AI Agent menu.

Try "Formula Assistant" and ask: "Get the total salary for Marketing" or "Find the average age in Column E".

🛠️ Tech Stack
Language: JavaScript (Google Apps Script)

AI Model: Gemini 3 Flash

APIs: Google Sheets API, UrlFetchApp

💡 One last suggestion for your GitHub:
Since you are using an API key, do not paste your real key into the code on GitHub! It’s best to keep a placeholder like const API_KEY = "YOUR_KEY_HERE"; so others can insert their own.
