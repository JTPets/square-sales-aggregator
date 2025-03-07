📊 Square Sales Aggregator
A Google Apps Script that integrates with the Square API to fetch and aggregate 91-day sales data, organize it in Google Sheets, and send email notifications upon completion. The script runs manually or automatically every 3 hours using time-based triggers.

🚀 Features
Fetches 91-day aggregated sales from Square for all locations.
Stores sales data in a dedicated Google Sheet (Sales-Aggregated).
Sends email notifications upon success or failure.
Provides a custom menu in Google Sheets for easy interaction.
Supports an automated 3-hour trigger to refresh data periodically.
🛠️ Installation & Setup
1️⃣ Copy the Script
Open your Google Sheet.
Click on Extensions > Apps Script.
Copy and paste the script into the editor.
Save the project.
2️⃣ Set Up API & Permissions
Go to Square Developer Dashboard and create a new app.
Get your Square API access token under the Credentials section.
In your Google Sheet, go to Square API > Set API Key and enter your Square API token.
(Optional) Set a notification email under Square API > Set Email Address to receive updates.
3️⃣ Run the Script
Click Square API > Start Aggregated Sales Processing to manually fetch and process data.
Click Square API > Set 3-Hour Timer to enable auto-refresh every 3 hours.
📌 Permissions Required
Google Sheets (to store sales data)
Google Apps Script Triggers (for automation)
Square API (to fetch sales data)
Google MailApp (to send email notifications)
🔧 Troubleshooting
Access Token Issues: Ensure your API token is active and correctly set via the script menu.
No Sales Data Found: Make sure you have completed orders in Square for the last 91 days.
Automation Not Running: Check your Apps Script Triggers in Extensions > Apps Script > Triggers.
📜 License
This project is licensed under the MIT License – feel free to use and modify it.
