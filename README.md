I created this automation because I needed a more efficient way to respond to students. Previously, the entire process was done manually, which was time-consuming and error-prone. So, I decided to automate it.

This Python script automates the following tasks:

Log in to a web-based chat system using Selenium and predefined credentials.

Navigate through chat pages and extract chat data (name, date, origin).

Filter chats based on a time limit (e.g., only chats from the last day).

Open each chat, extract the user's email (if available), and store the relevant information.

Send follow-up emails via the Gmail API to users from a specific origin ("EJA SEED") — unless it's associated with "SESI".

Save the collected data (with emails) in an Excel file formatted as a table.

Delete the Gmail token (for security) and close the browser at the end.
