# Gyanoday Vidyalaya Parent Messaging Repository

This repository contains two Python scripts designed to automate sending bulk text and image messages to the parents of Gyanoday Vidyalaya using WhatsApp.

Features:

Sends bulk messages containing both text and images.
Leverages the power of PyWhatKit for WhatsApp automation.
Employs Openpyxl for efficient data retrieval from Excel spreadsheets.
Requirements:

Python (version 3.x recommended)
PyWhatKit library: pip install PyWhatKit
Openpyxl library: pip install openpyxl
Getting Started:

Install dependencies: Open your terminal or command prompt and run the following commands:

Bash
pip install PyWhatKit
pip install openpyxl
Use code with caution.

 Configure scripts:

send_message_with_image.py:
Update the phone_numbers list with the WhatsApp phone numbers (in international format, including country code) of the parents.
Replace the placeholder path in image_path with the actual path to your image file.
Modify message to contain the text content you want to send along with the image.
send_message_only.py:
Update the phone_numbers list with the WhatsApp phone numbers of the parents.
Customize message to include the text you intend to send.
Data source (optional):

If you're using an Excel spreadsheet to manage parent data, make sure the scripts can access it. Ensure the phone numbers and message content are correctly formatted within the spreadsheet. Adapt the code to read from the relevant columns in your Excel file.
Run the scripts:

Navigate to the directory containing the scripts in your terminal or command prompt.

Execute the desired script using Python:

Bash
python send_message_with_image.py   # For messages with images
python send_message_only.py        # For text-only messages
Use code with caution.

 Important Notes:

WhatsApp Web: Ensure you have WhatsApp Web enabled on your computer and that your phone is connected to the internet for the initial login. Subsequent messages can be sent without the phone being nearby.
Safety & Privacy: Be mindful of privacy regulations when sending bulk messages. Obtain consent from parents before adding them to the recipient list.
Spam Prevention: Sending a large number of messages at once could trigger WhatsApp's spam detection. Consider sending messages in batches or with a delay between each message to avoid this.
Error Handling: The scripts may encounter errors due to invalid phone numbers or other issues. Consider adding error handling mechanisms to gracefully handle such situations.
Additional Considerations:

Personalize messages: For a better user experience, you might consider personalizing messages by including the parent's child's name or other relevant details. You can achieve this by reading data from additional columns in your Excel spreadsheet.
Schedule messages: If you need to send messages at specific times, explore libraries like schedule to create timed tasks within your scripts.
Logging: Implement logging mechanisms to track the success or failure of each message sent for better record-keeping and troubleshooting.
