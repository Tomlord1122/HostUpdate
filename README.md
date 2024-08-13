# Start Generation Here
# README.md

# Outlook Email Export Application

This is an application developed using Python and Tkinter, designed to export and summarize emails from Outlook. The application uses OpenAI's API to generate summaries of emails and provides filtering and date range selection features.

## Features

- **Collect Emails**: Collect today's emails from Outlook.
- **Filter Emails**: Filter emails based on subject or content.
- **Date Range Selection**: Select a specific date range to collect emails.
- **Export Emails**: Export the filtered emails.
- **Summarize Emails**: Generate summaries of emails using OpenAI's API.

## Environment Setup

1. Ensure Python 3.x is installed.
2. Install the required libraries:
   ```bash
   pip install tkinter tkcalendar pywin32 openai python-dotenv
   ```
3. Create a `.env` file and add your OpenAI API key:
   ```
   OPENAI_API_KEY=your_api_key_here
   ```

## Usage

1. Run the `app.py` file:
   ```bash
   python app.py
   ```
2. In the application, select the date range and filtering criteria, then click the "Collect" button to retrieve emails.
3. Use the "Summarize" button to generate summaries of the emails.

## Notes

- Ensure the Outlook application is installed and configured correctly.
- Please adhere to OpenAI's terms of use.

# Copyright Information

The copyright of this application belongs to the developer. Please follow the relevant open-source agreements.

# Contact Information

If you have any questions or suggestions, please contact the developer.

# Version

- Version 1.0.0
# End Generation Here
