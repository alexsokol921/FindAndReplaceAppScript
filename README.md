# FindAndReplaceAppScript
This Google Apps Script automates creating personalized Google Docs from data in a Google Sheet.
It copies a .docx template stored in Google Drive, converts it to a Google Doc, and replaces the placeholders (Name) and (Courses) for each row in your sheet.

A new, customized Google Doc is created for every row and saved to a folder of your choice in Google Drive.

---

## ✨ Features

- Reads data from a Google Sheet (Name and Courses columns)

- Uses a .docx template stored in Google Drive

- Converts the template automatically to Google Docs

- Replaces placeholders (Name) and (Courses)

- Creates one personalized document per row
---

## 🧩 Setup Instructions
### 1️⃣ Prepare Your Files
#### Template (.docx)

1. Create a Word or Google Docs file containing the placeholders (Name) and (Courses).

Example:

Dear (Name),

You are registered for the following courses:
(Courses)


2. Upload this .docx file to Google Drive.

---

#### Google Sheet

1. Create a new Google Sheet with columns like this:

Name	Courses
Alice	Math, Science
Bob	English, Art

The first row must contain the headers Name and Courses.

---

### 2️⃣ Add the Script to Google Sheets

1. In your Google Sheet, go to Extensions → Apps Script.

2. Delete any default code and paste the full script from code.gs
.

3. Update the configuration section near the top of the script:

const TEMPLATE_FILE_ID = 'YOUR_TEMPLATE_FILE_ID_HERE';
const OUTPUT_FOLDER_ID = 'YOUR_OUTPUT_FOLDER_ID_HERE';
const SHEET_NAME = 'Sheet1'; // or whatever your sheet tab name is

---

### 3️⃣ Enable the Google Drive v2 API

This script uses the Drive API to convert .docx files into Google Docs.

1. In the Apps Script editor, click the Services (+) icon on the left sidebar.

2. Scroll down, find Drive API, and click Add.

You should now see:

Drive   Google Drive API   v2
---

### 4️⃣ Get Your IDs

You’ll need two IDs from Google Drive: the template file ID and the output folder ID.
---

#### 📄 Template File ID

1. Open your .docx file in Google Drive.

The URL will look like:

https://drive.google.com/file/d/1AbCdEfGhIJkLmNopQr/view


The section between /d/ and /view is your template file ID:

1AbCdEfGhIJkLmNopQr
---

#### 📁 Folder ID

1. Open the Drive folder where you want new documents saved.

The URL will look like:

https://drive.google.com/drive/folders/2XyZaBcDeFGhIJkLm


The section after /folders/ is your folder ID:

2XyZaBcDeFGhIJkLm
---

### 5️⃣ Run the Script

1. Click the ▶ Run button in Apps Script.

2. On the first run, you’ll be asked to authorize the script to access your Drive and Sheets.

Once authorized, the script will:

Read each row of your sheet

Copy and convert the .docx template

Replace (Name) and (Courses)

Save the personalized file in your chosen folder
---

### 6️⃣ View the Results

1. Open your Drive output folder — you’ll find one document per row, for example:

Alice - Document
Bob - Document
---
## 🧰 Troubleshooting
### Issue	Solution
“The document is inaccessible”	Wait a few seconds and re-run — Drive conversion can take a moment.
Drive.Files.insert is not a function	The Drive API (v2) isn’t enabled — double-check Step 3.
Blank replacements	Ensure your sheet headers are exactly Name and Courses.
---
## 🪶 License

MIT License — feel free to fork, adapt, and extend.
---
