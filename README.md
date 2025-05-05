# Reverse ATS - ALPHA

A prototype or alpha version of a reverse-ATS, a tool to organize track what roles you've applied to, the companies those roles belong to, and interviews related to the role, and relevant contacts for each company.

## Overview

My default method to organize a job hunt was a mess of emails, calendar invites, along with notes scattered across paper, Notion, and Obsidian. I needed to bring some order to the chaos. I didn't want to immediately build another SaaS, so I'm starting with the basics, a collection of spreadsheets. Stay tuned for new iterations.

For obvious privacy reasons, I'm not sharing the workbook I'm using to track my job interviews. This repo is an approximation of it for anyone who might want to copy the system. There is a `python` script that will generate an Excel that you can use as a template. The file has also been committed and pushed to this repo as well.

The workbook contains the following sheets:
- **Companies**: Companies you've applied to
- **Roles**: Roles at Companies
- **Contacts**: Contacts at Companies
- **Interviews**: Interviews for each Role
- **InterviewsContacts**: Mapping Contacts to Interviews

## Requirements (for script)

- Python 3.8+
- Dependencies listed in `requirements.txt`:
  - pandas
  - openpyxl

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/rustyrohbot/reverse-ats-alpha.git
   cd reverse-ats-alpha
   ```

2. Create and activate a virtual environment (recommended):
   ```
   python3 -m venv venv
   source venv/bin/activate  # On Windows, use: venv\Scripts\activate
   ```

3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

Run the script to generate the Excel template:

```
python generate-template.py
```

This will create a file named `Reverse_ATS_Template.xlsx` in the current directory.

### Uploading to Google Sheets

The generated Excel file can be uploaded to Google Sheets:

1. Go to [Google Drive](https://drive.google.com)
2. Click **New** > **File upload**
3. Select the generated Excel file
4. Once uploaded, open it with Google Sheets

## Privacy

The generated template contains entirely fictional company names, recruiter information, and job details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.