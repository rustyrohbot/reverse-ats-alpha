from openpyxl import Workbook

# Define headers
sheet_headers = {
    "Companies": [
        "companyID", "name", "description", "url", "hqCity", "hqState"
    ],
    "Roles": [
        "roleID", "companyID", "name", "url", "description", "coverLetter",
        "applicationLocation", "appliedDate", "closedDate", "postedRangeMin",
        "postedRangeMax", "equity", "workCity", "workState", "location",
        "status", "discovery", "referral", "notes"
    ],
    "Interviews": [
        "interviewID", "roleID", "date", "start", "end", "notes", "type"
    ],
    "Contacts": [
        "contactID", "companyID", "firstName", "lastName", "role", "email",
        "phone", "linkedin", "notes"
    ],
    "InterviewsContacts": [
        "interviewsContactId", "interviewId", "contactId"
    ]
}

# Sample starter data showing relationships
sheet_data = {
    "Companies": [
        [1, "Acme Corp", "AI-driven analytics", "https://acme.example.com", "San Francisco", "CA"]
    ],
    "Roles": [
        [1, 1, "Machine Learning Engineer", "https://acme.example.com/careers/1",
         "Develop ML models for enterprise clients", "", "workday",
         "2025-05-01", "", 130000, 160000, True, "Remote", "Remote", "Remote",
         "Applied", "linkedin", False, "High growth team"]
    ],
    "Interviews": [
        [1, 1, "2025-05-03", "10:00", "10:30", "Initial phone screen with recruiter", "Phone Screen"]
    ],
    "Contacts": [
        [1, 1, "Jane", "Doe", "Technical Recruiter", "jane.doe@acme.example.com",
         "5551234567", "https://linkedin.com/in/janedoe", "Initial contact via LinkedIn"]
    ],
    "InterviewsContacts": [
        [1, 1, 1]
    ]
}

# Create workbook
wb = Workbook()

# Write each sheet with headers and example rows
for idx, (sheet_name, headers) in enumerate(sheet_headers.items()):
    ws = wb.active if idx == 0 else wb.create_sheet(title=sheet_name)
    ws.title = sheet_name
    ws.append(headers)
    for row in sheet_data[sheet_name]:
        ws.append(row)

# Save locally
wb.save("reverse-ats.xlsx")
print("Template generated: reverse-ats.xlsx")
