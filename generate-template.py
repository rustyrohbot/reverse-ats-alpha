import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Create a new workbook
wb = Workbook()

# Companies data
companies_data = [
    {"companyID": 1, "name": "Alpha Company", "description": "Enterprise software solutions", "url": "https://alphacompany.example.com", "hqCity": "Philadelphia", "hqState": "PA"},
    {"companyID": 2, "name": "Beta Industries", "description": "B2B SaaS platform provider", "url": "https://betaindustries.example.com", "hqCity": "New York", "hqState": "NY"},
    {"companyID": 3, "name": "Gamma Inc", "description": "AI and machine learning startup", "url": "https://gammainc.example.com", "hqCity": "San Francisco", "hqState": "CA"},
    {"companyID": 4, "name": "Delta Partners", "description": "Cybersecurity consulting firm", "url": "https://delta-partners.example.com", "hqCity": "Boston", "hqState": "MA"},
    {"companyID": 5, "name": "Epsilon Technologies", "description": "Cloud infrastructure provider", "url": "https://epsilontech.example.com", "hqCity": "Austin", "hqState": "TX"},
    {"companyID": 6, "name": "Zeta Solutions", "description": "Enterprise software development", "url": "https://zetasolutions.example.com", "hqCity": "Chicago", "hqState": "IL"},
    {"companyID": 7, "name": "Omega Systems LLC", "description": "Healthcare data analytics", "url": "https://omegasystems.example.com", "hqCity": "Seattle", "hqState": "WA"},
    {"companyID": 8, "name": "Sigma Group", "description": "Financial technology services", "url": "https://sigmagroup.example.com", "hqCity": "Denver", "hqState": "CO"}
]

# Roles data (shortened descriptions for clarity)
roles_data = [
    {"roleID": 1, "companyID": 1, "companyName": "Alpha Company", "name": "Senior Software Engineer",
     "url": "https://careers.alphacompany.example.com/jobs/1",
     "description": "We're looking for an experienced software engineer with 5+ years in cloud architecture",
     "cover": None, "applied": "workday", "date": "2025-04-08",
     "postedRangeMin": 130047, "postedRangeMax": 190736, "equity": False,
     "workCity": "Remote", "workState": "Remote", "location": "Remote",
     "status": "Applied", "discovery": "recruiter"},

    {"roleID": 2, "companyID": 1, "companyName": "Alpha Company", "name": "Engineering Manager I, DevOps",
     "url": "https://careers.alphacompany.example.com/jobs/2",
     "description": "Leading our DevOps transformation initiative",
     "cover": None, "applied": "workday", "date": "2025-04-08",
     "postedRangeMin": 128382, "postedRangeMax": 188294, "equity": False,
     "workCity": "Remote", "workState": "Remote", "location": "Remote",
     "status": "Applied", "discovery": "recruiter"},

    {"roleID": 3, "companyID": 2, "companyName": "Beta Industries", "name": "Sales Engineer",
     "url": "https://careers.betaindustries.example.com/jobs/3",
     "description": "Technical sales role bridging engineering and customers",
     "cover": "Cover letter text here", "applied": "company", "date": "2025-04-08",
     "postedRangeMin": 135000, "postedRangeMax": 150000, "equity": False,
     "workCity": "Remote", "workState": "Remote", "location": "Remote",
     "status": "Applied", "discovery": "wellfound"},

    {"roleID": 4, "companyID": 3, "companyName": "Gamma Inc", "name": "Full Stack Developer",
     "url": "https://careers.gammainc.example.com/jobs/4",
     "description": "We're looking for a full-stack developer with React and Node experience",
     "cover": None, "applied": "linkedin", "date": "2025-03-15",
     "postedRangeMin": 110000, "postedRangeMax": 145000, "equity": True,
     "workCity": "San Francisco", "workState": "CA", "location": "San Francisco, CA",
     "status": "Screening", "discovery": "linkedin"},

    {"roleID": 5, "companyID": 4, "companyName": "Delta Partners", "name": "DevOps Engineer",
     "url": "https://careers.deltapartners.example.com/jobs/5",
     "description": "Cloud infrastructure and CI/CD pipeline expertise needed",
     "cover": None, "applied": "indeed", "date": "2025-03-20",
     "postedRangeMin": 125000, "postedRangeMax": 165000, "equity": False,
     "workCity": "Remote", "workState": "Remote", "location": "Remote",
     "status": "Interview", "discovery": "job board"},

    {"roleID": 6, "companyID": 5, "companyName": "Epsilon Technologies", "name": "Data Scientist",
     "url": "https://careers.epsilontech.example.com/jobs/6",
     "description": "Experience with machine learning and large datasets required",
     "cover": None, "applied": "company website", "date": "2025-03-25",
     "postedRangeMin": 130000, "postedRangeMax": 170000, "equity": True,
     "workCity": "Austin", "workState": "TX", "location": "Austin, TX",
     "status": "Offer", "discovery": "referral"},

    {"roleID": 7, "companyID": 6, "companyName": "Zeta Solutions", "name": "Product Manager",
     "url": "https://careers.zetasolutions.example.com/jobs/7",
     "description": "Technical product manager for our enterprise platform",
     "cover": None, "applied": "referral", "date": "2025-02-28",
     "postedRangeMin": 120000, "postedRangeMax": 160000, "equity": False,
     "workCity": "Remote", "workState": "Remote", "location": "Remote",
     "status": "Rejected", "discovery": "linkedin"},

    {"roleID": 8, "companyID": 7, "companyName": "Omega Systems LLC", "name": "UX/UI Designer",
     "url": "https://careers.omegasystems.example.com/jobs/8",
     "description": "Design healthcare dashboards and user interfaces",
     "cover": None, "applied": "linkedin", "date": "2025-03-10",
     "postedRangeMin": 95000, "postedRangeMax": 125000, "equity": False,
     "workCity": "Seattle", "workState": "WA", "location": "Seattle, WA",
     "status": "Withdrawn", "discovery": "indeed"},

    {"roleID": 9, "companyID": 8, "companyName": "Sigma Group", "name": "Machine Learning Engineer",
     "url": "https://careers.sigmagroup.example.com/jobs/9",
     "description": "Developing ML algorithms for financial data analysis",
     "cover": None, "applied": "indeed", "date": "2025-04-01",
     "postedRangeMin": 140000, "postedRangeMax": 180000, "equity": True,
     "workCity": "Remote", "workState": "Remote", "location": "Remote",
     "status": "Screening", "discovery": "recruiter"},

    {"roleID": 10, "companyID": 2, "companyName": "Beta Industries", "name": "Frontend Developer",
     "url": "https://careers.betaindustries.example.com/jobs/10",
     "description": "React expertise needed for our consumer-facing applications",
     "cover": None, "applied": "company website", "date": "2025-03-05",
     "postedRangeMin": 100000, "postedRangeMax": 140000, "equity": False,
     "workCity": "New York", "workState": "NY", "location": "New York, NY",
     "status": "Interview", "discovery": "job board"}
]

# Contacts data (completely fictional)
contacts_data = [
    {"contactID": 1, "companyID": 1, "companyName": "Alpha Company", "firstName": "Farrah", "lastName": "Lundquist",
     "role": "Technical Recruiter", "email": "farrah.lundquist@alphacompany.example.com", "phone": "5105557123",
     "linkedin": "https://www.linkedin.com/in/farrah-lundquist-8832/", "notes": "Had a good initial chat about the role"},

    {"contactID": 2, "companyID": 4, "companyName": "Delta Partners", "firstName": "Miguel", "lastName": "Sanchez",
     "role": "Lead Technical Recruiter", "email": "miguel.sanchez@deltapartners.example.com", "phone": "6175559034",
     "linkedin": "https://www.linkedin.com/in/miguel-sanchez-5217/", "notes": "Very responsive to emails"},

    {"contactID": 3, "companyID": 2, "companyName": "Beta Industries", "firstName": "Emma", "lastName": "Johnson",
     "role": "Talent Acquisition Specialist", "email": "emma.johnson@betaindustries.example.com", "phone": "2125557890",
     "linkedin": "https://www.linkedin.com/in/emma-johnson-5462/", "notes": "Met Emma during the initial screening call"},

    {"contactID": 4, "companyID": 3, "companyName": "Gamma Inc", "firstName": "Noah", "lastName": "Williams",
     "role": "Technical Recruiter", "email": "noah.williams@gammainc.example.com", "phone": "4155552341",
     "linkedin": "https://www.linkedin.com/in/noah-williams-7821/", "notes": "Referred by a colleague"},

    {"contactID": 5, "companyID": 5, "companyName": "Epsilon Technologies", "firstName": "Olivia", "lastName": "Brown",
     "role": "HR Manager", "email": "olivia.brown@epsilontech.example.com", "phone": "5125559876",
     "linkedin": "https://www.linkedin.com/in/olivia-brown-3245/", "notes": "Connected at a virtual job fair"},

    {"contactID": 6, "companyID": 6, "companyName": "Zeta Solutions", "firstName": "Liam", "lastName": "Jones",
     "role": "Hiring Manager", "email": "liam.jones@zetasolutions.example.com", "phone": "3125554321",
     "linkedin": "https://www.linkedin.com/in/liam-jones-9012/", "notes": "Engineering director who manages the team"},

    {"contactID": 7, "companyID": 7, "companyName": "Omega Systems LLC", "firstName": "Ava", "lastName": "Garcia",
     "role": "Director of Talent", "email": "ava.garcia@omegasystems.example.com", "phone": "2065551234",
     "linkedin": "https://www.linkedin.com/in/ava-garcia-1289/", "notes": "Manages all technical recruiting"},

    {"contactID": 8, "companyID": 8, "companyName": "Sigma Group", "firstName": "William", "lastName": "Miller",
     "role": "Technical Sourcer", "email": "william.miller@sigmagroup.example.com", "phone": "3035557890",
     "linkedin": "https://www.linkedin.com/in/william-miller-5678/", "notes": "Initial outreach via LinkedIn"}
]

# Interviews data
interviews_data = [
    {"interviewID": 1, "roleID": 4, "date": "2025-03-18",
     "notes": "Initial conversation about my background and the role. Discussed salary expectations and timeline.",
     "type": "Phone Screen"},

    {"interviewID": 2, "roleID": 5, "date": "2025-03-24",
     "notes": "Covered algorithms, data structures, and system design questions. Asked about my experience with specific technologies.",
     "type": "Technical"},

    {"interviewID": 3, "roleID": 5, "date": "2025-03-27",
     "notes": "Met with team members to discuss work style, collaboration, and company values.",
     "type": "Culture Fit"},

    {"interviewID": 4, "roleID": 6, "date": "2025-03-30",
     "notes": "Asked about past experiences handling difficult situations, teamwork, and leadership examples.",
     "type": "Behavioral"},

    {"interviewID": 5, "roleID": 6, "date": "2025-04-02",
     "notes": "Designed a scalable architecture for their main product. Discussed trade-offs and optimizations.",
     "type": "System Design"},

    {"interviewID": 6, "roleID": 6, "date": "2025-04-05",
     "notes": "Met with senior leadership to discuss long-term goals and vision for the role.",
     "type": "Final Round"},

    {"interviewID": 7, "roleID": 9, "date": "2025-04-06",
     "notes": "Initial conversation about my background and the role. Discussed salary expectations and timeline.",
     "type": "Phone Screen"},

    {"interviewID": 8, "roleID": 10, "date": "2025-03-12",
     "notes": "Deep dive into my experience and how it relates to current team needs.",
     "type": "Hiring Manager"},

    {"interviewID": 9, "roleID": 10, "date": "2025-03-18",
     "notes": "Covered algorithms, data structures, and system design questions. Asked about my experience with specific technologies.",
     "type": "Technical"}
]

# InterviewsContacts data
interviews_contacts_data = [
    {"interviewContactId": 1, "interviewId": 1, "contactId": 4},
    {"interviewContactId": 2, "interviewId": 2, "contactId": 2},
    {"interviewContactId": 3, "interviewId": 3, "contactId": 2},
    {"interviewContactId": 4, "interviewId": 4, "contactId": 5},
    {"interviewContactId": 5, "interviewId": 5, "contactId": 5},
    {"interviewContactId": 6, "interviewId": 6, "contactId": 5},
    {"interviewContactId": 7, "interviewId": 7, "contactId": 8},
    {"interviewContactId": 8, "interviewId": 8, "contactId": 3},
    {"interviewContactId": 9, "interviewId": 9, "contactId": 3}
]

# Function to create and populate worksheets
def create_worksheet(df, sheet_name):
    # Create a new worksheet or select existing one
    if sheet_name == "Companies":
        ws = wb.active
        ws.title = sheet_name
    else:
        ws = wb.create_sheet(sheet_name)

    # Write the data
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    print(f"Populated {sheet_name} worksheet with {len(df)} records")

# Convert data to DataFrames
companies_df = pd.DataFrame(companies_data)
roles_df = pd.DataFrame(roles_data)
contacts_df = pd.DataFrame(contacts_data)
interviews_df = pd.DataFrame(interviews_data)
interviews_contacts_df = pd.DataFrame(interviews_contacts_data)

# Create the worksheets
create_worksheet(companies_df, "Companies")
create_worksheet(roles_df, "Roles")
create_worksheet(contacts_df, "Contacts")
create_worksheet(interviews_df, "Interviews")
create_worksheet(interviews_contacts_df, "InterviewsContacts")

# Save the workbook
output_filename = "Reverse_ATS_Template.xlsx"
wb.save(output_filename)

print(f"Excel file '{output_filename}' has been created successfully!")