# Job Search Assistant (India Edition)

An automated, CLI-based job search assistant and dynamic resume tailor designed for developers in India. 

The application searches major tech job boards (Naukri, Instahyre, LinkedIn, etc.), scores matching jobs against your profile, dynamically customizes your resume and cover letter for each specific job description, and logs all applications in an Excel-based tracker.

---

## Key Features

*   🔍 **Resilient Job Searcher**: Scrapes platforms using DuckDuckGo HTML and Lite fallbacks, rotating user-agents, and delay jitter to bypass strict anti-bot protections.
*   📊 **Relevance & Fitness Scorer**: Holistic fit calculations (0-100%) based on years of experience (YOE), technical skills overlap, seniority alignment, and project matching.
*   🧠 **Dynamic Document Tailoring**: Uses `python-docx` to dynamically re-order resume bullet points, achievements, and technical skill categories based on keywords found in the target job description.
*   📄 **Dynamic Cover Letters**: Automatically generates personalized, human-sounding cover letters (under 300 words) focusing on the exact technologies in the JD.
*   📈 **Safe Application Tracker**: Maintains a structured `job_tracker.xlsx` sheet with a permission-lock warning handler to prevent Excel crashes.

---

## Tech Stack

*   **Language**: Python 3.13+
*   **Libraries**: `python-docx` (Word generation), `openpyxl` (Excel automation), `requests` & `BeautifulSoup4` (HTML parsing)

---

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/suraj-kanse/job-search-assistant.git
   cd job-search-assistant
   ```

2. Install dependencies:
   ```bash
   pip install python-docx openpyxl beautifulsoup4 requests
   ```

3. Configure your details in `profile.json`.

---

## Usage Guide

### 1. Perform Automated Search
Scrape portals, evaluate fitness scores, generate customized documents for matches, and add them to the Excel sheet:
```bash
python main.py --search
```

### 2. Tailor for a Custom Job Description
If you find a live job posting online, copy the job description text and run the following to instantly score the role and output tailored docs:
```bash
python main.py --custom --company "Company Name" --role "Job Title" --location "Location" --jd "Paste the job description text here"
```

### 3. Check Application Status
Print a summary of your tracked applications directly on the console:
```bash
python main.py --status
```

### 4. Update Application Status
Update the status of a specific job in your Excel tracker (e.g. to "Applied", "Interview", etc.):
```bash
python main.py --update <JOB_ID> <STATUS>
```
