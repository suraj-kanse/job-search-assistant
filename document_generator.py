import os
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import json
import zipfile
import shutil
from profile_manager import load_profile

def create_resume(job, profile, output_path):
    """Generates a customized, ATS-optimized resume as a docx file."""
    doc = docx.Document()
    
    # Page setup - Standard 1 inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Base styling
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)

    # 1. Header (Centered Name & Contact)
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(profile["name"].upper())
    run.bold = True
    run.font.size = Pt(22)
    
    contact_p = doc.add_paragraph()
    contact_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact_info = f"{profile['phone']}  |  {profile['email']}  |  {profile['location']['current']}\nLinkedIn: {profile['linkedin']}  |  GitHub: {profile['github']}"
    run = contact_p.add_run(contact_info)
    run.font.size = Pt(10)

    # Helper function for section headings
    def add_section_heading(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(13)
        # Add bottom border/line if possible (using paragraph border is complex, so we will use bold headings and clean spacing)
        
    # 2. Professional Summary (Tailored to Job)
    add_section_heading("Professional Summary")
    summary_p = doc.add_paragraph()
    
    # Custom summary using non-AI-cliché human tone
    role_type = job["title"]
    company = job["company"]
    summary_text = (
        f"Information Technology undergraduate with hands-on experience building web applications "
        f"using Python, Django, React, and Node.js. Developed and deployed full-stack solutions "
        f"featuring REST APIs, secure user authentication, and automated data processing tools. "
        f"Eager to contribute technical skills and build reliable software as a {role_type} at {company}."
    )
    summary_p.add_run(summary_text)

    # 3. Technical Skills
    add_section_heading("Technical Skills")
    skills = profile["skills"]
    for cat, list_skills in skills.items():
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(2)
        run_cat = p.add_run(f"{cat.capitalize()}: ")
        run_cat.bold = True
        p.add_run(", ".join(list_skills))

    # 4. Experience
    add_section_heading("Experience")
    for exp in profile["experience"]:
        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_after = Pt(2)
        run_role = p_title.add_run(f"{exp['role']}  —  {exp['company']}")
        run_role.bold = True
        
        p_dates = doc.add_paragraph()
        p_dates.paragraph_format.space_after = Pt(4)
        run_dates = p_dates.add_run(exp["duration"])
        run_dates.font.italic = True
        
        # Customize bullet points based on JD if needed
        for detail in exp["details"]:
            p_bullet = doc.add_paragraph(style='List Bullet')
            p_bullet.paragraph_format.space_after = Pt(2)
            p_bullet.paragraph_format.left_indent = Inches(0.4)
            p_bullet.add_run(detail)

    # 5. Projects
    add_section_heading("Projects")
    for proj in profile["projects"]:
        p_proj = doc.add_paragraph()
        p_proj.paragraph_format.space_after = Pt(2)
        run_proj = p_proj.add_run(f"{proj['name']}  ({proj['duration']})")
        run_proj.bold = True
        
        p_tech = doc.add_paragraph()
        p_tech.paragraph_format.space_after = Pt(4)
        p_tech.paragraph_format.left_indent = Inches(0.2)
        run_tech = p_tech.add_run("Tech Stack: ")
        run_tech.bold = True
        p_tech.add_run(proj["tech_stack"])
        
        for detail in proj["details"]:
            p_bullet = doc.add_paragraph(style='List Bullet')
            p_bullet.paragraph_format.space_after = Pt(2)
            p_bullet.paragraph_format.left_indent = Inches(0.4)
            p_bullet.add_run(detail)

    # 6. Education
    add_section_heading("Education")
    for edu in profile["education"]:
        p_edu = doc.add_paragraph()
        p_edu.paragraph_format.space_after = Pt(2)
        run_deg = p_edu.add_run(f"{edu['degree']}  —  {edu['college']}")
        run_deg.bold = True
        
        p_details = doc.add_paragraph()
        p_details.paragraph_format.left_indent = Inches(0.2)
        p_details.add_run(f"Duration: {edu['duration']}  |  CGPA: {edu['cgpa']}")

    # 7. Certifications
    add_section_heading("Certifications")
    for cert in profile["certifications"]:
        p_cert = doc.add_paragraph(style='List Bullet')
        p_cert.paragraph_format.space_after = Pt(2)
        p_cert.paragraph_format.left_indent = Inches(0.4)
        p_cert.add_run(cert)

    # Save
    doc.save(output_path)


def create_cover_letter(job, profile, output_path):
    """Generates a customized, human-sounding cover letter as a docx file."""
    doc = docx.Document()
    
    # Page setup
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)

    # Date
    import datetime
    today = datetime.date.today().strftime("%B %d, %Y")
    doc.add_paragraph(today)
    
    # Applicant details
    p_applicant = doc.add_paragraph()
    p_applicant.paragraph_format.space_after = Pt(12)
    p_applicant.add_run(f"{profile['name']}\n{profile['location']['current']}\n{profile['email']} | {profile['phone']}")
    
    # Recipient details
    p_recipient = doc.add_paragraph()
    p_recipient.paragraph_format.space_after = Pt(12)
    p_recipient.add_run(f"Hiring Team\n{job['company']}\n{job['location']}")
    
    # Salutation
    p_salutation = doc.add_paragraph()
    p_salutation.paragraph_format.space_after = Pt(12)
    p_salutation.add_run("Dear Hiring Team,")
    
    # Custom, non-AI-speak body paragraphs under 300 words
    opening = (
        f"I saw your opening for a {job['title']} and wanted to reach out. "
        f"I am a final-year Information Technology student at Amrutvahini College of Engineering, "
        f"and I have been building full-stack web applications with Python, Django, React, and Node.js."
    )
    
    body = (
        f"Recently, I built a student counselling platform for my college using React, Node.js, and MongoDB, "
        f"which handles secure student session logs and support requests using role-based access control. "
        f"I also completed an internship where I built a Resume Skill Gap Analyzer using Django and Python. "
        f"Through these projects, I have learned how to write clean API endpoints, design structured SQL/NoSQL databases, "
        f"and secure applications using JWT and bcrypt. "
        f"Your team's work at {job['company']} aligns perfectly with my background, and I would love to bring my technical skills to the role."
    )
    
    closing = (
        f"Since I am in my final year, I am looking for both immediate internships and graduate roles starting in 2027. "
        f"I am ready to relocate to {job['location']} and am available to start immediately. "
        f"Thank you for your time and consideration. I hope we can discuss how my experience fits this role."
    )
    
    doc.add_paragraph(opening).paragraph_format.space_after = Pt(10)
    doc.add_paragraph(body).paragraph_format.space_after = Pt(10)
    doc.add_paragraph(closing).paragraph_format.space_after = Pt(18)
    
    # Sign off
    p_signoff = doc.add_paragraph()
    p_signoff.add_run(f"Sincerely,\n\n{profile['name']}")
    
    doc.save(output_path)


def generate_application_bundle(jobs, output_dir="dist"):
    """Generates resumes and cover letters for a list of jobs, zipping them by company."""
    profile = load_profile()
    
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)
    
    generated_files = []
    
    for job in jobs:
        company_clean = job["company"].replace(" ", "_").replace("/", "_")
        role_clean = job["title"].replace(" ", "_").replace("/", "_").replace("(", "").replace(")", "")
        job_dir = os.path.join(output_dir, f"{company_clean}_{job['location']}")
        os.makedirs(job_dir, exist_ok=True)
        
        name_clean = profile["name"].replace(" ", "_")
        resume_name = f"{name_clean}_Resume_{company_clean}_{role_clean}.docx"
        cl_name = f"{name_clean}_CoverLetter_{company_clean}.docx"
        
        resume_path = os.path.join(job_dir, resume_name)
        cl_path = os.path.join(job_dir, cl_name)
        
        create_resume(job, profile, resume_path)
        create_cover_letter(job, profile, cl_path)
        
        generated_files.append((job["company"], job_dir))
        print(f"Generated documents for {job['company']} in {job_dir}")
        
    # Zip the entire dist directory
    import datetime
    zip_name = f"{name_clean}_Applications_{datetime.date.today().strftime('%Y%m%d')}.zip"
    zip_path = os.path.join(os.path.dirname(output_dir), zip_name)
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # Save relative path in zip
                rel_path = os.path.relpath(file_path, output_dir)
                zipf.write(file_path, rel_path)
                
    print(f"Bundled all materials into: {zip_path}")
    return zip_path

if __name__ == "__main__":
    from search_engine import run_job_search
    print("Testing document generation on first search result...")
    jobs = run_job_search()
    if jobs:
        zip_res = generate_application_bundle([jobs[0]], output_dir="temp_dist")
        print(f"Success! Zip generated: {zip_res}")
        if os.path.exists("temp_dist"):
            shutil.rmtree("temp_dist")
        if os.path.exists(zip_res):
            os.remove(zip_res)
