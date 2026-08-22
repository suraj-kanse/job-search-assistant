import argparse
import sys
import os
from search_engine import run_job_search
from document_generator import generate_application_bundle
from tracker import add_jobs_to_tracker, update_job_status, get_tracker_summary, TRACKER_PATH

def format_report(jobs, zip_path):
    """Formats the job search report as markdown."""
    import datetime
    today = datetime.date.today().strftime("%d %b %Y")
    
    report = []
    report.append("=========================================")
    report.append(f"      DAILY JOB SEARCH REPORT - {today}")
    report.append("=========================================")
    report.append(f"Summary: Found {len(jobs)} new matches.")
    report.append("")
    report.append("TOP MATCHES (resumes & cover letters ready for submission):")
    report.append("-" * 75)
    
    # Table Header
    report.append(f"{'#':<3} | {'Company':<15} | {'Job Title':<30} | {'Location':<10} | {'Fit Score':<9} | {'Job ID':<15}")
    report.append("-" * 95)
    
    for idx, job in enumerate(jobs, 1):
        report.append(f"{idx:<3} | {job['company']:<15} | {job['title'][:30]:<30} | {job['location']:<10} | {job['fitness_score']:>8}% | {job['job_id']:<15}")
        
    report.append("-" * 95)
    report.append("")
    report.append("MATCH EXPLANATIONS:")
    for idx, job in enumerate(jobs, 1):
        report.append(f"{idx}. {job['company']} - {job['title']} - {job['fitness_score']}% Fit:")
        report.append(f"   Why: {job['fitness_reason']}")
        report.append(f"   Apply Link: {job['link']}")
        report.append(f"   Resumes: [Yes] Generated | Cover Letters: [Yes] Generated")
        report.append("")
        
    report.append(f"MATERIALS BUNDLED IN ZIP: {zip_path}")
    report.append(f"TRACKER UPDATED IN EXCEL: {TRACKER_PATH}")
    report.append("=========================================")
    return "\n".join(report)

def main():
    parser = argparse.ArgumentParser(description="Job Search Assistant (India Edition) CLI")
    parser.add_argument("--search", action="store_true", help="Run job search and generate tailored materials")
    parser.add_argument("--status", action="store_true", help="Display summary stats from application tracker")
    parser.add_argument("--update", nargs=2, metavar=("JOB_ID", "STATUS"), help="Update status of a specific job")
    parser.add_argument("--custom", action="store_true", help="Manually generate documents for a specific job description")
    parser.add_argument("--company", type=str, help="Company name (for --custom)")
    parser.add_argument("--role", type=str, help="Job title/role (for --custom)")
    parser.add_argument("--location", type=str, help="Location (for --custom)")
    parser.add_argument("--jd", type=str, help="Job description/snippet (for --custom)")
    
    args = parser.parse_args()
    
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(1)
        
    if args.search:
        print("Starting job search engine...")
        jobs = run_job_search()
        
        if not jobs:
            print("No jobs found matching your criteria.")
            sys.exit(0)
            
        print("Generating custom application documents...")
        zip_path = generate_application_bundle(jobs, output_dir="dist")
        
        print("Updating Excel application tracker...")
        add_jobs_to_tracker(jobs)
        
        print("\n" + format_report(jobs, zip_path))
        
    elif args.status:
        summary = get_tracker_summary()
        print("=========================================")
        print("       APPLICATION TRACKER STATUS")
        print("=========================================")
        print(f"Total Tracked Applications: {summary['total']}")
        print("Breakdown by status:")
        for status, count in summary["statuses"].items():
            print(f"  - {status}: {count}")
        print("=========================================")
        
    elif args.update:
        job_id, status = args.update
        update_job_status(job_id, status)
        
    elif args.custom:
        if not (args.company and args.role and args.location and args.jd):
            print("Error: --custom requires --company, --role, --location, and --jd to be specified.")
            sys.exit(1)
            
        from search_engine import calculate_fitness
        from profile_manager import load_profile
        
        profile = load_profile()
        # Generate a unique manual ID
        job_id = f"MAN-{args.company.upper()[:3]}-{str(abs(hash(args.jd)))[:6]}"
        
        job = {
            "title": args.role,
            "company": args.company,
            "job_id": job_id,
            "platform": "Manual",
            "location": args.location,
            "posted_date": "Today",
            "link": "Manual Input",
            "snippet": args.jd
        }
        
        score, reason = calculate_fitness(job, profile)
        job["fitness_score"] = score
        job["fitness_reason"] = reason
        
        print(f"Calculating fitness score... Match is: {score}% Fit")
        print("Generating customized application documents...")
        zip_path = generate_application_bundle([job], output_dir="dist")
        
        print("Updating Excel application tracker...")
        add_jobs_to_tracker([job])
        
        print("\n" + format_report([job], zip_path))

if __name__ == "__main__":
    main()
