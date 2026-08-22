import json
import os

PROFILE_PATH = os.path.join(os.path.dirname(__file__), "profile.json")

def load_profile():
    """Loads the user profile from profile.json."""
    if not os.path.exists(PROFILE_PATH):
        raise FileNotFoundError(f"Profile file not found at {PROFILE_PATH}")
    with open(PROFILE_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def save_profile(profile_data):
    """Saves the user profile to profile.json."""
    with open(PROFILE_PATH, "w", encoding="utf-8") as f:
        json.dump(profile_data, f, indent=2)

def print_summary():
    """Prints a pretty summary of the profile."""
    try:
        profile = load_profile()
        print("========================================")
        print("      JOB SEARCH PROFILE SUMMARY")
        print("========================================")
        print(f"Name: {profile.get('name')}")
        print(f"Email: {profile.get('email')}")
        print(f"Phone: {profile.get('phone')}")
        print(f"LinkedIn: {profile.get('linkedin')}")
        print(f"GitHub: {profile.get('github')}")
        print(f"Location: {profile.get('location', {}).get('current')}")
        print("----------------------------------------")
        print("Preferences:")
        prefs = profile.get("preferences", {})
        print(f"  Roles: {', '.join(prefs.get('roles', []))}")
        print(f"  Locations: {', '.join(prefs.get('locations', []))}")
        print(f"  Company Types: {', '.join(prefs.get('company_types', []))}")
        print(f"  Seniority: {prefs.get('seniority')}")
        print(f"  CTC Expectations: {prefs.get('ctc_expectations')}")
        print("----------------------------------------")
        print("Skills:")
        skills = profile.get("skills", {})
        for cat, list_skills in skills.items():
            print(f"  {cat.capitalize()}: {', '.join(list_skills)}")
        print("========================================")
    except Exception as e:
        print(f"Error printing profile summary: {e}")

if __name__ == "__main__":
    print_summary()
