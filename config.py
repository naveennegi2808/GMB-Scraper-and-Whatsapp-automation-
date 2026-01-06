# config.py

# 1. SEARCH KEYWORDS CATEGORIES (Target: Agentic AI Workshop Sales)
# REFRESHED LIST: Focusing on Active Student Communities & Societies
SEARCH_KEYWORDS = {
    # 1. GLOBAL STUDENT COMMUNITIES (Very High Intent)
    "STUDENT_COMMUNITIES": [
        "Google Developer Student Club",
        "GDSC College",
        "IEEE Student Branch",
        "ACM Student Chapter",
        "Computer Society of India Student Chapter",
        "CSI Student Branch",
        "CodeChef College Chapter",
        "GeeksforGeeks Student Chapter"
    ],

    # 2. INNOVATION & INCUBATION (Where budget exists)
    "INNOVATION_HUBS": [
        "Center of Excellence Artificial Intelligence",
        "IoT and Robotics Lab College",
        "Incubation and Innovation Cell",
        "Startup Cell Engineering College",
        "Research and Development Cell College",
        "Technology Business Incubator"
    ],

    # 3. SPECIFIC TECH NICHES (Fresh Angle)
    "NICHE_TECH_GROUPS": [
        "Cloud Computing Club",
        "Cyber Security Club",
        "Data Analytics Club",
        "Blockchain Society College",
        "Metaverse Lab College",
        "Game Development Club College"
    ],

    # 4. PRIVATE FINISHING SCHOOLS (Good for B2B upsell)
    "PRIVATE_TECH_SCHOOLS": [
        "Software Finishing School",
        "Industrial Training Institute Computer",
        "Advanced Computing Training Center",
        "Summer Training Engineering Students",
        "Winter Training Institute"
    ]
}

# 2. MATCHING KEYWORDS (RELEVANCE)
RELEVANCE_KEYWORDS = [
    # Core Tech
    "python", "java", "coding", "programming", "software", "developer",
    # AI Specific
    "artificial intelligence", "ai", "machine learning", "ml", "data science",
    "robotics", "automation", "tech", "innovation", "cloud", "cyber",
    # Communities
    "community", "society", "club", "chapter", "branch",
    # Educational Context
    "workshop", "seminar", "training", "course", "bootcamp",
    "student", "college", "institute", "placement", "skills"
]

# 3. DECISION MAKERS (Titles to find in "Deep Scraping")
DECISION_MAKERS = [
    "Faculty Advisor", "Faculty Coordinator", "Student Lead", "Chapter Lead",
    "President", "General Secretary", "Convenor",
    "HOD", "Head of Department",
    "Director", "Principal",
    "Incubation Manager"
]

# 5. LOCATIONS TO SEARCH
LOCATIONS = [
    # Tier 1
    "Bangalore", "Pune", "Mumbai", "Hyderabad", "Delhi", "Chennai", "Kolkata", 
    "Noida", "Gurgaon", "Ahmedabad",
    
    # Tier 2 (Education Hubs)
    "Jaipur", "Lucknow", "Indore", "Chandigarh", "Bhopal", "Coimbatore",
    "Visakhapatnam", "Nagpur", "Kochi", "Thiruvananthapuram", "Bhubaneswar",
    "Dehradun", "Patna", "Ludhiana", "Agra", "Nashik", "Surat", "Vadodara"
]

# 4. SHEET COLUMNS
SHEET_HEADERS = [
    "Name of Lead",             # col A
    "Contact number of lead",   # col B
    "Nature/Profession of Lead",# col C
    "Status",                   # col D
    "Email"                     # col E
]
