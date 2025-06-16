📄 AI-Powered Career Mentor: Personalized Job & Skill Recommendations
An intelligent, AI-driven web application that analyzes user resumes and profiles to suggest personalized job roles and required skill recommendations. The system leverages Natural Language Processing (NLP) and Machine Learning (ML) techniques to guide users in enhancing their career prospects.

📌 Table of Contents
About the Project

Features

Project Architecture

Tech Stack

Installation & Setup

Usage

Results

Future Enhancements

Contributing

License

📖 About the Project
The AI-Powered Career Mentor is designed to assist job seekers by:

Analyzing resumes using NLP

Matching user skills and experiences with job market demands

Recommending suitable job roles and essential skills to acquire

Providing ATS (Applicant Tracking System) resume scores to improve job application visibility

✨ Features
📄 Resume Parsing & Analysis

🧑‍💼 Personalized Job Role Recommendations

📚 Skill Gap Analysis & Recommendations

📊 ATS Resume Score Generation

📈 User-Friendly Dashboard Interface

🔍 Search & Filter for Desired Roles

📤 Upload Resume (PDF/Docx)

📑 Detailed Job & Skill Report Generation

🖥️ Project Architecture
pgsql
Copy
Edit
+---------------------+
|  User Uploads Resume |
+---------------------+
            |
            v
  +------------------+
  |  Resume Parser &  |
  |  Text Preprocessor|
  +------------------+
            |
            v
  +--------------------------+
  |   Skill Extraction & NLP  |
  +--------------------------+
            |
            v
  +--------------------------+
  | ML Model: Job Role Mapper |
  +--------------------------+
            |
            v
  +----------------------------+
  | ATS Score Calculator & UI  |
  +----------------------------+
🛠️ Tech Stack
Backend: Python, Flask

Frontend: HTML, CSS, Bootstrap

ML/NLP Libraries: Pandas, NumPy, scikit-learn, SpaCy, PyPDF2

Database: SQLite

Deployment: Localhost / Render / Heroku (as applicable)

🛠️ Installation & Setup
Clone the repository

bash
Copy
Edit
git clone https://github.com/Kinjal0706/AI-Powered_Career_Mentor_Personalized_Job_-_Skill_Recommendations.git
cd AI-Powered_Career_Mentor_Personalized_Job_-_Skill_Recommendations
Create a virtual environment and activate it

bash
Copy
Edit
python -m venv venv
source venv/bin/activate   # On Windows: venv\Scripts\activate
Install the required dependencies

bash
Copy
Edit
pip install -r requirements.txt
Run the application

bash
Copy
Edit
python app.py
Open your browser

arduino
Copy
Edit
http://localhost:5000
🚀 Usage
Upload your resume (PDF/Docx format)

View ATS score and resume analysis

Get recommended job roles and required skills

Improve your profile and apply confidently!

📊 Results
ATS Score Distribution
Resumes were scored and categorized:

ATS Score Range	Number of Resumes
90–100	2
80–89	4
70–79	3
Below 70	1

Skill Recommendations: Based on gaps identified between resume skills and job role requirements.

📈 Future Enhancements
Integrate real-time job listings via external APIs (e.g., LinkedIn, Indeed)

Implement user authentication and profile management

Enhance ATS scoring with industry benchmarks

Add multilingual resume support

Deploy the system to cloud platforms (Render/Heroku/AWS)

🤝 Contributing
Contributions are welcome! Follow these steps:

Fork the project

Create your feature branch git checkout -b feature/YourFeature

Commit your changes git commit -m 'Add SomeFeature'

Push to the branch git push origin feature/YourFeature

Open a pull request

📄 License
This project is open source and available under the MIT License.
