# College Recommendation System 🎓

An intelligent college recommendation web application that helps students discover suitable colleges based on their academic qualifications, preferences, location, fees, and career priorities.

Built using Python, Streamlit, SQL, and the K-Nearest Neighbors (KNN) algorithm, the system generates personalized college recommendations and downloadable reports for students.

## Features
- 🎯 Personalized college recommendations based on user profile
- 🧠 KNN-based location and preference filtering
- 🏫 Course and specialization-based college search
- 📍 Location preference filtering
- 💰 Fee-based filtering
- 📊 Weighted ranking system using:
  - Location preference
  - Fees
  - Placement rates
  - NIRF rankings
- 📄 Downloadable recommendation reports in Word format
- 💻 Interactive user-friendly Streamlit interface

## Tech Stack
- Python
- Streamlit
- SQL (MySQL)
- Scikit-learn (KNN)
- Pandas
- NumPy
- python-docx

## How It Works
Users enter:
- Academic background (12th / Diploma / 12 + Diploma)
- Marks and eligibility details
- Preferred degree/course/specialization
- Location preferences
- Fee budget

The system:
1. Filters eligible colleges from the database
2. Uses KNN for similarity-based recommendation
3. Applies weighted ranking for better personalization
4. Displays ranked college suggestions
5. Generates downloadable recommendation reports

## Project Structure
```bash
code_epe.py                # Main Streamlit application
college_database.sql    # College database (if included)
requirements.txt        # Project dependencies
```

## Installation

### Clone the repository
```bash
git clone https://github.com/your-username/college-recommendation-system.git
cd college-recommendation-system
```

### Create virtual environment (optional)
```bash
python -m venv venv
source venv/bin/activate      # Mac/Linux
venv\Scripts\activate         # Windows
```

### Install dependencies
```bash
pip install -r requirements.txt
```

## Database Setup
Configure your MySQL credentials in the project:

```python
db_config = {
 "host": "localhost",
 "user": "root",
 "password": "your_password",
 "database": "COLLEGE_DATA"
}
```

Import your college dataset into MySQL.

## Run the Application
```bash
streamlit run air1.py
```

## Machine Learning Approach
This project uses:
- **K-Nearest Neighbors (KNN)** for similarity-based college recommendations
- **Min-Max Normalization** for feature scaling
- **Weighted Ranking System** to prioritize:
  - Location
  - Fees
  - Placement opportunities
  - NIRF Ranking

## Use Cases
- College recommendation for students
- Personalized admission guidance
- Course and college exploration
- Educational decision support system

## Future Enhancements
- Cutoff prediction system
- College admission probability estimator
- Scholarship recommendations
- Support for all Indian states
- AI chatbot for admission guidance

