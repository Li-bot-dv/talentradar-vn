# 🎯 Vietnam Talent Recruiter AI Sourcing Agent (Mode A)

A specialized sourcing dashboard for talent mapping in **Vietnam**. 
Runs automated Google X-Ray searches company-by-company with **zero risk** to your LinkedIn account (no LinkedIn login required).

---

## ⚡ How to Run the Interactive UI

Double-click:
```bat
run_app.bat
```
Or run in terminal:
```bash
streamlit run app.py
```
This opens the Recruiter Agent interactive dashboard in your browser (`http://localhost:8501`).

---

## 📋 Sourcing Workflow

1. **Target Companies**: Enter your list of target companies (e.g. `MoMo`, `VNPay`, `ZaloPay`, `Shopee`, `Grab`).
2. **Relevant Titles**: Enter target job titles (e.g. `Senior Software Engineer`, `Tech Lead`, `Engineering Manager`).
3. **Background & Skills**: Enter relevant keywords (e.g. `Java`, `Golang`, `Payment`, `Microservices`).
4. **Geo-Location**: Automatically locked to **Vietnam** (`Ho Chi Minh City`, `Hanoi`, `Da Nang`).
5. **Start Automated Sourcing**:
   - The agent maps each company one-by-one via Google X-Ray.
   - Extracts candidate names, titles, companies, and direct LinkedIn profile links.
6. **Review Coverage Recap**:
   - Displays which companies have candidates vs. which companies had 0 results.
7. **Download Excel Recap**:
   - Generates a styled 2-sheet `.xlsx` file:
     - **Sheet 1 (`Sourced Candidates`)**: Name, Title, Industry/Background, Company, LinkedIn Link.
     - **Sheet 2 (`Search Coverage Recap`)**: Company, Status (Found / Not Found), Profile Count, Search Query.

---

## 🛡️ Why This is 100% Account Safe
- Mode A queries search engine index caches rather than logging into your personal LinkedIn account.
- Zero search limits, zero CAPTCHAs, and zero risk of LinkedIn account bans or restrictions.
