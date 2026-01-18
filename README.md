# How Much Is Each Hour Really Worth?

A simple web tool that helps you understand whether working more hours is actually worth it —  
or if additional hours stop paying off due to taxes and deductions.

The system analyzes real salary data and recommends whether you should:
- ➕ Work more hours
- ➗ Keep things as they are
- ➖ Reduce working hours

All decisions are based on actual numbers, not gut feeling.

---

## 🎯 Who Is This For?
- Employees who work overtime
- Anyone wondering “Is it worth staying another hour?”
- People who want to understand how taxes affect their hourly pay

No tax knowledge or advanced Excel skills are required.

---

## ⚙️ How to Use

### Option 1 – Upload an Excel File
1. Prepare an Excel file according to the structure below
2. Upload it to the system
3. View the analysis and recommendations per month

### Option 2 – Manual Entry
1. Select a month
2. Enter hours and salary data
3. Click **“Add Month”**

---

## 📊 Excel File Structure

The Excel file must include a **header row** and the following columns **in this exact order**:

| Column | Name                  | Required | Notes                    |
|------:|-----------------------|----------|--------------------------|
| A     | Month                 | ✅ Yes   | e.g. Aug-25              |
| B     | Regular Hours         | ✅ Yes   | Number                   |
| C     | Overtime 125%         | ✅ Yes   | Number                   |
| D     | Overtime 150%         | ✅ Yes   | Number                   |
| E     | Gross Salary          | ✅ Yes   | In local currency        |
| F     | Total Deductions      | ✅ Yes   | Taxes, insurance, etc.   |
| G–T   | Anything else         | ❌ No    | Ignored by the system    |

Only columns **A–F** are used for calculations.

---

## 🧮 How the Calculation Works (In Simple Terms)

- Working hours are **weighted**:
  - Regular hours = ×1.0  
  - Overtime 125% = ×1.25  
  - Overtime 150% = ×1.5
- Net salary = Gross salary − Deductions
- The system calculates:
  - **Average net pay per hour**
  - **Marginal net pay** (what the extra hours actually paid)
- A **personal target hourly net pay** is calculated automatically from “healthy” months
- Each new month is compared against this target and the previous month

---

## 📌 Recommendation Logic

- **➕ Add hours**  
  Extra hours pay better than your typical hourly value

- **➗ Keep as is**  
  No significant change in hourly profitability

- **➖ Reduce hours**  
  Extra hours pay significantly less (usually due to higher taxation)

---

## ⚠️ Notes & Limitations
- The analysis compares each month only to the previous one
- Months with retroactive salary corrections may appear unusual
- No data is stored or sent to a server — everything runs locally in your browser

---

## 🚀 Live Version
The app is deployed and accessible online via Netlify.

---

## 🛠 Tech Stack
- React + TypeScript
- Vite
- XLSX
- Netlify (deployment)

---

## 📄 License
This project is intended for personal and educational use.
