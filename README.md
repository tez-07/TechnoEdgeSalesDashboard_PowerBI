# Monthly Sales Data Automation with Power BI

## 📌 Project Overview
This project automates the processing and visualization of monthly sales data using **Power BI** and **Outlook Microsoft Exchange integration**.  
The sales data is received via email in CSV format and is automatically cleaned, combined, and visualized—eliminating the need for manual Excel work.

---

## 🛠 Features
- **Outlook Rule Setup**: Automatically moves all emails with a specific subject line into a dedicated "Monthly Sales Data" folder.
- **Direct Power BI Connection**: Uses *Microsoft Exchange* to pull data directly from the Outlook folder.
- **Data Cleaning & Transformation**: Removes the need for manual formula-based checks in Excel.
- **Automated Dashboard**: Generates interactive, real-time sales insights without daily/weekly/monthly manual updates.
- **Time Savings**: Significant reduction in reporting time and manual effort.

---

## 🔄 Previous Workflow vs. New Workflow

**Old Process:**
1. Receive monthly sales CSV via email.
2. Manually download the file.
3. Open in Excel and apply formulas to check/clean data.
4. Prepare daily, weekly, and monthly reports.

**New Automated Process:**
1. Outlook Rule filters relevant emails into a specific folder.
2. Power BI connects directly to the folder via Microsoft Exchange.
3. Data is automatically cleaned and combined.
4. Dashboard updates with the latest data instantly.

---

## 📂 Steps to Reproduce

1. **Create Outlook Rule**  
   - Go to `Rules > Create Rule`.
   - Set the rule to move all emails with the same subject to a new folder (e.g., "Monthly Sales Data").

2. **Connect Power BI to Microsoft Exchange**  
   - Open Power BI Desktop.
   - Select `Get Data > Exchange`.
   - Enter your email & password to connect.
   - Select the email folder you created.

3. **Data Transformation**  
   - Clean and combine CSV files within Power BI’s Power Query Editor.
   - Apply any necessary data formatting and transformations.

4. **Dashboard Creation**  
   - Build visualizations to display sales metrics.
   - Publish the dashboard to Power BI Service for real-time access.

---

## 📊 Benefits
- No need to manually download and process files.
- Real-time dashboard accessible to the organization.
- Eliminates repetitive daily, weekly, and monthly report creation.
- Improves accuracy and consistency.

---

## 🖼 Screenshot  
<img width="1331" height="562" alt="Dashboard" src="https://github.com/user-attachments/assets/881898a1-267a-4840-8cd1-18012613fd80" />


---

## 🧑‍💻 Tools & Technologies
- **Power BI Desktop**
- **Microsoft Exchange (Outlook)**
- **Power Query**
- **CSV Data Format**

---

## 📅 Project Status
✅ Completed – Fully automated monthly sales reporting in Power BI.

Credits: https://www.linkedin.com/in/pavanlalwani/
