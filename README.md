### **README: Group Analysis from Excel Sheet**  

#### **📌 Project Overview**  
This project reads data from an Excel file, extracts group names from the **"Additional comments"** column, and counts their occurrences. The output is saved in an Excel file (`group_counts.xlsx`).  

---

### **🛠️ Tools & Libraries Used**  
- **Python**: Programming language  
- **pandas**: For reading and processing Excel data  
- **openpyxl**: To handle Excel files  
- **re (Regex)**: To extract group names from text  
- **collections.Counter**: To count occurrences  

---

### **📂 File Structure**  
```
/group_analysis_project  
│── group_analysis.py        # Main Python script  
│── coding_challenge_test.xlsx  # Input Excel file  
│── group_counts.xlsx        # Output file (Generated)  
│── README.md                # Project documentation  
```

---

### **🚀 Installation & Setup**  

#### **1️⃣ Install Python (if not installed)**  
Download & install [Python](https://www.python.org/downloads/).  

#### **2️⃣ Install Required Libraries**  
Run the following command:  
```sh
pip install pandas openpyxl
```

#### **3️⃣ Place the Input File**  
Ensure the `coding_challenge_test.xlsx` file is inside the project folder.

---

### **📜 Steps to Execute the Code**  

#### **1️⃣ Load Excel File**  
- The script reads the Excel file using `pandas.ExcelFile()`.  
- If the file is missing, an error message is displayed.  

#### **2️⃣ Detect the Correct Sheet**  
- The script detects available sheets and selects the relevant one.  

#### **3️⃣ Find the "Additional comments" Column**  
- The script normalizes column names to handle spaces and encoding issues.  
- If the column is not found, it provides debugging information.  

#### **4️⃣ Extract Groups from Text**  
- Uses regex to find group names inside `[code]<I>...</I>[/code]`.  
- If multiple groups exist, they are split by commas.  

#### **5️⃣ Count Unique Groups**  
- A dictionary (`Counter`) stores the number of occurrences for each group.  

#### **6️⃣ Save Output to Excel**  
- The results are written to `group_counts.xlsx`.  

---

### **📌 Running the Script**  
Run the script using:  
```sh
python group_analysis.py
```

After execution, `group_counts.xlsx` will be created with the unique group names and their counts.  

---

### **🔍 Debugging Tips**  
- Ensure the Excel file name is correct.  
- Check the sheet name inside the script.  
- If the column is not detected, print `df.columns.tolist()` to verify its name.  

---

### **📧 Contact & Support**  
If you face any issues, feel free to reach out. 🚀