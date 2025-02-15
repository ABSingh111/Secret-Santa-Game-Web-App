import pandas as pd
import re
from collections import Counter

#  Excel File Name (Replace it with your actual Excel file name)
file_path = "coding_challenge_test.xlsx"

# Load the Excel File
try:
    xls = pd.ExcelFile(file_path)
    print("✅ Excel File Loaded Successfully!")
except FileNotFoundError:
    print(f"❌ Error: '{file_path}' not found! Please check the file name.")
    exit()
except Exception as e:
    print(f"❌ Error: {e}")
    exit()

#  Check Available Sheets
print("📌 Available Sheets:", xls.sheet_names)

#  Select the Correct Sheet Name
sheet_name = xls.sheet_names[0]  # Load the first sheet
print(f"✅ Using Sheet: {sheet_name}")

# Load Data from Excel
try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
except Exception as e:
    print(f"❌ Error: Unable to load sheet '{sheet_name}'. Error: {e}")
    exit()

#  Clean Column Names
df.columns = df.columns.str.strip()  # Remove extra spaces
df.columns = df.columns.str.replace("\xa0", " ", regex=True)  # Remove non-breaking spaces
df.columns = df.columns.str.encode('ascii', 'ignore').str.decode('utf-8')  # Remove hidden characters
df.columns = df.columns.str.lower().str.replace(" ", "")  # Remove spaces

#  Display Cleaned Column Names
print("📌 Cleaned Columns:", df.columns.tolist())

#  Find the "Additional comments" Column (Completely Cleaned Check)
comments_col = [col for col in df.columns if "additionalcomments" in col]
if not comments_col:
    print("❌ Error: 'Additional comments' column not found in the sheet!")
    print("🔍 Debugging: Please check column names manually.")
    exit()

comments_col = comments_col[0]  # First matching column
print(f"✅ Using Column: {comments_col}")

#  Regex Pattern to Extract Groups
group_pattern = re.compile(r"Groups\s*:\s*\[code\]<I>(.*?)<\/I>\[/code\]", re.IGNORECASE)

#  Extract Groups from Comments
group_list = []

for comment in df[comments_col].dropna():
    matches = group_pattern.findall(str(comment))
    for match in matches:
        groups = [g.strip() for g in match.split(",")]
        group_list.extend(groups)

#  Count Unique Groups
group_counts = Counter(group_list)

#  Convert to DataFrame
output_df = pd.DataFrame(group_counts.items(), columns=["Group name", "Number of occurrences"])
output_df = output_df.sort_values(by="Number of occurrences", ascending=False)

# Save to Excel
output_file = "group_counts.xlsx"
output_df.to_excel(output_file, index=False, engine="openpyxl")

print(f"🎉✅ Processing Done! Output saved as '{output_file}'")
