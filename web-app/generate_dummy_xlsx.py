
import pandas as pd
from datetime import datetime, timedelta

# Create a dummy Fusion Timesheet Export
data = {
    'Project Name': ['Project A', 'Project B', 'Project A', 'Project C'],
    'Work Code': ['Dev', 'Test', 'Meeting', 'Design'],
    'Total Hours': [20, 15, 5, 10]
}

# Generate dates for the current month
today = datetime.now()
current_month_start = today.replace(day=1)
# Create columns for dates
dates = []
for i in range(20): # 20 days
    d = current_month_start + timedelta(days=i)
    if d.weekday() < 5: # Weekday
        dates.append(d.strftime('%Y-%m-%d'))
        if len(dates) >= 10: break

# Add date columns to data
for date in dates:
    data[date] = [2, 1.5, 0.5, 1] # Dummy hours

df = pd.DataFrame(data)

# Add header rows expected by the parser
# Row 0: Report Date From: ...
# Row 1: Empty or Metadata
# Row 2: Headers (Project Name, Work Code, Dates...)

# specialized writer needed to format closely to what script.js expects?
# script.js expects:
# Row 0 starts with "Report Date From:"
# Then headers row with "Project Name"

# Let's create a list of lists to represent the sheet
sheet_data = []

# Row 0
start_date_str = current_month_start.strftime('%d/%m/%Y')
end_date_str = (current_month_start + timedelta(days=30)).strftime('%d/%m/%Y')
sheet_data.append([f"Report Date From: {start_date_str} to {end_date_str}"])

# Row 1 (Empty)
sheet_data.append([])

# Row 2 (Headers)
headers = ['Project Name', 'Work Code'] + [d.strftime('%d/%m/%Y') for d in [datetime.strptime(x, '%Y-%m-%d') for x in dates]]
sheet_data.append(headers)

# Data Rows
projects = [
    ('Project Alpha', 'Dev'),
    ('Project Beta', 'Test'),
    ('Astron DOMS & POS Ingesting', 'Support') # Test cleaning logic
]

import random

for proj, code in projects:
    row = [proj, code]
    for _ in dates:
        row.append(random.choice([0.5, 1.0, 2.0, 4.0, 0]))
    sheet_data.append(row)

# Create DataFrame from this list of lists? No, openpyxl is better for exact layout but pandas is easier if installed.
# We'll use pandas to write without header/index, just raw data.
df_final = pd.DataFrame(sheet_data)

output_path = 'dummy_timesheet.xlsx'
df_final.to_excel(output_path, index=False, header=False)
print(f"Created {output_path}")
