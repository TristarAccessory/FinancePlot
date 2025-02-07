import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as font_manager
import numpy as np
#%%
# Load the uploaded Excel file
file_path = "data/2024.xlsx"
excel_data = pd.ExcelFile(file_path)

# Function to extract totals for income and expense from each monthly sheet
def extract_totals_corrected(df: pd.DataFrame) -> tuple:
    # Identify rows containing '總收入' or '總支出' and extract the adjacent numeric values
    income_total = df[df.apply(lambda row: row.astype(str).str.contains('總收入', na=False).any(), axis=1)].apply(pd.to_numeric, errors='coerce').sum().max()
    expense_total = df[df.apply(lambda row: row.astype(str).str.contains('總支出', na=False).any(), axis=1)].apply(pd.to_numeric, errors='coerce').sum().max()
    return income_total, expense_total

# Extract income and expense for each month
monthly_income_expense = []

for sheet in [sheet for sheet in excel_data.sheet_names if sheet.startswith('2024.')]:
    monthly_df = excel_data.parse(sheet)
    income, expense = extract_totals_corrected(monthly_df)
    month = sheet.split('.')[1]
    monthly_income_expense.append((sheet, month, income, expense))

# Create a DataFrame for monthly totals
income_expense_df = pd.DataFrame(monthly_income_expense, columns=['Sheet', 'Month', 'Income', 'Expense'])
income_expense_df
#%%
# Add quarter information
income_expense_df['Quarter'] = income_expense_df['Month'].astype(int).apply(lambda x: f"2024-Q{(x - 1) // 3 + 1}")

# Summarize by quarter
quarterly_summary = income_expense_df.groupby('Quarter')[['Income', 'Expense']].sum().reset_index()

# Reshape data for plotting
quarterly_summary_melted = quarterly_summary.melt(id_vars=['Quarter'], value_vars=['Income', 'Expense'], var_name='Group', value_name='Amount')

quarterly_summary_melted
#%%
# Path to the custom font
font_path = 'font/TraditionalChinese.ttf'

# Add the custom font to the font manager
font_manager.fontManager.addfont(font_path)

# After adding the font, search for it by filename to get the correct font name
for font in font_manager.fontManager.ttflist:
    if font.fname == font_path:
        print(f"Found font: {font.name}")
        plt.rcParams['font.family'] = font.name
        break
#%%
# Adjusting the plot to match the given style
plt.figure(figsize=(12, 6))

# Colors and bar width for the side-by-side bars
bar_width = 0.35
colors = {'Income': '#f4a582', 'Expense': '#92c5de'}  # Soft pastel colors

# Compute positions for each bar group
quarters = np.arange(len(quarterly_summary['Quarter']))

# Plot income and expense bars side by side
plt.bar(quarters - bar_width / 2, quarterly_summary['Income'], width=bar_width, color=colors['Income'], label='收入', alpha=0.8)
plt.bar(quarters + bar_width / 2, quarterly_summary['Expense'], width=bar_width, color=colors['Expense'], label='支出', alpha=0.8)

# Add labels, title, and style adjustments
plt.title('公司收支狀況', fontsize=26, fontweight='bold', ha='center')
plt.xlabel('季度', fontsize=12)
plt.ylabel('金額', fontsize=12)
plt.xticks(quarters, quarterly_summary['Quarter'], rotation=45, fontsize=10)
plt.legend(title='', fontsize=10)
plt.grid(axis='y', linestyle='--', alpha=0.5)

plt.tight_layout()

# Save the plot to a file
plt.savefig('output/quarterly_income_expense.png', dpi=300)
plt.show()
plt.close()

# Plotting by month instead of by quarter
plt.figure(figsize=(12, 6))

# Compute positions for each bar group
months = np.arange(len(income_expense_df['Month']))

# Plot income and expense bars side by side
plt.bar(months - bar_width / 2, income_expense_df['Income'], width=bar_width, color=colors['Income'], label='收入', alpha=0.8)
plt.bar(months + bar_width / 2, income_expense_df['Expense'], width=bar_width, color=colors['Expense'], label='支出', alpha=0.8)

# Add labels, title, and style adjustments
plt.title('公司每月收支狀況', fontsize=18, fontweight='bold', ha='center')
plt.xlabel('月份', fontsize=12)
plt.ylabel('金額', fontsize=12)
plt.xticks(months, income_expense_df['Month'], rotation=45, fontsize=10)
plt.legend(title='', fontsize=10)
plt.grid(axis='y', linestyle='--', alpha=0.5)

plt.tight_layout()
# Save the plot to a file
plt.savefig('output/monthly_income_expense.png', dpi=300)
plt.show()
plt.close()
