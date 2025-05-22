import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import os


def aggregate_company_data(excel_file, company_list, data_column, year_column="Year"):
    xls = pd.ExcelFile(excel_file)
    yearly_totals = {}

    for company in company_list:
        if company in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=company)

            if data_column not in df.columns:
                print(f"Warning: Column '{data_column}' not found in sheet '{company}'")
                continue

            for _, row in df.iterrows():
                year = row[year_column]
                value = row[data_column]

                if pd.isna(year) or pd.isna(value):
                    continue

                if year in yearly_totals:
                    yearly_totals[year] += value
                else:
                    yearly_totals[year] = value

    result = pd.DataFrame.from_dict(
        yearly_totals, orient="index", columns=[data_column]
    )
    result.index.name = "Year"
    return result.sort_index(ascending=False)


# Configuration data (unchanged)
topics = [
    {"topic": "Climate", "column": "Climate Count"},
    {"topic": "Sustainable", "column": "Sustainable Count"},
    {"topic": "Environmental", "column": "Environmental Count"},
]

# File paths (unchanged)
excel_file = "C:/Users/alexi/OneDrive/Documents/10k-Climate-Scraping/Processed EDGAR API Links.xlsx"
output_file = "C:/Users/alexi/OneDrive/Documents/10k-Climate-Scraping/Final_Grouped_Companies.xlsx"

company_lists = {
    "Division F": ["McKesson", "Cencora", "Cardinal Health"],
    "Division B": ["Occidental Petroleum", "SLB", "EOG Resources"],
    "Division E": [
        "Verizon Communications",
        "Comcast",
        "United Parcel Service",
        "NextEra Energy",
        "FedEx",
        "Charter Communications",
        "Duke Energy",
        "Union Pacific",
        "Delta Air Lines",
    ],
    "Division G": [
        "Walmart",
        "Amazon",
        "CVS Health",
        "Costco",
        "Home Depot",
        "Lowe's",
        "Target",
        "McDonald's",
        "TJX Cos",
    ],
    "Division I": [
        "Alphabet",
        "Microsoft",
        "Meta",
        "Oracle",
        "Visa",
        "Walt Disney",
        "Salesforce",
        "HCA Healthcare",
        "Netflix",
        "PayPal",
        "Mastercard",
        "Fiserv",
        "Automatic Data Processing",
    ],
    "Division H": [
        "Berkshire Hathaway",
        "JP Morgan",
        "Morgan Stanley",
        "Citigroup",
        "American Express",
        "Elevance Health",
        "Cigna Group",
        "US Bancorp",
        "Capital One",
        "Apollo Global Management",
        "American International Group",
        "Charles Schwab",
        "Progressive",
        "PNC Financial Services",
        "MetLife",
        "KKR & Co. Inc",
        "Bank of New York Mellon Corp",
        "Prudential Financial",
        "Travelers",
        "Centene",
        "Aflac",
        "Marsh McLennan",
    ],
    "Division D": [
        "Apple",
        "Tesla",
        "Exxon Mobil",
        "Chevron",
        "Toyota Motor",
        "Johnson & Johnson",
        "Procter & Gamble",
        "General Motors",
        "Broadcom",
        "PepsiCo",
        "AbbVie",
        "Cisco Systems",
        "Coca-Cola",
        "IBM",
        "Caterpillar",
        "Intel",
        "Deere & Company",
        "ConocoPhillips",
        "Nvidia",
        "RTX",
        "Ford Motor",
        "Marathon Petroleum",
        "Phillips 66",
        "Merck & Co.",
        "Abbott Laboratories",
        "Dell Technologies",
        "Lockheed Martin",
        "Eli Lilly and Company",
        "Philip Morris International",
        "Valero Energy",
        "Honeywell International",
        "Qualcomm",
        "Amgen",
        "Danaher",
        "Mondelez International",
        "General Dynamics",
        "Nike",
        "Kraft Heinz",
        "Applied Materials",
        "Archer-Daniels-Midland",
        "Adient",
    ],
}

# Determine if file exists and set mode
if os.path.exists(output_file):
    mode = "a"
else:
    mode = "w"

# Process each company group
for group_name, company_list in company_lists.items():
    print(f"\nProcessing group: {group_name}")

    all_results = {}
    for topic_info in topics:
        topic = topic_info["topic"]
        column = topic_info["column"]

        print(f"Processing {topic}...")
        result = aggregate_company_data(
            excel_file=excel_file, company_list=company_list, data_column=column
        )
        all_results[topic] = result

    # Combine all topics for this group
    final_result = pd.concat(all_results.values(), axis=1)

    # Save with ExcelWriter
    with pd.ExcelWriter(
        output_file, engine="openpyxl", mode=mode, if_sheet_exists="replace"
    ) as writer:
        final_result.to_excel(writer, sheet_name=group_name)
        print(f"Saved results to sheet: {group_name}")

    # After first write, switch to append mode for subsequent groups
    mode = "a"

print("\nAll processing complete!")
