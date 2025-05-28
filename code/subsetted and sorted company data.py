import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import os

list_1 = [
    "Walmart",
    "Amazon",
    "Exxon Mobil",
    "McKesson",
    "Cencora",
    "Costco",
    "Microsoft",
    "Chevron",
    "JP Morgan",
    "Verizon Communications",
    "Procter and Gamble",
    "Pepsi",
    "Elevance Health",
    "Cisco Systems",
    "IBM",
    "Caterpillar",
    "Intel",
    "John Deere",
    "Nvidia",
    "RTX",
    "US Bancorp",
    "Capital One",
    "Progressive",
    "NextEra Energy",
    "Merck & Co.",
    "Abbott Laboratories",
    "FedEx",
    "Lockheed Martin",
    "ELI LILY AND COMPANY",
    "Prudential Financial",
    "Lowes",
    "Charter Communications",
    "Qualcomm",
    "HCA Healthcare",
    "Amgen",
    "Danaher",
    "Mondelez International",
    "Target",
    "Netflix",
    "Travelers",
    "Mcdonald's",
    "Mastercard",
    "General Dynamics",
    "Nike",
    "Aflac",
    "Fiserv",
    "SLB",
    "Applied Materials",
    "Archer Daniels Midland ",
]

list_2 = [
    "Tesla",
    "CVS Health",
    "Cardinal Health",
    "Toyota Motor",
    "Citigroup",
    "Comcast",
    "Johnson & Johnson",
    "American Express",
    "General Motors",
    "Home Depot",
    "Oracle",
    "Coca Cola",
    "ConocoPhillips",
    "Ford Motors",
    "Visa",
    "American International Group",
    "Charles Schwab",
    "PNC Financial Services",
    "United Parcel Service",
    "Metlife",
    "Salesforce",
    "KKR & Co. Inc",
    "Philip Morris International Inc",
    "Valero Energy",
    "Bank of New York Mellon Corp",
    "Honeywell International",
    "Duke Energy",
    "Union Pacific",
    "Centene",
    "Delta Air Lines",
    "Occidental Petroleum",
    "TJX Cos",
    "Automatic Data Processing",
    "EOG Resources",
    "Marsh McLennan",
]

list_3 = [
    "Apple",
    "Berkshire Hathaway",
    "Alphabet",
    "Meta",
    "Morgan Stanley",
    "Broadcom",
    "ABBvie",
    "Cigna Group",
    "Apollo Global Management",
    "Marathon Petroleum",
    "Walt Disney",
    "Phillips 66",
    "Dell Technologies",
    "Paypal",
    "Kraft Heinz Company",
    "Adient",
]


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


def process_group(group_name, company_list, topics, excel_file, output_file, mode):
    print(f"\nProcessing group: {group_name}")
    num_companies = len(company_list)

    all_results = {}
    for topic_info in topics:
        topic = topic_info["topic"]
        column = topic_info["column"]

        print(f"Processing {topic}...")
        result = aggregate_company_data(
            excel_file=excel_file, company_list=company_list, data_column=column
        )

        # 💥 Drop early years based on which list the group is from
        if "2010-2024" in group_name:
            result = result[result.index > 2009]
        elif "2002-2024" in group_name:
            result = result[result.index > 2001]

        if not result.empty:
            result[column] = result[column] / num_companies
            result[column] = round(result[column], 2)
            all_results[topic] = result

    if all_results:
        final_result = pd.concat(all_results.values(), axis=1)
        with pd.ExcelWriter(
            output_file, engine="openpyxl", mode=mode, if_sheet_exists="replace"
        ) as writer:
            final_result.to_excel(writer, sheet_name=group_name)
            print(f"Saved results to sheet: {group_name}")


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

if os.path.exists(output_file):
    mode = "a"
else:
    mode = "w"

# Process each company group
for group_name, company_list in company_lists.items():
    # Split companies into List 1 and List 2
    group_list1 = [c for c in company_list if c in list_1]
    group_list2 = [c for c in company_list if c in list_2]

    # Process List 1 companies
    if group_list1:
        process_group(
            f"{group_name} - 2002-2024",
            group_list1,
            topics,
            excel_file,
            output_file,
            mode,
        )

    # Process List 2 companies
    if group_list2:
        process_group(
            f"{group_name} - 2010-2024",
            group_list2,
            topics,
            excel_file,
            output_file,
            mode,
        )

    mode = "a"  # Ensure append mode for subsequent writes

print("\nAll processing complete!")
