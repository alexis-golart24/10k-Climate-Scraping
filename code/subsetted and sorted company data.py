import pandas as pd
from openpyxl import load_workbook
import os


import pandas as pd
from openpyxl import load_workbook
import os


def aggregate_company_data(excel_file, company_list, data_column, year_column="Year"):
    """
    Returns a DataFrame with yearly totals and counts of companies reporting data.
    Now also prints which companies reported data each year.
    """
    xls = pd.ExcelFile(excel_file)
    yearly_totals = {}
    yearly_counts = {}
    yearly_companies = {}

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
                    yearly_counts[year] += 1
                    yearly_companies[year].append(company)
                else:
                    yearly_totals[year] = value
                    yearly_counts[year] = 1
                    yearly_companies[year] = [company]

    # New: Print reporting companies by year
    print(f"\nCompanies reporting '{data_column}' by year:")
    for year in sorted(yearly_companies.keys()):
        print(f"{year}: {', '.join(yearly_companies[year])}")

    result = pd.DataFrame(
        {"total": yearly_totals, "count": yearly_counts}, index=yearly_totals.keys()
    )
    result.index.name = "Year"
    return result.sort_index(ascending=False)


def process_group(group_name, company_list, topics, excel_file, output_file, mode):
    print(f"\n{'='*50}")
    print(f"Processing group: {group_name}")
    print(f"Companies in division: {', '.join(company_list)}")
    all_results = {}

    for topic_info in topics:
        topic = topic_info["topic"]
        column = topic_info["column"]
        print(f"\nTopic: {topic}")

        result = aggregate_company_data(
            excel_file=excel_file, company_list=company_list, data_column=column
        )

        if not result.empty:
            result[column] = result["total"] / result["count"]
            result[column] = round(result[column], 2)
            all_results[topic] = result[[column]]

    if all_results:
        final_result = pd.concat(all_results.values(), axis=1)
        with pd.ExcelWriter(
            output_file, engine="openpyxl", mode=mode, if_sheet_exists="replace"
        ) as writer:
            final_result.to_excel(writer, sheet_name=group_name)
            print(f"\nSaved results to sheet: {group_name}")


topics = [
    {"topic": "climate change", "column": "climate change Count"},
    {"topic": "sustainable", "column": "sustainable Count"},
    {"topic": "sustainability", "column": "sustainability Count"},
    {"topic": "greenhouse gas", "column": "greenhouse gas Count"},
    {"topic": "environmental", "column": "environmental Count"},
    {"topic": "climate risk", "column": "climate risk Count"},
]

# File paths
excel_file = "C:/Users/alexi/Downloads/Processed EDGAR API Links_expanded.xlsx"
output_file = "C:/Users/alexi/OneDrive/Documents/10k-Climate-Scraping/Final_Expanded_Grouped_Companies.xlsx"

# Company divisions
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
        "Lowes",
        "Target",
        "Mcdonald's",
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

# Initialize output file mode
mode = "a" if os.path.exists(output_file) else "w"

# Process all divisions
for group_name, company_list in company_lists.items():
    process_group(
        group_name=group_name,
        company_list=company_list,
        topics=topics,
        excel_file=excel_file,
        output_file=output_file,
        mode=mode,
    )
    mode = "a"

print("\nAll processing complete!")
