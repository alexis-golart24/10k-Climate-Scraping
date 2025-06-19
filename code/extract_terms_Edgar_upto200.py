import pandas as pd
import openpyxl
import random
import requests
import time
import re
import os


def extract_sec_sections(ten_k_url, sections_to_extract, output_file):
    """
    Extract specified sections via SEC API; save to text file.
    
    Args:
        ten_k_url: URL to the SEC filing
        sections_to_extract: List of sections to extract
        output_file: File path where extracted text will be saved
    """
    # skip API call if output file exists
    if os.path.exists(output_file):
        print(f"{output_file} already extracted.")
        return

    # Initialize output file
    with open(output_file, "w", encoding="utf-8") as file:
        file.write("Extracted SEC Sections\n")
        file.write("=" * 30 + "\n\n")

    for section in sections_to_extract:
        api_url = "https://api.sec-api.io/extractor"
        params = {
            "url": ten_k_url,
            "item": section,
            "type": "text",
            "token": "7f6b79d4ae3ff8f336116ed5dcd8ce973bb670fcce1fbdd3086a94ae4c745830",
        }

        max_retries = 5
        retry_count = 0
        base_delay = 2  # Base delay in seconds

        while retry_count < max_retries:
            try:
                # Add delay between requests
                if retry_count > 0:
                    delay = base_delay * (2**retry_count) + random.uniform(0, 1)
                    print(f"Retrying in {delay:.2f} seconds...")
                    time.sleep(delay)
                else:
                    # Normal rate limiting delay
                    time.sleep(1)

                response = requests.get(api_url, params=params)

                if response.status_code == 200:
                    extracted_text = response.text

                    with open(output_file, "a", encoding="utf-8") as file:
                        file.write(f"Section {section}\n")
                        file.write("-" * 30 + "\n")
                        file.write(extracted_text + "\n\n")

                    print(f"Section {section} saved successfully.")
                    break

                elif response.status_code == 429:
                    retry_count += 1
                    print(f"Rate limit exceeded (429). Retry {retry_count}/{max_retries}")
                    # If server provides a Retry-After header, use that value
                    retry_after = response.headers.get("Retry-After")
                    if retry_after:
                        time.sleep(int(retry_after))
                    continue  # Retry after waiting
                else:
                    print(f"Error {response.status_code}: {response.text}")
                    break  # Don't retry for other error types

            except requests.exceptions.RequestException as e:
                print(f"Request failed: {e}")
                retry_count += 1
                if retry_count >= max_retries:
                    print(f"Max retries reached for section {section}. Moving to next section.")
                continue

        # If we've exhausted all retries
        if retry_count >= max_retries and not ("response" in locals() and response.status_code == 200):
            print(f"Failed to retrieve section {section} after {max_retries} attempts")



def process_text_files(company_name, start_year, num_entries, output_excel_path):
    """
    Process text files to extract climate-related references; save to excel file
    
    Args:
        company_name: Name of company/sheet
        start_year: most recent year
        num_entries: Number of entries/years to process
        output_excel_path: excel output
    """

    # search_words = ["climate change", "sustainable", "sustainability", "greenhouse gas", "environmental", "climate risk"]

    # Words that require exact matching (with word boundaries)
    exact_match_words = ["climate change", "greenhouse gas", "climate risk"]
    
    # Words that allow partial matching (no word boundaries on the end)
    partial_match_words = ["sustainab", "environmental"]
    
    # Combine both lists for processing
    search_config = {
        "exact": exact_match_words,
        "partial": partial_match_words
    }
    
    context_words = 50
    all_data = []

    for year in range(start_year, start_year - num_entries, -1):
        input_filename = f"extracted_{company_name}_sections_{year}.txt"

        if os.path.exists(input_filename):
            print(f"Processing {input_filename}...")
            extracted_data = extract_climate_references(
                input_filename, year, search_config, context_words, company_name
            )

            if extracted_data:
                all_data.append(extracted_data)
        else:
            print(f"File not found: {input_filename}, skipping...")

    if all_data:
        df = pd.DataFrame(all_data)

        file_exists = os.path.isfile(output_excel_path)

        if file_exists:
            # If file exists check if sheet needs to be replaced
            with pd.ExcelWriter(
                output_excel_path,
                engine="openpyxl",
                mode="a",
                if_sheet_exists="replace",
            ) as writer:
                df.to_excel(writer, index=False, sheet_name=company_name)
        else:
            # If no file create it
            with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name=company_name)

        print(f"Data successfully written to {output_excel_path}, sheet: {company_name}")
    else:
        print("No data extracted, Excel file not created.")


def extract_climate_references(input_txt_path, year, search_config, context_words, company_name):
    """
    Extract climate-related references from text files with surrounding context.
    
    Args:
        input_txt_path: Path to the text file
        year: Year of the filing
        search_config: Dictionary with 'exact' and 'partial' word lists
        context_words: Number of words to include before and after the match
        company_name: Name of the company
    """
    try:
        with open(input_txt_path, "r", encoding="utf-8") as file:
            extracted_text = file.read()

        cleaned_text = " ".join(extracted_text.splitlines())

        extracted_data = {
            "Company Name": company_name,
            "Year": year,
        }

        # Process exact match words
        for word in search_config["exact"]:
            pattern = rf"(?:\S+\s+){{0,{context_words}}}\b{re.escape(word)}\b(?:\s+\S+){{0,{context_words}}}"
            matches = re.findall(pattern, cleaned_text, flags=re.IGNORECASE)

            extracted_data[word] = " | ".join(matches) if matches else "No matches found"
            extracted_data[f"{word} Count"] = len(matches)

        # Process partial match words
        for word in search_config["partial"]:
            pattern = rf"(?:\S+\s+){{0,{context_words}}}\b{re.escape(word)}\w*(?:\s+\S+){{0,{context_words}}}"
            matches = re.findall(pattern, cleaned_text, flags=re.IGNORECASE)

            extracted_data[word] = " | ".join(matches) if matches else "No matches found"
            extracted_data[f"{word} Count"] = len(matches)

        return extracted_data

    except Exception as e:
        print(f"An error occurred while processing {input_txt_path}: {e}")
        return None


def run(file_path):
    """
    Main function to process excel file with SEC URLs
    
    Args:
        file_path: Excel file containing SEC URLs
    """
    sections_to_extract = (
        "1", "1A", "1B", "1C", "3", "4", "5", "6", "7", "7A",
        "8", "9", "9A", "9B", "10", "11", "12", "13", "14", "15"
    )
    output_excel_path = "Processed EDGAR API Links.xlsx"
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names

    for sheet_name in sheet_names:
        print(f"\n--- Processing {sheet_name} ---")
        
        # Read sheet data
        df = pd.read_excel(xls, sheet_name=sheet_name)
        most_recent_year = df.iloc[0, 0]
        num_entries = len(df)
        
        # (1) Extract sections for each year/filing
        for n in range(1, num_entries + 1):
            year = (most_recent_year + 1) - n
            
            # Get the 10-K URL 
            df_row = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            ten_k_url = df_row.iloc[n, 1]
            
            # Define output file for extracted sections
            output_file = f"extracted_{sheet_name}_sections_{year}.txt"
            
            # Extract sections 
            extract_sec_sections(ten_k_url, sections_to_extract, output_file)
        
        # (2) Process extracted text files
        process_text_files(sheet_name, most_recent_year, num_entries, output_excel_path)


if __name__ == "__main__":
    excel_file_path = "Edgar API Links.xlsx"
    run(excel_file_path)