import pandas as pd
import openpyxl
import requests
import time
import re
import os


def process_excel_file(file_path):
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names

    for sheet_name in sheet_names:
        current_sheet_name = sheet_name
        print(f"\n--- Processing {current_sheet_name} ---")

        df = pd.read_excel(xls, sheet_name=sheet_name)

        most_recent_year = df.iloc[0, 0]

        num_entries = len(df)
        # ---------------------------------------------------------

        for n in range(1, num_entries + 1):
            naming = (most_recent_year + 1) - (n)

            df = pd.read_excel(file_path, sheet_name=current_sheet_name, header=None)
            ten_k_url = df.iloc[n, 1]

            sections_to_extract = (
                "1",
                "1A",
                "1B",
                "1C",
                "3",
                "4",
                "5",
                "6",
                "7",
                "7A",
                "8",
                "9",
                "9A",
                "9B",
                "10",
                "11",
                "12",
                "13",
                "14",
                "15",
            )

            output_file = f"extracted_{current_sheet_name}_sections_{naming}.txt"

            with open(output_file, "w", encoding="utf-8") as file:
                file.write("Extracted SEC Sections\n")
                file.write("=" * 30 + "\n\n")

            for x in sections_to_extract:
                api_url = "https://api.sec-api.io/extractor"
                params = {
                    "url": ten_k_url,
                    "item": x,
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
                            time.sleep(2)

                        response = requests.get(api_url, params=params)

                        if response.status_code == 200:
                            extracted_text = response.text

                            with open(output_file, "a", encoding="utf-8") as file:
                                file.write(f"Section {x}\n")
                                file.write("-" * 30 + "\n")
                                file.write(extracted_text + "\n\n")

                            print(f"Section {x} saved successfully.")
                            break

                        elif response.status_code == 429:
                            retry_count += 1
                            print(
                                f"Rate limit exceeded (429). Retry {retry_count}/{max_retries}"
                            )
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
                            print(
                                f"Max retries reached for section {x}. Moving to next section."
                            )
                        continue

                # If we've exhausted all retries
                if retry_count >= max_retries and not (
                    "response" in locals() and response.status_code == 200
                ):
                    print(
                        f"Failed to retrieve section {x} after {max_retries} attempts"
                    )

        # --------------------------------------------------------------------------

        def extract_climate_references(
            input_txt_path, year, search_words, context_words
        ):
            try:

                with open(input_txt_path, "r", encoding="utf-8") as file:
                    extracted_text = file.read()

                cleaned_text = " ".join(extracted_text.splitlines())

                extracted_data = {
                    "Company Name": f"{current_sheet_name}",
                    "Year": year,
                }

                for word in search_words:
                    pattern = rf"(?:\S+\s+){{0,{context_words}}}\b{re.escape(word)}\b(?:\s+\S+){{0,{context_words}}}"
                    matches = re.findall(pattern, cleaned_text, flags=re.IGNORECASE)

                    extracted_data[word] = (
                        " | ".join(matches) if matches else "No matches found"
                    )

                    word_count = len(matches)
                    extracted_data[f"{word} Count"] = word_count

                return extracted_data

            except Exception as e:
                print(f"An error occurred while processing {input_txt_path}: {e}")
                return None

        input_folder = "C:/Users/alexi/OneDrive/Comp sci/ML/"
        output_excel_path = "C:/Users/alexi/OneDrive/Comp sci/ML/Climate Research/Processed EDGAR API Links.xlsx"

        start_year = most_recent_year

        search_words = [
            "artificial intelligence",
            "AI",
            "machine learning",
            "data protection",
            "data privacy",
            "data transfer",
            "cybersecurity",
            "data security",
            "data center",
            "DORA",
            "Digital Operational Resilience Act",
            "GDPR",
            "General Data Protection Regulation",
            "DMA",
            "Digital Markets Act",
            "DSA",
            "Digital Services Act",
            "AI Act",
            "EU AI Act",
            "DMCC",
        ]
        context_words = 50

        all_data = []

        for year in range(start_year, start_year - (num_entries), -1):
            input_filename = f"extracted_{current_sheet_name}_sections_{year}.txt"
            input_path = os.path.join(input_folder, input_filename)

            if os.path.exists(input_path):
                print(f"Processing {input_filename}...")
                extracted_data = extract_climate_references(
                    input_path, year, search_words, context_words
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
                book = pd.ExcelFile(output_excel_path)
                with pd.ExcelWriter(
                    output_excel_path,
                    engine="openpyxl",
                    mode="a",
                    if_sheet_exists="replace",
                ) as writer:
                    df.to_excel(writer, index=False, sheet_name=current_sheet_name)
            else:
                # If no file create it
                with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name=current_sheet_name)

            print(
                f"Data successfully written to {output_excel_path}, sheet: {current_sheet_name}"
            )

        else:
            print("No data extracted, Excel file not created.")


# ------------------------------------


if __name__ == "__main__":
    excel_file_path = "C:/Users/alexi/Downloads/Edgar API Links.xlsx"
    process_excel_file(excel_file_path)
