import openpyxl


def process_excel_column_to_E(
    file_path,
    variable1="Variable1",
    variable2="Variable2",
    variable3="Variable3",
    variable4="Variable4",
    variable5="Variable5",
    variable6="Variable6",
    variable7="Variable7",
    variable8="Variable8",
    variable9="Variable9",
):

    wb = openpyxl.load_workbook(file_path)
    sheet = wb.worksheets[0]

    if sheet["E1"].value is None:
        sheet["E1"] = "SIC Sorted Divisions "

    for row in range(2, sheet.max_row + 1):
        original_value = sheet[f"B{row}"].value

        try:
            num_value = float(original_value)
            if 1011 <= num_value <= 1499:
                sheet[f"E{row}"] = variable1
            elif 1521 <= num_value <= 1799:
                sheet[f"E{row}"] = variable2
            elif 2000 <= num_value <= 3999:
                sheet[f"E{row}"] = variable3
            elif 4011 <= num_value <= 4971:
                sheet[f"E{row}"] = variable4
            elif 5012 <= num_value <= 5199:
                sheet[f"E{row}"] = variable5
            elif 5211 <= num_value <= 5999:
                sheet[f"E{row}"] = variable6
            elif 6011 <= num_value <= 6799:
                sheet[f"E{row}"] = variable7
            elif 7011 <= num_value <= 8999:
                sheet[f"E{row}"] = variable8
            elif 9111 <= num_value <= 9999:
                sheet[f"E{row}"] = variable9
            else:

                sheet[f"E{row}"] = "Other"
        except (ValueError, TypeError):

            sheet[f"E{row}"] = "Invalid"

    wb.save(file_path)
    print(f"Successfully processed {sheet.max_row-1} rows in {file_path}")


process_excel_column_to_E(
    "C:/Users/alexi/OneDrive/Comp sci/ML/Climate Research/SIC_codes.xlsx",
    variable1="Division B: Mining",
    variable2="Division C: Construction",
    variable3="Division D: Manufacturing",
    variable4="Division E: Transportation, Communications, Electric, Gas, And Sanitary Services",
    variable5="Division F: Wholesale Trade",
    variable6="Division G: Retail Trade",
    variable7="Division H: Finance, Insurance, And Real Estate",
    variable8="Division I: Services",
    variable9="Division J: Public Administration",
)
