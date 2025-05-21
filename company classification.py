import pandas as pd

file_path = "C:/Users/alexi/OneDrive/Comp sci/ML/Climate Research/SIC_codes.xlsx"

df = pd.read_excel(file_path, sheet_name="SIC Keys")
df.columns = df.columns.str.strip()

group_count = (
    df.groupby("SIC Sorted Divisions")
    .agg(
        count=("Company Name", "size"),
        companies=("Company Name", lambda x: ", ".join(x.unique())),
    )
    .reset_index()
    .sort_values(by="count", ascending=False)
)


with pd.ExcelWriter(
    file_path, mode="a", engine="openpyxl", if_sheet_exists="replace"
) as writer:
    group_count.to_excel(writer, sheet_name="SIC Sorted Divisions", index=False)
