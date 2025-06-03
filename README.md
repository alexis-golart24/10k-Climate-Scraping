# 10k-Climate-Scraping
Analyzing 100+ corporate 10-K filings to identify climate funding trends and global investment disparities.

1. Create dataframe by copying 10k links from Edgar into “Edgar API Links” excel

    a. Run “final automated text extraction code”

    b. Update the search words & parameters to extract specific language
    
    c. Code produces “Processed Edgar API Links” excel sheet with extracted data
   
       **check to make sure the companies are in the SIC key codes excel file with their corresponding industry number - if not, add them because it's needed for step 3
  
3. Classify companies into divisions by running “Broader company classifications”

4. Use the division lists from Step 2. in the “subsetted and sorted data” code to sort and average your extracted data by division (subsets by division and by years: 2002/2010 start year)

      a. This code produces the “Final Grouped Companies” excel file

       **update topics, company division lists, 2002/2010 list, and input/output file links to match new data
  
6. You then go to the R code “Visualizing Individual companies”

      a. Run the very last code chunk “PDFs for Company Divisions”
  
      b. This produces the PDF with grouped visualizations “Final grouped Company reports”



## Other notable documents
### Code: 
- Company_Sorting.py : sorts the extracted data without subsetting by year (2002/2010)
- Company_Classifications.py: groups by specific SIC divisions rather than broad categories
- Professor_Park_10k_scraping.py: Early merged code from Professor Park’s edits & “final automated text extraction” file

### Visualizations:
- Finalized_Company_Reports: Individual company data (not grouped by SIC divisions)
- Visualizing_individual_companies.rmd: Code to produce PDFs (not updated on github, only locally - issues with merging when I tried to push updates - to be resolved later)

### Data:
- Edgar_API_Links.xlsx: original 100 Edgar links
- Edgar_API_Links(100+).xlsx: Currently being updated, expanded database to use
- Processed_Edgar_API_Links.xlsx: Extracted data from Edgar API Links, sorted by division and year in steps 1. and 2. 
- SIC_Codes.xlsx: Sorts companies on using SIC groupings (broad and narrow sorting but we use the broad “SIC Sorted Divisions” Sheet of the excel
- Final_Grouped_Companies.xlsx: Final averaged/sorted data that’s used in visualizations code

