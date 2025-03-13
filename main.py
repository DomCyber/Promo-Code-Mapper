import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

# File path
file_path = "Book3.xlsx" #Excel file with two sheets - one for promo codes, one for regular codes 

# Read Excel files
df_regular_codes = pd.read_excel(file_path, sheet_name="RegularCodes", engine="openpyxl") # Sheet name
df_promo_codes = pd.read_excel(file_path, sheet_name="PromoCodes", engine="openpyxl") # Sheet name

# Convert descriptions to lists and ensure no NaN or non-string values are included
regular_codes = df_regular_codes['Description'].dropna().astype(str).tolist() #Column we search for 
promo_codes = df_promo_codes['Description'].dropna().astype(str).tolist() #Column we search for

# List to hold matching results
matches = []

# Similarity threshold
threshold = 80

# Iterate over each regular code
for reg_code in regular_codes:
    match = process.extractOne(reg_code, promo_codes, scorer=fuzz.partial_ratio)
    
    if match and match[1] >= threshold:
        matches.append((reg_code, match[0], match[1]))

# Create a DataFrame for the matches
match_df = pd.DataFrame(matches, columns=["Regular Code", "Promo Code", "Similarity"])

# Specify the output CSV file path
output_file = "matched_codes.csv"

# Export to CSV
match_df.to_csv(output_file, index=False)

print(f"Matched codes have been saved to {output_file}")
