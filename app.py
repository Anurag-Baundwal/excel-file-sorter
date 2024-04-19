import pandas as pd
from rapidfuzz import fuzz
import time

def process_excel_file(file_path):
    start_time = time.time()

    # Read the Excel file
    df = pd.read_excel(file_path)

    # Sort by site name and create sheet 1
    df_site_name = df.sort_values('site_name')

    # Convert 'site_street' column to string before sorting
    df['site_street'] = df['site_street'].astype(str)

    # Sort by site address and create sheet 2
    df_site_address = df.sort_values('site_street')

    # Sort by latlong and create sheet 3
    df_latlong = df.sort_values(['site_latitude', 'site_longitude'])

    # Create Excel writer with context manager
    with pd.ExcelWriter('output.xlsx') as writer:
        df_site_name.to_excel(writer, sheet_name='Site Name', index=False)
        df_site_address.to_excel(writer, sheet_name='Site Address', index=False)
        df_latlong.to_excel(writer, sheet_name='Latlong', index=False)

    # Find exact duplicates
    exact_duplicates = df[df.duplicated(subset=['site_name'], keep=False)]

    # Function to find high matches for a given row
    def get_matches(row):
        matches = df[df['site_name'].apply(lambda x: fuzz.ratio(x, row['site_name']) >= 75)]
        matches = matches[matches.index != row.name]  # Exclude self-match
        return matches

    # Apply the function to each row to find high matches
    high_matches = df.apply(get_matches, axis=1)
    high_matches = high_matches[high_matches.apply(len) > 0]  # Filter rows with matches

    # Store exact duplicates in a text file
    with open('exact_duplicates.txt', 'w') as f:
        f.write("Exact Duplicates:\n")
        processed_pairs = set()  # Keep track of processed pairs to avoid duplicates

        for index, row in exact_duplicates.iterrows():
            # Get duplicate indices, excluding the current row
            duplicate_indices = df[(df['site_name'] == row['site_name']) & (df.index != index)].index.tolist()
            
            for dup_index in duplicate_indices:
                pair = tuple(sorted((index, dup_index)))  # Create a sorted tuple to represent the pair
                if pair not in processed_pairs:
                    processed_pairs.add(pair)  # Mark the pair as processed

                    f.write(f"Site 1 (Original Row {row.name + 1}): {row['site_name']}\n")  # Use original row number
                    f.write(f"Site 2 (Original Row {dup_index + 1}): {df.loc[dup_index, 'site_name']}\n")
                    f.write("Matching Attributes:\n")
                    f.write("- Site_Name\n")  # Always include site name for exact duplicates
                    for col in ['site_city', 'site_state', 'site_country']:
                        if len(df[df['site_name'] == row['site_name']][col].unique()) == 1:
                            f.write(f"- {col.title()}\n")
                    f.write("\n")

    # Store high matches in a text file (modified)
    with open('high_matches.txt', 'w') as f:
        f.write("High Matches:\n")
        processed_pairs = set()  # Keep track of processed pairs

        for index, matches in high_matches.items():
            if not matches.empty:
                for match_index, match_row in matches.iterrows():
                    pair = tuple(sorted((index, match_index)))
                    if pair not in processed_pairs:
                        processed_pairs.add(pair)

                        f.write(f"Site 1 (Original Row {index + 1}): {df.loc[index, 'site_name']}\n")
                        f.write(f"Site 2 (Original Row {match_index + 1}): {match_row['site_name']}\n")
                        f.write("Matching Attributes:\n")
                        for col in ['site_city', 'site_state', 'site_country']:
                            if df.loc[index, col] == match_row[col]:
                                f.write(f"- {col.title()}\n")
                        f.write("\n")

    end_time = time.time()
    execution_time = end_time - start_time
    print(f"Execution time: {execution_time:.2f} seconds")

# Example usage
file_path = 'input.xlsx'
process_excel_file(file_path)