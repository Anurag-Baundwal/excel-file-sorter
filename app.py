import pandas as pd
from fuzzywuzzy import fuzz

def process_excel_file(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Sort by site name and create sheet 1
    df_site_name = df.sort_values('site_name')
    df_site_name.to_excel('output.xlsx', sheet_name='Site Name', index=False)
    
    # Convert 'site_street' column to string before sorting
    df['site_street'] = df['site_street'].astype(str)
    
    # Sort by site address and create sheet 2
    df_site_address = df.sort_values('site_street')
    df_site_address.to_excel('output.xlsx', sheet_name='Site Address', index=False)
    
    # Sort by latlong and create sheet 3
    df_latlong = df.sort_values(['site_latitude', 'site_longitude'])
    df_latlong.to_excel('output.xlsx', sheet_name='Latlong', index=False)
    
    # Find exact duplicates and high matches
    exact_duplicates = []
    high_matches = []
    
    for i in range(len(df_site_name)):
        for j in range(i+1, len(df_site_name)):
            site1 = df_site_name.iloc[i]
            site2 = df_site_name.iloc[j]
            
            # Check for exact duplicates
            if site1['site_name'] == site2['site_name']:
                exact_duplicates.append((site1, site2))
            
            # Check for high matches (75% or more)
            elif fuzz.ratio(site1['site_name'], site2['site_name']) >= 75:
                high_matches.append((site1, site2))
    
    # Print exact duplicates and matching attributes
    print("Exact Duplicates:")
    for dup in exact_duplicates:
        site1, site2 = dup
        print(f"Site 1: {site1['site_name']}")
        print(f"Site 2: {site2['site_name']}")
        print("Matching Attributes:")
        if site1['site_city'] == site2['site_city']:
            print("- Site City")
        if site1['site_state'] == site2['site_state']:
            print("- Site State")
        if site1['site_country'] == site2['site_country']:
            print("- Site Country")
        print()
    
    # Print high matches and matching attributes
    print("High Matches:")
    for match in high_matches:
        site1, site2 = match
        print(f"Site 1: {site1['site_name']}")
        print(f"Site 2: {site2['site_name']}")
        print("Matching Attributes:")
        if site1['site_city'] == site2['site_city']:
            print("- Site City")
        if site1['site_state'] == site2['site_state']:
            print("- Site State")
        if site1['site_country'] == site2['site_country']:
            print("- Site Country")
        print()

# Example usage
file_path = 'input.xlsx'
process_excel_file(file_path)