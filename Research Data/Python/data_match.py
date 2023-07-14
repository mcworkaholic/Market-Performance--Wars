import sys
import pandas as pd

def match_and_write_data(source_file, destination_file, sheet_name):
    # Read the source and destination Excel files
    source_df = pd.read_excel(source_file)
    destination_df = pd.read_excel(destination_file, sheet_name=sheet_name, engine='openpyxl')

    # Convert the 'Date' column in the destination file to datetime64[ns] type
    destination_df['Date'] = pd.to_datetime(destination_df['Date'])

    # Check if the 'DJIA' column exists in both source and destination files
    if 'DJIA' in source_df.columns and 'DJIA' in destination_df.columns:
        # Match the dates and DJIA values based on the 'Date' column
        matched_data = source_df[['Date', 'DJIA']].dropna()

        # Merge the matched data with the destination file
        destination_df = pd.merge(destination_df, matched_data, on='Date', how='left')
        destination_df = destination_df.rename(columns={'DJIA': 'DJIA_source'})
    elif 'FTSE100' in source_df.columns and 'FTSE100' in destination_df.columns:
        # Convert the 'Date' column in the source file to datetime64[ns] type
        source_df['Date'] = pd.to_datetime(source_df['Date'])

        # Match the dates and FTSE100 values based on the 'Date' column
        matched_data = source_df[['Date', 'FTSE100']].dropna()

        # Merge the matched data with the destination file
        destination_df = pd.merge(destination_df, matched_data, on='Date', how='left')
        destination_df = destination_df.rename(columns={'FTSE100': 'FTSE_source'})

    # Save the updated destination file
    with pd.ExcelWriter(destination_file, engine='openpyxl', mode='a') as writer:
        destination_df.to_excel(writer, sheet_name=sheet_name, index=False)

if __name__ == '__main__':
    # Check if the required command-line arguments are provided
    if len(sys.argv) < 4:
        print("Usage: python data_match.py '<source_file_path>' '<destination_file_path>' '<sheet_name>'")
        sys.exit(1)

    # Extract the source and destination file paths and the sheet name from command-line arguments
    source_file_path = sys.argv[1]
    destination_file_path = sys.argv[2]
    sheet_name = sys.argv[3]

    # Call the match_and_write_data function
    match_and_write_data(source_file_path, destination_file_path, sheet_name)
