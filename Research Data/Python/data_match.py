import sys
import pandas as pd

def match_data(source_df, destination_df, column_names):
    # Convert the 'Date' column in the destination file to datetime64[ns] type
    destination_df['Date'] = pd.to_datetime(destination_df['Date'])

    # Convert the 'Date' column in the source file to datetime64[ns] type
    source_df['Date'] = pd.to_datetime(source_df['Date'])

    # Iterate over each column name and match the dates and values
    for column_name in column_names:
        if column_name in source_df.columns and column_name in destination_df.columns:
            matched_data = source_df[['Date', column_name]].dropna()

            # Merge the matched data with the destination file
            destination_df = pd.merge(destination_df, matched_data, on='Date', how='left')
            destination_df = destination_df.rename(columns={column_name: column_name + '_source'})

    return destination_df

def match_and_write_data(source_file, destination_file, sheet_name, column_names):
    # Read the source and destination Excel files
    source_df = pd.read_excel(source_file)
    destination_df = pd.read_excel(destination_file, sheet_name=sheet_name, engine='openpyxl')

    # Call the match_data function with the column names list
    destination_df = match_data(source_df, destination_df, column_names)

    # Save the updated destination file
    with pd.ExcelWriter(destination_file, engine='openpyxl', mode='a') as writer:
        destination_df.to_excel(writer, sheet_name=sheet_name, index=False)

if __name__ == '__main__':
    # Check if the required command-line arguments are provided
    if len(sys.argv) < 5:
        # python data_match.py "source_file.xlsx" "destination_file.xlsx" "Sheet1" "DJIA,FTSE100,S&P500"
        print("Usage: python data_match.py '<source_file_path>' '<destination_file_path>' '<sheet_name>' '<column_name1,column_name2,...>'")
        sys.exit(1)

    # Extract the source file path, destination file path, sheet name, and column names from command-line arguments
    source_file_path = sys.argv[1]
    destination_file_path = sys.argv[2]
    sheet_name = sys.argv[3]
    column_names = sys.argv[4].split(',')

    # Call the match_and_write_data function
    match_and_write_data(source_file_path, destination_file_path, sheet_name, column_names)


   
