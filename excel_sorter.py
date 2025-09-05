import pandas as pd
import re
import os
import sys
from urllib.parse import urlparse
from collections import defaultdict
import argparse

class ExcelSorter:
    def __init__(self):
        self.required_columns = ['reviews', 'website', 'rating']
    
    def extract_domain(self, url):
        """Extract domain name from URL"""
        if pd.isna(url) or url == '':
            return None
        
        try:
            # Clean the URL
            url = str(url).strip()
            
            # Add protocol if missing
            if not url.startswith(('http://', 'https://')):
                url = 'http://' + url
            
            # Parse URL
            parsed = urlparse(url)
            domain = parsed.netloc.lower()
            
            # Remove www. prefix
            if domain.startswith('www.'):
                domain = domain[4:]
            
            # Extract main domain name (before first dot)
            domain_parts = domain.split('.')
            if len(domain_parts) >= 2:
                return domain_parts[0]
            
            return domain
        except:
            return None
    
    def find_columns(self, df):
        """Find required columns in dataframe (case insensitive)"""
        column_mapping = {}
        df_columns_lower = [col.lower() for col in df.columns]
        
        for required_col in self.required_columns:
            found = False
            for i, col in enumerate(df_columns_lower):
                if required_col in col or col in required_col:
                    column_mapping[required_col] = df.columns[i]
                    found = True
                    break
            
            if not found:
                print(f"Warning: Column '{required_col}' not found in the data")
                return None
        
        return column_mapping
    
    def process_dataframe(self, df):
        """Process the dataframe according to requirements"""
        # Find required columns
        column_mapping = self.find_columns(df)
        if not column_mapping:
            return None
        
        # Create working copy
        df_work = df.copy()
        
        # Rename columns for easier processing
        website_col = column_mapping['website']
        reviews_col = column_mapping['reviews']
        rating_col = column_mapping['rating']
        
        # Convert reviews to numeric, handling errors
        df_work[reviews_col] = pd.to_numeric(df_work[reviews_col], errors='coerce').fillna(0)
        
        # Step 1: Separate rows with empty websites
        empty_website_mask = df_work[website_col].isna() | (df_work[website_col] == '') | (df_work[website_col] == 'nan')
        empty_website_rows = df_work[empty_website_mask].copy()
        non_empty_website_rows = df_work[~empty_website_mask].copy()
        
        # Step 2: Sort empty website rows by reviews (highest first)
        empty_website_rows = empty_website_rows.sort_values(by=reviews_col, ascending=False)
        
        # Step 3: Process non-empty website rows for domain extraction
        non_empty_website_rows['domain'] = non_empty_website_rows[website_col].apply(self.extract_domain)
        
        # Group by domain to find repeating businesses
        domain_groups = defaultdict(list)
        single_domain_rows = []
        
        for idx, row in non_empty_website_rows.iterrows():
            domain = row['domain']
            if domain:
                domain_groups[domain].append(idx)
            else:
                single_domain_rows.append(idx)
        
        # Separate repeated and single domain businesses
        repeated_businesses = []
        single_businesses = []
        
        for domain, indices in domain_groups.items():
            if len(indices) > 1:
                # Sort repeated businesses by reviews within each domain group
                domain_rows = non_empty_website_rows.loc[indices].sort_values(by=reviews_col, ascending=False)
                repeated_businesses.append(domain_rows)
            else:
                single_businesses.extend(indices)
        
        # Add single domain rows
        if single_domain_rows:
            single_businesses.extend(single_domain_rows)
        
        # Get single business rows and sort by reviews
        single_business_rows = non_empty_website_rows.loc[single_businesses].sort_values(by=reviews_col, ascending=False)
        
        # Step 4: Combine all data
        result_df = pd.concat([empty_website_rows, single_business_rows], ignore_index=True)
        
        # Add repeated businesses section
        if repeated_businesses:
            # Add separator row
            separator_row = pd.DataFrame([[''] * len(df.columns)], columns=df.columns)
            separator_row.iloc[0, 0] = 'Repeated Businesses'
            
            result_df = pd.concat([result_df, separator_row], ignore_index=True)
            
            # Add repeated business groups
            for domain_group in repeated_businesses:
                domain_group_clean = domain_group.drop('domain', axis=1)
                result_df = pd.concat([result_df, domain_group_clean], ignore_index=True)
        
        # Remove the temporary domain column if it exists
        if 'domain' in result_df.columns:
            result_df = result_df.drop('domain', axis=1)
        
        return result_df
    
    def load_file(self, file_path):
        """Load Excel or CSV file"""
        try:
            if file_path.lower().endswith('.csv'):
                return pd.read_csv(file_path)
            else:
                return pd.read_excel(file_path)
        except Exception as e:
            print(f"Error loading file {file_path}: {str(e)}")
            return None
    
    def process_single_file(self, input_file, output_dir=None):
        """Process a single file"""
        print(f"Processing: {input_file}")
        
        # Load file
        df = self.load_file(input_file)
        if df is None:
            return False
        
        # Process dataframe
        processed_df = self.process_dataframe(df)
        if processed_df is None:
            print(f"Could not process {input_file} - missing required columns")
            return False
        
        # Generate output filename
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        output_file = f"{base_name}_Cleaned.xlsx"
        
        if output_dir:
            output_file = os.path.join(output_dir, output_file)
        
        # Save processed file
        try:
            processed_df.to_excel(output_file, index=False)
            print(f"Saved: {output_file}")
            return True
        except Exception as e:
            print(f"Error saving {output_file}: {str(e)}")
            return False
    
    def process_multiple_files(self, input_files, output_file="Combined_Cleaned.xlsx"):
        """Process multiple files and combine into one"""
        print(f"Processing {len(input_files)} files...")
        
        all_data = []
        
        for file_path in input_files:
            print(f"Loading: {file_path}")
            df = self.load_file(file_path)
            if df is not None:
                processed_df = self.process_dataframe(df)
                if processed_df is not None:
                    all_data.append(processed_df)
                else:
                    print(f"Skipping {file_path} - missing required columns")
            else:
                print(f"Skipping {file_path} - could not load")
        
        if not all_data:
            print("No valid files to process")
            return False
        
        # Combine all dataframes
        combined_df = pd.concat(all_data, ignore_index=True)
        
        # Process the combined data again to handle cross-file duplicates
        final_df = self.process_dataframe(combined_df)
        
        # Save combined file
        try:
            final_df.to_excel(output_file, index=False)
            print(f"Combined file saved: {output_file}")
            return True
        except Exception as e:
            print(f"Error saving combined file: {str(e)}")
            return False

def main():
    parser = argparse.ArgumentParser(description='Excel/CSV Sorter Tool')
    parser.add_argument('files', nargs='+', help='Input files (Excel or CSV)')
    parser.add_argument('--combine', action='store_true', help='Combine multiple files into one')
    parser.add_argument('--output', help='Output file name (for combine mode) or directory')
    
    args = parser.parse_args()
    
    sorter = ExcelSorter()
    
    if args.combine or len(args.files) > 1:
        # Multiple files mode
        output_file = args.output if args.output else "Combined_Cleaned.xlsx"
        success = sorter.process_multiple_files(args.files, output_file)
    else:
        # Single file mode
        output_dir = args.output if args.output and os.path.isdir(args.output) else None
        success = sorter.process_single_file(args.files[0], output_dir)
    
    if success:
        print("Processing completed successfully!")
    else:
        print("Processing failed!")
        sys.exit(1)

if __name__ == "__main__":
    # If run without arguments, provide interactive mode
    if len(sys.argv) == 1:
        print("Excel/CSV Sorter Tool")
        print("=" * 50)
        
        # Get input files
        files_input = input("Enter file path(s) separated by commas: ").strip()
        files = [f.strip() for f in files_input.split(',')]
        
        # Check if files exist
        valid_files = []
        for file in files:
            if os.path.exists(file):
                valid_files.append(file)
            else:
                print(f"Warning: File not found - {file}")
        
        if not valid_files:
            print("No valid files found!")
            sys.exit(1)
        
        # Ask about combine mode
        if len(valid_files) > 1:
            combine = input("Combine all files into one? (y/n): ").lower().startswith('y')
        else:
            combine = False
        
        sorter = ExcelSorter()
        
        if combine:
            output_name = input("Enter output filename (default: Combined_Cleaned.xlsx): ").strip()
            if not output_name:
                output_name = "Combined_Cleaned.xlsx"
            success = sorter.process_multiple_files(valid_files, output_name)
        else:
            success = True
            for file in valid_files:
                if not sorter.process_single_file(file):
                    success = False
        
        if success:
            print("\nProcessing completed successfully!")
        else:
            print("\nSome files failed to process!")
    else:
        main()
