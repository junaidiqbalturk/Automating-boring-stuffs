import os
import pandas as pd
import argparse

#function to merge the excel files
def merge_excel_files(input_dir,output_dir,file_bases):
    base_data = {base: pd.DataFrame() for base in file_bases} #Creating a dictionary to store the dataframe for each file
    #Loop
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)
                for base in file_bases:
                    if base in file:
                        try:
                            print(f"Reading file: {file_path}")
                            temp_df = pd.read_excel(file_path, engine='openpyxl')
                            print(f"Merging data from file: {file_path} into base: {base}")
                            base_data[base] = pd.concat([base_data[base], temp_df],ignore_index=True)
                        except Exception as e:
                            print("Error Reading {file_path}: {e})")


        #saving the merge file now
        for base, df in base_data.items():
            output_file = os.path.join(output_dir, base + '.xlsx')
            df.to_excel(output_file, index=False)
            print(f'Consolidated file for {base} saved as {output_file}')


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Merge excel files')
    parser.add_argument('input_directory',type=str, help='Directory containing Excel files')
    parser.add_argument('output_directory',type=str, help='Directory to save the output files')
    parser.add_argument('file_bases', nargs='+', help='List of file bases to merge')

    args = parser.parse_args()
    merge_excel_files(args.input_directory,args.output_directory,args.file_bases)


