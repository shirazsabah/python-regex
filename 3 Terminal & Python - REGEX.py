## Terminal script

# Convert .docx to txt
textutil -convert txt FILE/PATH/docs/*.docx

# For .doc need to work around the 'argument list too long' error
# textutil -convert txt FILE/PATH/docs/*.doc
find FILE/PATH/docs/ \
    -maxdepth 1 \
    -type f \
    -name '*.doc' \
    -exec textutil -convert txt {} +


# Move .txt files to separate folder
setopt extended_glob
zmodload zsh/files
mv -- FILE/PATH/docs/*.txt FILE/PATH/txt/


# Search .txt files for selected regular expressions
grep -rilE \
    "(?i)(\bhip replacement\b.*\b(?:sciatic nerve|foot drop)\b|\b(?:sciatic nerve|foot drop)\b.*\bhip replacement\b|" \
    "\bTHR\b.*\b(?:sciatic nerve|foot drop)\b|\b(?:sciatic nerve|foot drop)\b.*\bTHR\b)" \
FILE/PATH/txt/ \
> FILE/PATH/matched_files.txt


# Check matched files
cd FILE/PATH/
open matched_files.txt

# Replace double forward slash with single
sed -i '' 's,//,/,g' matched_files.txt


# Concatenate all matching files into one document
while IFS= read -r filename; do
    echo -e "\n#############################################################################: \n $filename\n ############################################################################# \n" >> concatenated.txt
    cat "$filename" >> concatenated.txt
done < matched_files.txt


## Python

# Cross-ref between operations identified by RM and clinic letters

# Import RM data

import pandas as pd

excel_file = 'FILE/PATH/Data Request.xlsx'
sheet_name = 'Data' 
df_rm = pd.read_excel(excel_file, sheet_name=sheet_name)


# # Get list of column names
# df_rm.columns.tolist()
# df_rm['Patient Hospital Number']


file_path = 'FILE/PATH/matched_files.txt'

df_mf = pd.read_csv(file_path, header=None, names=['Patient Hospital Number'])

df_mf['Patient Hospital Number'] = df_mf['Patient Hospital Number'].str.strip('FILE/PATH/txt/')
df_mf['Patient Hospital Number'] = df_mf['Patient Hospital Number'].str.strip('.')



# Merge the two dataframes

# Check data types
df_rm['Patient Hospital Number'].dtype
df_mf['Patient Hospital Number'].dtype

# Convert df_mf to float64
df_mf['Patient Hospital Number'] = pd.to_numeric(df_mf['Patient Hospital Number'], errors='coerce')

# Merge with different options
df_merge = pd.merge(df_mf, df_rm, on='Patient Hospital Number', how='outer', indicator=True)    
df_merge.to_csv('FILE/PATH/df_merge.csv', index=False)

# # Explore merge
# tabulated = df_merge['_merge'].value_counts()
# print(tabulated)











