import pandas as pd

#File utilized to extract data from
file = r"c:\Users\logachoi\OneDrive - Advanced Micro Devices Inc\Desktop\Python Scripts - AEDT\Final Codebase\excel.csv"

df = pd.read_csv(file)

# Filter out rows where a certain parameter does not contain "Differential Net"
df = df[~df['Net'].str.contains('Differential Net')]
df.drop(columns=["Path", " Trace Separation (mm)", " IBIS", " Pin Source", "Layer", "Type", "Top Ref. Layer", "Bottom Ref. Layer"], inplace=True)
df.drop_duplicates(inplace=True)

df[""] = None
df['Small_Length(mm)'] = None
df['Small_Delay(ps)'] = None
df['Small Z0/Zdiff (ohms)'] = None

# Identify rows with the same net name but different lengths
duplicate_nets = df[df.duplicated(subset=['Net'], keep=False)]

# Iterate over the duplicate nets
for net, group in duplicate_nets.groupby('Net'):
    if len(group) == 2:  # Ensure there are only 2 rows for the same net
        # Get the rows with different lengths
        row1, row2 = group.itertuples(index=False)
        
        # Determine which row has the smaller length
        small_row = row1 if row1[1] < row2[1] else row2
        large_row = row1 if row1[1] >= row2[1] else row2
        
        df.loc[df.index[df['Net'] == net], 'Small_Length(mm)'] = small_row[1]  # Assuming the length is in the second column
        df.loc[df.index[df['Net'] == net], 'Small_Delay(ps)'] = small_row[2]  # Assuming the Delay (ps) is in the third column
        df.loc[df.index[df['Net'] == net], 'Small Z0/Zdiff (ohms)'] = small_row[3]  # Assuming the Z0/Zdiff (ohms) is in the fourth column

        # Drop the row with the smaller length
        indices_to_drop = df.index[(df['Net'] == net) & (df['Length (mm)'] == small_row[1])].tolist()
        df.drop(indices_to_drop, inplace=True)


# Create a new Excel file and add the filtered DataFrame as a new sheet
with pd.ExcelWriter('filtered_excel_file.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Filtered Sheet')

    # Enable filter options for all columns
    worksheet = writer.sheets['Filtered Sheet']
    worksheet.autofilter(0, 0, df.shape[0], df.shape[1]-1)
