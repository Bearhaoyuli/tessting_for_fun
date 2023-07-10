#importing libraries
# Please ensure the following packeages are installed via pip. 

import os, pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import seaborn as sns
from scipy import stats
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages
from scipy.stats import chi2_contingency
from pandas.api.types import CategoricalDtype
import xlsxwriter


cur_dir = os.getcwd()

file_name = 'People Analytics Data Science and Reporting - Case Study FINAL.xlsx'


os.makedirs(cur_dir+'/Q1_Outputs',exist_ok=True)
# os.makedirs(cur_dir+'/Q2_Outputs',exist_ok=True)
# os.makedirs(cur_dir+'/Q3_Outputs',exist_ok=True)



#reading the data into a dataframe
df_activity = pd.read_excel(cur_dir + '/' + file_name, sheet_name = "Recruiting Activity Data",skiprows=[1],header=1)

df_offer = pd.read_excel(cur_dir + '/' + file_name, sheet_name = "Offer Response Data")


#Get highest Degree ever obtained 
#Assumption: JD is considered Masters, so they have the same score with Master. 
degree_dict = {"PhD": 1, "Masters": 2, "JD" : 2, "Bachelors": 3}

def get_highest_degree_level(row):
    degree_levels = []
    for degree in ["Degree", "Degree.1", "Degree.2", "Degree.3"]:
        degree_value = row[degree]
        if not pd.isna(degree_value):
            degree_levels.append(degree_value)
    if not degree_levels:
        return None

    # Get degree with the lowest value in the dictionary(Note the highest is labled as 1 and lowest labled as 3)
    highest_degree = min(degree_levels, key=lambda x: degree_dict.get(x, 4), default=None)

    return highest_degree

df_activity["Highest Degree Level"] = df_activity.apply(get_highest_degree_level, axis=1)

# Get associated highest degree's School and Major.
def get_associated_school_major(row):
    degree_fields = ["Degree", "Degree.1", "Degree.2", "Degree.3"]
    school_fields = ["School", "School.1", "School.2", "School.3"]
    major_fields = ["Major", "Major.1", "Major.2", "Major.3"]

    highest_degree_level = row["Highest Degree Level"]

    for degree, school, major in zip(degree_fields, school_fields, major_fields):
        degree_value = row[degree]
        if degree_value == highest_degree_level or degree_value == "JD":
            return row[school], row[major]

    return None, None

df_activity[["Associated School", "Associated Major"]] = df_activity.apply(get_associated_school_major, axis=1, result_type="expand")



# Merge the two source & edit the Furthest Recruiting Stage Reached column so that the Offer Accepted is taken consideration
# (Other Offer decisions are not relevent to this particular analysis) 
df = pd.merge(df_activity, df_offer, on='Candidate ID Number', how='left')
mask = (df['Furthest Recruiting Stage Reached'] == 'Offer Sent') & (df['Offer Decision'] == 'Offer Accepted')
df.loc[mask, 'Furthest Recruiting Stage Reached'] = 'Offer Accepted'


df.to_csv(cur_dir+'/Q1_Outputs/1.Raw_data_output.csv', index=False)


df_tmp = df.copy(deep=True)

# Define the order of the stages
stage_order = ['new application', 'phone screen', 'in-house interview', 'offer sent', 'offer accepted']

# Normalize the 'Furthest Recruiting Stage Reached' column to lower case and remove any leading or trailing spaces
df_tmp['Furthest Recruiting Stage Reached'] = df_tmp['Furthest Recruiting Stage Reached'].str.lower().str.strip()

# Create a categorical type for stages with a specified order
df_tmp['Furthest Recruiting Stage Reached'] = pd.Categorical(df_tmp['Furthest Recruiting Stage Reached'], categories=stage_order, ordered=True)

# Replace String JD with Masters
df_tmp['Highest Degree Level'] = df_tmp['Highest Degree Level'].replace('JD', 'Masters')

# Create a pivot table for each department and highest degree level by furthest stage reached and count of candidate ID
pivot_df = df_tmp.pivot_table(index=['Department', 'Highest Degree Level', 'Furthest Recruiting Stage Reached'], 
                          values='Candidate ID Number', 
                          aggfunc='count').rename(columns={'Candidate ID Number': 'Applicants'})



# Update the count of "Offer Sent" applicants to include the count of "Offer Accepted" applicants
offer_accepted_index = pivot_df[pivot_df.index.get_level_values('Furthest Recruiting Stage Reached') == 'offer accepted'].index
offer_sent_index = pivot_df[pivot_df.index.get_level_values('Furthest Recruiting Stage Reached') == 'offer sent'].index
pivot_df.loc[offer_sent_index, 'Applicants'] += pivot_df.loc[offer_accepted_index, 'Applicants'].values

# Add Conversion Rate column
pivot_df['Conversion Rate'] = (pivot_df.groupby(['Department', 'Highest Degree Level'], group_keys=False)['Applicants']
                                    .apply(lambda x: x.div(x.shift()) * 100))
# Reset the index to make 'Furthest Recruiting Stage Reached' a regular column again
pivot_df = pivot_df.reset_index()

# Rename the stages for final output
pivot_df['Furthest Recruiting Stage Reached'] = pivot_df['Furthest Recruiting Stage Reached'].str.title()

# Sort the pivot_df by Department, Highest Degree Level, and Stage
pivot_df['Furthest Recruiting Stage Reached'] = pd.Categorical(pivot_df['Furthest Recruiting Stage Reached'], 
                                                               categories=[s.title() for s in stage_order], ordered=True)

pivot_df = pivot_df.sort_values(["Department", "Highest Degree Level", "Furthest Recruiting Stage Reached"])

pivot_df['Conversion Rate'] = pivot_df['Conversion Rate'].replace(np.inf, 100)

pivot_df.to_csv(cur_dir+'/Q1_Outputs/2.Funnel_output.csv', index=False)


# Function to format conversion rate
def format_conversion_rate(val):
    try:
        val = float(val)
        return f'{val:.2f}%'
    except ValueError:
        return val

# Apply function to Conversion Rate column
pivot_df['Conversion Rate'] = pivot_df['Conversion Rate'].apply(format_conversion_rate)

# Rename 'Furthest Recruiting Stage Reached' with 'Candidate Status'
pivot_df = pivot_df.rename(columns={'Furthest Recruiting Stage Reached': 'Candidate Status'})

# Define a custom ordering for 'Highest Degree Level' and 'Candidate Status'
degree_order = CategoricalDtype(['Bachelors', 'Masters', 'PhD'], ordered=True)
pivot_df['Highest Degree Level'] = pivot_df['Highest Degree Level'].astype(degree_order)

status_order = CategoricalDtype(['New Application', 'Phone Screen', 'In-House Interview', 'Offer Sent', 'Offer Accepted'], ordered=True)
pivot_df['Candidate Status'] = pivot_df['Candidate Status'].astype(status_order)

# Split the data by department and sort by degree level and candidate status
grouped = pivot_df.sort_values(['Department', 'Highest Degree Level', 'Candidate Status']).groupby('Department')

for name, group in grouped:
    group = group.copy()
    group['Department'] = group['Department'].mask(group['Department'].duplicated(), '')  # Only show department once per degree level
    fig, ax = plt.subplots(figsize=(12,8))
    ax.axis('tight')
    ax.axis('off')
    ax.set_title(f"{name}'s Recruiting Funnel", fontsize=20) # Add title

    # Create the table and add borders
    the_table = ax.table(cellText=group.values, colLabels=group.columns, cellLoc='center', loc='center')
    the_table.auto_set_font_size(False)
    the_table.set_fontsize(10)
    the_table.scale(1.2, 1.2)

    # Apply borders to cells
    for key, cell in the_table.get_celld().items():
        cell.set_edgecolor('black')

    # Save to PDF
    pp = PdfPages(cur_dir + f"/Q1_Outputs/{name} Department Funnel.pdf")
    pp.savefig(fig, bbox_inches='tight')
    pp.close()

print("Job finished, check the" + cur_dir + "/Q1_Outputs/*.pdf" + " for PDFs")
print("Check this file for Edited raw data(added Offer Accepted and Education)"+ cur_dir + '/Q1_Outputs/1.Raw_data_output.csv')
print("Check this file for Funnel output by department per highest Education" + cur_dir+'/Q1_Outputs/2.Funnel_output.csv')