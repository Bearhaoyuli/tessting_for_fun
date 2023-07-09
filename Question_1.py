#importing libraries
import os, pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import seaborn as sns
from scipy import stats
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages
from scipy.stats import chi2_contingency
cur_dir = os.getcwd()
cur_dir
file_name = 'People Analytics Data Science and Reporting - Case Study FINAL.xlsx'

#reading the data into a dataframe
df_activity = pd.read_excel(cur_dir + '/' + file_name, sheet_name = "Recruiting Activity Data",skiprows=[1],header=1)

df_offer = pd.read_excel(cur_dir + '/' + file_name, sheet_name = "Offer Response Data")



#Get highest Degree ever obtained 
#Assumption: JD is considered Masters, so they will be labled as Masters as well. 
degree_dict = {"PhD": 1, "Masters": 2,"JD":2, "Bachelors": 3}
# Create a new column called "Highest Degree Level"
def get_highest_degree_level(row):
    degree_levels = []
    for degree in ["Degree", "Degree.1", "Degree.2", "Degree.3"]:
        degree_value = row[degree]
        if not pd.isna(degree_value):
            degree_levels.append(degree_value)
    if not degree_levels:
        return None
    min_degree_value = min([degree_dict.get(degree_value, 4) for degree_value in degree_levels])
    min_degree_name = [degree_name for degree_name, degree_value in degree_dict.items() if degree_value == min_degree_value]
    if min_degree_name:
        return min_degree_name[0]
    else:
        return None

df_activity["Highest Degree Level"] = df_activity.apply(get_highest_degree_level, axis=1)


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



# merge the two source
df = pd.merge(df_activity, df_offer, on='Candidate ID Number', how='left')
mask = (df['Furthest Recruiting Stage Reached'] == 'Offer Sent') & (df['Offer Decision'] == 'Offer Accepted')
df.loc[mask, 'Furthest Recruiting Stage Reached'] = 'Offer Accepted'



# Define the order of the stages
stage_order = ['new application', 'phone screen', 'in-house interview', 'offer sent', 'offer accepted']

# Create a new list to store the recruiting funnel view
funnel_data = []

# Normalize the 'Furthest Recruiting Stage Reached' column to lower case and remove any leading or trailing spaces
df_tmp = df.copy(deep=True)
df_tmp['Furthest Recruiting Stage Reached'] = df_tmp['Furthest Recruiting Stage Reached'].str.lower().str.strip()

# Create a categorical type for stages with a specified order
df_tmp['Furthest Recruiting Stage Reached'] = pd.Categorical(df_tmp['Furthest Recruiting Stage Reached'], categories=stage_order, ordered=True)

# Group the dataframe by Department and Highest Degree Level
grouped_df = df_tmp.groupby(["Department", "Highest Degree Level"])

# Iterate over each group and calculate the funnel view
for group, data in grouped_df:
    department, degree_level = group
    
    # Find the application that reached the furthest recruiting stage for each candidate
    furthest_stage = data.groupby("Candidate ID Number")["Furthest Recruiting Stage Reached"].max()
    
    # Count the unique candidates who reached each recruiting stage
    funnel_data_group = furthest_stage.value_counts().sort_index().reset_index()
    funnel_data_group.columns = ["Stage", "Applicants"]
    
    # Update the count of "Offer Sent" applicants to include the count of "Offer Accepted" applicants
    if "offer accepted" in funnel_data_group["Stage"].values:
        offer_accepted_count = funnel_data_group.loc[funnel_data_group["Stage"] == "offer accepted", "Applicants"].values[0]
        funnel_data_group.loc[funnel_data_group["Stage"] == "offer sent", "Applicants"] += offer_accepted_count
    
    # Calculate the conversion rate
    funnel_data_group["Conversion Rate"] = (funnel_data_group["Applicants"].div(funnel_data_group["Applicants"].shift()) * 100)
    
    # Set the conversion rate of 'new application' as 'NA'
    funnel_data_group.loc[funnel_data_group["Stage"] == "new application", "Conversion Rate"] = np.nan
    
    # Append the department and degree level to the funnel data
    funnel_data_group["Department"] = department
    funnel_data_group["Highest_degree_level"] = degree_level
    
    # Append the funnel data to the main funnel data list
    funnel_data.append(funnel_data_group)

# Concatenate all the funnel data into a single dataframe
funnel_df = pd.concat(funnel_data, ignore_index=True)

# Sort the funnel dataframe by Department, Highest Degree Level, and Stage
funnel_df.sort_values(["Department", "Highest_degree_level", "Stage"], inplace=True)

# Fill NA conversion rate with "NA" string
funnel_df["Conversion Rate"] = funnel_df["Conversion Rate"].fillna("NA")

# Changing stage names back to title case for final output
funnel_df['Stage'] = funnel_df['Stage'].str.title()
funnel_df = funnel_df[["Department", "Highest_degree_level", "Stage", "Applicants", "Conversion Rate"]]



def format_conversion_rate(val):
    try:
        val = float(val)
        return f'{val:.2f}%'
    except ValueError:
        return val

# Apply function to Conversion Rate column
funnel_df['Conversion Rate'] = funnel_df['Conversion Rate'].apply(format_conversion_rate)

# group the data by department
grouped = funnel_df.groupby('Department')


for name, group in grouped:
    fig, ax = plt.subplots(figsize=(12,8))
    ax.axis('tight')
    ax.axis('off')
    ax.table(cellText=group.values, colLabels=group.columns, cellLoc='center', loc='center')

    pp = PdfPages(f"{name}.pdf")
    pp.savefig(fig, bbox_inches='tight')
    pp.close()

print("Job finished, check the folder for PDFs")