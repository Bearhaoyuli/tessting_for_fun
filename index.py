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
df_activity
df_offer = pd.read_excel(cur_dir + '/' + file_name, sheet_name = "Offer Response Data")
df_offer
##Get highest Degree ever obtained 
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

# Print the DataFrame
df_activity

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

# df = df_activity
df_activity[["Associated School", "Associated Major"]] = df_activity.apply(get_associated_school_major, axis=1, result_type="expand")


df_activity

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

# Split the data by department
grouped = funnel_df.groupby('Department')

for name, group in grouped:
    fig, ax = plt.subplots(figsize=(12,8))
    ax.axis('tight')
    ax.axis('off')
    ax.table(cellText=group.values, colLabels=group.columns, cellLoc='center', loc='center')

    pp = PdfPages(f"{name}.pdf")
    pp.savefig(fig, bbox_inches='tight')
    pp.close()


#Q2

# Step 1: Filter the dataframe
df_stats = df.copy(deep=True)


# Step 2: Preprocess the data
df_stats['Furthest Recruiting Stage Reached'] = df_stats['Furthest Recruiting Stage Reached'].str.lower()

# Step 3: Filter to relevant candidates
filtered_data = df_stats[df_stats['Furthest Recruiting Stage Reached'].isin(['in-house interview', 'offer sent', 'offer accepted', 'phone screen','new application'])]
filtered_data = filtered_data[filtered_data['Application Source'].isin(['Career Fair', 'Campus Event'])]

# Step 4: Calculate the in-house interview rates for each year
filtered_data['Year'] = pd.to_datetime(filtered_data['Date of Application']).dt.year
filtered_data['In-House'] = filtered_data['Furthest Recruiting Stage Reached'] == 'in-house interview'
filtered_data
# Step 5: Perform pairwise chi-squared tests and identify statistically significant differences
results = []
years = filtered_data['Year'].unique()

for i in range(len(years)):
    for j in range(i + 1, len(years)):
        year1_inhouse = filtered_data[(filtered_data['Year'] == years[i]) & filtered_data['In-House']].shape[0]
        year1_tot = filtered_data[(filtered_data['Year'] == years[i]) ].shape[0]
        year2_inhouse = filtered_data[(filtered_data['Year'] == years[j]) & filtered_data['In-House']].shape[0]
        year2_tot = filtered_data[(filtered_data['Year'] == years[j]) ].shape[0]
        
        contingency_table = pd.DataFrame(
            [[year1_inhouse, year1_tot],
             [year2_inhouse, year2_tot]],
            columns=['In-House', 'All Applicants'],
            index=[years[i], years[j]]
        )
        
        _, p_value, _, _ = chi2_contingency(contingency_table)
        results.append((years[i], years[j], p_value))

# Print results
results.sort(key=lambda x: x[2])
for result in results:
    year1, year2, p_value = result
    significance = 'Significant' if p_value < 0.05 else 'Not Significant'
    print(f"Comparison between {year1} and {year2}: p-value = {p_value:.4f} ({significance})")

#Q3 data prep


df_vis = df.copy(deep=True)

# Normalize the stages to lower case
df_vis['Furthest Recruiting Stage Reached'] = df_vis['Furthest Recruiting Stage Reached'].str.lower()

# Define the stages that involve human effort and bot stage
human_effort_stages = ['phone screen', 'in-house interview', 'offer sent', 'offer accepted']
bot_stage = ['new application']
team_effort_stages = ['in-house interview', 'offer sent', 'offer accepted']

# Add columns to mark if the furthest stage involves human effort, bot stage, team effort stage, and if offer is accepted
df_vis['Human_Effort_Involved'] = df_vis['Furthest Recruiting Stage Reached'].isin(human_effort_stages)
df_vis['Bot_Stage'] = df_vis['Furthest Recruiting Stage Reached'].isin(bot_stage)
df_vis['Team_Effort_Involved'] = df_vis['Furthest Recruiting Stage Reached'].isin(team_effort_stages)
df_vis['Offer_Accepted'] = (df_vis['Furthest Recruiting Stage Reached'] == 'offer accepted').astype(int)

# Calculate the metrics for each Application Source
grouped = df_vis.groupby('Application Source').agg(
    Total_Applicants=('Candidate ID Number', 'count'),
    Bot_Filtered=('Bot_Stage', 'sum'),
    Total_Candidates_Human_Effort=('Human_Effort_Involved', 'sum'),
    Team_Effort_Candidates=('Team_Effort_Involved', 'sum'),
    Accepted_Offers=('Offer_Accepted', 'sum')
).reset_index()

# Calculate effectiveness
grouped['Bot_Screen_rate'] = grouped['Bot_Filtered'] / grouped['Total_Applicants']
grouped['HR_Screen_rate'] = (grouped['Total_Candidates_Human_Effort']-grouped['Team_Effort_Candidates']) / grouped['Total_Candidates_Human_Effort']
grouped['Team_Effectiveness'] = grouped['Accepted_Offers'] / grouped['Team_Effort_Candidates']
grouped['Overall_Effectiveness'] = grouped['Accepted_Offers'] / grouped['Total_Candidates_Human_Effort']


grouped
grouped.to_csv(cur_dir +'/application_source_effectiveness.csv', index=False)