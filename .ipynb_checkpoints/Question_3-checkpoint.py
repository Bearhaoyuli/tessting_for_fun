# Importing libraries 
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

os.makedirs(cur_dir+'/Q3_Outputs',exist_ok=True)


df = pd.read_csv(cur_dir + "/Q1_Outputs/1.Raw_data_output.csv")


df_vis = df.copy(deep=True)


# Step 1. Calcualte the Effectiveness for each application source. 

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

# Calculate HR_Filtered and Team_effort_not_hired for each row
df_vis['HR_Filtered'] = df_vis['Human_Effort_Involved'] & (~df_vis['Team_Effort_Involved'])
df_vis['Team_effort_not_hired'] = df_vis['Team_Effort_Involved'] & (~df_vis['Offer_Accepted'])

# Convert boolean values to integers
df_vis['HR_Filtered'] = df_vis['HR_Filtered'].astype(int)
df_vis['Team_effort_not_hired'] = df_vis['Team_effort_not_hired'].astype(int)

# Define the list of item combinations for the for loop
item_combinations = [['Application Source'], ['Application Source', 'Department']]

for items in item_combinations:
    # Calculate the metrics for each Application Source
    grouped = df_vis.groupby(items).agg(
        Total_Applicants=('Candidate ID Number', 'count'),
        Bot_Filtered=('Bot_Stage', 'sum'),
        HR_Filtered = ('HR_Filtered','sum'),
        Team_effort_not_hired = ('Team_effort_not_hired','sum'),
        Total_Candidates_Human_Effort=('Human_Effort_Involved', 'sum'),
        Team_Effort_Candidates=('Team_Effort_Involved', 'sum'),
        Accepted_Offers=('Offer_Accepted', 'sum'),
        Avg_Years_Experience=('Years of Experience', lambda x: x[df_vis['Offer_Accepted'] == 1].mean())
    ).reset_index()

    # Calculate effectiveness
    grouped['Bot_Screen_rate'] = grouped['Bot_Filtered'] / grouped['Total_Applicants']
    grouped['HR_Screen_rate'] = grouped['HR_Filtered'] / grouped['Total_Candidates_Human_Effort']
    grouped['Team_Effectiveness'] = grouped['Accepted_Offers'] / grouped['Team_Effort_Candidates']
    grouped['Overall_Effectiveness'] = grouped['Accepted_Offers'] / grouped['Total_Candidates_Human_Effort']

    # Sum the specified columns for the 'Overall' row
    total_applicants = grouped['Total_Applicants'].sum()
    bot_filtered = grouped['Bot_Filtered'].sum()
    hr_filtered = grouped['HR_Filtered'].sum()
    team_effort_not_hired = grouped['Team_effort_not_hired'].sum()
    total_candidates_human_effort = grouped['Total_Candidates_Human_Effort'].sum()
    team_effort_candidates = grouped['Team_Effort_Candidates'].sum()
    accepted_offers = grouped['Accepted_Offers'].sum()

    # Calculate the rates for the 'Overall' row
    bot_screen_rate = bot_filtered / total_applicants
    hr_screen_rate = hr_filtered / total_candidates_human_effort
    team_effectiveness = accepted_offers / team_effort_candidates
    overall_effectiveness = accepted_offers / total_candidates_human_effort
    avg_years_experience = df_vis.loc[df_vis['Offer_Accepted'] == 1, 'Years of Experience'].mean()

    # Create a new DataFrame for the 'Overall' row
    overall = pd.DataFrame({
        items[0]: ['Overall'],
        'Total_Applicants': [total_applicants],
        'Bot_Filtered': [bot_filtered],
        'HR_Filtered':[hr_filtered],
        'Team_effort_not_hired' :[team_effort_not_hired],
        'Total_Candidates_Human_Effort': [total_candidates_human_effort],
        'Team_Effort_Candidates': [team_effort_candidates],
        'Accepted_Offers': [accepted_offers],
        'Avg_Years_Experience': [avg_years_experience],
        'Bot_Screen_rate': [bot_screen_rate],
        'HR_Screen_rate': [hr_screen_rate],
        'Team_Effectiveness': [team_effectiveness],
        'Overall_Effectiveness': [overall_effectiveness]
    })

    # Append the 'Overall' row to the grouped DataFrame
    grouped = pd.concat([grouped, overall], ignore_index=True)

    # Save the DataFrame to a CSV file
    filename = '_'.join(items).lower() + '_effectiveness.csv'
    grouped.to_csv(cur_dir + '/Q3_Outputs/' + filename, index=False)
    print(f"Finished groupby combination: {items}")

#Step 2 
#Data prep for Waterfall
df_wf = pd.read_csv(cur_dir+'/Q3_Outputs/application source_department_effectiveness.csv')
df_wf = df_wf[(df_wf['Application Source'] != 'Overall') & (df_wf['Department'] != 'Overall')]


# Select the relevant columns and apply transformations
df_wf_transformed = df_wf[['Application Source', 'Department', 'Total_Applicants', 'Bot_Filtered', 'HR_Filtered', 'Team_effort_not_hired', 'Accepted_Offers']].copy()
df_wf_transformed['Bot_Filtered'] = -df_wf_transformed['Bot_Filtered']
df_wf_transformed['HR_Filtered'] = -df_wf_transformed['HR_Filtered']
df_wf_transformed['Team_effort_not_hired'] = -df_wf_transformed['Team_effort_not_hired']

# Melt the dataframe
df_wf_melted = df_wf_transformed.melt(id_vars=['Application Source', 'Department'], var_name='Attribute', value_name='Measure')
df_wf_melted
# df_wf_melted.to_csv(cur_dir + '/Q3_Outputs/' + "Waterfall.csv", index=False)
df_wf_melted.to_excel(cur_dir+'/Q3_Outputs/Waterfall.xlsx', engine='xlsxwriter',index=False)  
print("Job Finished")