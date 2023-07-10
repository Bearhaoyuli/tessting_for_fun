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

os.makedirs(cur_dir+'/Q2_Outputs',exist_ok=True)


df = pd.read_csv(cur_dir + "/Q1_Outputs/1.Raw_data_output.csv")

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

        # Create a contigency table and fit a Chi-Squred test
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

block = """
\n
"""
conclusion = """
A chi-squared contingency test was performed on each pair of in-house candidate rates.\n
Results indicate that the in-house candidate rates between 2017 and 2018 are statistically significant \n
(p-value < 0.05). Meaning we can be 95% confident to say that the rate between the two years are different\n
Rates between 2016 & 2017 and 2016 & 2018 are not statistically significant.
"""

print(block,conclusion)