####################################################
# OSPAR Application of Management 
# Effectiveness of Protected and Conserved 
# Areas (MEPCA) Assessment from OSPAR 
# Management Reporting
####################################################
# Tool to read OSPAR MPA management reporting and 
# convert responses to MEPCA Compatible outputs
####################################################
# Author:  Chris Moulton
# Credits: Chris Moulton @ OSPAR
# License: CC0
# Version: 1.0.0
# Date: 20241020
####################################################

import os
import numpy as np
import pandas as pd
import glob
from plotly import express as px
from collections import Counter

# The user inputs the folder where the collated MPA management Excel sheet is stored. Code looks into `inputfolder` for xlsx files:
inputfolder = input('Enter input folder path : ')
filename = glob.glob(inputfolder)

# Defines a data frame for the Excel to be read into:
read_df = pd.DataFrame()

# For each Excel file in the folder, sheet `MPA_MGT` is read into the data frame `read_df`:
for filename in os.listdir(inputfolder):
    inputfile = os.path.join(inputfolder, filename)
    df = pd.DataFrame(pd.read_excel(inputfile, sheet_name='MPA_MGT', header=0))
print (df.iloc[:,[0,3,4,6,8,10,12]])

# Basic QC on the entries to make sure the right type is captured:
df_prep = df.replace(to_replace=["yes", "partial", "no", "no response", "unknown", "Not applicable", "high", "moderate"], value=["Yes", "Partial", "No", "No response", "Unknown", "Not Applicable", "High", "Moderate"])
df_prep.to_excel(inputfolder + "/2-Addendum3_Management_MEPCA-qc.xlsx", sheet_name = "mepca-qc") # Output QC'd data to a new workbook

# Find and replace the entries in the QC result `df_prep`, in accordance with the MEPCA guidance:  \
# | Reponse | Score |
# |---------|-------|
# | "Yes" | 2 |  |
# | "Partial" | 1 |
# | "No" / "Unknown" | 0 | 
# 
# \
# 
# | Confidence | Score |
# |---------|-------|
# | "High" | 3 |
# | "Moderate" | 2 |
# | "Low" | 1 |
# | "Not applicable" | 0 |

df_mepca = df_prep.replace(to_replace=["Yes", "Partial", "No", "No response", "Unknown", "Not Applicable", "High", "Moderate", "Low", np.nan], value=[2, 1, 0, 0, 0, 0, 3, 2, 1, 9999])
df_mepca.to_excel(inputfolder + "/2-Addendum3_Management_MEPCA-scoreconv.xlsx", sheet_name = "score-conv") # Output an Excel to check the conversion values are correct and feed into mepca-indicator-score

# List all headers in `df_mepca`
for col in df_mepca.columns:
    print(col)

# Count the results by Response:
df_countmatrix = df_prep
df_countmatrix = df.groupby(['Country', 'a) Management documented: Response', 'b)Measures implemented: Response ', 'c) Monitoring in place: Response', 'd)Moving towards objectives - Response', 'd)Moving towards objectives - Confidence score'], as_index=False).size()
df_countmatrix.to_excel(inputfolder + "/5-Addendum3_Management_MEPCA-countmatrix.xlsx")

for col in df_countmatrix.columns:
    print(col)

# ## Module to read `mepca-calc` and `mepca-indicator-score` outputs, and plot

# Import data from `mepca-calc`, output from QC:
# inputfolder = input('Enter input file path: ')
filename = ('2-Addendum3_Management_MEPCA-qc.xlsx')

# Create dataframe and read in data
plotdf = pd.DataFrame()
inputfile = os.path.join(inputfolder, filename)
plotdf = pd.DataFrame(pd.read_excel(inputfile, header=0))
#print(plotdf.columns) # Can be exposed to check the headers that read into the plot.

# Define colours and labels
# RAL values converted using https://rgb.to/
# 'Yes': 'Green': Hex: '#00b050', RAL: 000 176 080
# 'Partial': 'Orange': Hex:'#f79646', RAL: 247 150 070 
# 'No': 'Red': Hex:'#c0504d', RAL: 192 080 077
# 'No response': 'Dark grey': Hex: '#404040', RAL: 064 064 064
# 'Unknown': 'Light grey': Hex:'#a6a6a6', RAL: 166 166 166

pielabelmap = {'Yes': 'Yes',
               'Partial': 'Partial',
               'No': 'No',
               'No response': 'No response',
               'Unknown': 'Unknown'}

colours = {'Yes': '#00b050',
            'Partial': '#f79646',
            'No': '#c0504d',
            'No response': '#404040',
            'Unknown': '#a6a6a6',
            'High': '#00b050',
            'Moderate': '#f79646',
            'Low': '#c0504d',
            'Not Applicable': '#a6a6a6' }

wordcounts = ['Yes,', 'Partial', 'No', 'No response', 'Unknown', np.nan]

# Plot pies
a = Counter(plotdf['a) Management documented: Response'])
print(a.items())
labelsa = list(a.keys())
valuesa = list(a.values())
plota = px.pie(values=valuesa, names=labelsa, color=a.keys(), color_discrete_map=colours,title="'a) Management documented: Response'")
plota.update_traces(textinfo='percent+value')
plota.show()
plota.write_image(inputfolder + '/plota.png')

b = Counter(plotdf['b)Measures implemented: Response '])
print(b.items())
labelsb = list(b.keys())
valuesb = list(b.values())
plotb = px.pie(values=valuesb, names=labelsb, color=b.keys(), color_discrete_map=colours, title="'b)Measures implemented: Response'")
plotb.update_traces(textinfo='percent+value')
plotb.show()
plotb.write_image(inputfolder + '/plotb.png')

c = Counter(plotdf['c) Monitoring in place: Response'])
print(c.items())
labelsc = list(c.keys())
valuesc = list(c.values())
plotc = px.pie(values=valuesc, names=labelsc, color=c.keys(), color_discrete_map=colours, title="'c) Monitoring in place: Response'")
plotc.update_traces(textinfo='percent+value')
plotc.show()
plotc.write_image(inputfolder + '/plotc.png')

d = Counter(plotdf['d)Moving towards objectives - Response'])
print(d.items())
labelsd = list(d.keys())
valuesd = list(d.values())
plotd = px.pie(values=valuesd, names=labelsd, color=d.keys(), color_discrete_map=colours, title="'d)Moving towards objectives: Response'")
plotd.update_traces(textinfo='percent+value')
plotd.show()
plotd.write_image(inputfolder + '/plotd.png')

e = Counter(plotdf['d)Moving towards objectives - Confidence score'])
print(e.items())
labelse = list(e.keys())
valuese = list(e.values())
plote = px.pie(values=valuese, names=labelse, color=e.keys(), color_discrete_map=colours, title="'d)Moving towards objectives: Confidence score'")
plote.update_traces(textinfo='percent+value')
plote.show()
plote.write_image(inputfolder + '/plote.png')

# Module to read `mepca-calc` outputs and produce `mepca-indicator-score`
# Contextual information:
# The mapping of OSPAR Management Status Questions to MEPCA Indicator Metrics:
# 
# | OSPAR Management Status Question | MEPCA Indicator Metrics |
# |----------------------------------|-------------------------|
# | Question A - Is the MPA management documented?	|	c) Is information on the PCA for management available? |
# | Question B - Are the measures to achieve the conservation objectives being implemented? |		d) Are management measures being implemented for the area to achieve its outcomes for conservation?	|
# | Question C - Is monitoring in place to assess if measures are working?	|	e) Does monitoring take place which helps to assess progress towards achieving conservation outcomes? |
# | Question D - Is the MPA moving towards or has it reached its conservation objectives?	|	f) Is the PCA achieving its conservation outcomes?	|
# | Confidence scores	| g) What is the level of confidence in the data used to assess progress towards the achievement of conservation outcomes? |
# 
# Scoring conversions:  
# | Reponse | Score |
# |---------|-------|
# | "Yes" | 2 |  |
# | "Partial" | 1 |
# | "No" / "Unknown" | 0 |  \
# 
# Confidence scores:
# | Confidence | Score |
# |---------|-------|
# | "High" | 3 |
# | "Moderate" | 2 |
# | "Low" | 1 |
# | "Not applicable" | 0 |

# ### Formulas used to calculate the MEPCA indicator score
# 
# (((c * 0.15) + (d * 0.25) + (e * 0.25) + (h * 0.35) / 3.4) * 100)  \
# c = QstnA, d = QstnB, e = QstnC, h = (QstnD * Confidence scores)  \
# Pass = n% > 38.24  \
# Inadequate = n% < 38.24  \
# 
# PartOne = QstnA * 0.15  \
# PartTwo = QstnB * 0.25  \
# PartThree = QstnC * 0.25  \
# PartFour = QstnD * Confidence scores  \
# PartFive = PartFour * 0.35  \
# PartSix = PartFive / 3.4  \
# PartSeven = PartOne + PartTwo + PartThree + PartSix  \
# IndicatorScore = PartSeven * 100

# Data Processing:
# The user inputs the folder where the collated MPA management Excel sheet is stored. Code looks into `inputfolder` for xlsx files:
# inputfolder = input('Enter input file path: ')
filename = ('2-Addendum3_Management_MEPCA-scoreconv.xlsx')

# Defines a data frame for the Excel to be read into:
read_df = pd.DataFrame()

# Reads in the MEPCA Score Conversion Excel file from `mepca-calc`, sheet `score-conv` is read into the data frame `df_ind`:
inputfile = os.path.join(inputfolder, filename)
df_ind = pd.DataFrame(pd.read_excel(inputfile, sheet_name='score-conv', header=0))
print(df_ind.columns) # Can be uncommented to check the headers that read into the calculation.

# Calculating MEPCA indicator score:
df_ind['PartOne'] = df_ind['a) Management documented: Response'].multiply(0.15)
df_ind['PartTwo'] = df_ind['b)Measures implemented: Response '].multiply(0.25)
df_ind['PartThree'] = df_ind['c) Monitoring in place: Response'].multiply(0.25)
df_ind['PartFour'] = df_ind['d)Moving towards objectives - Response'].multiply(df_ind['d)Moving towards objectives - Confidence score'])
df_ind['PartFive'] = df_ind['PartFour'].multiply(0.35)
df_ind['PartSix'] = df_ind['PartOne'] + df_ind['PartTwo'] + df_ind['PartThree'] + df_ind['PartFive']
df_ind['PartSeven'] = df_ind['PartSix'].divide(3.4)
df_ind['IndicatorScore'] = df_ind['PartSeven'].multiply(100)
df_ind['IndicatorScoreFin'] = np.where(df_ind['d)Moving towards objectives - Confidence score'] > 0, df_ind['PartSeven'].multiply(100), 0)
df_ind.to_excel(inputfolder + "/2-Addendum3_Management_MEPCA-FinalScore.xlsx", sheet_name="MEPCA-Scores")

# MEPCA Indicator Final Scores Histogram - No Exclusions - No indicator score exclusion factors have been applied
plotfshis = px.histogram(x=df_ind['IndicatorScore'], title="MEPCA Indicator Final Scores", nbins=11, labels={'x':'MEPCA Indicator Score (%)'}, text_auto=True)
plotfshis.add_vline(x=38.24, line_dash='dash', annotation_text="38.24% minimum threshold")
plotfshis.update_xaxes(dtick=10)
plotfshis.update_layout(yaxis_title="Count of MPAs")
plotfshis.show()
plotfshis.write_image(inputfolder + '/plotfshis.png', width=1000, height=1000)

# MEPCA Indicator Final Scores Histogram - Exclusions - Zero Confidence scores given a zero indicator score
plotfshisfin = px.histogram(x=df_ind['IndicatorScoreFin'], title="MEPCA Indicator Final Scores - Zero Confidence scores given a '0' Indicator Score", nbins=11, labels={'x':'MEPCA Indicator Score (%)'}, text_auto=True)
plotfshisfin.add_vline(x=38.24, line_dash='dash', annotation_text="38.24% minimum threshold")
plotfshisfin.update_xaxes(dtick=10)
plotfshisfin.update_layout(yaxis_title="Count of MPAs")
plotfshisfin.show()
plotfshisfin.write_image(inputfolder + '/plotfshisfin.png', width=1000, height=1000)

# MEPCA Indicator Final Scores Bar Chart - No Exclusions - No indicator score exclusion factors have been applied
fsbar = Counter(df_ind['IndicatorScore'])
print(fsbar.items())
keysfsbar = list(fsbar.keys())
valuesfsbar = list(fsbar.values())
plotfsbar = px.bar(x=keysfsbar, y=valuesfsbar, title="MEPCA Indicator Final Scores", labels={'x':'MEPCA Indicator Score (%)', 'y':'Count of MPAs'}, text_auto=True, barmode='stack')
plotfsbar.add_vline(x=38.24, line_dash='dash', annotation_text="38.24% minimum threshold")
plotfsbar.update_xaxes(dtick=5)
plotfsbar.show()
plotfsbar.write_image(inputfolder + '/plotfsbar.png', width=1200, height=1000)

# MEPCA Indicator Final Scores Bar Chart - Exclusions - Zero Confidence scores given a zero indicator score
fsbarfin = Counter(df_ind['IndicatorScoreFin'])
print(fsbarfin.items())
keysfsbarfin = list(fsbarfin.keys())
valuesfsbarfin = list(fsbarfin.values())
plotfsbarfin = px.bar(x=keysfsbarfin, y=valuesfsbarfin, title="MEPCA Indicator Final Scores - Zero Confidence scores given a '0' Indicator Score", labels={'x':'MEPCA Indicator Score (%)', 'y':'Count of MPAs'}, text_auto=True, barmode='stack')
plotfsbarfin.add_vline(x=38.24, line_dash='dash', annotation_text="38.24% minimum threshold")
plotfsbarfin.update_xaxes(dtick=5)
plotfsbarfin.show()
plotfsbarfin.write_image(inputfolder + '/plotfsbarfin.png', width=1200, height=1000)

bindata = df_ind['IndicatorScore']

# Define custom bin edges
bin_edges = [0, 20, 40, 60, 80, 100]
 
# Use numpy's histogram function with custom bins
hist, bins = np.histogram(bindata, bins=bin_edges)
 
# Print the result
print("Bin Edges:", bins)
print("Histogram Counts:", hist)

# MEPCA Indicator Final Scores Pie Chart - No Exclusions
plotbin = px.pie(values=hist, names=['0-20', '20-40', '40-60', '60-80', '80-100'], color_discrete_map=colours,title="MEPCA Indicator Final Scores - No Exclusions")
plotbin.update_traces(textinfo='percent+value')
plotbin.show()
plotbin.write_image(inputfolder + '/plotbin.png')

binfindata = df_ind['IndicatorScoreFin']

# Define custom bin edges
binfin_edges = [0, 38.24, 60, 80, 100]
 
# Use numpy's histogram function with custom bins
hist, bins = np.histogram(binfindata, bins=binfin_edges)
 
# Print the result
print("Bin Edges:", bins)
print("Histogram Counts:", hist)

# MEPCA Indicator Final Scores Pie Chart - Exclusions - Zero Confidence scores given a zero indicator score
plotbinfin = px.pie(values=hist, names=['0-38.24', '38.25-60', '60-80', '80-100'], color_discrete_map=colours,title="MEPCA Indicator Final Scores - Zero Confidence scores given a '0' Indicator Score")
plotbinfin.update_traces(textinfo='percent+value')
plotbinfin.show()
plotbinfin.write_image(inputfolder + '/plotbinfin.png')