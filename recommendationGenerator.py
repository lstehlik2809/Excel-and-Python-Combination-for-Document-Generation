# -*- coding: utf-8 -*-
"""
Created on Jan 7 2023
RECOMMENDATION GENERATOR
@author: Ludek Stehlik (ludek.stehlik@gmail.com)
"""
# importing libraries
import os
import pandas as pd
import docx
from docx.shared import Pt

# setting working directory
os.chdir("path to your working directory with the excel file")

# uploading excel file with diagnoses and recommendations
recFile = pd.read_excel(
    './recommendationGenerator.xlsm', 
    sheet_name = 'Generator', 
    skiprows=[0,1,2,3,4,5],
    header = 0, 
    engine='openpyxl'
    )

# filtering selected diagnoses
selectedDiagnoses = recFile[(recFile["PRESENT"] == "Yes") |  (recFile["PRESENT"] == "yes")]

# resetting indexes
selectedDiagnoses.reset_index(inplace=True, drop=True)

# shell df
allTextsDf = pd.DataFrame(columns=["allTexts"])

# merging texts only from columns with some text in them and putting empty space between them
for i in range(0,len(selectedDiagnoses)):
    
    dfSupp = selectedDiagnoses.loc[[i]]
    dfSupp.dropna(axis=1, inplace=True)
    dfSupp.drop(columns = ["PRESENT"], inplace=True)
    # removing empty lines in individual cells
    dfSupp = dfSupp.applymap(lambda x: os.linesep.join([s for s in x.splitlines() if s.strip()]))
    dfSupp["allTexts"] = dfSupp[dfSupp.columns].apply(lambda x: "\n\n".join(x), axis =1)
    allTextsDf = pd.concat([allTextsDf, dfSupp[["allTexts"]]], axis=0)
    

# final text that includes all selected diagnoses
finalText = '\n\n\n\n'.join(allTextsDf["allTexts"])

# removing carriage return from the final text
finalText = finalText.replace("\r", "")
    
# creating report
doc = docx.Document()
# adding text
doc.add_paragraph(finalText)
# setting font
style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(11)
# saving report
doc.save('./recommendatios.docx')
        

        
    
