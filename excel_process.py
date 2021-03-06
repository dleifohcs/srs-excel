#!/Users/steven/opt/anaconda3/bin/python3

####################################################################################################
# This code is for working with Excel sheets created by Microsoft Forms
# Steven R. Schofield, University College London, Apr. 2022
####################################################################################################

# Load required packages
import os
import sys
import numpy as np
import pandas as pd

# load main functions
import SRSExcel as ms

# Define a default filename to work with in case none is specified on the command line.
applicationsFileName = 'APPLICATIONS-DO-NOT-EDIT/PGTA_applications_2022-23 1.xlsx'
applicationsSheetName = 'Form1'

# Module information spreadsheet
moduleInformationFileName = 'MANUAL_INPUT-DO-NOT-CHANGE-HEADERS/Module_Information_2021-22.xlsx'
moduleInformationSheetName = 'Sheet1'

######################################################################################################
# Begin Programme
######################################################################################################

# Display programme start message.
ms.startscript()
print()

# read the excel file - the applications
dfApp = pd.read_excel(applicationsFileName,sheet_name=applicationsSheetName,header=[0])

# Process the applications spreadsheet to a new data frame
dfApp = ms.processRawApplications(dfApp)

# write the processed data frame to excel sheet.
dfApp.to_excel('PROCESSED_FILES_DO-NOT-EDIT/processed_applications.xlsx')

# read the excel file - the module details
dfModule = pd.read_excel(moduleInformationFileName,sheet_name=moduleInformationSheetName,header=[0])

# remove trailing spaces that can occur in the excel entries - only does the header, might need to do entries later??
dfModule = ms.cleanColumnHeader(dfModule)

# this is simple lookup function
#ms.who(dfProcessed,'PHAS0022','Selected') #can be 'Selected', 'Experience', or 'Matched'

dfModule, dfAssigned = ms.makeAssignments(dfModule,dfApp)

# Write the module information to a processed version
dfModule.to_excel('PROCESSED_FILES_DO-NOT-EDIT/processed_module_information.xlsx')

# write the processed data frame to excel sheet.
dfApp.to_excel('PROCESSED_FILES_DO-NOT-EDIT/processed_applications.xlsx')
