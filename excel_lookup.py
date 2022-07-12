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
applicationsFileName = 'PROCESSED_FILES_DO-NOT-EDIT/processed_applications.xlsx'
applicationsSheetName = 'Sheet1'

# Module information spreadsheet
moduleInformationFileName = 'PROCESSED_FILES_DO-NOT-EDIT/processed_module_information.xlsx'
moduleInformationSheetName = 'Sheet1'

######################################################################################################
# Begin Programme
######################################################################################################

# Display programme start message.
ms.startscript()
print()

# read the excel file - the applications
dfApp = pd.read_excel(applicationsFileName,sheet_name=applicationsSheetName,header=[0])

# read the excel file - the module details
dfModule = pd.read_excel(moduleInformationFileName,sheet_name=moduleInformationSheetName,header=[0])

# look up the desired informaiton
ms.who(dfApp,'PHAS0022','Selected') #can be 'Selected', 'Experience', or 'Matched'

