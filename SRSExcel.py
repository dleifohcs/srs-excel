#!/Users/steven/opt/anaconda3/bin/python3

####################################################################################################
# This code is for working with Excel sheets created by Microsoft Forms
# Steven R. Schofield, University College London, Apr. 2022
# GNU General Public License v3.0
####################################################################################################


####################################################################################################
def startscript():
  """This provides a header output when the programme starts
  """
  print("SRS MSForm tools ***************************************************************************")
  print("Steven R. Schofield, University College London (April 2022, GNU General Public License v3.0)")


####################################################################################################  
def cleanColumnHeader(df):
    """removes any trailing spaces on column headers
    """
    
    col=df.columns
    newcols = {}
    for word in col:
        newcols[word] = word.rstrip()
    df.rename(columns=newcols,inplace=True)
    
    return df


####################################################################################################  
def dfLookupSingle(df,a,b,c):
    """Find 'b' in column 'a' and return the corresponding 'c'
    """

    i = df[df[a] == b].index[0]
    ans = df.loc[i,c]
    
    return ans

####################################################################################################  
def dfLookupSubDF(df,a,b):
    """Return a data frame that is a subset of the original matching 'b' in category 'a'
    """

    subDF = df[df[a] == b]

    return subDF


####################################################################################################  
def processModuleSelection(modulesStr):
    """Takes in the string that MS Forms generated from the module selection dialogue. This consists
       of a module code and a description. This function throws away everything exept the module
       code
    """
    modulesStr = str(modulesStr)
    # replace the ';' separators with spaces and ':' with nothing.
    modulesStr = modulesStr.replace(';',' ')
    modulesStr = modulesStr.replace(',',' ')
    modulesStr = modulesStr.replace(':','')
    
    modulesList = []
    for word in modulesStr.split():
        if word[0:4] == 'PHAS':
            modulesList.append(word.replace(":",""))
            
    return modulesList


####################################################################################################  
def processRawApplications(df):
    """Takes in a data frame of the original MS Forms excel spreadsheet data and processes this 
       into a form that we will use to make the PGTA assignments. Returns a new data frame.
    """

    import pandas as pd
    
    nameList = []
    emailList = []
    supervisorList = []
    roleList = []
    modList = []
    expList = []
    matchList = []
    assignedList = []

    # Go through the Module data frame loaded from the application spreadsheet and make lists to write
    # a new dataframe. This loop also compares the modules that were selected to those indicated
    # as past experience. When there is a match, this is written to a new list.  All the data is
    # added to a new data frame and this is written to an excel sheet.
    for i, row in df.iterrows():
        # Name
        nameStr = str(row['Your FIRST NAME'])+" "+str(row['Your LAST NAME'])
        # Email
        emailStr = row['Email address (must be a UCL email address if you have one)']
        # Supervisor
        supervisorStr = row['Line manager or supervisor']
        # Role
        roleStr = row['Current status']
        # Modules requested
        modSelStr = row['Please select all modules you would be willing to teach']
        modSelStr = processModuleSelection(modSelStr)
        # Modules experience
        modExpStr = row['If you have PREVIOUSLY TAUGHT any of these modules, please select those here']
        modExpStr = processModuleSelection(modExpStr)

        matchStr = ''
        # Check if selected modules match experience in a module
        if len(modSelStr) != 0:
            for mod in modSelStr:
                for modExp in modExpStr:
                    if mod == modExp:
                        matchStr = matchStr+mod+';'
            matchStr = processModuleSelection(matchStr)

        # append the values to the lists
        nameList.append(nameStr)
        emailList.append(emailStr)
        roleList.append(roleStr)
        supervisorList.append(supervisorStr)
        modList.append(' '.join(modSelStr))
        expList.append(' '.join(modExpStr))
        matchList.append(' '.join(matchStr))
        assignedList.append('')

        # Create a python set for the gathered data
        data = {'Name': nameList, 'Email': emailList, 'Role': roleList, 'Supervisor': supervisorList,
                    'Selected': modList, 'Experience': expList, 'Matched': matchList,
                    'Assigned': assignedList}

        # Make the set into a pandas data frame
        dfProcessed = pd.DataFrame(data)
        
    return dfProcessed

####################################################################################################  
def moduleList(df):
    """Takes in a data frame of the original spread sheet of the module details. Returns a list
    of modules and the number of PGTAs required.
    """

    moduleList = []
    requiredList = []

    for i, row in df.iterrows():
        codeStr = row['Code']
        targetStr = row['PGTA target']

        moduleList.append(codeStr)
        requiredList.append(targetStr)
        
    return moduleList, requiredList

####################################################################################################  
def who(df,moduleName,criteria):
    """ Takes in the processed data frame of the applications, the name of the module of interest
    and the criteria of interest - i.e., "Selected", "Experience" or "Matched"
    """
    
    names = df.loc[df[criteria].str.contains(moduleName, case=True)]['Name']
    email = df.loc[df[criteria].str.contains(moduleName, case=True)]['Email']
    supervisor = df.loc[df[criteria].str.contains(moduleName, case=True)]['Supervisor']

    print('List of people on module',moduleName, 'matching the criteria:',criteria,':')
    print()
    i = 0
    for name in names:
        i+=1
        print(i, ' ', name)
    print()
    return 


####################################################################################################  
def makeAssignments(dfModule,dfApplications):
    """ Takes in the processed data frame of the applications, the name of the module of interest
    and the criteria of interest - i.e., "Selected", "Experience" or "Matched"
    """
    # Loop over the modules from the info DF, read PGTAs from applications DF and write
    # back to the info DF with the selected PGTAs
    for i, rowModule in dfModule.iterrows():
        # get the name of the module to work with on this loop iteration
        moduleCode = rowModule['Code']
        # get the number of PGTAs required for this module
        requiredNum = rowModule['PGTA target']
        # pgtaList is the list of PGTAs for this module
        pgtaList = []
        # number of PGTAs so far assigned to this module
        assignedNum = 0
        # iterate through the PGTA applications to find PGTAs
        for j, rowApplications in dfApplications.iterrows():
            # look in the Matched column. If the module code is there, this will be a positive integer
            ismatched = rowApplications['Matched'].find(moduleCode)
            # if ismatched is -1 then there was no match, but otherwise...
            if ismatched != -1 and rowApplications['Assigned'] == "":
                pgtaName = rowApplications['Name']
                pgtaList.append(pgtaName)
                dfApplications.at[j,'Assigned'] = moduleCode
                assignedNum += 1
            # If we have assigned the required number of PGTAs then stop the loop
            if assignedNum == requiredNum:
                break

        # Turn the list into a semicolon separated string of names for writing to the DF
        pgtaListStr = ''
        for name in pgtaList:
            pgtaListStr = pgtaListStr + name + ';'
        dfModule.at[i,'AssignedPGTAs'] = pgtaListStr
        
    print(dfModule.head())
    print(dfApplications.head())
    return dfModule, dfApplications
