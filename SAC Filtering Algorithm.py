from itertools import filterfalse
from pandas.core.api import Index
import pandas as pd

#INSIDE THE QUOTES, PUT THE NAME OF THE FILE YOU ARE FILTERING. MAKE SURE YOU INCLUDE THE ".xlsx"
#CHANGE THE NUMBER IN SHEET_NAME ONLY IF YOUR EXCEL HAS MULTIPLE TABS. THis NUMBER SPECIFIES THE TAB YOU WANT TO FILTER: FIRST TAB BEING = 0, SECOND TAB =  1, ETC.
df_main = pd.read_excel('Middlesex3.xlsx', sheet_name= 0) 

#This should not be changed; this is the information for the reference file
df_reference = pd.read_excel('sheet.xlsx', sheet_name= 0)

#This should not be changed; this is the information for the new filtered file
df_new = pd.read_excel('Starter Datasheet Template.xlsx', sheet_name=0) 

last_name_previous = ""
last_name = ""

#Change this to the total number of data entries minus 1
for x in range(29578): 
  status = True
  exitkey = 14
  strOriginal = df_main.loc[x, "Owner's Name"]
  if (exitkey == 14):
    if (strOriginal.find(',') > -1):
      count = strOriginal.count(",")
      if (count > 1):
        status = False
      parts = strOriginal.split(",")
      if len(parts) == 2:
        last_name_previous = last_name
        last_name = parts[0]
        first_name = parts[1]
    else:
      last_name = strOriginal
      status = False

    #Do not change this; it traverses the reference file
    for y in range(149): 
      strCompare = df_reference.loc[y, 'Name']
      if (last_name.lower() == strCompare.lower()):
        if (status):
          if (last_name_previous != last_name):
            if (exitkey == 14):
              new_row = {'Last Name': df_main.loc[x, "Owner's Name"], 'Address': df_main.loc[x, "Owner's Mailing Address"], 'Town': df_main.loc[x, "City/State/Zip"]}
              df_new.loc[len(df_new)] = new_row
              y = 250
              df_main = df_main.drop(x)



# This is the name of the filtered file. Replace the name with the name of your town and make sure you finish the name with a ".xlsx"
df_new.to_excel('Woodbridge Filtered.xlsx', index=False) 