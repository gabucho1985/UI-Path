#!/usr/bin/env python
# coding: utf-8

# In[ ]:





# In[ ]:


#this code takes folders names and add them into a list(source_dir) then takes ralph list and compare. If there are any matches
#then extract excel file from those matches using source_dir and paste them in dest(path)
import numpy as np
import pandas as pd
import os, sys
import shutil
import random
import re  
source_dir ="\\\\isb-51JFKFS\\BSA\\BSA_EDD\\UiPath\\Output Results\\\KYC Cases Dec 19 Re-run"
#source_dir = "\\\\isb-51JFKFS\\BSA\\BSA_EDD\\UiPath\\Output Results\\December 2020- KYC Cases"
#dest = "\\\\isb-51jfkfs\\BSA\\BSA_EDD\\KYC Population December 2020\\KYC Population 12-18-2020"
dest="\\\\isbnj-corp\\PUBLICSHARE\\KYC Reporting QS\\Comind\\KYC Production\\12-24-2020\\UiPath files"
group_cases_CRA = pd.read_excel(r'C:\Users\GMacias\Documents\CRA Cases 09.01.2020.xlsx', dtype=str) #ralph_list
mylist = group_cases_CRA['TIN'].tolist()
result= []

def convert_to_lists(source_cases, ralph_list,res):
    dirs = os.listdir(source_cases) 
    ralph_list = mylist
    result_not_match= list(set(ralph_list) - set(dirs))
    res=list(set(ralph_list) & set(dirs)) #keep matches
    res = [case for case in res] #convert to lists
    print("List of matches: " + "\n" + "Total: " + str(len(res)) + '\n' + str(res) + '\n'
          "No match in folder: " + "\n"+ "Total: " + str(len(result_not_match)) + '\n' + str(result_not_match))
    return res


def copy_to_destination_folder(src, dest):
    raz=convert_to_lists(source_dir, mylist,result)
    for root, dirs, files in os.walk(src):
        for i in raz:
            #if root.endswith('EDD Forms'):
            g=os.path.join(root, str(i))
            add_path=os.path.join(g, 'EDD Forms')
            for file in os.listdir(add_path):
                jamon=os.path.join(add_path, file)
                newPath = shutil.copy(jamon, dest)
                    
#convert_to_lists(source_dir, mylist,result)
copy_to_destination_folder(source_dir,dest)


# In[ ]:


#compare CRA Cases list of folders against sourc_dir path folders and find total cases, matches and missing cases
import os, sys
import shutil
import random
import re
import xlsxwriter
import re
source_dir = "\\\\isb-51JFKFS\\BSA\\BSA_EDD\\UiPath\\Output Results\\December 2020- KYC Cases"
# dest = "\\\\isb-51jfkfs\\BSA\\BSA_EDD\\GabrielFolder" 
group_cases_CRA = pd.read_excel(r'C:\Users\GMacias\Documents\CRA Cases 09.01.2020.xlsx', dtype=str)
mylist = group_cases_CRA['TIN'].tolist()
#path = os.getcwd() 
dirs = os.listdir(source_dir) 
# print(dirs)
print(len(dirs))

result_match = mylist
lista=(list(dirs))
#listas = re.findall("\d+", lista)[0]
#print(lista)
result_not_match= list(set(result_match) - set(lista))
result_match= list(set(result_match) & set(lista)) 
#print("List of matches: " + str(len(result_match)))
print("No match in folder: " + str(result_not_match))
print("Matches in folder: " + str(result_match))

workbook = xlsxwriter.Workbook('EDD CASES DECEMBER 18TH 2020.xlsx')
worksheet = workbook.add_worksheet()      

List  = [mylist,result_match, result_not_match]

row_num = 1 

for col_num, data in enumerate(List):
    worksheet.write_column(row_num, col_num, data)
worksheet.write(0, 0, "TOTAL CASES")
worksheet.write(0, 1, "MATCHES")
worksheet.write(0, 2, "MISSING CASES")
workbook.close()


# In[ ]:


#this code is the 2nd part from code above. I compare source_dir against matches(using MATCH column) 
#then write to batch file(myBat) in order to move folders a lot quicker using powershell
import numpy as np
import pandas as pd
import os, sys
import shutil
import random
import re 
source_dir = "\\\\isb-51JFKFS\\BSA\\BSA_EDD\\UiPath\\Output Results\\December 2020- KYC Cases"
source_dir = os.listdir(source_dir) #big list
matches = pd.read_excel(r"C:\\Users\\GMacias\\EDD CASES DECEMBER\\EDD CASES DECEMBER 18TH 2020.xlsx", dtype=str) #ralph_list
mylist  = matches['MATCHES'].tolist() #list you substract from always check
mylist  = [x for x in mylist if str(x) != 'nan']
result_match= list(set(source_dir) & set(mylist))
final_list= []
print(result_match)
print(len(result_match))

# file_path = 'C:\\\\Users\\GMacias\\Documents\\bash_exec\\Dados.bat'
# os.remove(file_path)


myBat = open(r'C:\Users\GMacias\Documents\bash_exec\Dados.bat','w+')
#myBat.write("/s/q \\isb-51jfkfs\BSA\BSA_EDD\Gabriel33"'\\'+ str(i))
#line to delete paths below
for path in result_match:
    line = "powershell -Command " +'"'+ "Copy-Item \\\\isb-51JFKFS\BSA\BSA_EDD\Gabriel33"'\\'+ str(path) + " C:\\Users\\GMacias\\newfolder"+'"' + '\n'
    line2 ="powershell -Command Robocopy " + "'"+"\\\\isb-51JFKFS\BSA\BSA_EDD\Gabriel33"+'\\'+ str(path)+ "'" + " '"+"C:\\Users\\GMacias\\newfolder"+'\\'+ str(path)+ "'" + " /S /E /MT:32 /NFL /NDL /NJH /NJS" + '\n'
   
    #line = "rmdir /s/q \\\isb-51jfkfs\BSA\BSA_EDD\Gabriel"'\\'+ str(path) + '\n'
    #line = "move \\isb-51JFKFS\BSA\BSA_EDD\Gabriel33"'\\'+ str(path) +" " +"\\isb-51JFKFS\BSA\BSA_EDD\Gabriel"'\n'
    myBat.write(line)
    myBat.write(line2)
    print(line)
    print(line2)  
    #line =Robocopy '\\isb-51jfkfs\BSA\BSA_EDD\UiPath\Output Results\Orig EDD CasesNovember-2020' '\\isb-51JFKFS\BSA\BSA_EDD\Gabriel2' /S /E /MT:32 /NFL /NDL /NJH /NJS
myBat.close()







# In[ ]:


# This is the third part where we call Dados.bat file to move matching folders 
#to another destination path C:\\Users\\GMacias\\newfolder
import subprocess
subprocess.call([r'C:\Users\GMacias\Documents\bash_exec\Dados.bat'])

