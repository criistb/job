import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook

df = pd.read_excel("../01. MDB/01 MDB 2018.xlsx", sheet_name="MASTER DATA BASES")

master = df.copy()

master[['Pos Emp Group','Employee Group']] = master[['Pos Emp Group','Employee Group']].replace(
    {'EMPLOYEE':'Colaborador','TEMPORARY':'Temporal','INTERN':'Aprendiz','CO-OP':'Aprendiz','APPRENTICE':'APRENDIZ'})

master[['Emp Gender']] = master[['Emp Gender']].replace({'F':'Mujeres','M':'Hombres'})

master[['Entity']] = master[['Entity']].replace('CEM-SUMMA','CEM')

cap = ['Parent Position','Job Level', 'Job', 'Division', 'Business Unit', 'Company', 'Position', 'Employee Subgroup']

for x in cap:
    master[x] = master[x].str.capitalize()
    
master = master.rename(columns={'Entity':'entitytemp'})

master['Entidad'] = np.where(master.entitytemp == 'CEM','Cementos',
                            np.where(master.entitytemp == 'GRA','Grupo Argos',
                                     np.where(master.entitytemp == 'ODI','Odinsa',
                                              np.where(master.entitytemp == 'CLS','Celsia',
                                                       np.where(master.entitytemp == 'SUMMA','Summa',"")))))

cols = ["Month", "Entidad", "Pos Code", "Position", "Pos Emp Group", 
        "Country", "Company Code", "Company", "Business Unit Code", "Business Unit",
        "Division Code", "Division", "Department Code", "Department", "Job Code", "Job", "Grade", "Salary Grade", "Job Level",
        "Unioniosed", "Parent Position Code", "Parent Position","Hr First Name", "Hr Last Name", "Hr Manager Pos Code",
        "Person Id","Payroll Id 1", "Payroll Id 2", "User Id", "Employee Status", "Employee Group", 
        "Emp Date Of Birth", "Emp Hire Date", "Emp Original Start Date", "Emp Gender", 
        "Emp Marital Status", "Location Code","Union 1","Union 2", "Emp First Name","Emp Last Name"]

master.drop(columns=[col for col in master if col not in cols], inplace=True)

master = master[cols]

cols = ["Entidad", "Position", "Country", "Company", "Business Unit", "Division",
"Job", "Parent Position", "Payroll Id 1", "Payroll Id 2", "User Id", "Emp Date Of Birth",
"Emp Gender", "Emp Marital Status", "Grade", "Job Level"]

#Master

masterof = master.copy()

masterof = masterof.drop(columns=[col for col in masterof if col in cols])

masterof['HR Manager'] = masterof['Hr First Name'].str.cat(masterof['Hr Last Name'])

masterof.drop(columns=['Hr First Name','Hr Last Name'], inplace=True)

cols = ["Month", "Pos Code", "Pos Emp Group", "Company Code", "Business Unit Code",
"Division Code", "Department Code", "Department", "Grouping Process", "Job Code",
"Salary Grade", "Unioniosed", "Parent Position Code", "HR Manager", 
"Hr Manager Pos Code", "Person Id", "Employee Status", "Employee Subgroup", 
"Employee Group", "Emp Hire Date", "Contract Type", "Work Relationship", "Benefits Type", 
"Emp Original Start Date", "Location Code", "Union 1", "Union 2"]

masterof = masterof.drop(columns=[col for col in masterof if col not in cols])

#Personas

mdbpersonas = master.copy()

cols = ["Month", "Person Id", "Payroll Id 1", "Payroll Id 2", "User Id", "Emp First Name",
       "Emp Last Name", "Emp Date Of Birth", "Emp Hire Date", "Emp Original Start Date",
       "Emp Gender", "Emp Marital Status"]

mdbpersonas = mdbpersonas.drop(columns=[col for col in mdbpersonas if col not in cols])

mdbpersonas = mdbpersonas.dropna(subset = 'Person Id')

mdbpersonas = mdbpersonas.sort_values('Month', ascending=False)

mdbpersonas = mdbpersonas.drop_duplicates(subset='Person Id')

mdbpersonas = mdbpersonas.drop(columns='Month')

cap = ['Emp Marital Status','Emp First Name', 'Emp Last Name']

for x in cap:
    mdbpersonas[x] = mdbpersonas[x].str.capitalize()


#Compañia

mdbcompany = master.copy()

cols = ["Month", "Entidad", "Country", "Company Code", "Company"]

mdbcompany = mdbcompany.drop(columns=[col for col in mdbcompany if col not in cols])

mdbcompany = mdbcompany.sort_values('Month', ascending=False)

mdbcompany = mdbcompany.drop_duplicates(subset='Company Code')

mdbcompany = mdbcompany.drop(columns='Month')

cols = ['Company Code','Entidad', 'Country', 'Company']

mdbcompany = mdbcompany[cols]

#Vicepresidencia

mdbvice = master.copy()

cols = ["Month", "Business Unit Code", "Business Unit"]

mdbvice = mdbvice.drop(columns=[col for col in mdbvice if col not in cols])

mdbvice = mdbvice.sort_values('Month', ascending=False)

mdbvice = mdbvice.drop_duplicates(subset='Business Unit Code')

mdbvice = mdbvice.drop(columns='Month')

#Division

mdiv = master.copy()

cols = ["Month", "Division Code", "Division"]

mdiv = mdiv.drop(columns=[col for col in mdiv if col not in cols])

mdiv = mdiv.sort_values('Month', ascending=False)

mdiv = mdiv.drop_duplicates(subset='Division Code')

mdiv = mdiv.drop(columns='Month')

#Job code

mdbjob = master.copy()

cols = ["Month", "Job Code", "Job"]

mdbjob = mdbjob.drop(columns=[col for col in mdbjob if col not in cols])

mdbjob = mdbjob.sort_values('Month', ascending=False)

mdbjob = mdbjob.drop_duplicates(subset='Job Code')

mdbjob = mdbjob.drop(columns='Month')

#Department

mddep = master.copy()

cols = ["Month", "Department Code", "Department"]

mddep= mddep.drop(columns=[col for col in mddep if col not in cols])

mddep = mddep.sort_values('Month', ascending=False)

mddep = mddep.drop_duplicates(subset='Department Code')

mddep = mddep.drop(columns='Month')

dfs = [masterof,mdbpersonas,mdbcompany,mdbvice,mdiv,mdbjob,mddep]
sheets = ["Hoja2","Hoja3","Hoja4","Hoja5","Hoja6","Hoja1","Hoja7"]
names = ['_01_Master','_02_Personas','_03_Company','_04_Vicepresidencia','_05_Division','_06_Job_Code','_07_Department']

with pd.ExcelWriter('../01. MDB/lol.xlsx') as writer:
    masterof.to_excel(writer, sheet_name=sheets[0],index=False)
    mdbpersonas.to_excel(writer, sheet_name=sheets[1],index=False)
    mdbcompany.to_excel(writer, sheet_name=sheets[2],index=False)
    mdbvice.to_excel(writer, sheet_name=sheets[3],index=False)
    mdiv.to_excel(writer, sheet_name=sheets[4],index=False)
    mdbjob.to_excel(writer, sheet_name=sheets[5],index=False)
    mddep.to_excel(writer, sheet_name=sheets[6],index=False)
    wb = writer.book
    for x,y,z in zip(dfs,sheets,names):
        tab = openpyxl.worksheet.table.Table(displayName=z, ref=f'A1:{chr(len(x.columns)+64)}{len(x)+1}')
        wb[y].add_table(tab)
    writer.save