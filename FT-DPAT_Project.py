import pandas as pd
import numpy as np
import os
from pathlib import Path

print("FT-DPAT Project \n By Dimas Emiliano Trejo \n March, 2023")
path = Path(input('Enter FT-SPEC path file: '))
path = Path(path)
# C:\Users\trejode\Desktop\PDSE\FT-DPATs\53735-21\FT-SPEC-53735-21G_RevA.xlsx
# C:\Users\trejode\Desktop\PDSE\FT-DPATs\53812-24\221122\53812-24 FT-SPEC-53812-24G_RevD_FNX.xlsx
# C:\Users\trejode\Desktop\PDSE\FT-DPATs\53834-14\DRG_FT-SPEC-53834-14_I\DRG_FT-SPEC-53834-14_I.xlsx
# C:\Users\trejode\Desktop\specprueba.xlsx
dataframe = pd.read_excel(path, header = 1)
df = dataframe.iloc[:, [dataframe.columns.get_loc('Test #'),dataframe.columns.get_loc('Test Name'),dataframe.columns.get_loc('Units'),dataframe.columns.get_loc('Bin 1'),dataframe.columns.get_loc('Bin 1')+1,dataframe.columns.get_loc('SWBIN Name'),dataframe.columns.get_loc('PAT Limit Constraints')-1,dataframe.columns.get_loc('PAT Limit Constraints'),dataframe.columns.get_loc('PAT Limit Constraints')+1,dataframe.columns.get_loc('Disable Test')]]
df.columns = ['Test #', 'TestName', 'Units', 'Bin1_LSL', 'Bin1_USL','SoftwareBin','Skip?','Constraint_LSL','Constraint_USL','Disable Test']
df.drop([0],axis="index", inplace = True)
df.loc[(df['Disable Test'].str.capitalize() =='No'), 'Skip?']='No'
#df['Constraint_LSL','Constraint_USL']=['','']

#NO CONSTRAINTS - SoftwareBin / CONTINUITY, TRIM, SWITCH, REGISTER, ISO
df.loc[(df['Disable Test'].str.capitalize() =='No') & ((df.SoftwareBin.str.contains('CONT')) | (df.SoftwareBin.str.contains('Continuity')) ), ['Skip?','Constraint_LSL','Constraint_USL']]=['Yes','','']
df.loc[(df['Disable Test'].str.capitalize() =='No') & ( (df.SoftwareBin.str.contains('TRIM')) | (df.SoftwareBin.str.contains('Trim')) ), ['Skip?','Constraint_LSL','Constraint_USL']]=['Yes','','']
df.loc[(df['Disable Test'].str.capitalize() =='No') & ( (df.SoftwareBin.str.contains('SW')) | (df.SoftwareBin.str.contains('Sw')) | (df.SoftwareBin.str.contains('READ')) | (df.SoftwareBin.str.contains('ISO')) ), ['Skip?','Constraint_LSL','Constraint_USL']]=['Yes','','']

#NO CONSTRAINTS - TestName / PI, PO, P1DB, ORL, RL, IL, MSW G6_IDD, G6_G
df.loc[(df['Disable Test'].str.capitalize() =='No') & ( (df.TestName.str.contains('PO')) | (df.TestName.str.contains('PI')) | (df.TestName.str.contains('P1DB')) | (df.TestName.str.contains('ORL')) ), ['Skip?','Constraint_LSL','Constraint_USL']]=['Yes','','']
df.loc[(df['Disable Test'].str.capitalize() =='No') & ( (df.TestName.str.contains('_IL')) | (df.TestName.str.contains('_RL')) | (df.TestName.str.contains('MSW')) ), ['Skip?','Constraint_LSL','Constraint_USL']]=['Yes','','']
df.loc[(df['Disable Test'].str.capitalize() =='No') & ( (df.TestName.str.startswith('G6_')) &  ( (df.TestName.str.contains('IDD_')) | (df.TestName.str.contains('G')) )), ['Skip?','Constraint_LSL','Constraint_USL']]=['Yes','','']

#YES - G0/1/2/3_G/IDD
#df.loc[(df['Disable Test'].str.capitalize() =='No') & ( ((df.TestName.str.contains('G0')) | (df.TestName.str.contains('G1')) | (df.TestName.str.contains('G2')) | (df.TestName.str.contains('G3'))) & (df.TestName.str.contains('_G_') | df.TestName.str.contains('IDD'))), 'Skip?']='No'
df.loc[(df['Disable Test'].str.capitalize() =='No') & ( ((df.TestName.str.contains('G0')) | (df.TestName.str.contains('G1')) | (df.TestName.str.contains('G2')) | (df.TestName.str.contains('G3'))) & (df.TestName.str.contains('_G_') | df.TestName.str.contains('IDD'))), 'Constraint_LSL']= df['Bin1_LSL']+0.5
df.loc[(df['Disable Test'].str.capitalize() =='No') & ( ((df.TestName.str.contains('G0')) | (df.TestName.str.contains('G1')) | (df.TestName.str.contains('G2')) | (df.TestName.str.contains('G3'))) & (df.TestName.str.contains('_G_') | df.TestName.str.contains('IDD'))), 'Constraint_USL']= df['Bin1_USL']-0.5

#YES - G4/5/6/7_G/IDD
#df.loc[(df['Disable Test'].str.capitalize() =='No') & ( ((df.TestName.str.contains('G4')) | (df.TestName.str.contains('G5')) | (df.TestName.str.contains('G6')) | (df.TestName.str.contains('G7'))) & (df.TestName.str.contains('_G_') | df.TestName.str.contains('IDD'))), 'Skip?']='No'
df.loc[(df['Disable Test'].str.capitalize() =='No') & ( ((df.TestName.str.contains('G4')) | (df.TestName.str.contains('G5')) | (df.TestName.str.contains('G6')) | (df.TestName.str.contains('G7'))) & (df.TestName.str.contains('_G_') | df.TestName.str.contains('IDD'))), 'Constraint_LSL']= df['Bin1_LSL']+0.5
df.loc[(df['Disable Test'].str.capitalize() =='No') & ( ((df.TestName.str.contains('G4')) | (df.TestName.str.contains('G5')) | (df.TestName.str.contains('G6')) | (df.TestName.str.contains('G7'))) & (df.TestName.str.contains('_G_') | df.TestName.str.contains('IDD'))), 'Constraint_USL']= df['Bin1_USL']-0.5

#YES - IIP3
#df.loc[(df['Disable Test'].str.capitalize() =='No') & df.TestName.str.contains('IIP3'), 'Skip?'] = 'No'
df.loc[(df['Disable Test'].str.capitalize() =='No') & df.TestName.str.contains('IIP3'), 'Constraint_LSL'] = df['Bin1_LSL']+2
#df.loc[(df['Disable Test'].str.capitalize() =='No') & df.TestName.str.contains('IIP3')] 

#YES - NF
#df.loc[(df['Disable Test'].str.capitalize() =='No') & df.TestName.str.contains('NF'), 'Skip?'] = 'No'
df.loc[(df['Disable Test'].str.capitalize() =='No') & df.TestName.str.contains('NF'), 'Constraint_LSL'] = df['Bin1_LSL']
df.loc[(df['Disable Test'].str.capitalize() =='No') & df.TestName.str.contains('NF'), 'Constraint_USL'] = df['Bin1_USL']-0.2
#df.loc[(df['Disable Test'].str.capitalize() =='No') & df.TestName.str.contains('NF')]

# YES, Leakage
df.loc[((df['Disable Test'].str.capitalize() =='No')) & ((df.SoftwareBin.str.contains('LEAK')) | (df.SoftwareBin.str.contains('Leakage')) | (df.SoftwareBin.str.contains('LKG')) | (df.TestName.str.contains('LKG'))),'Constraint_LSL']=df['Bin1_LSL']
df.loc[((df['Disable Test'].str.capitalize() =='No')) & ((df.SoftwareBin.str.contains('LEAK')) | (df.SoftwareBin.str.contains('Leakage')) | (df.SoftwareBin.str.contains('LKG')) | (df.TestName.str.contains('LKG'))),'Skip?']='No'

#Drop NaN rows
df.dropna(axis="index", inplace = True)

#for col, row in df.iterrows():
    #print(row['TestName'])
    #dataframe.loc[dataframe['Test Name']==row['TestName'],'PAT Limit Constraints']=row['Constraint_LSL']
    #dataframe.loc[dataframe['Test Name']==row['TestName'],dataframe.columns.get_loc('PAT Limit Constraints')+1]=row['Constraint_USL']
    #dataframe.loc[dataframe['Test Name']==row['TestName'],dataframe.columns.get_loc('PAT Limit Constraints')-1]=row['Skip?']

print("Report Done!\nSaved as "+ os.path.expanduser("~/Desktop/FT-DPAT_ConstraintsReport.xlsx"))
df.to_excel(os.path.expanduser("~/Desktop/FT-DPAT_ConstraintsReport.xlsx"),index=False)

