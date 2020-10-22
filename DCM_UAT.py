import pandas as pd
import numpy as np
import glob
import os
import os, sys
from pandas import ExcelWriter
from pandas import ExcelFile


path_new = r'C:\Users\nxf46768\Desktop\DCM_UAT\DCM1.1' # use your path
path_old = r'C:\Users\nxf46768\Desktop\DCM_UAT\DCM1.0' # use your path
path_MM = r'C:\Users\nxf46768\Desktop\DCM_UAT' # use your path

new_files = glob.glob(os.path.join(path_new + "/in_fsl_codp*.csv"))
old_files = glob.glob(os.path.join(path_old + "/in_fsl_codp*.csv"))
df_MM = pd.read_excel('in_fsl_pep_xref_Item_class.xlsx', sheet_name='Sheet1', converters={'ITEM':str , 'NXP_ITEM' : str,'ITEMCLASS' : str})

df_new_files = (pd.read_csv(f,index_col=None,low_memory=False) for f in new_files)
concatenated_dfnew = pd.concat(df_new_files, ignore_index=True, sort=False).reset_index()
#concatenated_dfnew.drop(['RECORD_NUMBER'], axis=1, inplace=True)
concatenated_dfnew.drop(['index'], axis=1, inplace=True)
concatenated_dfnew.drop(concatenated_dfnew.columns[concatenated_dfnew.columns.str.contains('unnamed', case=False)],axis=1, inplace=True)
concatenated_dfnew = pd.merge(concatenated_dfnew,df_MM, left_on=['ITEM'],right_on=['ITEM'], how='left')
#concatenated_dfnew.to_csv('504_new_output.csv', header= True,index=False)
#concatenated_dfnew.drop(['ITEM_BOM_RT_ID'], axis=1, inplace=True)
df_old_files = (pd.read_csv(f,index_col=None,low_memory=False) for f in old_files)
concatenated_dfold = pd.concat(df_old_files, ignore_index=True, sort=False).reset_index()
#concatenated_dfold.drop(['RECORD_NUMBER'], axis=1, inplace=True)
concatenated_dfold.drop(['index'], axis=1, inplace=True)
concatenated_dfold.drop(concatenated_dfold.columns[concatenated_dfold.columns.str.contains('unnamed', case=False)],axis=1, inplace=True)
concatenated_dfold = pd.merge(concatenated_dfold,df_MM, left_on=['ITEM'],right_on=['ITEM'], how='left')
#concatenated_dfold.to_csv('504_old_output.csv', header= True,index=False)
#concatenated_dfold.drop(['ITEM_BOM_RT_ID'], axis=1, inplace=True)
df_compare = pd.concat([concatenated_dfold, concatenated_dfnew],sort=False,keys=['DCM1.0', 'DCM1.1'],names=['DCM_Version']).reset_index(level=1, drop=True)
#df_compare = df_compare.reset_index(drop=True)
df_diff = df_compare.drop_duplicates(keep=False)
print(df_diff.head())
df_diff.to_csv(r'C:\Users\nxf46768\Desktop\DCM_UAT\Report_Output\Case43_CODP_0827.csv')