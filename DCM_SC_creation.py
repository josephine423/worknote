import pandas as pd
import numpy as np
import datetime
import time
import glob
import os
import os, sys
import re
from pandas import ExcelWriter
from pandas import ExcelFile
import xlsxwriter




path_old = r'C:\Users\nxf46768\Desktop\DCM_UAT\DCM1.0' # use your path
path_new = r'C:\Users\nxf46768\Desktop\DCM_UAT\DCM1.1' # use your path

path_MM = r'C:\Users\nxf46768\Desktop\DCM_UAT' # use your path
df_MM = pd.read_excel('in_fsl_pep_xref_Item_class.xlsx', sheet_name='Sheet1', converters={'ITEM':str , 'NXP_ITEM' : str,'ITEMCLASS' : str})
df_noSAPNXP_MM = pd.read_excel('Item_ItemSites_not_in SAPNXP- WK34.xlsx', sheet_name='ItemMaster', converters={'ITEM':str , 'ITEMCLASS' : str})
df_noSAPNXP_LMM = pd.read_excel('Item_ItemSites_not_in SAPNXP- WK34.xlsx', sheet_name='ItemSiteMaster', converters={'ITEM':str , 'ITEMDESC' : str,'SITEID' : str})

DCM0_SC = glob.glob(os.path.join(path_old + "/in_fsl_supplychain*.csv"))
DCM1_SC = glob.glob(os.path.join(path_new + "/in_fsl_supplychain*.csv"))
df_SC_DCM0 = (pd.read_csv(f,index_col=None,low_memory=False) for f in DCM0_SC)
df_SC_DCM0 = pd.concat(df_SC_DCM0, ignore_index=True, sort=False).reset_index()
df_SC_DCM0.drop(['index'], axis=1, inplace=True)
df_SC_DCM0.drop(df_SC_DCM0.columns[df_SC_DCM0.columns.str.contains('unnamed', case=False)],axis=1, inplace=True)

DCM0_BOM = glob.glob(os.path.join(path_old + "/in_fsl_bom*.csv"))
df_BOM_DCM0 = (pd.read_csv(f,index_col=None,low_memory=False) for f in DCM0_BOM)
df_BOM_DCM0 = pd.concat(df_BOM_DCM0, ignore_index=True, sort=False).reset_index()
df_BOM_DCM0.drop(['index'], axis=1, inplace=True)
df_BOM_DCM0.drop(df_BOM_DCM0.columns[df_BOM_DCM0.columns.str.contains('unnamed', case=False)],axis=1, inplace=True)

df_SC_Base = pd.merge(df_SC_DCM0,df_BOM_DCM0[['BOM_NAME','COMPONENT_ITEM','COMPONENT_QUANTITY']], left_on=['BOM_NAME'],right_on=['BOM_NAME'], how='left')
df_SC_Base = pd.merge(df_SC_Base,df_MM, left_on=['ITEM'],right_on=['ITEM'], how='left')
df_SC_Base.rename(columns={'NXP_ITEM': 'Produced_Item','ITEMCLASS' : 'Prod_Category'}, inplace=True)

df_SC_Base = pd.merge(df_SC_Base,df_MM, left_on=['COMPONENT_ITEM'],right_on=['ITEM'], how='left')
df_SC_Base.drop(['ITEM_y'], axis=1, inplace=True)
df_SC_Base.rename(columns={'NXP_ITEM': 'Consumed_Item','ITEMCLASS' : 'Cons_Category','ITEM_x':'FSL_PROD_ITEM','COMPONENT_ITEM':'FSL_CONS_ITEM'}, inplace=True)

df_SC_Base['Make or Move'] = np.where(df_SC_Base['Produced_Item'] == df_SC_Base['Consumed_Item'], 'Move', 'Make')
DCM0_RT = glob.glob(os.path.join(path_old + "/in_fsl_routing*.csv"))
df_RT_DCM0 = (pd.read_csv(f,index_col=None,low_memory=False) for f in DCM0_RT)
df_RT_DCM0 = pd.concat(df_RT_DCM0, ignore_index=True, sort=False).reset_index()
df_RT_DCM0.drop(['index'], axis=1, inplace=True)
df_RT_DCM0.drop(df_RT_DCM0.columns[df_RT_DCM0.columns.str.contains('unnamed', case=False)],axis=1, inplace=True)

df_SC_Base = pd.merge(df_SC_Base,df_RT_DCM0[['ROUTING_NAME','OPERATIONS']], left_on=['ROUTING_NAME'],right_on=['ROUTING_NAME'], how='left')
new_op = df_SC_Base["OPERATIONS"].str.split('; |, |\~|\~', expand = True)

df_SC_Base["OPERATION_NAME_1"]= new_op[1]
df_SC_Base["OPERATION_NAME_2"]= new_op[3]  
df_SC_Base["OPERATION_NAME_3"]= new_op[5]  

DCM0_OP = glob.glob(os.path.join(path_old + "/in_fsl_operation*.csv"))
df_OP_DCM0 = (pd.read_csv(f,index_col=None,low_memory=False) for f in DCM0_OP)
df_OP_DCM0 = pd.concat(df_OP_DCM0, ignore_index=True, sort=False).reset_index()
df_OP_DCM0.drop(['index'], axis=1, inplace=True)
df_OP_DCM0.drop(df_OP_DCM0.columns[df_OP_DCM0.columns.str.contains('unnamed', case=False)],axis=1, inplace=True)

df_SC_Base = pd.merge(df_SC_Base,df_OP_DCM0[['OPERATION_NAME','BOR_NAME']], left_on=['OPERATION_NAME_1'],right_on=['OPERATION_NAME'], how='left')
df_SC_Base.rename(columns={'BOR_NAME': 'BOR_NAME_1'}, inplace=True)
df_SC_Base.drop(['OPERATION_NAME'], axis=1, inplace=True)
df_SC_Base = pd.merge(df_SC_Base,df_OP_DCM0[['OPERATION_NAME','BOR_NAME']], left_on=['OPERATION_NAME_2'],right_on=['OPERATION_NAME'], how='left')
df_SC_Base.rename(columns={'BOR_NAME': 'BOR_NAME_2'}, inplace=True)
df_SC_Base.drop(['OPERATION_NAME'], axis=1, inplace=True)
df_SC_Base = pd.merge(df_SC_Base,df_OP_DCM0[['OPERATION_NAME','BOR_NAME']], left_on=['OPERATION_NAME_3'],right_on=['OPERATION_NAME'], how='left')
df_SC_Base.rename(columns={'BOR_NAME': 'BOR_NAME_3'}, inplace=True)
df_SC_Base.drop(['OPERATION_NAME'], axis=1, inplace=True)

# DCM0_SCOPOVR = glob.glob(os.path.join(path_old + "/in_fsl_scopovr*.csv"))
# df_SCOPOVR_DCM0 = (pd.read_csv(f,index_col=None,low_memory=False) for f in DCM0_SCOPOVR)
# df_SCOPOVR_DCM0 = pd.concat(df_SCOPOVR_DCM0, ignore_index=True, sort=False).reset_index()
# df_SCOPOVR_DCM0.drop(['index'], axis=1, inplace=True)
# df_SCOPOVR_DCM0.drop(df_SCOPOVR_DCM0.columns[df_SCOPOVR_DCM0.columns.str.contains('unnamed', case=False)],axis=1, inplace=True)
# df_SCOPOVR_DCM0 = df_SCOPOVR_DCM0[['ROUTING_NAME','OPERATION_NAME']]
# df_SCOPOVR_DCM0 = df_SCOPOVR_DCM0.drop_duplicates(keep="first")
# df_SC_Base = pd.merge(df_SC_Base,df_SCOPOVR_DCM0[['ROUTING_NAME','OPERATION_NAME']], left_on=['ROUTING_NAME'],right_on=['ROUTING_NAME'], how='left')
#df_SC_Base.drop(['BOM_NAME_y'], axis=1, inplace=True)
#print(df_SC_Base.head(100))
#df_SC_Base.to_csv(r'C:\Users\nxf46768\Desktop\DCM_UAT\Report_Output\SC_Output.csv',index=False)
#df_SC_DCM0 = pd.merge(df_SC_DCM0,df_MM, left_on=['ITEM'],right_on=['ITEM'], how='left')

df_SC_DCM = (pd.read_csv(f,index_col=None,low_memory=False) for f in DCM1_SC)
df_SC_DCM = pd.concat(df_SC_DCM, ignore_index=True, sort=False).reset_index()
df_SC_DCM.drop(['index'], axis=1, inplace=True)
df_SC_DCM.drop(df_SC_DCM.columns[df_SC_DCM.columns.str.contains('unnamed', case=False)],axis=1, inplace=True)

DCM1_BOM = glob.glob(os.path.join(path_new + "/in_fsl_bom*.csv"))
df_BOM_DCM1 = (pd.read_csv(f,index_col=None,low_memory=False) for f in DCM1_BOM)
df_BOM_DCM1 = pd.concat(df_BOM_DCM1, ignore_index=True, sort=False).reset_index()
df_BOM_DCM1.drop(['index'], axis=1, inplace=True)
df_BOM_DCM1.drop(df_BOM_DCM1.columns[df_BOM_DCM1.columns.str.contains('unnamed', case=False)],axis=1, inplace=True)

df_SC_DCM = pd.merge(df_SC_DCM,df_BOM_DCM1[['BOM_NAME','COMPONENT_ITEM','COMPONENT_QUANTITY']], left_on=['BOM_NAME'],right_on=['BOM_NAME'], how='left')
df_SC_DCM = pd.merge(df_SC_DCM,df_MM, left_on=['ITEM'],right_on=['ITEM'], how='left')
df_SC_DCM.rename(columns={'NXP_ITEM': 'Produced_Item','ITEMCLASS' : 'Prod_Category'}, inplace=True)

df_SC_DCM = pd.merge(df_SC_DCM,df_MM, left_on=['COMPONENT_ITEM'],right_on=['ITEM'], how='left')
df_SC_DCM.drop(['ITEM_y'], axis=1, inplace=True)
df_SC_DCM.rename(columns={'NXP_ITEM': 'Consumed_Item','ITEMCLASS' : 'Cons_Category','ITEM_x':'FSL_PROD_ITEM','COMPONENT_ITEM':'FSL_CONS_ITEM'}, inplace=True)

df_SC_DCM['Make or Move'] = np.where(df_SC_DCM['Produced_Item'] == df_SC_DCM['Consumed_Item'], 'Move', 'Make')
DCM1_RT = glob.glob(os.path.join(path_new + "/in_fsl_routing*.csv"))
df_RT_DCM1 = (pd.read_csv(f,index_col=None,low_memory=False) for f in DCM1_RT)
df_RT_DCM1 = pd.concat(df_RT_DCM1, ignore_index=True, sort=False).reset_index()
df_RT_DCM1.drop(['index'], axis=1, inplace=True)
df_RT_DCM1.drop(df_RT_DCM1.columns[df_RT_DCM1.columns.str.contains('unnamed', case=False)],axis=1, inplace=True)

df_SC_DCM = pd.merge(df_SC_DCM,df_RT_DCM1[['ROUTING_NAME','OPERATIONS']], left_on=['ROUTING_NAME'],right_on=['ROUTING_NAME'], how='left')
new_op = df_SC_DCM["OPERATIONS"].str.split('; |, |\~|\~', expand = True)

df_SC_DCM["OPERATION_NAME_1"]= new_op[1]
df_SC_DCM["OPERATION_NAME_2"]= new_op[3]  
df_SC_DCM["OPERATION_NAME_3"]= new_op[5]  

DCM1_OP = glob.glob(os.path.join(path_new + "/in_fsl_operation*.csv"))
df_OP_DCM1 = (pd.read_csv(f,index_col=None,low_memory=False) for f in DCM1_OP)
df_OP_DCM1 = pd.concat(df_OP_DCM1, ignore_index=True, sort=False).reset_index()
df_OP_DCM1.drop(['index'], axis=1, inplace=True)
df_OP_DCM1.drop(df_OP_DCM1.columns[df_OP_DCM1.columns.str.contains('unnamed', case=False)],axis=1, inplace=True)

df_SC_DCM = pd.merge(df_SC_DCM,df_OP_DCM1[['OPERATION_NAME','BOR_NAME']], left_on=['OPERATION_NAME_1'],right_on=['OPERATION_NAME'], how='left')
df_SC_DCM.rename(columns={'BOR_NAME': 'BOR_NAME_1'}, inplace=True)
df_SC_DCM.drop(['OPERATION_NAME'], axis=1, inplace=True)
df_SC_DCM = pd.merge(df_SC_DCM,df_OP_DCM1[['OPERATION_NAME','BOR_NAME']], left_on=['OPERATION_NAME_2'],right_on=['OPERATION_NAME'], how='left')
df_SC_DCM.rename(columns={'BOR_NAME': 'BOR_NAME_2'}, inplace=True)
df_SC_DCM.drop(['OPERATION_NAME'], axis=1, inplace=True)
df_SC_DCM = pd.merge(df_SC_DCM,df_OP_DCM1[['OPERATION_NAME','BOR_NAME']], left_on=['OPERATION_NAME_3'],right_on=['OPERATION_NAME'], how='left')
df_SC_DCM.rename(columns={'BOR_NAME': 'BOR_NAME_3'}, inplace=True)
df_SC_DCM.drop(['OPERATION_NAME'], axis=1, inplace=True)


## to map the list that is already identified by IT
df_SC_Base = pd.merge(df_SC_Base,df_noSAPNXP_MM[['ITEM','ITEMCLASS']], left_on=['Consumed_Item'],right_on=['ITEM'], how='left')
df_SC_Base.rename(columns={'ITEM': 'Cons_ITEM_SAPNXP','ITEMCLASS' : 'Cons_SAPNXP'}, inplace=True)
df_SC_Base = pd.merge(df_SC_Base,df_noSAPNXP_MM[['ITEM','ITEMCLASS']], left_on=['Produced_Item'],right_on=['ITEM'], how='left')
df_SC_Base.rename(columns={'ITEM': 'Prod_ITEM_SAPNXP','ITEMCLASS' : 'Prod_SAPNXP'}, inplace=True)


df_SC_DCM = pd.merge(df_SC_DCM,df_noSAPNXP_MM[['ITEM','ITEMCLASS']], left_on=['Consumed_Item'],right_on=['ITEM'], how='left')
df_SC_DCM.rename(columns={'ITEM': 'Cons_ITEM_SAPNXP','ITEMCLASS' : 'Cons_SAPNXP'}, inplace=True)
df_SC_DCM = pd.merge(df_SC_DCM,df_noSAPNXP_MM[['ITEM','ITEMCLASS']], left_on=['Produced_Item'],right_on=['ITEM'], how='left')
df_SC_DCM.rename(columns={'ITEM': 'Prod_ITEM_SAPNXP','ITEMCLASS' : 'Prod_SAPNXP'}, inplace=True)

df_compare = pd.concat([df_SC_Base, df_SC_DCM],sort=False,keys=['DCM1.0', 'DCM1.1'],names=['DCM_Version']).reset_index(level=1, drop=True)
df_compare.drop(['BOM_NAME','SC_NAME'], axis=1, inplace=True)
#df_compare = df_compare.reset_index(drop=True)
df_diff = df_compare.drop_duplicates(keep=False)


df_SC_Base = pd.merge(df_SC_Base,df_noSAPNXP_LMM[['ITEM','ITEMDESC','SITEID']], left_on=['Consumed_Item'],right_on=['ITEM'], how='left')
df_SC_Base.rename(columns={'ITEM': 'Cons_ITEMLOC_SAPNXP','ITEMDESC' : 'Cons_ITEMLOC_SAPNXP_DES','SITEID' : 'ConsLoc_ITEMLOC_SAPNXP'}, inplace=True)
df_SC_Base = pd.merge(df_SC_Base,df_noSAPNXP_LMM[['ITEM','ITEMDESC','SITEID']], left_on=['Produced_Item'],right_on=['ITEM'], how='left')
df_SC_Base.rename(columns={'ITEM': 'Prod_ITEMLOC_SAPNXP','ITEMDESC' : 'Prod_ITEMLOC_SAPNXP_DES','SITEID' : 'ProdLoc_ITEMLOC_SAPNXP'}, inplace=True)

df_SC_DCM = pd.merge(df_SC_DCM,df_noSAPNXP_LMM[['ITEM','ITEMDESC','SITEID']], left_on=['Consumed_Item'],right_on=['ITEM'], how='left')
df_SC_DCM.rename(columns={'ITEM': 'Cons_ITEMLOC_SAPNXP','ITEMDESC' : 'Cons_ITEMLOC_SAPNXP_DES','SITEID' : 'ConsLoc_ITEMLOC_SAPNXP'}, inplace=True)
df_SC_DCM = pd.merge(df_SC_DCM,df_noSAPNXP_LMM[['ITEM','ITEMDESC','SITEID']], left_on=['Produced_Item'],right_on=['ITEM'], how='left')
df_SC_DCM.rename(columns={'ITEM': 'Prod_ITEMLOC_SAPNXP','ITEMDESC' : 'Prod_ITEMLOC_SAPNXP_DES','SITEID' : 'ProdLoc_ITEMLOC_SAPNXP'}, inplace=True)


start = time.time()
TodaysDate = time.strftime("%Y-%m-%d")
excelfilename = "Case68_69_73_supplychain_AllSAP_"+TodaysDate+ ".xlsx"
writer = pd.ExcelWriter(excelfilename, engine='xlsxwriter')
#store your dataframes in a dict, where the key is the sheet name you want
frames = {'Diff_DCM1.0_DCM1.1': df_diff, 'DCM1.0': df_SC_Base,'DCM1.1': df_SC_DCM}
#now loop thru and put each on a specific sheet
for sheet, frame in frames.items(): # .use .items for python 3.X
    frame.to_excel(writer, sheet_name = sheet)
#critical last step
writer.save()
end = time.time()
print((end - start)/60)

#df_diff.to_csv(r'C:\Users\nxf46768\Desktop\DCM_UAT\Report_Output\Case68_69_73_supplychain_0819.csv')