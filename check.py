import pandas as pd

df_pdq213 = pd.read_csv('pdq213_PG_SUPPLY_CHAIN_with_Component_Info.csv', converters={'CONSUMED_LOCATION': str},low_memory=False)
df_pdq213.head()

df_pdq505 = pd.read_csv('pdq505_PG_ITEM_LOCATION_Attributes.csv', converters={'LOCATION_NAME':str},low_memory=False)
#df_pdq505['ITEM_NAME'] = pd.to_numeric(df_pdq505['ITEM_NAME'],errors='coerce')
#df_pdq505['ITEM_NAME'] = df_pdq505['ITEM_NAME'].str.strip()

df_pdq505.head()

df_com=pd.merge(df_pdq213,df_pdq505[['ITEM_NAME', 'LOCATION_NAME','RELEASE_FENCE','RELEASE_FENCE_UOM']], left_on=['CONSUMED_ITEM', 'CONSUMED_LOCATION'],right_on=['ITEM_NAME', 'LOCATION_NAME'], how='left')
df_com.to_csv('pdq213_505.csv')
#print(df_com.head())