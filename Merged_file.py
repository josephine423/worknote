import pandas as pd

df_MARC = pd.read_csv('MARC_0529.csv', converters={'Plnt':str}, low_memory=False)
df_MARC.Material = df_MARC.Material.astype(str)

df_MARC.head()

df_pdq505 = pd.read_csv('pdq505_PG_ITEM_LOCATION_Attributes.csv', converters={'LOCATION_NAME':str},low_memory=False)
df_pdq505.ITEM_NAME = df_pdq505.ITEM_NAME.astype(str)
df_pdq505.head()


#print (df_pdq505.dtypes)
#print(df_pdq505.head())

df_com=pd.merge(df_MARC,df_pdq505, left_on=['Material'],right_on=['ITEM_NAME'], how='left')

df_com.to_csv('output3.csv')

#print(df_com.head())
