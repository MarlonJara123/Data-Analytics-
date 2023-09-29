import pandas as pd
import pip
import numpy as np
from datetime import datetime

# first time running the code is importat to run the following code to install the PIP import pip /  pip.main(["install"
# ,
# "openpyxl"])

#Open Up both root files from the source
df1 = pd.read_excel(
    r'C:\Users\marlon.jaramillo\OneDrive - Thermo Fisher Scientific\Desktop\Alteryx\Ricardo\.HU PO created.xlsx')
df2 = pd.read_excel(
    r'C:\Users\marlon.jaramillo\OneDrive - Thermo Fisher Scientific\Desktop\Alteryx\Ricardo\Chart of accounts list.xlsx',
    sheet_name="Sheet2")

# Columns between both DF had a different name due to caps, line below to correct.
df2 = df2.rename(columns={"Chart of accounts": "Chart Of Accounts"})

# Display All columns
pd.set_option("display.max_columns", None)
pd.set_option('display.expand_frame_repr', False)
##
# Dropped the duplicated US and CA purchases
df1 = df1[df1["Chart Of Accounts"].str.contains("US") == False]
df1 = df1[df1["Chart Of Accounts"].str.contains("CA") == False]

# Combine both DFs into one, left join
df3 = pd.merge(df1, df2[["Wave", "Numeric code", "Chart Of Accounts"]], on=["Chart Of Accounts"], how="left")

# Remove duplicates from the DF
df3.drop_duplicates(subset="PO Number (Header)", keep="first", inplace=True)

# Extract the Country initials
df3["Country"] = df3["Chart Of Accounts"].str.split("_").str[0]

# List of Country POs that need to be different
excep_country_POs = ["DD_UK_GB", "DD_1030_DE", "UK60_GB", "DD_UK60_GB", "DD_UK50_GB", "DD_UK01_GB", "DD_NL60_NL"]

# for loop  and Loc to locate strings in Supplier that contain the 7 codes we need to change in Country
for x in excep_country_POs:
    df3.loc[df3["Supplier"].str.contains(x), "Country"] = x

# Create a new column where if the value is Null then 0 else 1
df3["Catalog POs"] = np.where(df3["Contract"].isnull(), 0, 1)

# create a column to place the run date
df3["Run date"] = pd.Timestamp.today().strftime("%m/%d/%Y")

# remove unnecesary columns
df3.drop(columns=["Numeric code", "Chart Of Accounts", "Commodity", "Account"], inplace=True)

# print to excel
df3.to_excel("output3.xlsx", index=False)

print(df3)
