import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns


#Completion_rate = Comr# Loading the dataset for Primary
Comr_df = pd.read_excel(r'C:\Users\USER\Downloads\Completion_rate_2022_formatted(2).xlsx', sheet_name = 'Primary', header = [0])

Comr_df

Comr_df.head(10)

Comr_df = Comr_df.drop(columns = [('Data source'), ('Population data'), ('Unnamed: 18'), ('Unnamed: 19'), ('Unnamed: 20'), ('Unnamed: 21'), ('Unnamed: 22')])

Comr_df.info()

# First, listing all the numeric columns I want to convert
numeric_cols = ['Total', 'Female', 'Male', 'Rural', 'Urban', 
                'Poorest', 'Second', 'Middle', 'Fourth', 'Richest', 'Time period']

# Then converting them to numeric, forcing errors to NaN
Comr_df[numeric_cols] = Comr_df[numeric_cols].apply(pd.to_numeric, errors='coerce')

# Optional: rounding and converting to Int64 (nullable integer type)
Comr_df[numeric_cols] = Comr_df[numeric_cols].round(0).astype('Int64')


Comr_df

missing_percent = Comr_df.isnull().mean() * 100

# For columns with <5% missing, dropping the rows where those are missing
columns_to_check = missing_percent[missing_percent < 5].index
Comr_df.dropna(subset=columns_to_check, inplace=True)

Comr_df.info()

Comr_df

cols = Comr_df.columns[Comr_df.columns.get_loc('Male'):Comr_df.columns.get_loc('Time period')+1]
Comr_df[cols].isnull().all(axis=1).sum()

Comr_df = Comr_df[~Comr_df[cols].isnull().all(axis=1)]

Comr_df

#Getting the percentage of missing values in each columns
missing_percent = Comr_df.isnull().mean() * 100
print(missing_percent.sort_values(ascending=False))

# Getting the index of the row where Country is Zimbabwe
zimbabwe_index = Comr_df[Comr_df['Country'] == 'Zimbabwe'].index[0]

# Keeping all rows up to and including Zimbabwe (i.e., drop rows after it)
Comr_df = Comr_df.loc[:zimbabwe_index]

Comr_df

missing_percent = Comr_df.isnull().mean() * 100
print(missing_percent.sort_values(ascending=False))

cols_to_fill = ['Poorest', 'Second', 'Middle', 'Fourth', 'Richest']
Comr_df.loc[:, cols_to_fill] = Comr_df[cols_to_fill].fillna(Comr_df[cols_to_fill].median())

Comr_df

missing_percent = Comr_df.isnull().mean() * 100
print(missing_percent.sort_values(ascending=False))

Comr_df

#Completion_rate = Comr#  Loading the dataset for Lower secondary
Comr_df_1 = pd.read_excel(r'C:\Users\USER\Downloads\Completion_rate_2022_formatted(2).xlsx', sheet_name = 'Lower secondary', header = [0])

Comr_df_1

Comr_df_1 = Comr_df_1.drop(columns = [('Data source'), ('Population data'), ('Unnamed: 18'), ('Unnamed: 19'), ('Unnamed: 20'), ('Unnamed: 21'), ('Unnamed: 22')])

Comr_df_1

Comr_df_1.info()

# First, list all numeric columns you want to convert
numeric_cols_1 = ['Total', 'Female', 'Male', 'Rural', 'Urban', 
                'Poorest', 'Second', 'Middle', 'Fourth', 'Richest', 'Time period']

# Then convert them to numeric, forcing errors to NaN
Comr_df_1[numeric_cols_1] = Comr_df_1[numeric_cols_1].apply(pd.to_numeric, errors='coerce')

# Optional: round and convert to Int64 (nullable integer type)
Comr_df_1[numeric_cols_1] = Comr_df_1[numeric_cols_1].round(0).astype('Int64')

Comr_df_1

missing_percent_1 = Comr_df_1.isnull().mean() * 100

# For columns with <5% missing, dropping the rows where those are missing
columns_to_check_1 = missing_percent_1[missing_percent_1 < 5].index
Comr_df_1.dropna(subset=columns_to_check_1, inplace=True)

Comr_df_1

cols_1 = Comr_df_1.columns[Comr_df_1.columns.get_loc('Male'):Comr_df_1.columns.get_loc('Time period')+1]
Comr_df_1[cols_1].isnull().all(axis=1).sum()

Comr_df_1 = Comr_df_1[~Comr_df_1[cols_1].isnull().all(axis=1)]

Comr_df_1

missing_percent_1 = Comr_df_1.isnull().mean() * 100
print(missing_percent_1.sort_values(ascending=False))

# Get the index of the row where Country is Zimbabwe
zimbabwe_index_1 = Comr_df_1[Comr_df_1['Country'] == 'Zimbabwe'].index[0]

# Keeping all rows up to and including Zimbabwe (i.e., drop rows after it)
Comr_df_1 = Comr_df_1.loc[:zimbabwe_index_1]

Comr_df_1

missing_percent_1 = Comr_df_1.isnull().mean() * 100
print(missing_percent_1.sort_values(ascending=False))

cols_to_fill_1 = ['Poorest', 'Second', 'Middle', 'Fourth', 'Richest']
Comr_df_1.loc[:, cols_to_fill_1] = Comr_df_1[cols_to_fill_1].fillna(Comr_df_1[cols_to_fill_1].median().round(0).astype('Int64'))

Comr_df_1.info()

Comr_df_1

missing_percent_1 = Comr_df_1.isnull().mean() * 100
print(missing_percent_1.sort_values(ascending=False))

Comr_df_1

#Completion_rate = Comr# Loading the dataset for Upper secondary
Comr_df_2 = pd.read_excel(r'C:\Users\USER\Downloads\Completion_rate_2022_formatted(2).xlsx', sheet_name = 'Upper secondary', header = [0])

Comr_df_2

Comr_df_2 = Comr_df_2.drop(columns = [('Data source'), ('Population data'), ('Unnamed: 18'), ('Unnamed: 19'), ('Unnamed: 20'), ('Unnamed: 21'), ('Unnamed: 22')])

Comr_df_2

Comr_df_2.info()

# First, list all numeric columns you want to convert
numeric_cols_2 = ['Total', 'Female', 'Male', 'Rural', 'Urban', 
                'Poorest', 'Second', 'Middle', 'Fourth', 'Richest', 'Time period']

# Then convert them to numeric, forcing errors to NaN
Comr_df_2[numeric_cols_2] = Comr_df_2[numeric_cols_2].apply(pd.to_numeric, errors='coerce')

# Optional: round and convert to Int64 (nullable integer type)
Comr_df_2[numeric_cols_2] = Comr_df_2[numeric_cols_2].round(0).astype('Int64')

Comr_df_2

cols_2 = Comr_df_2.columns[Comr_df_2.columns.get_loc('Male'):Comr_df_2.columns.get_loc('Time period')+1]
Comr_df_2[cols_2].isnull().all(axis=1).sum()

missing_percent_2 = Comr_df_2.isnull().mean() * 100

# For columns with <5% missing, dropping the rows where those are missing
columns_to_check_2 = missing_percent_2[missing_percent_2 < 5].index
Comr_df_2.dropna(subset=columns_to_check_2, inplace=True)

Comr_df_2 = Comr_df_2[~Comr_df_2[cols_2].isnull().all(axis=1)]

Comr_df_2

# Get the index of the row where Country is Zimbabwe
zimbabwe_index_2 = Comr_df_2[Comr_df_2['Country'] == 'Zimbabwe'].index[0]

# Keeping all rows up to and including Zimbabwe (i.e., drop rows after it)
Comr_df_2 = Comr_df_2.loc[:zimbabwe_index_2]

Comr_df_2

missing_percent_2 = Comr_df_2.isnull().mean() * 100
print(missing_percent_2.sort_values(ascending=False))

cols_to_fill_2 = ['Poorest', 'Second', 'Middle', 'Fourth', 'Richest']
Comr_df_2.loc[:, cols_to_fill_2] = Comr_df_2[cols_to_fill_2].fillna(Comr_df_2[cols_to_fill_2].median().round(0).astype('Int64'))

Comr_df_2

missing_percent_2 = Comr_df_2.isnull().mean() * 100
print(missing_percent_2.sort_values(ascending=False))

Comr_df_2

with pd.ExcelWriter(r'C:\Users\USER\Downloads\Completion_rate_2022_cleaned.xlsx') as writer:
    Comr_df.to_excel(writer, sheet_name='Primary', index=False)
    Comr_df_1.to_excel(writer, sheet_name='Lower secondary', index=False)
    Comr_df_2.to_excel(writer, sheet_name='Upper secondary', index=False)


