import pandas as pd


'''
DONE - input_brokers_df - collect data from brokers
DONE - input_core_df - collect data from core
DONE - output_upload_df - output data which will be ready for uploading
DONE - output_brokers_filter - output data which will filter rows from data given by brokers for calculation purpose
DONE - output_brokers_unavailable - output data which are not available in core but available in brokers' data
DONE - output_brokers_duplicate - output data which are duplicate entries in brokers data
'''

# read brokers' data
input_brokers_df = pd.read_excel('Book1.xlsx', engine='openpyxl')

# read core data
input_core_df = pd.read_excel('EXCELFORMAT.xls','GPA_UPLOAD')

# Converting Relation to Title case and removing any spaces in the values
input_brokers_df['Relation'] = input_brokers_df['Relation'].str.title()
input_brokers_df['Relation'] = input_brokers_df['Relation'].str.replace(' ','')

# create dataframe of entries where the relation is self
df_self_emp = input_brokers_df[input_brokers_df['Relation'] == 'Self']

# create dataframe of self employee IDs which repeat in brokers' data
df_duplicate_emp = df_self_emp[df_self_emp[['Employee ID']].duplicated()]

list_unique_emp_id = df_duplicate_emp['Employee ID'].unique()

list_repeat_data = []

# creating dataframe of entries which have repeating employee IDs in brokers' data
for i in list_unique_emp_id:
    df_repeat = input_brokers_df[input_brokers_df['Employee ID'] == i]
    list_repeat_data.append(df_repeat)

# outputting dataframe of employees ID which have duplicate self entries
df_repeat_data = pd.concat(list_repeat_data)
df_repeat_data.to_excel("output-brokers data - repeating employees ID.xlsx",index=False)

# dropping repeating dataframe from our input_brokers_df dataframe
df_not_repeat_data = input_brokers_df.drop(df_repeat_data.index)

list_not_repeat_emp_id = df_not_repeat_data['Employee ID'].unique()

list_final_output = []

for i in list_not_repeat_emp_id:
    df_core_data = input_core_df[input_core_df['EmploymentCode'] ==i]
    list_final_output.append(df_core_data)

# outputting core format where employees ID are available in Core as well as brokers' data
df_final_output = pd.concat(list_final_output)

# setting flagstatus to D
df_final_output["FLAGSTATUS"] = "D"

#writer_1 = pd.ExcelWriter("output.xls",engine='xlsxwriter')
#df_final_output.to_excel(writer_1,sheet_name="GPA_UPLOAD",index=False)

df_final_output.to_excel("output-core_format - deletion endorsement.xls",engine='xlwt',sheet_name="GPA_UPLOAD",index=False)

list_final_output_emp_available = []

list_emp_id_available = df_final_output['EmploymentCode'].unique()

for i in list_emp_id_available:
    df_brokers_data_available = input_brokers_df[input_brokers_df['Employee ID'] ==i]
    list_final_output_emp_available.append(df_brokers_data_available)

# outputting brokers data which can be used for deletion endorsement calculation
df_brokers_calc_data = pd.concat(list_final_output_emp_available)
df_brokers_calc_data.to_excel("output-calculation-brokers-data.xlsx",index=False)

# filtering employees not available in core from entire brokers' data
df_brokers_data_unavailable = df_not_repeat_data.drop(df_brokers_calc_data.index)
df_brokers_data_unavailable.to_excel("output-employees-not-available in core but in brokers data.xlsx",index=False)
