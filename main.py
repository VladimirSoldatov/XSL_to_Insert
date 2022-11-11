import numpy as np
import pandas as pd
import openpyxl

# Load the xlsx file
path_dict = {"path1": r"C:\temp\ДЗО Альянс 720 дней 11.11.2022 .xlsx"}
excel_data = pd.read_excel('{}'.format(path_dict["path1"]))
# Read the values of the file in the dataframe
data = pd.DataFrame(excel_data)

##data.insert(2, "New_1", np.NAN)
data.insert(5, "New_2", 0)
data.insert(6, "New_3", "getdate()")
data.insert(7, "New_4", 'otrs 110172067')
data.insert(8, "New_5", 'ДЗО Альянс 720 дней 11.11.2022 .xlsx')
print(type(data['OptionsParamValue']))
data['OptionsCodeDefault'] = 'NULL'
data['TariffOptionsID'] = 'NULL'
data = data.astype(str)
# Print the content
print(data.head())
data.to_excel('1.xlsx', index=False)
##('sn', '123000091278', NULL,  NULL, '24', '0', getdate(), 'otrs 110172067', 'ДЗО Альянс 720 дней 11.11.2022 .xlsx'),
