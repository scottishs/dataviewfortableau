# import codecs
# import xlsxwriter
# import re
import os
import pandas as pd
from tableaudocumentapi import Workbook

# from tableaudocumentapi import Datasource
# from tableaudocumentapi import ConnectionParser
# from tableaudocumentapi import Connection
# from tableaudocumentapi import dbclass


# Setup
sourceWB = Workbook('tableauworkbook.twb')
printDatasource = False
printWorksheet = False
saveExcel = False


def get_calc_resolved_calculation(fieldObj, allFieldsObj):
    """ returns calculation value of a field, with resolved nested calculations """
    if fieldObj.calculation is None:
        return None

    mappedField = fieldObj.calculation
    while mappedField.find('[Calculation_') >= 0:
            n = mappedField.find('[Calculation_')
            toMap = mappedField[n:mappedField.find(']', n)+1]
            mappedField = mappedField.replace(toMap, allFieldsObj[toMap].calculation)
    return mappedField.encode('ascii', 'xmlcharrefreplace').strip()


def get_name_resolved_calculation(fieldObj, allFieldsObj):
    """ returns calculation value of a field, with resolved nested calculations """
    if fieldObj.calculation is None:
        return None

    mappedField = fieldObj.calculation
    while mappedField.find('[Calculation_') >= 0:
            n = mappedField.find('[Calculation_')
            toMap = mappedField[n:mappedField.find(']', n)+1]
            mappedField = mappedField.replace(toMap, allFieldsObj[toMap].caption)
    return mappedField



# Loop through all files with the same file extension #
d = []

for fname in os.listdir(os.getcwd()):
    if fname.endswith('twb') or fname.endswith('twbx') or fname.endswith('tds'):
        sourceWB = Workbook(fname)
        for sourceTDS in sourceWB.datasources:
            for count, field in enumerate(sourceTDS.fields.values()):

                d.append({'Field': field.name.encode('ascii', 'xmlcharrefreplace'),
                        'Type': field.datatype,
                        'Default Aggregation': field.default_aggregation,
                        'Field Calculation': 'field'#get_calc_resolved_calculation(field, sourceTDS.fields)
                        })

d = pd.DataFrame(d)
d.to_csv('Python Output File.csv') # Output File


#==============================================================================
# if printWorksheet:
#     for ws in sourceWB.datasources[2:3]:
#         for f in ws.fields.values():
#             print ('######### GETTING: {}'.format(f.name))
#             print (get_name_resolved_calculation(f, ws.fields))
#             # print (get_nested_calculation(ws.fields['CY Label (Percentage)'], ws.fields))
#
#             # print (f.worksheets)
#             # print (f.id, f.name, f.caption)
#             # print (f.calculation)
#
#
# if printDatasource:
#     for sourceTDS in sourceWB.datasources[2:3]:
#         ############################################################
#         # Print out all of the fields and what type they are
#         ############################################################
#         print('----------------------------------------------------------')
#         print('--- fields for {} [{}]'.format(sourceTDS.caption, sourceTDS.name))
#         print('--- {} total fields in this datasource'.format(len(sourceTDS.fields)))
#         print('----------------------------------------------------------')
#         for count, field in enumerate(sourceTDS.fields.values()):
#             print('{:>4}: {} is a {}'.format(count+1, field.name, field.datatype))
#             blank_line = False
#             if field.calculation:
#                 print('      the formula is {}'.format(field.calculation.encode('utf-8').strip()))
#                 blank_line = True
#             if field.default_aggregation:
#                 print('      the default aggregation is {}'.format(field.default_aggregation))
#                 blank_line = True
#             if field.description:
#                 print('      the description is {}'.format(field.description))
#
#             if blank_line:
#                 print('')
#         print('----------------------------------------------------------')
#
#
# if saveExcel:
#     book = xlsxwriter.Workbook('pythonexcel.xlsx')
#     sh = book.add_worksheet()
#     sh.set_column('A:A',20)
#     sh.write(0,0,'Calculation Name')
#     for ws in sourceWB.datasources[2:3]:
#         row = 1
#         for f in ws.fields.values():
#             sh.write(row,0,str(f.worksheets))
#             sh.write(row,1,f.name)
#             sh.write(row,2,f.caption)
#             sh.write(row,3,f.calculation)
#             row+=1
#
#     book.close()
#==============================================================================
