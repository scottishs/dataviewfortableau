
import os
import pandas as pd
import numpy as np
import unicodedata
from tableaudocumentapi import Workbook
# from tableaudocumentapi import Connection
# from tableaudocumentapi import dbclass


# Setup
#sourceWB = Workbook('tableauworkbook.twb')
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
    return mappedField.encode('utf-8').strip() #unicodedata.normalize('NFKD', mappedField)


def get_name_resolved_calculation(fieldObj, allFieldsObj):
    """ returns calculation caption of a field, with resolved nested calculations """
    if fieldObj.calculation is None:
        return None

    mappedField = fieldObj.calculation
    while mappedField.find('[Calculation_') >= 0:
            n = mappedField.find('[Calculation_')
            toMap = mappedField[n:mappedField.find(']', n)+1]
            mappedField = mappedField.replace(toMap, allFieldsObj[toMap].caption)
    return mappedField

def _resolve_calculation_parts(fieldObj, allFieldsObj, resolutionType):
    #TODO need to make this loop through Calculation_ and other fields that have names, but need to be resolved

    #this is the target field which we're resoliving calculation names to
    mappedField = fieldObj.calculation
    while mappedField.find('[') >= 0:
        n = mappedField.find('[') #starting point of calc
        toMap = mappedField[n:mappedField.find(']', n)+1] #find the end of the calc name

        #set output to the specified resolutionType
        if resolutionType is None or resolutionType == 'calc':
            rfield = allFieldsObj[toMap].calculation
        elif resolutionType == 'caption':
            rfield = allFieldsObj[toMap].caption
        elif resolutionType == 'calcwithref':
            rfield = '{[{}] {}}'.format(allFieldsObj[toMap].caption, allFieldsObj[toMap].calculation)
        else:
            rfield = 'could not resolve type'
            print ("invalid parameter [{}] passed for 'resolutionType' ".format(resolutionType))

        mappedField = mappedField.replace(toMap, rfield)
    return mappedField


def _get_calc_value(fieldObj, type):
    return getattr(fieldObj, type)

def _field_has_calculation(fieldObj):
    return fieldObj.calculation is not None




def get_resolved_calculation(fieldObj, allFieldsObj, resolutionType=None):
    """ returns calculation caption of a field, with resolved nested calculations """
    """ options: calc (default); caption; calcwithref """
    if not _field_has_calculation(fieldObj):
        return None

    print 'checker: ', fieldObj.name, _field_has_calculation(fieldObj)

    ## Resolve non-named calculated fields
    # look for the fields with 'Calculation_' ids
    #TODO make the function call to resolve the names
    #this is the target field which we're resoliving calculation names to
    mappedField = fieldObj.calculation
    while mappedField.find('[') >= 0:
        n = mappedField.find('[') #starting point of calc
        toMap = mappedField[n:mappedField.find(']', n)+1] #find the end of the calc name

        #set output to the specified resolutionType
        #TODO need to check if the field is a calculation or just an attributes
        if resolutionType is None or resolutionType == 'calc':
            # if toMap == '[Returned]':
            #     print ('RETURN FOUND: PRINTING OBJECTS IN LIST')
            #     for x in allFieldsObj:
            #         print x, x[0]
            # print "GETTING: ", toMap, ' ,', allFieldsObj[toMap]," | RESULT: ", _get_calc_value(allFieldsObj[toMap], 'name')
            rfield = _get_calc_value(allFieldsObj[toMap], 'name')
            # rfield = allFieldsObj[toMap].calculation
        # elif resolutionType == 'caption':
        #     rfield = allFieldsObj[toMap].caption
        # elif resolutionType == 'calcwithref':
        #     rfield = '{[{}] {}}'.format(allFieldsObj[toMap].caption, allFieldsObj[toMap].calculation)
        # else:
        #     rfield = 'could not resolve type'
        #     print ("invalid parameter [{}] passed for 'resolutionType' ".format(resolutionType))
        # print toMap
        mappedField = mappedField.replace(toMap, 'rfield')
    return mappedField
    # return _resolve_calculation_parts(fieldObj, allFieldsObj, resolutionType)




# Loop through all files with the same file extension #
d = []
dw = []

##### TODO Re-enable this
path = ""
if path == "":
    path = os.path.realpath(os.getcwd())

for fname in os.listdir(path):
    if fname == 'Sample Superstore Workbook ALT.twb':
        if fname.endswith('twb') or fname.endswith('twbx') or fname.endswith('tds'):
            sourceWB = Workbook(fname)
            for sourceTDS in sourceWB.datasources:
                if sourceTDS.name != 'Parameters':
                    # for x in sourceTDS.fields:
                    #     print "RAISING 1: TDS Name: {}, Caption: {}, Value: {}".format(sourceTDS.name, sourceTDS.caption, x)

                    for count, field in enumerate(sourceTDS.fields.values()):
                        print "WORKBOOK: ", sourceWB.filename, " DATASOURCE: ", sourceTDS.name, " FIELD: ", field.name
                        d.append({'Workbook': sourceWB.filename,
                                'Field': field.name,
                                'Type': field.datatype,
                                'Default Aggregation': field.default_aggregation,
                                'Field Calculation': get_resolved_calculation(field, sourceTDS.fields)
                                })

d = pd.DataFrame(d)

########################################
# De-Duplication Unique List of Fields #
########################################
#TODO Re-enable this
# d1 = d # Create a Copy
# d1['Field'] = d1['Field'].str.strip('[]') # Remove Brackets [] from field names
# d1 = d1.drop_duplicates(subset=['Default Aggregation', 'Field', 'Field Calculation','Type'])  # Ignore Workbook
# d1 = d1.sort_values('Field') # Sort for convenience
# d1['freq'] = d1.groupby('Field')['Field'].transform('count') # Add Frequency
# d1_1 = d1.query('freq > 1') # Partition Fields with duplicates
# d1_2 = d1.query('freq == 1') # Partition Fields without duplicates
#
# d1_1['Field_Calc_Count'] = d1_1.groupby(['Field'])['Field Calculation'].transform('count')
#
# d1_1_1 = d1_1.query('Field_Calc_Count > 1') # Multipe fields, diff calcs
# d1_1_1['Dup_Calc'] = 1
#
# d1_1_2 = d1_1.query('Field_Calc_Count <=1') # Non duplicate field calcs
#
# d1_1_2['Agg_Null'] = np.where(d1_1_2['Default Aggregation'].isnull(),1,0)
# d1_1_2 = d1_1_2.query('Agg_Null==0') # Non Duplicates
#
# # Join final partitions together #
# final = pd.concat([d1_2, d1_1_2, d1_1_1])
#
# # Check for Same Calc with different field names #
# final['calc_freq'] = final.groupby('Field Calculation')['Field Calculation'].transform('count')
# final['Dup_Field'] = np.where(final['calc_freq']>1,1,"")
#
# final.to_csv('Python Output File.csv'
#         , columns=[#'Workbook',
#         'Field', 'Type', 'Default Aggregation', 'Field Calculation',
#         'Dup_Calc','Dup_Field']
#         ) # Output File







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
