# import codecs
import xlsxwriter
from tableaudocumentapi import Workbook
# from tableaudocumentapi import Datasource
# from tableaudocumentapi import ConnectionParser
# from tableaudocumentapi import Connection
# from tableaudocumentapi import dbclass

# Setup
sourceWB = Workbook('tableauworkbook.twb')
printDatasource = False
printWorksheet = True
saveExcel = True

if printDatasource:
    for sourceTDS in sourceWB.datasources[2:3]:
        ############################################################
        # Print out all of the fields and what type they are
        ############################################################
        print('----------------------------------------------------------')
        print('--- fields for {} [{}]'.format(sourceTDS.caption, sourceTDS.name))
        print('--- {} total fields in this datasource'.format(len(sourceTDS.fields)))
        print('----------------------------------------------------------')
        for count, field in enumerate(sourceTDS.fields.values()):
            print('{:>4}: {} is a {}'.format(count+1, field.name, field.datatype))
            blank_line = False
            if field.calculation:
                print('      the formula is {}'.format(field.calculation.encode('utf-8').strip()))
                blank_line = True
            if field.default_aggregation:
                print('      the default aggregation is {}'.format(field.default_aggregation))
                blank_line = True
            if field.description:
                print('      the description is {}'.format(field.description))

            if blank_line:
                print('')
        print('----------------------------------------------------------')

if printWorksheet:
    for ws in sourceWB.datasources[2:3]:
        for f in ws.fields.values():
            print f.worksheets
            # print ws.fields[f]



if saveExcel:
    book = xlsxwriter.Workbook('pythonexcel.xlsx')
    sh = book.add_worksheet()
    sh.set_column('A:A',20)
    sh.write(0,0,'Calculation Name')
    for ws in sourceWB.datasources[2:3]:
        row = 1
        for f in ws.fields.values():
            sh.write(row,0,str(f.worksheets))
            sh.write(row,1,f.name)
            sh.write(row,2,f.caption)
            sh.write(row,3,f.calculation)
            row+=1

    book.close()
