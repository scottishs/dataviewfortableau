import os
from tableaudocumentapi import Datasource
from tableaudocumentapi import Workbook

path = os.path.realpath(os.getcwd())
print(path)
# sourceTDS = Datasource.from_file('{}/tableauworkbook.tds'.format(path))

############################################################
# Step 3)  Print out all of the fields and what type they are
############################################################
fname = '{}/tableauworkbook.twb'.format(path)
if fname.endswith('twb') or fname.endswith('twbx'): # or fname.endswith('tds'):

	    sourceWB = Workbook(fname)
	    for sourceTDS in sourceWB.datasources:
	
			print('--- Fields in {}'.format(sourceTDS.name))
			print('----------------------------------------------------------')
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
			# print('----------------------------------------------------------')