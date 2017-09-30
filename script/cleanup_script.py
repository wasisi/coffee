import openpyxl
import datetime
import csv

NoneString=""

#load document...TODO: name of document and path to the document
workbook = openpyxl.load_workbook('TransactionListingSale30.xlsx')

#the sheets in the workbook
sheets = workbook.get_sheet_names()

#how many sheets the document has
nsheets = len(sheets)


months={'01':'Jan','02':'Feb','03':'Mar',
        '04':'Apr','05':'May','06':'June',
        '07':'July','08':'Aug','09':'Sept',
        '10':'Oct','11':'Nov','12':'Dec'}

#print the sheet names
print("Workbook sheets: ",workbook.get_sheet_names())

#the columns of a row that we are interested in
#These correspond to excel sheet column letters
#
#Field-Column map
#A=Transanr,        B=Lotnr,        C=Marks,
#D=Grade,           E=Bags,         F=Weight,
#G=Saleno,          H=BagsBought    I=WeightBought
#J=BuyerCode,       K=Price,        L=SeatNr,
#M=AUTCODE,         N=Status,       O=Datum,        P=TijD

interestCols='ABCDEFGHIJKLMNO'

#get the sheet. Assume that the workbook has
#only one sheet? Validate?
sheet = workbook.get_sheet_by_name(sheets[0])

def correct_mark_format(cell_value):

	if cell_value==None:
		return NoneString

	#begin with removing the spaces in the given
	#string. Strings are immutables in Python
	marks2 = cell_value.replace(" ","")

	# the string using / as delimeter 
	separate_list = marks2.split('/')

	#no factory number and I don't
	#know what to do
	if len(separate_list) <= 2:
		return [cell_value,marks2]

	#get the factory number
	factory_num = separate_list[2]

	#find and replace the O to zero string
	factory_num.replace('O','0')

	separate_list[2]=factory_num

	marks2=""
	#merge the list and return the value
	for i in range(len(separate_list)):
		marks2 += separate_list[i]+'/'

	#remove full stops
	marks2 = marks2.replace('.',"")
	marks2_split = marks2.split('/')
	return (cell_value,marks2,marks2_split[2],marks2_split[2],marks2_split[0])


def process_datum(cell_value):
    """
    Proceccess the Datum column to produce an ISODate
    and Season
    """

    if cell_value==None:
		    return NoneString

    value = str(cell_value.value.date())
    value_list = value.split('-')
    year =  value_list[0]
    month = value_list[1]
    day = value_list[2]
    yearint = int(year)
    yearprev = yearint-1
    #print(" %s %s %s "%(day,month,year))
    #print(type(cell_value.value))

    return (str(day)+'-'+months[str(month)]+'-'+str(year),str(yearprev)+'-'+str(year))
 
#we will output the needed data into a CVS file
#which will be our clean file to upload to DB
outputname = 'output.csv'
csv_file = open(outputname,'wt')

#try to write data into the CSV file
try:
    csvwriter = csv.writer(csv_file,lineterminator='\n')

    #write header with column info
    #csvwriter.writerow(('#TRANSNR','LOTNT','MARKS', 'MARKS2','REF','REF2', 'BAGMARK','GRADE-GR',
    #                    'BAGSNR','WEIGHT-Kgr','SALENO','BAGSBOUGHTNR','WEIGHTBOUGHT-Kgr',
    #                    'BUYERCODE','PRICE','SEATNR','AUCTCODE','STATUS','ISODATE', 'SEASON', 'TIJD','VALUE'))

    csvwriter.writerow(('#TRANSNR','LOTNT','MARKS', 'MARKS2','REF','REF2', 'BAGMARK','GRADE-GR',
                        'BAGSNR','WEIGHT-Kgr','SALENO','BAGSBOUGHTNR','WEIGHTBOUGHT-Kgr',
                        'BUYERCODE','PRICE','SEATNR','AUCTCODE','STATUS','ISODATE', 'SEASON','VALUE'))

    for row in range(2,sheet.max_row):
        row_vals=[]
        weightBought=0.0
        price = 0.0
        for col in interestCols:
            cell = "{}{}".format(col,row)

            if col == 'C': #this is the marks column
                cell_value = correct_mark_format(sheet[cell].value)
                for val in cell_value:
                    row_vals.append(val)
            elif col=='O':
                values = process_datum(sheet[cell])
                for val in values:
                    row_vals.append(val)
                #row_vals.append(val)
            elif col=='I':

                #cache the weight bought for value calculation
                weightBought = float(sheet[cell].value)
                row_vals.append(sheet[cell].value)
            elif col=='K':
                #cache the price for value calculation
                price = float(sheet[cell].value)
                row_vals.append(sheet[cell].value)
            else:
                row_vals.append(sheet[cell].value)

        #we processed the sheet let's calculate the value
        value = (weightBought/50.)*price
        row_vals.append(value)

        #make a tuple from the list values
        row_vals = tuple(row_vals)

			  #let's write the row in the specified
        csvwriter.writerow(row_vals)
finally:
    csv_file.close()
	
