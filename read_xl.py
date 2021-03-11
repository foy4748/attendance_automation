from openpyxl import load_workbook as lw
from datetime import datetime
# Attendance.xlsx

try: 
    from openpyxl.cell import get_column_letter
except ImportError:
    from openpyxl.utils import get_column_letter

wb = lw('Attendance.xlsx', data_only= True)

#Getting sheet names
sheets = wb.sheetnames
print(sheets)

w_sh = wb[sheets[0]]

#for i in range(1,7):
#	print(w_sh.cell(row=1,column=i).value , get_column_letter(i))

collection = list()
fh = open('./store.txt')
i = 2
#print(w_sh.cell(row=i,column=2).value)
while w_sh.cell(row=i,column=2).value != None:
	collection.append( (str(w_sh.cell(row=i,column=2).value.strip()), str(w_sh.cell(row=i,column=6).value.strip())) )
	#fh.write(str(w_sh.cell(row=i,column=2).value) + str(w_sh.cell(row=i,column=6).value) + '\n')
	i = i + 1
	




#Checking Attended Pupils
attended = list()

for line in fh:
	line = line.strip()
	
	
	A = line.split('-')
	for i in collection:
		roll_ = i[1].split('-')
		if A[-1] == roll_[-1]:
			attended.append(i)
	
	for i in collection:
		name_ = i[0].upper()
		if name_ == line.upper():
			attended.append(i)		
fh.close()
#Writing Attendance

#Switching sheet
w_sh = wb[sheets[1]]

i = 5

while w_sh.cell(row=2, column=i).value != None:
	i = i + 1
	verify = w_sh.cell(row=2, column=i).value
	current_date = datetime.today().strftime('%d-%m-%Y')
	
	if  verify == None and verify != current_date:
		col_letter = get_column_letter(i)
		index = col_letter + '2'
		print(index)
		w_sh[index] = current_date
		wb.save('Attendance.xlsx')
		break
	elif verify == current_date:
		print('Attendance is Up to Date')
		break
	






