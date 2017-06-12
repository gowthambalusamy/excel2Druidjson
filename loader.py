import os
import json
import sys
import xlrd
import csv
import shutil

newpath = "output"
if not os.path.exists(newpath):
    os.makedirs(newpath)
else :
	shutil.rmtree(newpath)
	os.makedirs(newpath)

from pprint import pprint

with open('config.json') as data_file:    
    data1 = json.load(data_file)

xls_file=data1['config']['xls']
#json_file=data1['config']['json']
sheet=data1['config']['sheet']

workbook = xlrd.open_workbook(xls_file)
worksheet = workbook.sheet_by_name(sheet)

data = []
keys = [v.value for v in worksheet.row(0)]
for row_number in range(worksheet.nrows):
    if row_number == 0:
        continue
    row_data = {}
    for col_number, cell in enumerate(worksheet.row(row_number)):
        row_data[keys[col_number]] = cell.value
    data.append(row_data)

with open("output/output.json", 'w') as json_file:
	json_file.write(json.dumps(data))

with open("output/output.json", 'r+') as jsonFile:
    data = json.load(jsonFile)
    a=data[0]
    c=data

ki={}
headers={}
keys=[]
samp=[]
keys=a.keys()
headers=keys
li=len(keys)
l=len(c)
for j in range(0,l):
	b=data[j]
	for i in range(0,li):
		string=b[keys[i]]
		if type(string) != 'int' and type(string) != 'float':
			print type(212116.8)

		else:
			print "hi"
		if b[keys[i]]=="":
			b[keys[i]]=None


		samp=b[keys[i]]
		ky=keys[i]
		ki[ky]=samp				
		
	# print ki

	with open("output/output1.json", 'a') as json_file:
			# json.dump(ki,json_file)
			json_file.write(json.dumps(ki))
			json_file.write("\n")
# unique files maving

with open('config.json') as uniq_file:    
    uniq1 = json.load(uniq_file)
# print uniq1['config']['Uniq Files']

fd=uniq1['config']['Uniq Files']
vi={}
with open("output/output.json", 'r+') as jsonFile:
    uniq = json.load(jsonFile)
    a=uniq[0]
    # print a
    c=uniq
uy=[]
cs={}
count=0
lo=len(c)
p=len(fd)
os.remove("output/output.json")
for i in range (0,p):
	value=fd[i]
	for j in range(0,lo):
		vi=uniq[j][value]
		
		uy=vi
		# print uy
		with open("output/"+value+".txt", 'a') as json_file:
			# json.dump(vi,json_file)
				json_file.write(vi)
				json_file.write("\n")
	from collections import OrderedDict
	with open("output/"+value +".txt") as fin:
    		lines = (line.rstrip() for line in fin)
    		unique_lines = OrderedDict.fromkeys( (line for line in lines if line))
	cs= unique_lines.keys()
	le=len(cs)
	os.remove("output/"+value +".txt")
	
	with open("output/"+value+".csv", 'a') as json_file:
			# json.dump(vi,json_file)
				json_file.write("ID"+"	"+"NAME"+"	"+"CODE"+"	"+"DESC"+"	"+"SYN1"+"	"+"SYN2"+"	"+"SYN3")
	id=200001	
			
	for i in range(0,le):
	
		with open("output/"+value+".csv", 'a') as json_file:
			# json.dump(vi,json_file)
				json_file.write("\n")
				json_file.write(str(id)+"	"+cs[i]+"	"+cs[i]+"	"+cs[i])
		id=id +1
		#print id
	id =200001
		
	with open("output/headers.json", 'w') as json_file:
			# json.dump(vi,json_file)
			json_file.write("\n")
			json_file.write(json.dumps(headers))

				
				
				
				






			

