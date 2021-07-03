import os,sys,time,re,xlrd,xlwt,copy

def sorted_nicely(l): 
	convert=lambda text:int(text) if text.isdigit() else text 
	alphanum_key=lambda key:[ convert(c) for c in re.split('([0-9]+)', key)] 
	return sorted(l,key=alphanum_key)

def get_data(sheet,column_num,column_title):

	data=[]
	column_read=False
	for row in range(sheet.nrows):
		cell_value=sheet.cell_value(row,column_num)
		if column_read==True:
			data.append(cell_value)
		if cell_value==column_title:
			column_read=True

	return data

type_final_fn='Replicates'

stop=False

#prefix='Myanmar study'

input_subdir=sys.argv[1]

if input_subdir[-1]!='/':
	input_subdir+='/'

if not os.path.exists(input_subdir):
	print('\n\n'+' '*4,'> The path you used to indicate the input subdirectory doesn\'t exist!\n')
	print(' '*4,'> "'+input_subdir+'" ? Please check again in case there is a typo\n\n')
	sys.exit()

xls_files=[input_subdir+x for x in os.listdir(input_subdir) if x.endswith('.xls') or x.endswith('.xlsx')]
xls_files=[x for x in xls_files if not x.startswith('.')]

if len(xls_files)==0:
	print('\n\n'+' '*4,'> I couldn\'t find any .xls files in input subdirectory you indicated :-/\n\n')
	sys.exit()

print('\n','='*50)
print('\n\nI found', str(len(xls_files)), '.xls files in the "' + input_subdir + '" input_subdirectory (subfolder):\n')

i=0
for x in xls_files:
	i+=1
	print(' '*4,str(i)+'.'+' "'+x[len(input_subdir):]+'"')

print()

xls_data=[]

for x in xls_files:

	book = xlrd.open_workbook(x)
	sheet = book.sheet_by_index(0)

	sample_name_data=get_data(sheet,1,'Sample Name')
	target_name_data=get_data(sheet,2,'Target Name')
	ct_replicate_data=get_data(sheet,6,'Cт')
	ct_mean_data=get_data(sheet,7,'Cт Mean')
	ct_sd_data=get_data(sheet,8,'Cт SD')
	quantity_data=get_data(sheet,9,'Quantity')
	quantity_mean_data=get_data(sheet,10,'Quantity Mean')
	quantity_sd_data=get_data(sheet,11,'Quantity SD')

	if stop:
		sample_name_data=sample_name_data[:-4]
		target_name_data=target_name_data[:-4]
		ct_mean_data=ct_mean_data[:-4]
	
	for i in range(len(sample_name_data)):
		l=[sample_name_data[i],target_name_data[i],ct_replicate_data[i],ct_mean_data[i],ct_sd_data[i],quantity_data[i],quantity_mean_data[i],quantity_sd_data[i]]
		xls_data.append(l)

xls_replicate_data=[]

my_max=0
for x in xls_data:

	if len(str(x[0]))==0:
		xls_data.remove(x)
		continue

	rpl_ct_values=[x[2]]
	rpl_qnt_values=[x[5]]

	status=False
	for y in xls_data:

		if y[0]==x[0] and y[1]==x[1]:
			if x!=y:
				status=True
				rpl_ct_values.append(y[2])
				rpl_qnt_values.append(y[5])
				xls_data.remove(y)

	#replicate_ct_values=['' if len(str(my_value))==0 else my_value for my_value in replicate_ct_values]
	#replicate_qnt_values=['' if len(str(my_value))==0 else my_value for my_value in replicate_qnt_values]

	if status:
		my_list=[x[0],x[1],rpl_ct_values[0],rpl_ct_values[1],x[3],x[4],rpl_qnt_values[0],rpl_qnt_values[1],x[6],x[7]]
	else:
		my_list=[x[0],x[1],rpl_ct_values[0],'replicate not found',x[3],x[4],rpl_qnt_values[0],'replicate not found',x[6],x[7]]

	if len(my_list)>my_max:
		my_max=len(my_list)
	
	my_list=[re.sub('undetermined','',str(z).lower()) for z in my_list]

	xls_replicate_data.append(my_list)

sample_name_data_to_sort=[]

k=0
for xx in xls_replicate_data:
	x=str(xx[0])
	sample_name_data_to_sort.append(x+'_'+str(k))
	k+=1

sorted_snd_indices=[sample_name_data_to_sort.index(x) for x in sorted_nicely(sample_name_data_to_sort)]

book = xlwt.Workbook(type_final_fn)
sh = book.add_sheet(type_final_fn)

other_fields=['Sample Name','Target Name','Cq Rpl 1','Cq Rpl 2','Cq Mean','Cq SD','Qnt Rpl 1','Qnt Rpl 2','Quantity Mean','Quantity SD']

for i in range(len(other_fields)):
	sh.write(0, i, other_fields[i])

k=0
for i in sorted_snd_indices:

	if len(str(xls_replicate_data[i][0]))>0:

		k+=1
		for j in range(my_max):
			try:
				sh.write(k,j,xls_replicate_data[i][j])
			except:
				break

timestr = time.strftime("%d-%m-%Y_%H-%-M-%S")

output_subdir=input_subdir.lstrip('input/')
output_subdir='output/'+output_subdir

if not os.path.exists(output_subdir):
	os.makedirs(output_subdir)

filename=output_subdir + type_final_fn + ' - ' + timestr + '.xls'
book.save(filename)

print('\nThe results were saved at:\n')
print(' '*4,'>>','"'+filename+'"\n\n')
print('='*50,'\n')