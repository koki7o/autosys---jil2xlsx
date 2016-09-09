import re, xlsxwriter

from Tkinter import Tk
from tkFileDialog import askopenfilename, asksaveasfile, asksaveasfilename

file = askopenfilename()

excel_file = asksaveasfile(mode='wb', defaultextension=".xlsx")
workbook = xlsxwriter.Workbook(excel_file)
worksheet = workbook.add_worksheet()

list_with_parameters = ['insert_job','job_type','box_name','command','machine','owner','permission','date_conditions','days_of_week','start_times','run_window','condition','description','timezone','run_calendar','n_retrys','term_run_time','box_terminator','watch_interval', 'watch_file','job_terminator','std_out_file','std_err_file','min_run_alarm','max_run_alarm','alarm_if_fail','max_exit_status','chk_files','profile','job_load','priority','auto_delete','group','application', 'exclude_calendar']


#Removes the empty lines
lines = [i for i in open(file, 'r') if i[:-1]]


def read_jil(parameter, column):

		
	row = 0
	
	if parameter == 'insert_job':
	
		worksheet.write_string(row, column, "job_name")
		
	else:
	
		worksheet.write_string(row, column, parameter)
	
	for line in lines:
		
		line = line.strip()
			
		if ':' in line:
		
			if line.startswith('insert_job'):
				
				string = parameter + ':' + ' (\S+)'
				
			else:
			
				string = parameter + ':' + ' (.+)'
			
			x = re.findall(string, line)
			
					
			if len(x) > 0 :
				
				
				worksheet.write_string(row, column, x[0])
			
				
		else:
	
			row += 1
			continue
				
			

col = 0


				 
for i in list_with_parameters:

			
	read_jil(i, col)
	col += 1
	

workbook.close()