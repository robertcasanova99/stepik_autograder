import subprocess
import os
import sys
import math
import xlwt
import time 
from xlwt import Workbook

def convert(seconds): 
    seconds = seconds % (24 * 3600) 
    hour = seconds // 3600
    seconds %= 3600
    minutes = seconds // 60
    seconds %= 60
      
    return "%d:%02d:%02d" % (hour, minutes, seconds)

verbose = False
extensions = {'c++': '.cpp', 'java': '.java', 'python': '.py'}
lang = ""
if len(sys.argv) > 1:
	if '-v' in sys.argv:
		verbose = True
	user_requested_lang = sys.argv[-1].lower()
	p = subprocess.Popen(['submitter','lang'], shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
	possible_langs = []
	for line in p.stdout.readlines():
		possible_langs.append(line[:-1].decode('utf-8'))
	if user_requested_lang in possible_langs:
		lang = user_requested_lang

	if lang == "":
		print('Not a supported language for your currently selected problem')
		sys.exit()
else:
	print('Please specify a language')
	sys.exit()
			

extension = extensions[lang]
submission_files = os.listdir('submissions/')
code_files = []

#constructs a list of file names for all files that end in .extension
for file in submission_files:
	if len(file) > len(extension) and file[-1*(len(extension)):] == extension:
		code_files.append(file)

wb = Workbook()
sheet = wb.add_sheet('Student Grades') 
sheet.write(0,0, 'Student Name')
sheet.write(0,1, 'Result')
sheet.write(0,2, 'Percentage')

if verbose:
	print('Submitting ' + str(len(code_files)) + " " + extension + ' files for evaluation...')
			
start = time.time()

for i, file in enumerate(code_files):
	name = file[:file.find('_')]
	sheet.write(i + 1, 0, name)

	args = ['submitter', 'submit', '-l', lang, 'submissions/' + file]
	p = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
	result = p.stdout.readlines()[-1][:-3].decode('utf-8')
	if verbose:
		print('Result for ' + name + ': ' + result)

	#try again
	if 'error' in result.lower() or 'id' in result.lower() or result == '':
		if verbose:
			print('\nError in the latest submission, trying again...')
		p = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
		result = p.stdout.readlines()[-1][:-3].decode('utf-8')
		if 'error' in result.lower() or 'clientid' in result.lower() or result == '':
			if verbose:
				print('Error persists. Manually review the submission of ' + name + '\n')
			result = 'Needs manual revision'

	sheet.write(i + 1, 1, result)

	try:
		first_space = result.find(" ")
		second_space = result.find(" ", result.find('of'))
		third_space = result.find("test") - 1
		score = int(result[:first_space])
		total = int(result[second_space + 1:third_space])
		percentage = str(math.ceil((score/total * 100))) + "%"
		sheet.write(i + 1, 2, percentage)
	except:
		sheet.write(i + 1, 2, 'NaN')


wb.save('stepik_results.xls')
end = time.time()

print('Results saved in this directory in \'stepik_results.xls\', process duration: ' + convert(end-start))