#!usr/bin/python
class Csv_config():
	def __init__(self):
		self.dictionary = {}
			
	def searchDOD(self, operation,code,dictionary):
		#This function searches the dictionary passed to it for the specified operation and short code
		#The return value is the column number it can be found
		if code == None:
			return dictionary[operation]
		else:
			for each_dict in dictionary :
				if operation.lower() == each_dict.lower() :
					for key in dictionary[each_dict] :
						if code.lower() == key.lower() :
							return dictionary[each_dict][key]
					
	def buildDictionary(self, csvfile, wksht):
		#gets the csv format for the csv file
		#The return value is a dictionary of dictionaries	
		strlist = []
		#Here the configuration file is opened and each line is appended to strlist[]
		#Its limited to two lines for guaranteed efficiency, for now
		f = open(csvfile,'r')
		for line in f:
			if (line.strip('\n')=='' and strlist != []) or len(strlist) > 3:
				break
			elif line.rstrip('\n') == wksht or strlist != []:
				strlist.append(line.strip('\n'))
			else:
				continue
		if strlist != []:
			# remove worksheet name of header csv list
			strlist = strlist[1:]
			#print strlist
			
			#This is the number of columns in our spreadsheet
			numCol = len(strlist[len(strlist) - 1].split(","))
			
			if len(strlist) == 1:
				headers = strlist[0].split(',')
				col = 1
				dictionary = {}
				for col in range(numCol):
					dictionary[headers[col]] = col+1
				self.dictionary = dictionary
				return self.dictionary
			elif len(strlist) == 2:
				headers = strlist[0]
				headers = headers.strip("\n")
				list_of_op = headers.split(",")
				headers = headers.split(",")

				sub_headers = strlist[1]
				#sub_headers = sub_headers.strip("\n")
				sub_headers = sub_headers.split(",")
				#At this point, headers is a list of values representing row 1
				#sub_headers is a list of short codes representing row 2

				#This is to remove all empty strings from the headers list
				# leaving behind only the the operation headings
				#The exception handling was done since after removing all empty strings, the loop still continues
				#Honestly, i didn't know how to make it stop. Its cool aaa
				try:
					for i in range(len(headers)) :
						headers.remove("")
				except :
					pass	

				lc = []
				#This portion of the code gets the number of sub headers under each main header
				#Hence, the number of short codes under each operation
				for i in range(len(headers)) :
					lc.append(0)
					slyce = list_of_op.index(headers[i])
					for char in list_of_op[slyce: ] :
						try:
							if char == headers[i] :
								lc[i]+= 1
							elif char == headers[i+1] :
									break		
							elif char == "" :
								lc[i] += 1
						except IndexError:
							lc[i] += 1	

				num_of_subheaders = lc
				dictionary = {}

				i=1
				#Here our dictionary is FINALLY built. Do well to understand it.
				#When u do, come and explain it to me. :D
				#Seriously, look through it. Am sure it can be simplified
				while i < numCol:
					for j in range(len(headers)) :
						dictionary[headers[j]] = {}
						for k in range(num_of_subheaders[j]) :
							dictionary[headers[j]][sub_headers[i]] = i+1
							i +=1
				self.dictionary = dictionary
				return self.dictionary
		else:
			# empty strlist -> headers not found
			print 'Invalid Worksheet name. Headers not found'
"""
operation = raw_input("Input the operation to be searched: ")
code = raw_input("Input the short code to be searched: ")

print "This is the column for it: ",
print searchDOD(operation,code,dictionary)
"""
