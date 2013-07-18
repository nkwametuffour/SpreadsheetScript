#!usr/bin/python
class Csv_config():
	def __init__(self):
		self.dictionary = {}
	#This function searches the dictionary passed to it for the specified operation and short code
	#The return value is the column number it can be found			
	def searchDOD(self, operation,code,dictionary) :
		
		for each_dict in dictionary :
			if operation == each_dict :
				for key in dictionary[each_dict] :
					if code == key :
						return dictionary[each_dict][key]
	#The return value is a dictionary of dictionaries					
	def buildDictionary(self, csvfile):
		strlist = []
		#Here the configuration file is opened and each line is appended to strlist[]
		#Its limited to two lines for guaranteed efficiency, for now
		with open(csvfile,"r") as f:
			for line in f :
				strlist.append(line)

		#This is the number of columns in our spreadsheet
		numCol = len(strlist[1].split(","))

		headers = strlist[0]
		headers = headers.strip("\n")
		list_of_op = headers.split(",")
		headers = headers.split(",")

		sub_headers = strlist[1]
		sub_headers = sub_headers.strip("\n")
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
"""
operation = raw_input("Input the operation to be searched: ")
code = raw_input("Input the short code to be searched: ")

print "This is the column for it: ",
print searchDOD(operation,code,dictionary)
"""
