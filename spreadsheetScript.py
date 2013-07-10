#!/usr/bin/python

import gdata
import gdata.client
import gdata.docs.client
import gdata.spreadsheet.service
import gspread
import getopt
import sys

# Spreadsheet Class
class SpreadsheetScript():
	
	def __init__(self, email, password, src='Default'):
		# create gdata spreadsheet client instance and login with email and password
		self.client = gdata.spreadsheet.service.SpreadsheetsService()
		self.client.email = email
		self.client.password = password
		#self.client.source = src
		self.client.ProgrammaticLogin()
		self.sheet_key = ''
		self.wksht_id = ''
		
		# create gspread client instance and login
		self.gs_client = gspread.login(email, password)
		
		# create google docs client and login
		self.docs_client = gdata.docs.client.DocsClient(source=src)
		self.docs_client.client_login(email, password, source=src, service='writely')

	# create a new Google spreadsheet in Drive
	def createSpreadsheet(self, email, pswd, src, doc):
		# create spreadsheet
		document = gdata.docs.data.Resource(type='spreadsheet', title=doc)
		document = self.docs_client.CreateResource(document)
		print 'Spreadsheet created'
		
	def deleteSpreadsheet(self, doc):
		# delete a spreadsheet document by passing Title of document
		for spdsht in self.docs_client.GetAllResources():
			if spdsht.title.text == doc:
				cnfrm = raw_input('Confim deleting '+doc+' (y/n): ')
				if cnfrm == 'y':
					self.docs_client.DeleteResource(spdsht, force=True)
					print 'Delete successful'
					break
				else:
					print 'Delete unsuccessful'
					break
		
	#Prints the data in the worksheet
	def printData(self):
		feed = self.client.GetListFeed(self.sheet_key, self.wksht_id)
		# print field titles of data
		for row in feed.entry:
			for key in row.custom:
				print key+'\t\t',
			print
			break
		
		# print records in each row
		for row in feed.entry:
			for key in row.custom:
				print "%s" % (row.custom[key].text)+'\t\t',
			print
	
	# return a list of spreadsheets titles
	def getSpreadsheetsTitles(self):
		feed = self.client.GetSpreadsheetsFeed()
		spdsheets = []
		for doc in feed.entry:
			spsheets.append(doc.title.text)
		return spdsheets
	
	# print titles of the spreadsheets in Drive
	def printSpreadsheetsList(self):
		docs = self.getSpreadsheetsTitles()
		for doc in docs:
			print doc
	
	#Returns the spreadsheet key for docTitle
	def getSpreadsheetKey(self, docTitle):
		feed = self.client.GetSpreadsheetsFeed()
		for doc in feed.entry:
			if doc.title.text == docTitle:
				self.sheet_key = doc.id.text.rsplit('/', 1)[1]
				return self.sheet_key
		return ''
	
	# returns all worksheet IDs of a spreadsheet as a list
	def getWorksheetIds(self, s_key):
		feed = self.client.GetWorksheetsFeed(s_key)
		wksheets = []
		for sht in feed.entry:
			wksheets.append(sht.id.text.rsplit('/',1)[1])
		return wksheets
		
	#Returns a list of worksheet titles
	def getWorksheetTitles(self, s_key):
		feed = self.client.GetWorksheetsFeed(s_key)
		wksheets = []
		for sht in feed.entry:
			wksheets.append(sht.title.text)
		return wksheets
	
	# select from a list of worksheets	
	def selectWorksheet(self, s_key, index):
		return self.getWorksheetIds(s_key)[index]
		
	def deleteRecord(self, row):
		cnfrm = raw_input('Confim deleting record '+str(row)+' (y/n): ')
		if cnfrm.lower() == 'y':
			feed = self.client.GetListFeed(self.sheet_key, self.wksht_id)
			self.client.DeleteRow(feed.entry[row-1]) # user enters from 1, but records are numbered from 0
			print 'Record delete successful'
		else:
			print 'Record delete unsuccessful'
		
	#Takes in the document name, checks if it exists and asks the user for the worksheet to work with.
	def flow(self, docmnt, rm_doc, delete=False, prnt=False):
		# if rm_doc is not an empty string, delete spreadsheet with that title
		if rm_doc != '':
			self.deleteSpreadsheet(rm_doc)
			sys.exit(2)	# may be taken out, we may ask user if they want to create a new spreadsheet or work with existing spreadsheet
		
		self.sheet_key = self.getSpreadsheetKey(docmnt)
		if self.sheet_key == '':
			print docmnt,"doesn't exist"
			sys.exit(2)
		wkshts = self.getWorksheetTitles(self.sheet_key)
		for i in range(len(wkshts)):
			print i+1, wkshts[i]
		wkid = input("Select worksheet: ")
		self.wksht_id = self.selectWorksheet(self.sheet_key, wkid-1)
		
		# if delete is set to True, do delete operation
		if delete:
			row = input('Enter row to delete: ')
			self.deleteRecord(row)
		#if prnt is set to True, prints the contents of the specified worksheet to the screen
		if prnt:
			self.printData()
		
	def updateCell(self, docName, row, col, new_value, wks = 0):
		#Overwrites the value in the cell specified with new_value
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = spreadsheet.get_worsheet(wks)
		print "This is the value in that current cell :",
		print self.worksheet.cell(row,col).value
		self.worksheet.update_cell(row,col,new_value)
	
	def updateRow(self, docName, row, wks = 0):
		#Overwrites the values in the row with the given values
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = spreadsheet.get_worsheet(wks)
		list_of_values=self.worksheet.row_values(row)
		print "These are the current values in that row: "
		for each in list_of_values :
			print each," ",
		print "\n",	
		for i in range(1,len(list_of_values)+1):
			self.worksheet.update_cell(row,i,raw_input("New:",))
	
	def updateCol(self,docName, col, wks = 0):
		#Overwrites the values in the column with the given values
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = spreadsheet.get_worsheet(wks)
		print "This would empty the specified column of these values:\n",
		list_of_values=self.worksheet.col_values(col)
		print list_of_values
		for i in range(1,len(list_of_values)+1):
			self.worksheet.update_cell(i,col,raw_input("New:",))
	
	def deleteCellValue(self, docName, x, y, wks = 0):
		#Puts an empty string in the specified cell
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = spreadsheet.get_worsheet(wks)
		print "This would empty your specified cell of this current value :",
		print self.worksheet.cell(x,y).value
		self.worksheet.update_cell(x,y,"")
	
	def deleteRowValues(self, docName, row, wks = 0):
		#Puts an empty string in the cells on the specified row
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = spreadsheet.get_worsheet(wks)
		print "This would empty the specified row of these values:\n",
		list_of_values=self.worksheet.row_values(row)
		print list_of_values
		for i in range(1,len(list_of_values)):
			self.worksheet.update_cell(row,i,"")
	
	def deleteColValues(self, docName, col, wks = 0):
		#Puts an empty string in the cells on the specified column
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = spreadsheet.get_worsheet(wks)
		print "This would empty the specified column of these values:\n",
		list_of_values=self.worksheet.col_values(col)
		print list_of_values
		for i in range(1,len(list_of_values)):
			self.worksheet.update_cell(i,col,"")
			
	# prints script documentation
	@staticmethod
	def getHelp():
		print """SPREADSHEET SCRIPT


NAME
	spreadsheetScript - a Python script which interacts with Google Spreadsheets.

USAGE
	python spreadsheetScript [OPTIONS]

OVERVIEW
	spreadsheetScript script enables access to the spreadsheet files in a Google Drive account using the user's email and password.

OPTIONS
Generic Script Information
	--help
		Displays this documentation of the script, and then exits
	
Main Options
	--user
		Provide a username with which to log into Google Drive.
	--pwd
		Provide a password for the username to log in.
	--src
		Provide the spreadsheet source name (like a project name, not spreadsheet name).
	--docName
		Enter the title of an existing spreadsheet document to work with.
	--new
		Creates a new spreadsheet document and names it with the string provided after --new.
	--rmv
		Remove the spreadsheet with the title provided, from Drive.
	--del
		Delete cells, columns or entire rows from the spreadsheet.
	--print
		Does not take any character after it, if specified will print the contents of the document name provided.
	--help
		Prints this screen
	"""
	
def main():
	# check if user has entered the correct options
	try:
		opts, args = getopt.getopt(sys.argv[1:], "", ["user=", "pwd=", "src=", "docName=", "print", "del", "help", "new=", "rmv="])
	except getopt.GetoptError, e:
		print "python spreadsheetScript.py --help. For help:", e, "\n"
		sys.exit(2)
	
	# default username and password to empty strings
	user = ''
	pwd =''
	src = 'Default'
	doc = ''
	new_doc = ''
	rm_doc = ''
	delete = ''	# delete option set to empty string by default, sets (row, column) to delete
	prnt = False	# print option set to False by default
	hlp = False	# help option set to False by default
	
	# get user and pwd values provided by user into their respective variables	
	for opt, val in opts:
		if opt == "--user":
			user = val
		elif opt == "--pwd":
			pwd = val
		elif opt == "--src":
			src = val
		elif opt == "--docName":
			doc = val
		elif opt == "--del":
			delete = True	# delete option sets row to delete to (row, column)
		elif opt == "--print":
			prnt = True	# print option set to true, if the option is added
		elif opt == "--help":
			hlp = True	# help option set to true, if the option is added
		elif opt == "--new":
			new_doc = val
		elif opt == "--rmv":
			rm_doc = val	# title of document to remove
	
	# validate user and pwd values, if any is empty, terminate script		
	if (user == '' or pwd == '') and not hlp:
		print "python spreadsheetScript.py --user [username] --pwd [password] --src [source] --docName [document Title]"
		sys.exit(2)
	elif hlp == True:	# if user added the help option, display script documentation
		SpreadsheetScript.getHelp()
		sys.exit(2)
	elif new_doc != '':
		# create a new spreadsheet document with string in new_doc variable
		smclient = SpreadsheetScript(user, pwd, src)
		smclient.createSpreadsheet(user, pwd, src, new_doc)
		doc = new_doc
	else:
		# create SpreadsheetScript instance with user and pwd fetched	
		smclient = SpreadsheetScript(user, pwd, src)
	
	# passing document title, delete, and prnt options to method
	smclient.flow(doc, rm_doc, delete, prnt)

# if script is being run as a standalone application, its name attribute is __main__
if __name__ == '__main__':
	# execute script
	main()
