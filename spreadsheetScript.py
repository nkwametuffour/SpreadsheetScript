#!/usr/bin/python

import gdata.spreadsheet.service
import gdata
import gdata.client
import gdata.docs.client
import gdata.spreadsheets.client
import gspread
import getopt
import sys
import os
import webbrowser
import smtplib
from Crypto.Cipher import AES
import base64
import os

# Spreadsheet Class
class SpreadsheetScript():
	
	CLIENT_ID = '498758732944.apps.googleusercontent.com'
	CLIENT_SECRET = 'A_KMQ3yGorVSuTT-3P7DuHRC'
	SCOPE = 'https://spreadsheets.google.com/feeds/'
	USER_AGENT = 'Spreadsheet'
	
	def __init__(self, src='Default'):
		
		try:
			f = open(os.environ['HOME']+'/.hide/tokens.txt','r+')
			cred = f.read().split('\n')
			print cred[:]
			try:
				self.__create_clients(self.__decrypt(cred[0]), self.__decrypt(cred[1]), src)
			except Exception, e:
				print e
		except:
			try:
				user = raw_input('Enter username: ').strip()
				pwd = raw_input('Enter password: ').strip()
				# validate user and password
				try:
					self.__create_clients(user, pwd, src)
				except Exception, e:
					print 'Login failed.',e
					
				self.__store_cred([user, pwd])
			except Exception, e:
				print 'Program execution failed:',e
		#token = gdata.gauth.OAuth2Token(client_id=self.CLIENT_ID, client_secret=self.CLIENT_SECRET, scope=self.SCOPE, user_agent=self.USER_AGENT, access_token=tkn[1], refresh_token=tkn[0])
		# create gdata spreadsheet client instance and login with email and password
		#self.client = gdata.spreadsheets.client.SpreadsheetsClient()
		#token.authorize(self.client)
		self.sheet_key = ''
		self.wksht_id = ''	
		
	def __create_clients(self, user, pswd, src):
		self.client = gdata.spreadsheet.service.SpreadsheetsService()
		self.client.email = user
		self.client.password = pswd
		self.client.ProgrammaticLogin()
		
		# create gspread client instance and login
		self.gs_client = gspread.login(user, pswd)
		
		# create google docs client and login
		self.docs_client = gdata.docs.client.DocsClient(source=src)
		self.docs_client.client_login(user, pswd, source=src, service='writely')
		
	def __get_access_auth(self):
		token = gdata.gauth.OAuth2Token(client_id=self.CLIENT_ID, client_secret=self.CLIENT_SECRET, scope=self.SCOPE, user_agent=self.USER_AGENT)
		webbrowser.open(token.generate_authorize_url(redirect_uri='urn:ietf:wg:oauth:2.0:oob'))
		code = raw_input('What is the verification code? ').strip()
		token.get_access_token(code)
		return [token.refresh_token, token.access_token]
		
	def __store_token(self, accs_data):
		os.system("mkdir "+os.environ['HOME']+"/.hide")
		os.system("touch tokens.txt")
		f = open(os.environ['HOME']+'/.hide/tokens.txt','w+')
		for i in range(len(accs_data)):
			f.write(accs_data[i] + '\n')
		f.read()
		f.close()
	
	def __encrypt(self, string):
		BLOCK_SIZE = 32
		PADDING = '{'
		pad = lambda s: s + (BLOCK_SIZE - len(s) % BLOCK_SIZE) * PADDING
		EncodeAES = lambda c, s: base64.b64encode(c.encrypt(pad(s)))
		DecodeAES = lambda c, e: c.decrypt(base64.b64decode(e)).rstrip(PADDING)
		# generate a random secret key
		secret = os.urandom(BLOCK_SIZE)
		# create a cipher object using the random secret
		cipher = AES.new(secret)
		# encoded string
		encoded = EncodeAES(cipher, string)
		return encoded
	
	def __decrypt(self, enc):
		return DecodeAES(cipher, encoded)
		
	def __store_cred(self, log_cred):
		os.system("mkdir "+os.environ['HOME']+"/.hide")
		os.system("touch tokens.txt")
		f = open(os.environ['HOME']+'/.hide/tokens.txt','w+')
		for i in range(len(log_cred)):
			f.write(self.__encrypt(log_cred[i]) + '\n')
		f.read()
		f.close()

	# create a new Google spreadsheet in Drive
	def createSpreadsheet(self, doc):
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
				print doc.id.text
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
			feed = self.client.get_list_feed(self.sheet_key, self.wksht_id)
			self.client.DeleteRow(feed.entry[row-1]) # user enters from 1, but records are numbered from 0
			print 'Record delete successful'
		else:
			print 'Record delete unsuccessful'
			
	#Sends mail to the user
	def sendMail(success = True):
		server = smtplib.SMTP()
		server.connect('smtp.gmail.com', 587)
		server.ehlo()
		server.starttls()
		server.ehlo()
		server.login("rancardinterns2013@gmail.com","nopintern2013")
		if success:
			message ="SpreadsheetScript has been successful"
		else:
			message = "SpreadsheetScript was NOT successful"
			
		server.sendmail('rancardinterns2013@gmail.com', self.client.email, message)
		
	#Takes in the document name, checks if it exists and asks the user for the worksheet to work with.
	def flow(self, docmnt, rm_doc, delete=False, prnt=False, edit=False, delVal=False):
		# if rm_doc is not an empty string, delete spreadsheet with that title
		if rm_doc != '':
			self.deleteSpreadsheet(rm_doc)
			sys.exit(2)	# may be taken out, we may ask user if they want to create a new spreadsheet or work with existing spreadsheet
		
		# if docmnt passed is not an empty string, get its key for use
		if docmnt != '':
			self.sheet_key = self.getSpreadsheetKey(docmnt)
		
		if self.sheet_key == '':
			print "File doesn't exist"
			sys. exit(2)
		wkshts = self.getWorksheetTitles(self.sheet_key)
		for i in range(len(wkshts)):
			print i+1, wkshts[i]
		wkid = input("Select worksheet: ")
		self.wksht_id = self.selectWorksheet(self.sheet_key, wkid-1)
		
		#if prnt is set to True, prints the contents of the specified worksheet to the screen
		if prnt and docmnt != '':
			self.printData()

		#if delete is set to True, do delete operation
		if delete:
			row = input('Enter row to delete: ')
			self.deleteRecord(row)
		#if edit is set to True,
		if edit:
			choice = (raw_input('Enter row to edit a row, col to edit a col and cell to edit a cell: ')).lower().strip()
			if choice == 'row':
				row = input('Enter row number: ')
				values = raw_input('Enter values in order (Example 1,2,3): ')
				values = values.split(',')
				self.updateRow(docmnt, row, values, wkid -1)
				self.printData()
			elif choice == 'col':
				col = input('Enter column number: ')
				values = raw_input('Enter values in order (Example 1,2,3): ')
				values = values.split(',')
				self.updateCol(docmnt, col, values, wkid -1)
				self.printData()
			elif choice == 'cell':
				cell = raw_input('Enter row,column,value in that order: ')
				#value = raw_input('Cell value: ')
				cell = cell.split(',')
				self.updateCell(docmnt, cell[0],cell[1],cell[2], wkid -1)
				self.printData()
		
		if delVal:
			cell = raw_input("Enter the cell's row, column in that order.(Example 2,3): ")
			cell = (cell.strip()).split(',')
			sef.deleteCellValue(docmnt, cell[0], cell[1], wkid-1)
				
		
	def updateCell(self, docName, row, col, new_value, wks = 0):
		#Overwrites the value in the cell specified with new_value
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = self.spreadsheet.get_worksheet(wks)
		#print "This is the value in that current cell :",
		#print self.worksheet.cell(row,col).value
		self.worksheet.update_cell(row,col,new_value)
	
	def updateRow(self, docName, row, values, wks = 0):
		#Overwrites the values in the row with the given values
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = self.spreadsheet.get_worksheet(wks)
		#list_of_values=self.worksheet.row_values(row)
		#print "These are the current values in that row: "
		#for each in list_of_values :
		#	print each," ",
		#print "\n",	
		for i in range(len(values)):
			self.worksheet.update_cell(row,i+1,values[i])
	
	def updateCol(self,docName, col, values, wks = 0):
		#Overwrites the values in the column with the given values
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = self.spreadsheet.get_worksheet(wks)
		#print "This would empty the specified column of these values:\n",
		#list_of_values=self.worksheet.col_values(col)
		#print list_of_values
		for i in range(len(values)):
			self.worksheet.update_cell(i+2,col,values[i])
	
	def deleteCellValue(self, docName, row, col, wks = 0):
		#Puts an empty string in the specified cell
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = self.spreadsheet.get_worksheet(wks)
		#print "This would empty your specified cell of this current value :",
		#print self.worksheet.cell(x,y).value
		self.worksheet.update_cell(row,col,"")
	
	def deleteRowValues(self, docName, row, wks = 0):
		#Puts an empty string in the cells on the specified row
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = self.spreadsheet.get_worksheet(wks)
		print "This would empty the specified row of these values:\n",
		list_of_values=self.worksheet.row_values(row)
		print list_of_values
		for i in range(1,len(list_of_values)):
			self.worksheet.update_cell(row,i,"")
	
	def deleteColValues(self, docName, col, wks = 0):
		#Puts an empty string in the cells on the specified column
		self.spreadsheet = self.gs_client.open(docName)
		self.worksheet = self.spreadsheet.get_worksheet(wks)
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
	spreadsheetScript script enables access to the spreadsheet files in a Google Drive account using OAuth2.0 authentication.

OPTIONS
Generic Script Information
	--help
		Displays this documentation of the script, and then exits
	
Main Options
	--src
		Provide the spreadsheet source name (like a project name, not spreadsheet name).
	--docName
		Enter the title of an existing spreadsheet document to work with.
	--new
		Creates a new spreadsheet document and names it with the string provided after --new.
	--edit
		Edit column, row and cell values
	--rmv
		Remove the spreadsheet with the title provided, from Drive.
	--del
		Delete cells entire rows from the spreadsheet.
	--delVal
		Deletes the value in a cell
	--print
		Does not take any character after it, if specified will print the contents of the document name provided.
	--help
		Prints this screen
	"""
	
def main():
	# check if user has entered the correct options
	try:
		opts, args = getopt.getopt(sys.argv[1:], "", ["src=", "docName=", "print", "del", "help", "new=", "rmv=", "edit", "delVal"])
	except getopt.GetoptError, e:
		print "python spreadsheetScript.py --help. For help:", e, "\n"
		sys.exit(2)
	
	src = 'Default'
	doc = ''
	new_doc = ''
	rm_doc = ''
	delete = False	# delete option set to empty string by default, sets (row, column) to delete
	prnt = False	# print option set to False by default
	hlp = False	# help option set to False by default
	edit = False
	delVal = False
	# get user and pwd values provided by user into their respective variables	
	for opt, val in opts:
		if opt == "--src":
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
		elif opt == "--edit":
			edit = True
		elif opt == "--delVal":
			delVal = True
	
	# validate user and pwd values, if any is empty, terminate script
	if hlp == True:	# if user added the help option, display script documentation
		SpreadsheetScript.getHelp()
		sys.exit(2)
	elif new_doc != '':
		# create a new spreadsheet document with string in new_doc variable
		smclient = SpreadsheetScript(src)
		smclient.createSpreadsheet(new_doc)
		doc = new_doc
	else:
		# create SpreadsheetScript instance with user and pwd fetched	
		smclient = SpreadsheetScript(src)
	
	loop  = True
	
	# passing document title, delete, and prnt options to method
	while loop:
		smclient.flow(doc, rm_doc, delete, prnt, edit, delVal)
		choice = (raw_input('Continue? y/n : ')).lower()
		if not choice == 'y':
			loop = False
		else:
			title = (raw_input('Different spreadsheet? y/n: ')).lower()
			if title == 'y':
				title = raw_input('Spreadsheet Title: ')
				doc = title

# if script is being run as a standalone application, its name attribute is __main__
if __name__ == '__main__':
	# execute script
	main()
