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
			user = raw_input('Enter username: ').strip()
			pwd = raw_input('Enter password: ').strip()
			# validate user and password
			try:
				self.__create_clients(user, pwd, src)
			except Exception, e:
				print 'Login failed.',e
				
			#self.__store_cred([user, pwd])
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
		try:
			token.get_access_token(code)
		except :
			print "Access Denied"
			sys.exit(0)	
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

	def deleteRecord(self, rows):
		#cnfrm = raw_input('Confim deleting record '+str(row)+' (y/n): ')
		if cnfrm.lower() == 'y':
			for row in rows:
				row = int(row)
				feed = self.client.GetListFeed(self.sheet_key, self.wksht_id)
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
			sys.exit(0)	# may be taken out, we may ask user if they want to create a new spreadsheet or work with existing spreadsheet
		
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
		try:
			self.wksht_id = self.selectWorksheet(self.sheet_key, wkid-1)
		except IndexError:
			print "That index is out of range. Please input again"
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
			if choice.strip() == 'row':
				row = input('Enter row number: ')
				values = raw_input('Enter values in order (Example 1,2,3): ')
				values = values.split(',')
				self.updateRow(row, values)
				self.printData()
			elif choice.strip() == 'col':
				col = input('Enter column number: ')
				values = raw_input('Enter values in order (Example 1,2,3): ')
				values = values.split(',')
				self.updateCol( col, values)
				self.printData()
			elif choice.strip() == 'cell':
				cell = raw_input('Enter row,column,value in that order: ')
				#value = raw_input('Cell value: ')
				cell = cell.split(',')
				self.updateCell(cell[0], cell[1], cell[2])
				self.printData()
		
		if delVal:
			cell = raw_input("Enter the cell's row, column in that order.(Example 2,3): ")
			cell = (cell.strip()).split(',')
			sef.deleteCellValue(docmnt, cell[0], cell[1], wkid-1)
				
		
	def updateCell(self, cellAndVal):
		#Overwrites the value in the cell specified with new_value
		for each in cellAndVal:
			cell = each.split(',')
			self.client.UpdateCell(row = int(cell[0]), col = int(cell[1]), inputValue = cell[2],  key = self.sheet_key, wksht_id = self.wksht_id )
	
	def updateRow(self, rowAndVal):
		#Overwrites the values in the row with the given values
		for i in range(len(rowAndVal)):
			row = int(rowAndVal[i[0]])
			for h in range(1, len(rowAndVal[i])):
				self.client.UpdateCell(row = row, col = i, inputValue = rowAndVal[i[h]], key = self.sheet_key, wksht_id = self.wksht_id )

	def updateCol(self, colAndVal):
		#Overwrites the values in the column with the given values
		for i in range(len(colAndVal)):
			col = int(colAndVal[i[0]])
			for h in range(1, len(colAndVal[i])):
				self.client.UpdateCell(row = i+1, col = col, inputValue = colAndVal[i[h]], key = self.sheet_key, wksht_id = self.wksht_id )
	
	def deleteCellValue(self, cells):
		#Puts an empty string in the specified cell
		for cell in cells:
			cell = cell.split(',')
			self.client.UpdateCell(row = int(cell[0]), col = int(cell[0]), inputValue = None, key = self.sheet_key, wksht_id = self.wksht_id )
	
	def deleteRowValues(self, rows):
		#Puts an empty string in the cells on the specified row
		for row in rows:
			row = int(row)
			list_of_values=self.worksheet.row_values(row)
			#print list_of_values
			for i in range(1,len(list_of_values)):
				self.worksheet.update_cell(row,i,"")
				self.client.UpdateCell(row = row, col = i, inputValue = None, key = self.sheet_key, wksht_id = self.wksht_id )
	#on hold for now
	'''	def deleteColValues(self, docName, cols, wks = 0):
		#Puts an empty string in the cells on the specified column
		for col in cols:
			col = int(col)
			list_of_values=self.worksheet.col_values(col)
			#print list_of_values
			for i in range(1,len(list_of_values)):
				self.worksheet.update_cell(i,col,"")
	'''			
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
	src = False
	docName = False
	worksheet = False
	prnt = False
	hlp = False
	new = False
	rmv = False
	inRow = False
	inCol = False
	inCell = False
	delRow = False
	delRowVal = False
	delColVal = False
	delCellVal= False
	
	# check if user has entered the correct options
	try:
		opts, args = getopt.getopt(sys.argv[1:], "", ["src=", "docName=", "worksheet=", "print", "help", "new=", "rmv=", "inRow=", "inCol=", "inCell=", "delRow=", "delRowVal=", "delColVal=", "delCellVal=", ])
	except getopt.GetoptError, e:
		print "python spreadsheetScript.py --help. For help:", e, "\n"
		sys.exit(2)
	
	
	# get user and pwd values provided by user into their respective variables	
	worksheetVal = 0
	srcVal = "Default"
	for opt, val in opts:
		if opt == "--src":
			src = True
			srcVal = val
		elif opt == "--docName":
			docName = True
			docNameVal = val
		elif opt == "--worksheet":
			worksheet = True	# delete option sets row to delete to (row, column)
			try:
				worksheetVal = int(val)
			except:
				print "--worksheet accepts only integers"
				SpreadsheetScript.getHelp()
				sys.exit()
		elif opt == "--print":
			prnt = True	# print option set to true, if the option is added
		elif opt == "--help":
			hlp = True	# help option set to true, if the option is added
		elif opt == "--new":
			new = True
			newVal = val
		elif opt == "--rmv":
			rmv = True	
			rmvVal = val # title of document to remove
		elif opt == "--inRow":
			inRow = True
			inRowVal = val
		elif opt == "--inCol":
			inCol = True
			inColVal = val
		elif opt == "--inCell":
			inCell = True
			inCellVal = val
		elif opt == "--delRow":
			delRow = True
			delRowVal = val
		elif opt == "--delRowVal":
			delRowVal = True
			delRowValVal = val
		elif opt == "--delColVal":
			delColVal = True
			delColValVal = val
		elif opt == "--delCellVal":
			delCellVal = True
			delCellValVal = val
	
	
	if hlp == True:
		SpreadsheetScript.getHelp()
		sys.exit(0)
	else:
		if src == True:
			if not (docName == True and worksheet == True):
				print "You have to specify a document name and worksheet"
				SpreadsheetScript.getHelp()
				sys.exit()
		if worksheet == True or prnt==True or inRow==True or inCol==True or inCell==True or delRow==True or delRowVal==True or delColVal==True or delCellVal==True:
			if docName==False:
				print "Please specify a document Title"
				SpreadsheetScript.getHelp()
				sys.exit()
			
	if new == True and docName == False:
		docNameVal = newVal
	
	client = SpreadsheetScript(srcVal)
	
	
	if prnt == True:
		client.printData()
	if new == True:
		client.createSpreadsheet(newVal)
	if rmv == True:
		client.deleteSpreadsheet(rmvVal)
	if inRow == True:
		pass
	if inCol == True:
		pass
	if inCell == True:
		inCellVal = inCellVal.split(';')
		client.updateCell(docNameVal, inCellVal, worksheetVal)
	if delRow == True:
		delRowVal = delRowVal.split(';')
		client.deleteRecord(delRowVal)
	if delRowVal == True:
		pass
	if delColVal == True:
		pass
	if delCellVal == True:
		inCellVal == inCellVal.split(';')
		client.deleteCellValue(docNameVal, inCellVal, worksheetVal)
		
# if script is being run as a standalone application, its name attribute is __main__
if __name__ == '__main__':
	# execute script
	main()
