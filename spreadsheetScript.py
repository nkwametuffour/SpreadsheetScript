#!/usr/bin/python
import csv_config
import gdata.spreadsheet.service
import gdata
import gdata.client
import gdata.docs.client
import gdata.spreadsheets.client
import getopt
import sys
import os
import datetime
import webbrowser
from email.MIMEText import MIMEText
import smtplib
from Crypto.Cipher import AES
import base64

# Spreadsheet Class
class SpreadsheetScript():
	
	CLIENT_ID = '498758732944.apps.googleusercontent.com'
	CLIENT_SECRET = 'A_KMQ3yGorVSuTT-3P7DuHRC'
	SCOPE = 'https://spreadsheets.google.com/feeds/'
	USER_AGENT = 'Spreadsheet'
	
	def __init__(self, email, password, src='Default'):
		f = open("editlog.txt","w")
		f.close()
		self.config = csv_config.Csv_config()
		#csv_config is the string of the file path of the configuration file
		self.spreadsheetDict = self.config.buildDictionary("config.csv")
		user = email
		pwd = password
		self.today = self.getDate()
		try:
			self.__create_clients(user, pwd, src)
		except Exception, e:
			print 'Login failed.',e
		
		self.sheet_key = ''
		self.wksht_id = ''	
		
	def __create_clients(self, user, pswd, src):
		self.client = gdata.spreadsheet.service.SpreadsheetsService()
		self.client.email = user
		self.client.password = pswd
		self.client.ProgrammaticLogin()
		
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
		# generate a random secret key
		secret = os.urandom(BLOCK_SIZE)
		# create a cipher object using the random secret
		cipher = AES.new(secret)
		# encoded string
		encoded = EncodeAES(cipher, string)
		return encoded
	
	def __decrypt(self, enc):
		BLOCK_SIZE = 32
		PADDING = '{'
		# generate a random secret key
		secret = os.urandom(BLOCK_SIZE)
		# create a cipher object using the random secret
		cipher = AES.new(secret)
		DecodeAES = lambda c, e: c.decrypt(base64.b64decode(e)).rstrip(PADDING)
		return DecodeAES(cipher, enc)
		
	def __store_cred(self, log_cred):
		os.system("mkdir "+os.environ['HOME']+"/.hide")
		os.system("touch tokens.txt")
		f = open(os.environ['HOME']+'/.hide/tokens.txt','w+')
		for i in range(len(log_cred)):
			f.write(self.__encrypt(log_cred[i]) + '\n')
		f.read()
		f.close()

	def getDate(self) :
		yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
		yesterday = yesterdat.strftime("%d-%b-%Y")
		return yesterday.lstrip('0')

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
	
	# upload a spreadsheet from from local machine, to Drive with format converted to Drive spreadsheet format
	def upload_spreadsheet(self, filepath, doctitle):
		document = gdata.docs.data.Resource(type='spreadsheet', title=doctitle)
		media = gdata.data.MediaSource()
		media.SetFileHandle(filepath, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
		create_uri = gdata.docs.client.RESOURCE_UPLOAD_URI + '?convert=true'
		upload_doc = self.docs_client.CreateResource(document, create_uri=create_uri, media=media)
	
	# add worksheet with specified size to current spreadsheet document
	def addWorksheet(self, name, row_size=20, col_size=20):
		if name in self.getWorksheetTitles(self.sheet_key):
			print "Please choose a unique worksheet name\n"
		else:
			self.client.AddWorksheet(name, row_size, col_size, self.sheet_key)
			print 'Worksheet added.'
	
	# delete an entire worksheet from the current spreadsheet
	def deleteWorksheet(self, workshts):
		for wkid in workshts:
			worksht_id = self.getWorksheetIdByName(wkid)
			if worksht_id != None:
				worksheet = self.client.GetWorksheetsFeed(self.sheet_key, worksht_id)
				self.client.DeleteWorksheet(worksheet_entry = worksheet)
				print 'Worksheet deleted.'
			else:
				print "Worksheet not found."
	
	#Prints the data in the worksheet
	def printData(self) :
		feed = self.client.GetListFeed(self.sheet_key, self.wksht_id)
		# print field titles of data
		for row in feed.entry :
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
		
	# function to exit program if there self.sheet_key is an empty string
	def exit_if_no_key(self):
		if self.sheet_key == '':
			print "No Spreadsheet selected. Exiting..."
			sys.exit(2)
	
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
		
	def getWorksheetIdByName(self, name):
		wk_feed = self.client.GetWorksheetsFeed(self.sheet_key)
		for wksht in wk_feed.entry:
			if name == wksht.title.text:
				return wksht.id.text.rsplit('/', 1)[1]
		return None	# worksheet with title, name, not found

	def deleteRecord(self, rows):
		cnfrm = raw_input('Confim deleting record '+str(rows)+' (y/n): ')
		if cnfrm.lower() == 'y':
			for row in rows:
				row = int(row)
				feed = self.client.GetListFeed(self.sheet_key, self.wksht_id)
				self.client.DeleteRow(feed.entry[row-1]) # user enters from 1, but records are numbered from 0
			print 'Record delete successful'
		else:
			print 'Record delete unsuccessful'
			
	def getOperationColumnNumber(self, operation, code):
		col_number = self.config.searchDOD(operation, code, self.spreadsheetDict)
		return str(col_number)
	#took out the string variable that was below
	#its replaced by self.today			
	def getRowNumber(self):
		row_entry = self.client.GetListFeed(self.sheet_key, self.wksht_id)
		row_ct = 2
		for entry in row_entry.entry:
			if self.today == entry.title.text:
				return row_ct
			row_ct += 1
			
	#Sends mail to the user
	def sendMail(self, success = True):
		with open("editlog.txt","rb") as text :
			msg = MIMEText(text.read())
			
		msg['From'] = "rancardinterns2013@gmail.com"
		msg['To'] = self.client.email
		msg['Subject'] = "Update on SpreadSheet "
		msg.preamble = "Edit Log"	
		server = smtplib.SMTP()
		server.connect('smtp.gmail.com', 587)
		server.ehlo()
		server.starttls()
		server.ehlo()
		server.login("####@gmail.com","password")
		message = msg.as_string()
		server.sendmail('#from_email', self.client.email, message)
		os.remove("editlog.txt")
				
	#Takes in the document name, checks if it exists and asks the user for the worksheet to work with.
	def flow(self):
#		doc = docmnt
		con = True
		while con == True:
			command = raw_input('>>> ')
			if command[0] == 'i':
				if command[1:len('CellVal')+1].lower() == 'CellVal'.lower():
					val = command[command.find('(')+1:command.find(')')]
					val = val.split(';')
					try:
						self.updateCell(val)
					except :	
						with open("editlog.txt","a") as log :
							log.write("\nInserting of Values into cell was Unsuccessful")
					else :
						with open("editlog.txt","a") as log :
							log.write("\nInserting of Values into cell was Successful")
					
				elif command[1:len('RowVal')+1].lower() == 'RowVal'.lower():
					val = command[command.find('(')+1:command.find(')')]
					val = val.split(';')
					
					try:
						self.updateRow(val)
					except :	
						with open("editlog.txt","a") as log :
							log.write("\nInserting of Values into rows was Unsuccessful")
					else :
						with open("editlog.txt","a") as log :
							log.write("\nInserting of Values into rows was Successful")
				elif command[1:len('ColVal')+1].lower() == 'ColVal'.lower():
					val = command[command.find('(')+1:command.find(')')]
					val = val.split(';')
					
					try:
						self.updateCol(val)
					except :	
						with open("editlog.txt","a") as log :
							log.write("\nInserting of Values into columns was Unsuccessful")
					else :
						with open("editlog.txt","a") as log :
							log.write("\nInserting of Values into columns was Successful")
				else:
					print 'Cannot find command: '+command
			elif command[0] == 'd':
				if command[1:len('CellVal')+1].lower() == 'CellVal'.lower():
					val = command[command.find('(')+1:command.find(')')]
					val = val.split(';')
					
					try:
						self.deleteCellValue(val)
					except :	
						with open("editlog.txt","a") as log :
							log.write("\nDeleting of Values in Cells was Unsuccessful")
					else :
						with open("editlog.txt","a") as log :
							log.write("\nDeleting of Values in Cells was Successful")
				elif command[1:len('RowVal')+1].lower() == 'RowVal'.lower():
					val = command[command.find('(')+1:command.find(')')]
					val = val.split(';')
										
					try:
						self.deleteRowValues(val)
					except :	
						with open("editlog.txt","a") as log :
							log.write("\nDeleting of Values in Rows was Unsuccessful")
					else :
						with open("editlog.txt","a") as log :
							log.write("\nDeleting of Values in Rows was Successful")
				elif command[1:len('ColVal')+1].lower() == 'ColVal'.lower():
					#val = command[command.find('(')+1:command.find(')')]
					#val = val.split(';')
					#self.deleteColValues(docmnt, val)
					pass
				elif command[1:len('Row')+1].lower() == 'Row'.lower():
					val = command[command.find('(')+1:command.find(')')]
					val = val.split(';')
					
					try:
						self.deleteRecord(val)
					except :	
						with open("editlog.txt","a") as log :
							log.write("\nDeleting of Record was Unsuccessful")
					else :
						with open("editlog.txt","a") as log :
							log.write("\nDeleting of Record was Successful")

				elif command[1:len('WS')+1].lower() == 'WS'.lower():
					val = command[command.find('(')+1:command.find(')')]
					val = val.split(';')
					
					try:
						self.deleteWorksheet(val)
					except :	
						with open("editlog.txt","a") as log :
							log.write("\nDeleting of Worksheet was Unsuccessful")
					else :
						with open("editlog.txt","a") as log :
							log.write("\nDeleting of Worksheet was Successful")
				elif command[1:len('SS')+1].lower() == 'SS'.lower():
					val = command[command.find('(')+1:command.find(')')]
					val = val.split(';')
					
					try:
						self.deleteSpreadsheet(val)
					except :	
						with open("editlog.txt","a") as log :
							log.write("\nDeleting of Spreadsheet was Unsuccessful")
					else :
						with open("editlog.txt","a") as log :
							log.write("\nDeleting of Spreadsheet was Successful")
				else:
					print 'Cannot find command: '+command
			elif command[0] == 'c':
				if command[1:len('WS')+1].lower() == 'WS'.lower():
					val = command[command.find('(')+1:command.find(')')].strip()
					#val = val.split(';')
					client.wksht_id = client.getWorksheetIdByName(val)
				elif command[1:len('SS')+1].lower() == 'SS'.lower():
					val = command[command.find('(')+1:command.find(')')].strip()
					#val = val.split(';')
					client.sheet_key = client.getSpreadsheetKey(val)
					client.exit_if_no_key()
				elif command[0:len('clear')+1].lower() == 'clear'.lower():
					os.system("clear")
				else:
					print 'Cannot find command: '+command
			elif command[0] =='n':
				if command[1:len('WS')+1].lower() == 'WS'.lower():
					val = command[command.find('(')+1:command.find(')')].strip()
					val = val.split(',')
					
					try:
						if len(val) == 3:
							self.addWorksheet(val[0], val[1], val[2])
						else:
							self.addWorksheet(val[0])
					except :	
						with open("editlog.txt","a") as log :
							log.write("\nCreation of new worksheet was Unsuccessful")
					else :
						with open("editlog.txt","a") as log :
							log.write("\nCreation of new worksheet was Successful")
				elif command[1:len('SS')+1].lower() == 'SS'.lower():
					val = command[command.find('(')+1:command.find(')')].strip()
					
					try:
						self.createSpreadsheet(val)
					except :	
						with open("editlog.txt","a") as log :
							log.write("\nCreation of new spreadsheet was Unsuccessful")
					else :
						with open("editlog.txt","a") as log :
							log.write("\nCreation of new spreadsheet was Successful")
				else:
					print 'Cannot find command: '+command
			else:
				if command[0:len('print')+1].lower() == 'print'.lower():
					self.printData()
				elif command[0:len('help')+1].lower() == 'help'.lower():
					self.getHelp()
				elif command[0:len('exit')+1].lower() == 'exit'.lower():
					self.sendMail()
					sys.exit(2)
				else:
					print 'Cannot find command: '+command
		
	def updateCell(self, cellAndVal):
		#Overwrites the value in the cell specified with new_value
		for each in cellAndVal:
			cell = each.split(',')
			self.client.UpdateCell(row = int(cell[0]), col = int(cell[1]), inputValue = cell[2],  key = self.sheet_key, wksht_id = self.wksht_id )
	
	def updateRow(self, rowAndVal):
		#Overwrites the values in the row with the given values
		for i in range(len(rowAndVal)):
			row = int(rowAndVal[i][0])
			for h in range(1, len(rowAndVal[i])):
				self.client.UpdateCell(row = row, col = i, inputValue = rowAndVal[i][h], key = self.sheet_key, wksht_id = self.wksht_id )

	def updateCol(self, colAndVal):
		#Overwrites the values in the column with the given values
		for i in range(len(colAndVal)):
			col = int(colAndVal[i][0])
			for h in range(1, len(colAndVal[i])):
				self.client.UpdateCell(row = i+1, col = col, inputValue = colAndVal[i][h], key = self.sheet_key, wksht_id = self.wksht_id )
	
	def deleteCellValue(self, cells):
		#Puts an empty string in the specified cell
		for cell in cells:
			cell = cell.split(',')
			self.client.UpdateCell(row = int(cell[0]), col = int(cell[1]), inputValue = None, key = self.sheet_key, wksht_id = self.wksht_id )
	
	# prints script documentation
	@staticmethod
	def getHelp():
		print """SPREADSHEET SCRIPT


NAME
	spreadsheetScript - a Python script which interacts with Google Spreadsheets.

USAGE
	python spreadsheetScript [OPTIONS]

OVERVIEW
	spreadsheetScript script enables users access to their Google Drive accounts, where they can have access to spreadsheet documents
	in their Drive. Users can also create and delete spreadsheets, create and delete worksheets, insert and delete values in the 
	spreadsheet cells.

OPTIONS
Generic Script Information
	--help
		Displays this documentation of the script, and then exits
	
Main Options
	--user
		Takes in the email of the user.
	--pwd
		Takes in the password of the user.
	--src
		Provide the spreadsheet source name (like a project name, not spreadsheet name).
	--docName
		Enter the title of an existing spreadsheet document to work with.
	--worksheet
		Takes the title of a worksheet in a particular spreadsheet. If not specified, the first worksheet is used.
	--nSS
		Creates a new spreadsheet document and names it with the string provided after --nSS.
	--nWS
		Creates a new worksheet in the current spreadsheet and names it with the string provided after --nWS.
	--dSS
		Remove the spreadsheet with the title provided, from Drive.
	--dWS
		Remove the worksheet with the title provided, from the current spreadsheet.
	--iRowVal
		Inserts values into a row; cell after cell, specified by row number, column number, and values as comma-separated values.
		Multiple insertions are distinguished by ; separation.
	--iColVal
		Inserts values into a column; cell after cell, specified by row number, column number, and values as comma-separated values.
		Multiple insertions are distinguished by ; separation.
	--iCellVal
		Inserts a value in a cell, specified by row number, column number, and value as comma-separated values.
		Multiple insertions are distinguished by ; separation.
	--print
		Does not take any character after it, if specified will print the contents of the document name provided.
	--help
		Prints this screen
	--exit
		Exits script execution after all commands specified in the initial command string have been executed.
		Enables one time script execution.
	"""
	
def main():
	user = False
	pwd = False
	insert = False
	src = False
	docName = False
	worksheet = False
	prnt = False
	hlp = False
	#iRow = False
	iRowVal = False
	iColVal = False
	iCellVal = False
	dRowVal = False
	dColVal = False
	dCellVal= False
	dRow = False
	dWS = False
	dSS = False
	#cWS = False
	#cSS = False
	nSS = False
	nWS = False
	ext = False
	
	# check if user has entered the correct options
	try:
		opts, args = getopt.getopt(sys.argv[1:], "", ["user=", "pwd=", "insert=", "src=", "docName=", "worksheet=", "print", "help", "iRowVal=", "iColVal=", \
		"iCellVal=", "dCellVal=", "dRow=", "dWS=", "dSS=", "nSS=", "nWS=", "exit"])
	except getopt.GetoptError, e:
		print "python spreadsheetScript.py --help. For help:", e, "\n"
		sys.exit(2)
	
	
	# get user and pwd values provided by user into their respective variables	
	worksheetVal = 0
	srcVal = "Default"
	for opt, val in opts:
		if opt == "--user":
			user = True
			userVal = val
		elif opt == "--pwd":
			pwd = True
			pwdVal = val
		elif opt == "--insert":
			insert = True
			insertVal = val
		if opt == "--src":
			src = True
			srcVal = val
		elif opt == "--docName":
			docName = True
			docNameVal = val
		elif opt == "--worksheet":
			worksheet = True
			worksheetVal = val
		elif opt == "--print":
			prnt = True
		elif opt == "--help":
			hlp = True
		elif opt == "--exit":
			ext = True
		elif opt == "--nWS":
			nWS = True
			nWSVal = val
		elif opt == "--nSS":
			nSS = True
			nSSVal = val
		elif opt == "--iRowVal":
			iRowVal = True
			iRowValVal = val
		elif opt == "--iColVal":
			iColVal = True
			iColValVal = val
		elif opt == "--iCellVal":
			iCellVal = True
			iCellValVal = val
		elif opt == "--dRowVal":
			dRowVal = True
			dRowValVal = val
		elif opt == "--dColVal":
			dColVal = True
			dColValVal = val
		elif opt == "--dCellVal":
			dCellVal = True
			dCellValVal = val
		elif opt == "--dRow":
			dRow = True	
			dRowVal = val
		elif opt == "--dWS":
			dWS = True	
			dWSVal = val
		elif opt == "--dSS":
			dSS = True	
			dSSVal = val
		
	
	#worksheet = True
	
	if hlp == True:
		SpreadsheetScript.getHelp()
		sys.exit(0)
	else:
		if user == False or pwd == False:
			print "python spreadsheetScript.py --user email --pwd password"
			print "python spreadsheetScript.py --help"
			sys.exit(0)
		if docName == False and nSS == False:
			print "You have to specify a document or create a new Spreadsheet to work with"
		if nSS == True and docName == False:
			docNameVal = nSSVal
		if src == True:
			if nSS == False:
				print "You have to specify a new document"
				SpreadsheetScript.getHelp()
				sys.exit()
	
	client = SpreadsheetScript(userVal, pwdVal, srcVal)

	log = open("editlog.txt","a+")
	
	if nSS == True:
		client.createSpreadsheet(nSSVal)
	client.sheet_key = client.getSpreadsheetKey(docNameVal)
	client.exit_if_no_key()
	if nWS == True:
		nWSVal = nWSVal.split(',')
		try:
			if len(nWSVal) == 3:
				client.addWorksheet(nWSVal[0], nWSVal[1], nWSVal[2])
			else:
				client.addWorksheet(nWSVal[0])
		except :	
			log.write("\nCreation of new worksheet was Unsuccessful")
		else :
			log.write("\nCreation of new worksheet was Successful")
			
		
	if worksheet == True :
		client.wksht_id = client.getWorksheetIdByName(worksheetVal)
	elif nWS == True:
		client.wksht_id = client.getWorksheetIdByName(nWSVal)
	else:
		client.wksht_id = client.selectWorksheet(client.sheet_key, worksheetVal)
		
	if insert == True and docName == True:
		row = str(client.getRowNumber())
		set_of_values = insertVal.split(";")
		valueToPass = []
		for each_set in set_of_values :
			spl_val = each_set.split(",")
			col = str(client.getOperationColumnNumber(spl_val[0], spl_val[1]))
			valueToPass.append(row+","+col+","+str(spl_val[2]))
		
		client.updateCell(valueToPass)
	
	if iRowVal == True:
		iRowValVal = iRowValVal.split(';')
		val = []
		try:
			for i in range(len(iRowValVal)):
				h = iRowValVal[i].split(',')
				val.append(h)
			client.updateRow(val)
		except :	
			log.write("\nRow Update was Unsuccessful")
		else :
			log.write("\Row Update was Successful")	
	if iColVal == True:
		iColValVal = iColValVal.split(';')
		val = []
		try:
			for i in range(len(iColValVal)):
				h = iColValVal[i].split(',')
				val.append(h)
			client.updateRow(val)
		except :	
			log.write("\nColumn Update was Unsuccessful")
		else :
			log.write("\Column Update was Successful")		
	if iCellVal == True:
		iCellValVal = iCellValVal.split(';')
		try:
			client.updateCell(iCellValVal)
		except :	
			log.write("\nCell Updates was Unsuccessful")
		else :
			log.write("\Cell Updates was Successful")		
	if dRow == True:
		delRowVal = delRowVal.split(';')
		try:
			client.deleteRecord(delRowVal)
		except :	
			log.write("\nRecord DElete was Unsuccessful")
		else :
			log.write("\Record Delete was Successful")		
	if dRowVal == True:
		#dRowValVal = dRowValVal.split(';')
		#client.deleteRowValues(dRowValVal)
		pass
	if dColVal == True:
		#dColValVal = dColValVal.split(';')
		#client.deleteColValues(dRowValVal)
		pass
	if dCellVal == True:
		dCellValVal = dCellValVal.split(';')
		try:
			client.deleteCellValue(dCellValVal)
		except :	
			log.write("\nCell Delete was Unsuccessful")
		else :
			log.write("\Cell Delete was Successful")		
	if dSS == True:
		try:
			client.deleteSpreadsheet(dSSVal)
		except :	
			log.write("\nSpreadsheet Delete was Unsuccessful")
		else :
			log.write("\Spreadsheet Delete was Successful")	
	if dWS == True:
		val = []
		val.append(dWSVal)
		try:
			client.deleteWorksheet(val)
		except :	
			log.write("\nWorksheet Delete was Unsuccessful")
		else :
			log.write("\Worksheet Delete was Successful")	
	if prnt == True:
		client.printData()
	if ext == True:
		client.sendMail()
		sys.exit(0)

	if ext == False:
		client.flow()
		
# if script is being run as a standalone application, its name attribute is __main__
if __name__ == '__main__':
	# execute script
	main()
