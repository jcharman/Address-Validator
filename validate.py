#Build addresses from house numbers and postcodes stored in Google Sheets
#Written by Jake Charman

import pip, requests, json, os, sys
def setup():
	#Print program information.
	print('##############################')
	print('# Address Validator 0.1      #')
	print('# By Jake Charman            #')
	print('# www.jakecharman.co.uk      #')
	print('# software@jakecharman.co.uk #')
	print('##############################')
	#Import required libraries.
	print('Importing libraries')
	try:
		import gspread
	except(ImportError):
		print("Failed to import gspread. Installing")
		pip.main(['install', 'gspread'])
		import gspread
		setup.gspread = gspread

	try:
		from oauth2client.service_account import ServiceAccountCredentials
	except(ImportError):
		print("Failed to import oauth2client. Installing")
		pip.main(['install', 'oauth2client'])
		from oauth2client.service_account import ServiceAccountCredentials

	try:
		globals()["tqdm"] = __import__("tqdm.tqdm")
	except(ImportError):
		print("Failed to import tqdm. Installing")
		pip.main(['install', 'tqdm'])
		globals()["tqdm"] = __import__("tqdm")


	#Define the getaddress.io API key.
	setup.apiKey = 'ngqPvrgD_E2CwjOAZ_I6zQ7705'

	#Create a Google Drive API client.
	print('Creating Google API client')
	scope = ['https://spreadsheets.google.com/feeds']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
	setup.client = gspread.authorize(creds)

def validateAddresses(sheetName):
		#Create an empty list.
		output = []

		print('Opeining sheet "%(sheet)s"'%{"sheet": (sheetName)})
		#Find the sheet by name.
		sheet = setup.client.open(sheetName).sheet1

		# Extract the required values from the sheet.
		print('Pulling data from sheet')
		postcodes = sheet.col_values(6)
		postcodes = [p for p in postcodes if p]
		addresses = sheet.col_values(5)

		print('Validating Addresses')
		#For each postcode, call the API and retrieve data.
		for i in tqdm.tqdm(range(1,len(postcodes))):
				#Omit blank cells.
				if postcodes[i] != "" or addresses[i] != "":
						#Take only the first word of the address string.
						number = addresses[i].split(' ', 1)[0]
						#Remove spaces from the postcode.
						postcode = postcodes[i].replace(" ", "")
						#Print the postcode.
						#print(postcode)
						#Query the API.
						response = requests.get("https://api.getaddress.io/v2/uk/%(postcode)s/%(number)s?api-key=%(apiKey)s"%{"postcode": (postcode), "number": (number), "apiKey": (setup.apiKey)})
						#Retrieve and format the data.
						try:
								data = (json.loads(response.text)['Addresses'][0].replace(", ", "\n"))
								data = os.linesep.join([s for s in data.splitlines() if s])
						except:
							data = ("BAD INPUT")

						#Print the address.
						#print(data)
						#Append the data to the list.
						output.append(data + "\n" + postcode)
		#Return the list of data.
		return (output)

def uploadOutput(sheetName, output):
	print('Opeining sheet "%(sheet)s"' % {"sheet": (sheetName)})
	#Find the sheet by name.
	sheet = setup.client.open(sheetName)

	print('Creating new worksheet "Validated Addresses"')
	#Create a new worksheet.
	worksheet = sheet.add_worksheet(title="Validated Addresses", rows=str(len(output)), cols="1")

	print("Opening worksheet")
	#Open the newly created worksheet.
	worksheet = sheet.worksheet("Validated Addresses")

	print("Adding values to worksheet")
	#Add the data to the newly created cells.
	for i in tqdm.tqdm(range(0, len(output))):
			worksheet.update_cell(i+1, 1, output[i])

#UI.
try:
	#Run the program.
	if sys.argv[1] == "-r":
			try:
				#Pass the second argument to the program provided it is a string.
				if isinstance(sys.argv[2], basestring):
					#Run setup
					setup()
					#Validate the addresses and upload them.
					uploadOutput(sys.argv[2], validateAddresses(sys.argv[2]))
				else:
					print("INVALID INPUT! Use -h for help")
			except IndexError:
				print("INVALID INPUT! Use -h for help")
		#Display help.
        elif sys.argv[1] == "-h":
				print("Usage: -r <sheet>")
				print(" ")
				print("INSTRUCTIONS FOR USE:")
				print("1) Share the Google Sheet with 'formatting@s6c-address-formatting.iam.gserviceaccount.com'")
				print("2) Run this program with the usage above")

except IndexError:
		print("INVALID INPUT! Use -h for help")

                        
			  
