import os
import re
import email
import imaplib
import urllib.request

def Main():
	## Define user credentials
	username = "USERNAME"
	password = "PASSWORD"

	## Define mail server
	server = "Outlook.Office365.com"

	## Define service
	service = "SERVICE"

	## Create folder for service if necessary
	folderPath = "FOLDERPATH"
	folderPath = os.path.join(folderPath,service,"")
	if not os.path.isdir(folderPath):
		os.mkdir(folderPath)

	## Define providers
	provs = ["GENERIC_SERVICE_PROVIDERS"]

	## Define years
	years = ["2020"]
	
	## Define months 1 - 12
	dates = [1,2,3,4,5,6,7,8,9,10]

	## Login to mail
	mail = login(username, password, server)

	## Find all mailboxes
	mailboxes = mailBoxes(mail)

	## Search for files in all mailboxes
	for mailbox in mailboxes:
		mail.select(mailbox)
		search(mail, years, dates, provs, folderPath)

	## Logout
	mail.logout()

def login(username, password, server):
	mail = imaplib.IMAP4_SSL(server)
	mail.login(username,password)
	return mail

def mailBoxes(mail):
	mailboxes = []
	for i in mail.list()[1]:
		l = i.decode().split(' "/" ')
		mailboxes.append(str(l[1])))
	return mailboxes

def search(mail, years, dates, provs, folderPath):
	for prov in provs:
		## Create provider folder if necessary
		provPath = os.path.join(folderPath,prov,"")
		if not os.path.isdir(provPath):
			os.mkdir(provPath)

		## Iterate through given years
		for year in years:
			
			## Create months dictionary with current year
			months = {1:[f'1-Jan-{year} BEFORE 1-Feb-{year}','01.Jan'],
				2:[f'1-Feb-{year} BEFORE 1-Mar-{year}','02.Feb'],
				3:[f'1-Mar-{year} BEFORE 1-Apr-{year}','03.Mar'],
				4:[f'1-Apr-{year} BEFORE 1-May-{year}','04.Apr'],
				5:[f'1-May-{year} BEFORE 1-Jun-{year}','05.May'],
				6:[f'1-Jun-{year} BEFORE 1-Jul-{year}','06.Jun'],
				7:[f'1-Jul-{year} BEFORE 1-Aug-{year}','07.Jul'],
				8:[f'1-Aug-{year} BEFORE 1-Sep-{year}','08.Aug'],
				9:[f'1-Sep-{year} BEFORE 1-Oct-{year}','09.Sep'],
				10:[f'1-Oct-{year} BEFORE 1-Nov-{year}','10.Oct'],
				11:[f'1-Nov-{year} BEFORE 1-Dec-{year}','11.Nov'],
				12:[f'1-Dec-{year} BEFORE 1-Jan-{int(year)+1}','12.Dec']
				}

			## Create year folder if necessary
			yearPath = os.path.join(provPath,year,"")
			if not os.path.isdir(yearPath):
				os.mkdir(yearPath)

			## Iterate through given dates
			for date in dates:
				try:
					## Select month for folder name
					month = months[date][1]

					## Create necessary month folder, pdf and xml folders
					monthPath = os.path.join(yearPath,month,"")
					pdfPath = os.path.join(monthPath,'mail_pdf',"")
					xmlPath = os.path.join(monthPath,'mail_xml',"")

					if not os.path.isdir(monthPath):
						os.mkdir(monthPath)

					if not os.path.isdir(pdfPath):
						os.mkdir(pdfPath)

					if not os.path.isdir(xmlPath):
						os.mkdir(xmlPath)

					## Search all item uids in mailbox for current date and condition
					result, data = mail.uid('search', None, '(ALL SINCE ' +months[date][0] +' TEXT "'+prov+'")')
					data = data[0].decode("utf-8")
				
					## Create uids list with found items
					uids = data.split(" ")

					## If any uids are found, call getFiles() function to retrieve files in emails
					if uids[0] != '':
						getFiles(mail,uids,month,pdfPath,xmlPath)

				except Exception as e:
					print(e)

def getFiles(mail,uids,month,pdfPath,xmlPath):
	## Iterate through all uids
	for uid in uids:
		try:
			print("UID " + uid)
			## Fetch and decode email with specified uid
			result2, email_data = mail.uid('fetch', uid, '(RFC822)')
			raw_email = email_data[0][1].decode("utf-8")
			email_message = email.message_from_string(raw_email)

			## Iterate through email parts
			for part in email_message.walk():
				## Try to find a file
				fileName = part.get_filename()
				filePath = None
				if bool(fileName):
					print(fileName)
					## If a file was found and it is an xml or pdf, save into respective folder if it doesn't exist already
					if "xml" in fileName:
						filePath = os.path.join(xmlPath,fileName)
					elif "pdf" in fileName or "PDF" in fileName:
						filePath = os.path.join(pdfPath,fileName)

					if filePath != None:
						if not os.path.isfile(filePath):
							data = part.get_payload(decode=True)
							with open(filePath, 'wb') as f:
								f.write(data)

		except Exception as e:
			## Exception for emails with download link instead of file
			print(e)
			try:
				## Fetch email with specified uid
				result2, email_data = mail.uid('fetch', uid, '(RFC822)')

				## Search for specified link pattern in email data
				link_pattern = re.compile('<a[^>]+href=\"(https://cfdi.alerta.com.mx/api/GenerarComprobantes/generar/.*?)\">.*?</a>')
				search = link_pattern.findall(str(email_data[0][1]))

				for url in search:
					download = re.search("/[1-2]/1$", url)
					if download:
						response = urllib.request.urlopen(url)

						## Retrieve file name from response and give appropiate format
						fileName = (response.info()['Content-Disposition']).split('filename=')[1]
						fileName = fileName.replace('"','')
						filePath = None
						if "xml" in fileName:
							filePath = os.path.join(xmlPath,fileName)
						elif "pdf" in fileName or "PDF" in fileName:
							filePath = os.path.join(pdfPath,fileName)

						if filePath != None:
							if not os.path.isfile(filePath):
								data = response.read()
								with open(filePath, 'wb') as f:
									f.write(data)
			except Exception as e:
				print(e)

if __name__ == "__main__":
	Main()