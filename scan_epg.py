#!/usr/bin/python

import ConfigParser
import os
import MySQLdb
import datetime
import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#Globals
ConfigFile = 'scan_epg.conf'

class General_Settings(object):

	def __init__(self):
		self.KeyWordFile = ''
		self.ResultsFile = ''

class MediaPortal_Settings(object):

	def __init__(self):
		self.MPSQLServer = ''
		self.MPSQLUser = ''
		self.MPSQLPW = ''
		self.MPSQLDB = '' 
		self.MPWebInterface = ''
		
class Email_Settings(object):

	def __init__(self):
		self.username = ''
		self.password = ''
		recipients = []
		sender = ''
		subject = ''
		server = ''
		port = 25


		

def GetConfig():
	
	config = ConfigParser.ConfigParser()
	
	# Check config file exists and can be accessed, then open
	try:
		dir = os.path.dirname(os.path.abspath(__file__))
		filepath = dir + '/' + ConfigFile

		if not os.path.isfile(filepath):
			print ("Error - Missing Config File: %s" % (ConfigFile))
			raise IOError('Config file does not exist')
			
		config.read(filepath)
		
	except:
		print ("Error - Unable to access config file: %s" % (ConfigFile))
		exit()

	
	gen_settings = General_Settings()
	mp_settings = MediaPortal_Settings()
	email_settings = Email_Settings()
	
	#Get config		
	gen_settings.KeyWordFile = config.get('General', 'KeyWordFile')
	gen_settings.ResultsFile = config.get('General', 'ResultsFile')

	mp_settings.MPSQLServer = config.get('MediaPortal', 'MPSQLServer')
	mp_settings.MPSQLUser = config.get('MediaPortal', 'MPSQLUser')
	mp_settings.MPSQLPW = config.get('MediaPortal', 'MPSQLPW')
	mp_settings.MPSQLDB = config.get('MediaPortal', 'MPSQLDB')
	mp_settings.MPWebInterface = config.get('MediaPortal', 'MPWebInterface')

	email_settings.username = config.get('Email', 'username')
	email_settings.password = config.get('Email', 'password')
	email_settings.recipients = config.get('Email', 'recipients').split(',')
	email_settings.sender = config.get('Email', 'sender')
	email_settings.subject = config.get('Email', 'subject')
	email_settings.server = config.get('Email', 'server')
	email_settings.port = int(config.get('Email', 'port'))
	
		
	return gen_settings, mp_settings, email_settings

	

def GetListFromTextFile(config):

	kwfile = None
	
	#Open keyword file
	try:
		kwfile = open(config.KeyWordFile, 'r')
	except:
		print 'Unable to open keyword file'
		exit()
	
	lst = []
	
	for line in kwfile:
		lst.append(line.rstrip('\r''\n'))
		
	return lst


def GenerateSQLQuery(terms):
	
	query = 'SELECT * FROM (select * from program where <QUERY>) as tmp ORDER BY startTime ASC;'
	
	q = ''
	
	max = len(terms)
	
	for i in range(0, max):

		q += "title like '%%%s%%'"" or description like '%%%s%%'" % (terms[i], terms[i])
		
		if i < max-1:
			q += ' or '
	
	query = query.replace('<QUERY>', q)
	
	return query
	


def ConnectToSQL(config):
	
	d=None
	
	try:
		d = MySQLdb.connect(host=config.MPSQLServer,
							user=config.MPSQLUser,
							passwd=config.MPSQLPW,
							db=config.MPSQLDB)
	except:
		print "Unable to connect to SQL DB"
		exit()
	
	return d


	
def ExecuteSQLQuery(db, query):
	
	cur = db.cursor()
	
	cur.execute(query)
	
	return cur.fetchall()


	
def CheckTime(tyme):
		
	if tyme < datetime.datetime.now():
		return None
		
	return str(tyme)


	
def OrganiseInfo(table, db):
	
	dict = {}
	ChannelNames = {}
	
	
	for row in table:
		
		ID = row[0]
		channel_code = row[1]
		date = CheckTime(row[2])
		title = row[4]
		description = row[5]
		
		channel_name = None
		
		if channel_code in ChannelNames:
			channel_name = ChannelNames[channel_code]
		else:			
			isVisable = ExecuteSQLQuery(db, "select visibleInGuide from mptvdb.channel where idChannel = '%d'" % (channel_code))[0][0]
			
			if isVisable == '\x01':
				channel_name = ExecuteSQLQuery(db, "select displayName from mptvdb.channel where idChannel = '%d'" % (channel_code))[0][0]
				ChannelNames[channel_code] = channel_name
			else:
				ChannelNames[channel_code] = None
				
		
		if date and channel_name:
		
			if title in dict:
				tmp = dict[title]
				tmp.append({'Date': date, 'Channel': channel_name, 'ID': ID, 'Description': description})
				dict[title] = tmp
			
			else:
				dict[title] = [{'Date': date, 'Channel': channel_name, 'ID': ID, 'Description': description}]
	
	
	return dict

	
def HighLight(str, wordlist):
		
	for w in wordlist:
		
		start = str.lower().find(w.lower())
		end = start + len(w)
		
		if start != -1:
			tmp = str[:start]
			tmp += '<mark>'
			tmp += str[start:end]
			tmp += '</mark>'
			tmp += HighLight(str[end:], [w])

			str = tmp
			
	return str
	
	
def GenerateHTMLOutput(d, wordlist, config):

	html = '<html>\r\n<head>\r\n</head>\r\n<body>\r\n<font face="verdana">\r\n'

	for key in d:
	
		titleStr = HighLight(key, wordlist)
		descriptionStr = HighLight(d[key][0]['Description'], wordlist)
		
		html += "\t<b>%s</b><br>\r\n\t" % (titleStr)

		for episode in d[key]:
			
			idStr = str(episode['ID'])
			dateStr = str(episode['Date'])
			channelStr = episode['Channel']
			
			html += "<font size=\"2\"><a href=\"http://%s/Television/ProgramDetails?programId=%s\">%s %s</a></font> " % (config.MPWebInterface, idStr, dateStr, channelStr)

		html += "\r\n\t<br>%s<br><br>\r\n\r\n" % (descriptionStr)

	
	html += "</font>\r\n</body>\r\n</html>"
	
	return html
	
	
	
	
def WriteToTextFile(output, config):

	try:
		f = open(config.ResultsFile, 'w')
		f.write(output)

	except:
	
		print 'Unable to write to ResultsFile'
	
	

def SendEmail(MessageText, config):
	
	msg = MIMEMultipart('alternative')
	
	msg['To'] = ", ".join(config.recipients)
	msg['From'] = config.sender
	msg['Subject'] = config.subject
	
	msg.attach(MIMEText(MessageText, 'html'))
	
	try:
		s = smtplib.SMTP(config.server, config.port)
		s.login(config.username, config.password)
		s.sendmail(msg['From'], msg['To'], msg.as_string())
		
	except:
		print 'Unable to send email'
		
	finally:
		s.quit()
	


#MAIN
if __name__ == "__main__":
	
	#Get settings
	gen_config, mp_config, email_config = GetConfig()
	
	#Convert keyword text file into a List
	KeyWordList = GetListFromTextFile(gen_config)
	
	#Generate SQL query from keywords
	SQLQuery = GenerateSQLQuery(KeyWordList)

	#Connect to MySQL
	database = ConnectToSQL(mp_config)	
	
	#Get Tuple containing response from SQL query
	programs = ExecuteSQLQuery(database, SQLQuery)

	#Organise Tuple into a Dictionary
	TitleDict = OrganiseInfo(programs, database)

	#Generate some HTML to send as email
	HyperText = GenerateHTMLOutput(TitleDict, KeyWordList, mp_config)

	#Output HTML to text file
	WriteToTextFile(HyperText, gen_config)

	#Email HTML
	SendEmail(HyperText, email_config)
	
