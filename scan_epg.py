#!/usr/bin/python

"""

Takes a list of keywords
Scans Mediaportal Database for TV program titles and descriptions
containing those keywords
Sends list of those programs as email

"""

import MySQLdb
import ConfigParser
import os
import datetime
import collections
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#Globals
CONFIG_FILE = 'scan_epg.conf'

#Container to store general settings
GeneralSettings = collections.namedtuple(
    'GeneralSettings',
    'keyword_file results_file'
    )

#Container to store mediaportal mysql db connection settings
MediaPortalSettings = collections.namedtuple(
    'MediaPortalSettings',
    'mp_sql_server mp_sql_user mp_sql_pw mp_sql_db mp_web_interface'
    )

#Container to store email settings
EmailSettings = collections.namedtuple(
    'EmailSettings',
    'username password recipients sender subject server port'
    )


#Functions

def get_config():

    """Get config from config file and store settings into namedtuples"""

    config = ConfigParser.ConfigParser()

    # Check config file exists and can be accessed, then open
    try:
        this_dir = os.path.dirname(os.path.abspath(__file__))
        filepath = this_dir + '/' + CONFIG_FILE

        if not os.path.isfile(filepath):
            print "Error - Missing Config File: %s" % (CONFIG_FILE)
            raise IOError('Config file does not exist')

        config.read(filepath)

    except IOError:
        print "Error - Unable to access config file: %s" % (CONFIG_FILE)
        exit()


    #Get config

    gen_settings = GeneralSettings(
        keyword_file=config.get('General', 'KeyWordFile'),
        results_file=config.get('General', 'ResultsFile')
        )

    mp_settings = MediaPortalSettings(
        mp_sql_server=config.get('MediaPortal', 'MPSQLServer'),
        mp_sql_user=config.get('MediaPortal', 'MPSQLUser'),
        mp_sql_pw=config.get('MediaPortal', 'MPSQLPW'),
        mp_sql_db=config.get('MediaPortal', 'MPSQLDB'),
        mp_web_interface=config.get('MediaPortal', 'MPWebInterface')
        )

    email_settings = EmailSettings(
        username=config.get('Email', 'username'),
        password=config.get('Email', 'password'),
        recipients=config.get('Email', 'recipients').split(','),
        sender=config.get('Email', 'sender'),
        subject=config.get('Email', 'subject'),
        server=config.get('Email', 'server'),
        port=int(config.get('Email', 'port'))
        )

    return gen_settings, mp_settings, email_settings



def get_list_from_text_file(config):

    """
    Get list of keywords from text file
    Convert keyword list into a List
    """

    kwfile = None

    #Open keyword file
    try:
        kwfile = open(config.keyword_file, 'r')
    except IOError:
        print 'Unable to open keyword file'
        exit()

    lst = []

    for line in kwfile:
        lst.append(line.rstrip('\r''\n'))

    return lst


def generate_sql_query(terms):

    """return a SQL query from list of keywords"""

    num_terms = len(terms)

    query_contents = ''

    max = len(terms)
    
    for i in range(0, max):

        query_contents += ("title like '%%%s%%'"" or description like '%%%s%%'"
                           % (terms[i], terms[i]))

        if i < num_terms-1:
            query_contents += ' or '

    query = ("SELECT * FROM (select * from program where %s)"
             "as tmp ORDER BY startTime ASC;" % query_contents)

    return query



def connect_to_sql(config):

    """
    Connect to sql server
    Return connection
    """

    connection = None

    try:
        connection = MySQLdb.connect(host=config.mp_sql_server,
                                     user=config.mp_sql_user,
                                     passwd=config.mp_sql_pw,
                                     db=config.mp_sql_db)
    except:
        print "Unable to connect to SQL DB"
        exit()

    return connection



def execute_sql_query(d_base, query):

    """
    Execute a SQL query against the server and return the result as a Tuple
    """

    cur = d_base.cursor()

    cur.execute(query)

    return cur.fetchall()



def check_time(tyme):

    """Return None if argument is older than current time"""

    if tyme < datetime.datetime.now():
        return None

    return str(tyme)



def organise_info(table, d_base):

    """
    Function to clean up Tuple returned from DB query
    and return as a dictionary

    Remove programs from channels which are disabled
    Remove programs that started in the past
    Replace channel code with name

    Args:
    Tuple of all returned results from DB query
    Database connection

    Returns:
    Dictionary
    """

    program_list = {}
    channel_names = {}


    for row in table:

        program_id = row[0]
        channel_code = row[1]
        date = check_time(row[2])
        title = row[4]
        description = row[5]

        channel_name = None

        if channel_code in channel_names:
            channel_name = channel_names[channel_code]

        else:
            is_visable = (execute_sql_query(
                d_base,
                "select visibleInGuide from mptvdb.channel where idChannel = '%d'"
                % (channel_code))[0][0])

            if is_visable == '\x01':
                channel_name = (execute_sql_query(
                    d_base,
                    "select displayName from mptvdb.channel where idChannel = '%d'"
                    % (channel_code))[0][0])

                channel_names[channel_code] = channel_name

            else:
                channel_names[channel_code] = None


        if date and channel_name:

            if title in program_list:
                tmp = program_list[title]
                tmp.append(
                    {'Date': date, 'Channel': channel_name,
                     'ID': program_id, 'Description': description}
                    )

                program_list[title] = tmp

            else:
                program_list[title] = [
                    {'Date': date, 'Channel': channel_name,
                     'ID': program_id, 'Description': description}
                    ]

    return program_list


def highlight(html_str, wordlist):

    """Highlight words in HTML"""

    for word in wordlist:

        start = html_str.lower().find(word.lower())
        end = start + len(word)

        if start != -1:
            tmp = html_str[:start]
            tmp += '<mark>'
            tmp += html_str[start:end]
            tmp += '</mark>'
            tmp += highlight(html_str[end:], [word])

            html_str = tmp

    return html_str


def generate_html_output(program_dict, wordlist, config):

    """Convert dictionary into HTML for viewing"""

    html = '<html>\r\n<head>\r\n</head>\r\n<body>\r\n<font face="verdana">\r\n'

    for key in program_dict:

        title_str = highlight(key, wordlist)
        description_str = highlight(program_dict[key][0]['Description'], wordlist)

        html += "\t<b>%s</b><br>\r\n\t" % (title_str)

        for episode in program_dict[key]:

            id_str = str(episode['ID'])
            date_str = str(episode['Date'])
            channel_str = episode['Channel']

            html += ("<font size=\"2\">"
                     "<a href=\"http://%s/Television/ProgramDetails?programId=%s\">"
                     "%s %s</a></font> "
                     % (config.mp_web_interface, id_str, date_str, channel_str))

        html += "\r\n\t<br>%s<br><br>\r\n\r\n" % (description_str)


    html += "</font>\r\n</body>\r\n</html>"

    return html




def write_to_text_file(output, config):

    """Write output to text file"""

    try:
        output_file = open(config.results_file, 'w')
        output_file.write(output)

    except IOError:

        print 'Unable to write to results_file'



def send_email(message_contents, config):

    """Send output as email"""

    msg = MIMEMultipart('alternative')

    msg['To'] = ", ".join(config.recipients)
    msg['From'] = config.sender
    msg['Subject'] = config.subject

    msg.attach(MIMEText(message_contents, 'html'))

    try:
        smtp_email = smtplib.SMTP(config.server, config.port)
        smtp_email.login(config.username, config.password)
        smtp_email.sendmail(msg['From'], msg['To'], msg.as_string())

    except:
        print 'Unable to send email'

    finally:
        smtp_email.quit()



#MAIN
if __name__ == "__main__":

    #Get settings
    GEN_CONFIG, MP_CONFIG, EMAIL_CONFIG = get_config()

    #Convert keyword text file into a List
    KEYWORDLIST = get_list_from_text_file(GEN_CONFIG)

    #Generate SQL query from keywords
    SQL_QUERY = generate_sql_query(KEYWORDLIST)

    #Connect to MySQL
    DATABASE = connect_to_sql(MP_CONFIG)

    #Get Tuple containing response from SQL query
    PROGRAMS = execute_sql_query(DATABASE, SQL_QUERY)

    #Organise Tuple into a Dictionary
    TITLE_DICT = organise_info(PROGRAMS, DATABASE)

    #Generate some HTML to send as email
    HYPERTEXT = generate_html_output(TITLE_DICT, KEYWORDLIST, MP_CONFIG)

    #Output HTML to text file
    write_to_text_file(HYPERTEXT, GEN_CONFIG)

    #Email HTML
    send_email(HYPERTEXT, EMAIL_CONFIG)
