
'''
Looks for any new reports (not existing in the database). Saves reports to xml (one report per file) and returns a list of their locations.
'''

import easygui
from pysimplesoap.client import SoapClient
import base64
from xml.etree import ElementTree as ET
import datetime
import os

class ScoreReport:

    def __init__(self, name, directory, loc, date_str):
        """
        A collection of attributes related to a score report.
        :param name: the file's name
        :param directory: the directory the file lives in
        :param loc: the file's full path
        :param date_str: the date of the report
        """

        self.name = name
        self.directory = directory      #the directory it lives in
        self.loc = loc                  #its full location or path
        self.date_str = date_str
        self.date = datetime.datetime.strptime(date_str, '%m/%d/%Y')
        self.age = (datetime.datetime.now() - self.date).days


def GetReportingDates(url, user, passwd):
    """
    Fetches reporting dates as XML
    :param url: url to send request
    :param user: user name
    :param passwd: password
    :return: a list of all dates of reports available.
    """
    try:
        client = SoapClient(wsdl=url,trace=False)
        response = client.GetReportingDates(user, passwd)
        result = response['GetReportingDatesResult']
        reply = base64.decodestring(result)
    except Exception as ex:
        err = "<Error>"+ex.message+"</Error>"
        raise BaseException(err)

    return ParseReportingDates(reply)

def ParseReportingDates(xml_str):
    """

    :param xml_str: an XML string of a SOAP response containing all reporting dates.
    :return: list of dates representing all reporting dates.
    """

    root = ET.fromstring(xml_str)
    date_tags = root.findall('./Date/REPORT_DATE')
    return [datetime.datetime.strptime(i.text, '%m/%d/%Y').date() for i in date_tags]


def FindMissingReports(cursor, dates):
    """
    Finds all report dates in the database, returns any reports ETS has that the database doesn't have
    :param cursor: a database cursor object
    :param dates: a list of all dates available
    :return: a list of dates not in the database currently
    """
    ''''''

    SQL = """SELECT *
    FROM tblPraxisUpdate
    """

    cursor.execute(SQL)
    results = cursor.fetchall()
    db_dates = [r.ReportDate.date() for r in results]
    missing_dates = []
    for d in dates:
        if d not in db_dates:
            missing_dates.append(d)

    return missing_dates

'''Score reports'''

def GetScoreReport(url, user, passwd, rpt_date):
    """
    Fetches a score report given a report date.
    :param url: request url
    :param user: user name
    :param passwd: password
    :param rpt_date: report date to fetch
    :return: a ScoreReport object.
    """
    try:
        print 'Fetching report for', rpt_date,
        client = SoapClient(wsdl=url,trace=False)
        response = client.GetScoreReportsGivenReportingDate(user, passwd, rpt_date, 'P')
        result = response['GetScoreReportsGivenReportingDateResult']
        reply = base64.decodestring(result)
    except Exception as ex:
        err = "<Error>"+ex.message+"</Error>"
        raise BaseException(err)

    #write XML to disc
    report_loc = r"\\fileshare\Academic\Academic\Education\Department Only\Continuous tracking\New Student Database\Back Stage\Praxis Converter\Reports\{}".format(datetime.datetime.now().date().strftime('%m_%d_%Y'))

    #make the directory if it doesn't exist
    if not os.path.exists(report_loc):
        os.makedirs(report_loc)

    file_name = 'praxis_report_{}.xml'.format(rpt_date.replace('/', '_'))
    with open(os.path.join(report_loc, file_name), 'w+') as xml_file:
        xml_file.write(reply)
        print '-> Created file', file_name
    return ScoreReport(file_name, report_loc, os.path.join(report_loc, file_name), rpt_date)

def SOAPHandler(cursor):
    """
    driver function to make a SOAP request
    :param cursor: a database cursor object
    :return: new XML reports
    """

    url = 'https://datamanager.ets.org/edmwebservice/edmpraxis.wsdl'

    #gather credentials
    msg = "Enter ETS logon information"
    title = "ETS Web Service"
    fieldNames = ["User ID", "Password"]
    field_values = easygui.multpasswordbox(msg=msg,title=title, fields=fieldNames, values=('stoebelj', ''))      #using 'stoebelj' as default user name

    # make sure that none of the fields was left blank
    while True:
        if field_values == None: break       #cancel was pressed.
        errmsg = ""
        for i in range(len(fieldNames)):
            if field_values[i].strip() == "":
                errmsg = errmsg + ('"%s" is a required field. \n' % fieldNames[i])
        if errmsg == "": break # no problems found
        print errmsg
        field_values = easygui.multpasswordbox(errmsg, title, fieldNames, field_values)

    login, pw = tuple(field_values)

    dates = GetReportingDates(url, login, pw)
    missing_dates = FindMissingReports(cursor, dates)

    if not missing_dates:
        print 'No new reports found.'
        return

    new_xmls = []
    for d in missing_dates:
        date_str = d.strftime('%m/%d/%Y')
        new_x = GetScoreReport(url, login, pw, date_str)        #returns full path of new xml created
        new_xmls.append(new_x)
    return new_xmls