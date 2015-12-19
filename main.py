#-------------------------------------------------------------------------------
# Name:        praxis
# Purpose: A script to handle uploading praxis scores into your data managemnet system
#
# Author:      stoebelj
#
# See read me for more details
#------------------------------>-------------------------------------------------


import sys
sys.path.insert(0, '.')
import soap
import praxis_email
import easygui
import jacob
import datetime
import roman
from xml.etree.ElementTree import iterparse
import pyodbc


class Student:
    def __init__(self, bnum, name):
        """
        an object for representing students.
        :param bnum: student's id number
        :param name: student's full name
        """
        self.bnum=bnum
        self.name=name.title()
        self.tests=[]

    def compute_pass_status(self):
        """
        look at all tests objects taken by this student and see if any are failing.
        :return: sets self.all_pass to the appropriate value
        """

        self.all_pass=True
        for test in self.tests:
            if test.cut_score>test.score:
                self.all_pass=False

    def fetch_email_info(self, cursor):
        """
        sets parameters needed to compose an email about this student.
        :param cursor: a cursor object for the database
        :return: sets parameters needed to compose an email.
        """

        '''Fetch a student's advisor's info and first and last name (for preping an email).'''

        SQL="""SELECT tblBasicInfo.FirstName as FirstName, tblBasicInfo.PreferredFirstName as PrefFirst, tblBasicInfo.LastName as LastName, tblBasicInfo.Email as Stu_Email, tblBasicInfo.EnrollmentStatus as EnrollmentStatus,
        tblAdvisors.Advisor as Advisor, tblAdvisors.AdvisorEmail as Ad_Email, tblAdvisors.Salutation as Ad_Sal
        FROM tblAdvisors RIGHT JOIN tblBasicInfo ON tblAdvisors.Advisor = tblBasicInfo.EDSAdvisor
        WHERE tblBasicInfo.Bnum=?
        """
        cursor.execute(SQL, self.bnum)

        result=cursor.fetchone()
        self.stu_first=result.FirstName

        if result.PrefFirst:        #use the prefered name field, unless its blank
            self.pref_first=result.PrefFirst
        else:
            self.pref_first=result.FirstName

        self.stu_last=result.LastName
        self.stu_email=result.Stu_Email
        self.enrollment_status=result.EnrollmentStatus
        self.advisor=result.Advisor
        self.advisor_email=result.Ad_Email
        self.advisor_salutation=result.Ad_Sal

class Test:
    def __init__(self, test_date, test_code, test_name, test_score, best_score, cut_score):
        """
        an object to represent a single test
        :param test_date: date test was taken
        :param test_code: the test's unique code
        :param test_name: the name of the test
        :param test_score: the score recived on this attempt
        :param best_score: the best score recived by this student on this test.
        :param cut_score: the score needed to pass
        """
        self.date=test_date
        self.code=test_code
        self.name=test_name
        self.score=test_score
        self.best_score=best_score
        self.cut_score=cut_score

        if self.score>=self.cut_score:

            self.passing=True
        else:
            self.passing=False

        self.subtests=[]

    def build_html(self):
        """
        build this the html to report on this test for us in email.
        post: the html for this test is set
        """


        if self.passing:
            self.text='<font color="black">'
        else:
            self.text='<font color="red">'

        #needs to be wrapped in <ol> tags
        #<ol>
        self.text+='''<li>{test_name}: {score}, cut score: {cut_score}
        <br>
        <ul>
        '''.format(test_name=self.name, score=self.score, cut_score=self.cut_score)

        for sub in self.subtests:
            self.text+='''<li>{sub_name}: {sub_score} out of {pts_possible} points possible.<br>
            Performance of middle 50%: {lower}-{upper}
            <br>'''.format(sub_name=sub.name, sub_score=sub.raw_score, pts_possible=sub.pts_aval, lower=sub.avg_low, upper=sub.avg_high)
        self.text+='</ul></font>'
        #</ol>

class Subtest:
    def __init__(self, sub_name, num, raw_score, pts_aval, avg_high, avg_low):
        """
        a subtest object. Each test has several sub-tests.
        :param sub_name: name of the sub test
        :param num: the order this sub-test falls within the test
        :param raw_score: number of points earned
        :param pts_aval: number of points available
        :param avg_high: the upper bound score of the middle 50%
        :param avg_low: the lower bound score of the middle 50%
        """
        self.name=sub_name
        self.num=num
        self.raw_score=raw_score
        self.pts_aval=pts_aval
        self.avg_high=avg_high
        self.avg_low=avg_low

def QMarks(qty):
    """
    :param qty: number of question marks to insert
    :return: a string of question marks seperated by commas for use with blind parameters
    """

    return ', '.join(['?' for i in range(qty)])

def FetchBnum(ssn, cursor, name):
    """
    Looks up student id (called a B number at Berea College)  based on SSN
    :param ssn: a social security number
    :param cursor: a databaes cursor object
    :param name: a student's name in case a match can't be found
    :return: student's Bnumber is returned
    """

    SQL_bnum="""SELECT Bnum
    FROM tblBasicInfo
    WHERE SSN=?
    """
    cursor.execute(SQL_bnum, ssn)
    result = cursor.fetchone()
    if result:
        return result.Bnum
    else:
        bnum = easygui.enterbox("Couldn't find a B# for {name}: {ssn}. Please enter a correct B# or cancel to skip name".format(name=name, ssn=ssn))
        if bnum:
            return bnum

def AddNewTest(test, cursor):
    """
    adds a new test to the database if it doesn't not currently exist.
    :param test: a test code not currently in the database.
    :param cursor: a database cursor object
    :return: none
    """

    current_test = test.find('currenttestinfo')
    test_dict = current_test.attrib
    test_name = test_dict['test_name']
    test_code = test_dict['test_code']
    test_number = test_code[1:]
    while True:
        try:
            cut_score = int(raw_input('What is the cut score for this test?'))
            break
        except ValueError:      #catching invalid entries
            easygui.msgbox('Invalid entry. Please enter an integer')

    #build sub test info
    sub_names = []
    sub_tests = current_test.findall('currenttestcategoryinfo')
    for st in sub_tests:
        st_info = st.attrib['testcategory'].split(' ')
        st_name = ' '.join(st_info[2:])
        sub_names.append(st_name)


    params = [test_code, test_number, test_number, False]
    test_family = easygui.choicebox('Which test family does this test belong to? {0} ({1})'.format(test_name, test_code), choices=('Praxis I', 'Praxis II'))
    if test_family == 'Praxis I':

        empty_subs = 4-len(sub_names)
        params += sub_names + [None for i in range(empty_subs)] + [cut_score]
        SQL_insert = """INSERT INTO tblPraxisITests
        VALUES ({})
        """.format(QMarks(len(params)))
        cursor.execute(params)

    elif test_family == 'Praxis II':
        empty_subs = 7 - len(sub_names)
        params += sub_names + [None for i in range(empty_subs)] + [cut_score]
        SQL_insert = """INSERT INTO tblPraxisITests
        VALUES ({})
        """.format(QMarks(len(params)))
        cursor.execute(params)


def FetchCutScore(test, cursor):
    """
    fetches a test's cut score
    :param test: a praxis test code
    :param cursor: a dataabse cursor object
    :return: none
    """

    SQL_PIcut="""SELECT CutScore
    FROM tblPraxisITests
    WHERE TestID=?
    """
    cursor.execute(SQL_PIcut, test)
    result=cursor.fetchone()
    if result:
        return result.CutScore
    else:   #it wasn't in Praxis I, look in Praxis II
        SQL_PIICut="""SELECT CutScore
        FROM tblPraxisIITests
        WHERE TestID=?
        """
        cursor.execute(SQL_PIICut, test)
        result=cursor.fetchone()
        if result:
            return result.CutScore
        else:
            #here's where to add a new test
            add = easygui.ynbox('Test Code {} was not found. Would you like to add it?'.format(test_code))
            if add:
                AddNewTest(test, cursor)
            else:
                raise BaseException ('No Cut score found for {}. Operation stopped by user.'.format(test_code))

def TestFamilies(cursor):
    """
    Builds a dictionary of test families:  All tests belonging to Praxis I are mapped to the integer 1 tests belonging to Praxis II references to the integer 2
    :param cursor: a database cursor objet
    :return: a dictionary mapping all known tests to their test family
    """

    test_families={
    }
    SQL_P1="""SELECT TestID
    FROM tblPraxisITests
    """
    cursor.execute(SQL_P1)
    for row in cursor.fetchall():
        test_families[row.TestID]=1

    SQL_P2="""SELECT TestID
    FROM tblPraxisIITests
    """
    cursor.execute(SQL_P2)
    for row in cursor.fetchall():
        test_families[row.TestID]=2

    return test_families

def RemoveLeadZeros(date_str):
    """
    :param date_str: a date in format MM/DD/YYYY (string)
    :return: a date in with leading zero's removed.
    """

    date_list=date_str.split('/')
    m=str(int(date_list[0]))
    d=str(int(date_list[1]))
    y=(date_list[2])

    return '/'.join([m,d,y])

def hghstestDict(hghst_test, cursor):
    """
    Returns a dictionary of best results for each test student has taken.
    :param hghst_test: a <hghsttest> element in a praxis report
    :param cursor: a database cursor object
    :return: a dictionary containing all scores this student has taken mapped to their highest score.
    """

    best_dict={}
    tests=hghst_test.findall('hghsttestinfo')
    for test in tests:
        test_info=test.attrib
        test_code=(test_info['testname'][:test_info['testname'].find(' ')])
        if test_code=='5031':     #5031 is the elem ed multiple subjects. Its not a real test, so just skip it.
            continue

        try:
            best_score=int(test_info['testscore'])
        except ValueError:
            best_score=None

        try:
            cut_score=int(test_info['requiredscoreasofreportdate'])
        except ValueError:
            cut_score=FetchCutScore(test_code, cursor)

        best_dict[test_code]=(best_score, cut_score)

    return best_dict


def AlreadyIn(report, cursor):
    """
    :param report: a ScoreReport object
    :param cursor: a database cursor object
    :return: returns if this score report is already in the database.
    """

    SQL = """SELECT *
    FROM tblPraxisUpdate
    WHERE ReportDate=?
    """
    cursor.execute(SQL, report.date)
    if cursor.fetchone():
        print 'Report date', report_date, 'has already been processed. Emails will not be sent.'
        return True
    else:
        return False

def UpdateReportLog(report_date, cursor):
    """
    updates the database to reflect having processed a report. Raises an error if the report date causes an IntegrityError
    :param report_date: date of score report
    :param cursor: a database cursor object
    """

    SQL = """INSERT INTO tblPraxisUpdate
    VALUES (?,Now())
    """
    try:
        cursor.execute(SQL, report_date)
        cursor.commit()
        print 'Praxis log updated.'
    except pyodbc.IntegrityError:
        raise pyodbc.IntegrityError('Report date {} already exists in the database'.format(report_date))

def ProcessXML(xml_loc, cursor):
    """
    Reads an XML document and returns a list of student objects.
    :param xml_loc: a location of an XML score report
    :param cursor: a database connect object
    :return: a list of Student objects which in turn refernce Test objects.
    """

    all_students=[]

    for event, sroot in iterparse(xml_loc):

        if sroot.tag=='scorereport':
            hght_test=sroot.find('hghsttest')
            best_dict=hghstestDict(hght_test, cursor)

            candidate_info=sroot.find('candidateinfo')
            ssn=candidate_info.find('ssn').text.replace('-','')
            name=candidate_info.find('name').text
            bnum=FetchBnum(ssn, cursor, name)
            if not bnum:        #if no Bnum could be found, move to next record
                continue

            student=Student(bnum, name)

            current_test=sroot.find('currenttest')
            tests=current_test.findall('currenttestinfo')
            for troot in tests:     #iterate over all tests
                test_info=troot.attrib
                test_code=(test_info['testname'][:test_info['testname'].find(' ')])
                test_name=test_info['testname'][test_info['testname'].find(' '):]
                test_score=int(test_info['testscore'])
                test_date_str=test_info['testdate']
                test_date= datetime.datetime.strptime(test_date_str, '%m/%d/%Y')

                best_score, cut_score=best_dict[test_code]

                test=Test(test_date, test_code, test_name, test_score, best_score, cut_score)

                sub_tests=troot.findall('currenttestcategoryinfo')
                for sub_root in sub_tests:  #iterate over all sub tests.
                    sub_info=sub_root.attrib
                    sub_name_raw=sub_info['testcategory']
                    sub_name=' '.join(sub_name_raw.split(' ')[1:])      #changed on 11/2/15, new format is "I. United States History; Government; Citizenship"

                    raw_num = sub_name_raw.split(' ')[0]    #raw sub test number
                    sub_num = roman.fromRoman(raw_num.replace('.', ''))     #convert to integer.

                    pts_earned=int(sub_info['pointsearned'])
                    pts_aval=int(sub_info['pointavailable'])
                    avg_low, avg_high=tuple( sub_info['avgperformancerange'].split(' - '))

                    sub_test=Subtest(sub_name, sub_num, pts_earned, pts_aval, avg_high, avg_low)

                    #append subtest to test
                    test.subtests.append(sub_test)

                #append test to student
                student.tests.append(test)

            #append student to all students
            all_students.append(student)

    return all_students


def AddRecords(all_students, cursor):
    """
    Inserts all student records into the database.
    :param all_students: a list of Student objects
    :param cursor: a database cursor object
    :return: records from all Student objects are inserted into the database.
    """

    test_families= TestFamilies(cursor) #creates a dictionary of tests, grouped by family
    for student in all_students:
        for test in student.tests:

            try:
                family=test_families[test.code]
            except KeyError:    #test isn't in either test family. Let's not worry about it.
                continue

            date_str=RemoveLeadZeros(datetime.datetime.strftime(test.date, '%m/%d/%Y'))

            test_id='-'.join([student.bnum, test.code, date_str])

            #---------------#
            #BUILD PRAXIS I
            #---------------#

            if family==1:

                package=[test_id, student.bnum, student.name, test.date, test.code, test.name, test.score, test.best_score, test.cut_score]     #build a package for insert first
                SQL_insert="""INSERT INTO tblPraxisI
                (NewTestID, Bnum, FileAs, PIDate, TestNumber, TestCode, PIScore, PIBestScore, PINeeded, """

                #add in sub tests dynamically.
                sub_count=len(test.subtests)
                for s in range(1,sub_count+1):
                    SQL_insert+='PISub{0}, PIRaw{0}, PIPtsAval{0}, Avg{0}High, Avg{0}Low, '.format(s)     #adds to the SQL
                    sub_test=test.subtests[s-1]         #access the sub test
                    package+=[sub_test.name, sub_test.raw_score, sub_test.pts_aval, sub_test.avg_high, sub_test.avg_low]        #add sub test elements into the package
                #remove the last ', '
                SQL_insert=SQL_insert[:-2]
                SQL_insert+=""")
                VALUES ({})
                """.format(QMarks(len(package)))      #finish the SQL

                try:    #try to insert
                    cursor.execute(SQL_insert, package)
                    cursor.commit()
                    print 'INSERTED Praxis I test', test_id
                except pyodbc.IntegrityError:       #assemble an update querry
                    package=[test.score, test.best_score, test.cut_score]
                    SQL_update="""UPDATE tblPraxisI
                    SET PIScore=?, PIBestScore=?, PINeeded=?, """
                    for s in range(1, sub_count+1):
                        SQL_update+='PISub{0}=?, PIRaw{0}=?, PIPtsAval{0}=?, Avg{0}High=?, Avg{0}Low=?, '.format(s)
                        sub_test=test.subtests[s-1]
                        package+=[sub_test.name, sub_test.raw_score, sub_test.pts_aval, sub_test.avg_high, sub_test.avg_low]

                    SQL_update = SQL_update[:-2]    #remove the last comma
                    SQL_update+="""
                    WHERE NewTestID=?
                    """
                    package.append(test_id)
                    cursor.execute(SQL_update, package)
                    cursor.commit()
                    print 'UPDATED Praxis I test', test_id

            #---------------#
            #BUILD PRAXIS II
            #---------------#

            elif family==2:

                package=[test_id, student.bnum, student.name, test.date, test.code, test.name, test.score, test.best_score, test.cut_score]     #build a package for insert first
                SQL_insert="""INSERT INTO tblKey4
                (NewTestID, Bnum, FileAs, PIIDate, TestCode, TestName, PIIScore, PIIBestScore, PIINeeded, """

                #add in sub tests dynamically.
                sub_count=len(test.subtests)
                for s in range(1,sub_count+1):
                    SQL_insert+='PIISub{0}, PIIRaw{0}, PIIPtsAval{0}, PIIAvgHigh{0}, PIIAvgLow{0}, '.format(s)     #adds to the SQL
                    sub_test=test.subtests[s-1]         #access the sub test
                    package+=[sub_test.name, sub_test.raw_score, sub_test.pts_aval, sub_test.avg_high, sub_test.avg_low]        #add sub test elements into the package

                #remove the last ', '
                SQL_insert=SQL_insert[:-2]
                SQL_insert+=""")
                VALUES ({})
                """.format(QMarks(len(package)))      #finish the SQL

                try:    #try to insert
                    cursor.execute(SQL_insert, package)
                    cursor.commit()
                    print 'INSERTED Praxis II test', test_id
                except pyodbc.IntegrityError:       #assemble an update querry
                    package=[test.score, test.best_score, test.cut_score]
                    SQL_update="""UPDATE tblKey4
                    SET PIIScore=?, PIIBestScore=?, PIINeeded=?, """
                    for s in range(1, sub_count+1):
                        SQL_update+='PIISub{0}=?, PIIRaw{0}=?, PIIPtsAval{0}=?, PIIAvgHigh{0}=?, PIIAvgLow{0}=?, '.format(s)
                        sub_test=test.subtests[s-1]
                        package+=[sub_test.name, sub_test.raw_score, sub_test.pts_aval, sub_test.avg_high, sub_test.avg_low]

                    SQL_update = SQL_update[:-2]    #remove the last comma
                    SQL_update+="""
                    WHERE NewTestID=?
                    """.format(QMarks(len(package)))
                    package.append(test_id)
                    cursor.execute(SQL_update, package)
                    cursor.commit()
                    print 'UPDATED Praxis II test', test_id

def Main():
    '''
    program flow:
        look for reports not in database
        for each report
            create xml file
            ask user if email should be sent
            process as before
    '''

    #CHANGE THE LINE BELOW TO CREATE YOUR OWN CONNECTION AND CURSOR OBJECTS.
    cnxn,cursor= CreateConnectionAndCursor()

    new_xmls = soap.SOAPHandler(cursor)     #returns info for all new xmls

    if not new_xmls:    #no new XMLs were found. Stop here
        print 'No new reports found.'
        return

    #revise to automatically send emails if report hasn't been processed unless:
        #user doesn't want any emails sent
        #report is more than 30 days old (print a warning if so)
    report_names = [r.name for r in new_xmls]
    send_emails = bool(easygui.ynbox('The following new reports were found {}\n. Would you like to send emails if required? If not, score reports will be processed but not marked as done.'.format(', '.join(report_names))))
    if send_emails:
        email_password=easygui.passwordbox(msg='Please enter email password', title='Email password')

    for i in new_xmls:

        print 'Processing {}'.format(i.name)
        all_students=ProcessXML(i.loc, cursor)
        AddRecords(all_students, cursor)

        print '{} done'.format(i.name)

        #send email if required
        if send_emails and not AlreadyIn(i, cursor):        #email will never send if 1) user disallows or 2) report is already in.
            if i.age > 30:
                if bool(easygui.ynbox('Report {} is more than 30 days old. Are you sure you want to send emails for this report?'.format(i.name))):
                    praxis_email.EmailHandler(all_students, cursor, email_password, go=True)
                    UpdateReportLog(i.date, cursor)
            else:       #yougner than 30 days
                praxis_email.EmailHandler(all_students, cursor, email_password, go=True)
                UpdateReportLog(i.date, cursor)

if __name__ == '__main__':
    Main()
