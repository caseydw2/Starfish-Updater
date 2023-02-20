# -*- coding: utf-8 -*-
"""
Created on Mon Oct 24 13:03:01 2022

Reads Excel Spreadsheets and creates list of Students in excel sheet

@author: Casey Wheaton-Werle
"""

import datetime as dt
import pandas as pd
from GUI.GUI_Utilities import get_file, which_subj


# create a Student class for each row of Student spreadsheet
class Student:
    def __init__(self, Name:str, ID, Course:str, Flag:str, emailnum:int, needsemail,camp):
        self.FN = Name.split(" ")[0]
        self.LN = Name.split(" ")[1]
        self.id = ID
        self.email = "S" + str(ID) + "@student.mcckc.edu"
        self.course = Course.split(";")[0]
        self.flag = Flag.split(";")[0]
        self.camp = camp
        self.emailnum = emailnum
        self.needsemail = needsemail

    def __str__(self):
        return str(self.FN) + " " + str(self.LN) + " (ID#: S" + str(self.id) + ")"


# make dict of all the short coursenames with their full course names
courses = {"ENGL101": "Composition and Reading I",
           "ENGL102": "Composition and Reading II",
           "ENGL90": "Foundations of College Writing II",
           "ENGL215": "Technical Writing",
           "MATH31": "Pre-College Mathematics",
           "MATH115": "Statistics",
           "MATH95": "Algebra Principles",
           "MATH120": "College Algebra",
           "MATH150": "PreCalculus",
           "MATH119": "Mathematical Reasoning and Modeling"}

subjdict = {
    "MATH": {"bcc": "casey.wheaton-werle@mcckc.edu", "subject": "Flag Notification: Visit Math Lab", "ssheet": "MATH",
             "rsheet": "MATH Replies"},
    "ENGL": {"bcc": "leslie.newton@mcckc.edu", "subject": "Flag Notification: Visit Writing Studio", "ssheet": "WS",
             "rsheet": "WS Replies"}}

campcont = {
    "PV": {"phone":"816-604-4292","loc":"LR, 2nd floor","email":"pvssc.reservation@mcckc.edu"},
    "LV": {"phone":"816-604-2205","loc":"LR 225","email":"lvlr.studentsuccess@mcckc.edu"},
    "ON": {"phone":"","loc":"","email":""},
    "MW": {"phone":"","loc":"","email":"mw.learningcenter@mcckc.edu"},
    "BR": {"phone":"816-604-6770","loc":"CC 142","email":"br.learningservices@mcckc.edu"}
}

# get the body paragraph for the student based on their flag and studen


def GetBody(studento: Student, Rep):
    # cycle through the rows and compare the flag raised with the first column
    for i, Flag in enumerate(Rep):
        # grab the body that matches the flag and the number of times emailed
        if str(studento.flag) == str(Flag[0]):
            body = str(Flag[studento.emailnum])
            body = body.replace(
                "[NAME]", studento.FN).replace("[COURSE]", courses[str(studento.course)]).replace("[FLAG]",studento.flag.lower())
            return body
    else:
        raise Exception(
            f"{studento} has flag {studento.flag}, which cannot be found in the flag responses.")
    # replace the name spaces with the students information


def DaysSLEmail(one, two, three):
    """
    Takes three dates and returns the number of days since the newest date or
    1000 if all three inputs are not date.time types.

    Parameters
    ----------
    one : datetime.datetime
        First day Emailed.
    two : datetime.datetime
        Second day Emailed.
    three : datetime.datetime
        Third day emailed.

    Returns
    -------
    int
        The number of days from today since the last day emailed or 1000.

    """
    if type(one) == dt.datetime:
        a = (dt.datetime.today() - one).days
    else:
        a = 1000
    if type(two) == dt.datetime:
        b = (dt.datetime.today() - two).days
    else:
        b = 1000
    if type(three) == dt.datetime:
        c = (dt.datetime.today() - three).days
    else:
        c = 1000
    return min(a, b, c)

def istest(bool:bool,subj:str):
    if bool:
        bccemail = input("Would you like to bcc an email? Type 'YES' to add a bcc email").upper()
        if bccemail == "YES":
            bccemail = input("which email would you like to bcc?(MINE to send to work email) ").upper()
            if bccemail == "MINE":
                bccemail = "casey.wheaton-werle@mcckc.edu"
        else:
            bccemail = " "
        subject = "This is a Test"
        bccemail = bccemail
    else:
        subject = subjdict[subj]["subject"]
        bccemail = subjdict[subj]["bcc"]
    return subject,bccemail
        
    

def count_email():
    pass

# MAKE IT LOOK BETTER
def ListStudents() -> list[Student]:
    """
    Returns
    -------
    students : list
        List of student objects based on the information provided in the spreadsheet.
    """
    # initialize path to spreadsheet
    path = get_file("Main Starfish tracking worksheet")

    # Ask if we are wanting to test and what subject
    Subj = which_subj("Which subject would you like to create a list for?")
    isTest = input("Is this a test? ('Y' for yes)").upper()
    subject,bccemail = istest(isTest == "Y",Subj)
    ssheet = subjdict[Subj]["ssheet"]
    rsheet = subjdict[Subj]["rsheet"]

    Stud = pd.read_excel(path, ssheet)
    Stud = Stud.to_numpy()
    Stud = Stud[2:]

    # Get the spreadsheet for responses
    Rep = pd.read_excel(path, rsheet)
    # col 0 is flag, col 1 is first email, col 2 is second email...
    Rep = Rep.to_numpy()

    NameCol, IDCol, ClassCol, CampCol, FlagCol, DateRaised, FirstEmailCol, NeedsEmailCol, SecondEmailCol, ThirdEmailCol = 0, 1, 2, 3, 4, 5, 6, 7,8,9
    WithdrawCol = -1

    numneedemail = str(Stud[:, NeedsEmailCol]).upper().count("Y")
    studcount = len(Stud)

    students = []
    onetwothreeemail = [0, 0, 0]
    for i, student in enumerate(Stud):

        if student[NeedsEmailCol]== "Y":
            emailnum = 4 - \
                       (str(student[FirstEmailCol]).count("N") +
                        str(student[SecondEmailCol]).count("N") +
                        str(student[ThirdEmailCol]).count("N"))

            onetwothreeemail[emailnum - 1] += 1

            studento = Student(student[NameCol], student[IDCol],
                               student[ClassCol], student[FlagCol], emailnum, student[NeedsEmailCol],student[CampCol])

            studento.subject = subject
            studento.bccemail = bccemail

            studento.b = GetBody(studento, Rep)
            studento.withdrawn = (student[WithdrawCol] == "Withdrew")

            # maybe add a date of last email to avoid quick email turn around
            studento.sincelastemail = DaysSLEmail(
                student[FirstEmailCol], student[SecondEmailCol], student[ThirdEmailCol])

            students.append(studento)
    print("There are ", studcount, " student(s) in this spreadsheet, with ", numneedemail,
          " needing to be emailed. These ", numneedemail, " students are in this list to be emailed.")
    print("\nThere are ", onetwothreeemail[0], " to recieve their first email.")
    print("\nThere are ", onetwothreeemail[1], " to recieve their second email.")
    print("\nThere are ", onetwothreeemail[2], " to recieve their third email.")
    return students


if __name__ == '__main__':
    x = ListStudents()
    for student in x:
        print("\n")
        print(student.__dict__)
 