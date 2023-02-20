# -*- coding: utf-8 -*-
"""
Created on Mon Oct 24 16:16:40 2022

Sends Email

@author: Casey Wheaton-Werle
"""

import win32com.client as win32
from ListStudentsCOPYv2 import ListStudents





def SendMassEmail(lstudents:list[object], emailnum:int):
    '''
    Sends emails to the list of students in lstudents with the

    Parameters
    ----------
    lstudents : list
        List of student objects from ListStudents.
    emailnum : int
        Email number the student is on.

    Returns
    -------
    str
        "Number of emails sent: ". String for readability.
    count : int
        Number of emails sent.
    '''

    count = 0
    for student in lstudents:
        if (student.emailnum == emailnum) and (student.needsemail == "Y") and (student.withdrawn == False):
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)  # creates an email object
            mail.SentOnBehalfOfName = "lvlr.studentsuccess@mcckc.edu"
            mail.BCC = student.bccemail
            mail.To = student.email
            mail.Subject = student.subject
            mail.Body = student.b

            mail.Send()

            count += 1
            print("Email Sent to " + str(student) + " to their student email:\n " +
                  str(student.email) + ". \n" + str(student.bccemail) + " was bcc'd.")

    return ("Number of emails sent: ", count)



if __name__ == "__main__":
    students = ListStudents()
    for student in students:
        print("\n")
        print(student.__dict__)
    SendMassEmail(students,1)

