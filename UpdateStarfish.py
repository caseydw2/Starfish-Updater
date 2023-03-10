# -*- coding: utf-8 -*-
"""
Created on Wed Nov 16 09:08:38 2022



@author: E1448105
"""

import datetime as dt
import shutil

import pandas as pd
from openpyxl import load_workbook

from ListStudentsCOPYv2 import subjdict
from updatevisits import update_visit_log
from GUI.GUI_Utilities import get_file, yes_no

# from GetPath import getfilepathfromfolder

#maybe change it to simply drop the file in a dir?

starfish_tracker_edits = "./Starfish Tracking Spring 2023.xlsx"

#COMPARE STARFISH SPREASHEET WITH STARFISH TRACKING SPREADSHEET

flags_to_track = ["Missing Assignments","Academic Concern", "Course Withdrawal", "Academics - Attend Tutoring-Closable","In Danger of Earning Under a C","I Need Help", "I Need Help In A Course"]

def prep_starfish_sheet(sheet:str) -> pd.DataFrame:
    """Makes dataframe with 'First Name' and 'Last Name' column added from starfish's 'trackingitems {subj}'

    Args:
        sheet (str): Path to 'tracking item's csv'

    Returns:
        DataFrame: trackingitems {subj} DataFrame containing "First Name", "Last Name", "studentExtId","name","category","courseId","raiseDate"
        of students in starfish raised flags.
    """

    df = pd.read_csv(sheet)
    #Create a column for first and last name
    df["names"] = df["studentName"].str.split(",")

    df[["Last Name","First Name"]] = df["names"].tolist()

    df["raiseDate"] = pd.to_datetime(df["raiseDate"])
    df = df.loc[df["name"].isin(flags_to_track)] # Changed 3/6/2023. Hopefully to drop any flag/to-do we don't need
    df= df[["First Name", "Last Name", "studentExtId","name","category","courseId","raiseDate"]].astype({"studentExtId" : "str"})
    return df

def prep_tracking_sheet(sheet:str, subj:str) -> pd.DataFrame:
    sheet_name = subjdict[subj]["ssheet"]
    df = pd.read_excel(sheet, sheet_name,header = 2).fillna("N").astype({"ID#" : "str"})
    return df


def tilldate(star_df:pd.DataFrame,date:dt.datetime) -> pd.DataFrame:
    format_str = "%m/%d/%Y"

    date = dt.datetime.strptime(date, format_str)
    if date > dt.datetime.today():
        date = input("Invalid date. Please reenter. ")

    star_df = star_df.loc[star_df.raiseDate >= date]
    return star_df



#NEED TO TEST
def getCourse(course:str):
    subj = course[3:7]
    num = course[7:10]
    if num[0] == "0":
        num = num[1:]
    camp = course[:2]
    return [camp,subj + num]



#NEED TO EDIT
def AddStudentToTopRow(student:pd.DataFrame,s) -> None:
    s.insert_rows(4)
    if type(student["First Name"]) != float:
        s["A4"] = (student["First Name"] + " " + student["Last Name"]).strip()
    s["B4"] = student["studentExtId"].strip()
    s["D4"] = getCourse(student["courseId"])[0]
    s["C4"] = getCourse(student["courseId"])[1]
    s["E4"] = student["name"].strip()
    s["F4"] = student["raiseDate"]
    s["G4"] = "N"
    s["I4"] = "N"
    s["J4"] = "N"
    if student["category"] == "FLAG" or student['category'] == 'TO_DO':
        s["H4"]= "Y"
    else:
        s["H4"]= "N"

def get_withdrawls(df:pd.DataFrame):
    df_w = df.loc[df["name"] == "Course Withdrawal"]
    df_nw  = df.loc[df["name"] != "Course Withdrawal"].drop_duplicates(("studentExtId","category"))
    return df_w,df_nw

def update_student_row(row:int,since_date:dt.datetime,student:pd.DataFrame,s):
    recent_raise_date = student["raiseDate"]
    is_after_update = (since_date == None) or (recent_raise_date > since_date)
    is_same_flag = s["F" + str(row)].value == recent_raise_date
    is_flag_of_concern = ((str(student["category"]) in ["FLAG" , "TO_DO"]) and (student["name"] not in ["Attendance Concern","Lack of Class Participation","Electronic Distraction"]))
    if is_after_update and is_flag_of_concern and not is_same_flag:
        s["E" + str(row)] = student["name"] #Add the new flag name in flag column
        s["H" + str(row)] = "Y"
        s["F" + str(row)] = recent_raise_date

#Clean Up
def addupdate_students(tracker0,subj:str) -> None:
    add_sheet = "Yes"
    change = False
    while add_sheet == "Yes":
        base_row = 4
        starfish = prep_starfish_sheet(get_file(f"{subj} data sheet pulled from starfish"))
        tracker = prep_tracking_sheet(tracker0,subj)
        ids_stripped = [s.strip() for s in tracker["ID#"].values]

        sheet = subjdict[subj]["ssheet"]
        book = load_workbook(starfish_tracker_edits)
        s = book[sheet]
        starfish_w, starfish_nw = get_withdrawls(starfish)


        if input(f"Would you like to add a 'since' date for {subj}? If so type 'Y'. Otherwise, press Enter.").upper() == "Y":
            since_date = dt.datetime.strptime(input("please input since date (mm/dd/yyyy)"),"%m/%d/%Y")
        else:
            since_date = s["A2"].value


        
        for index,student in starfish_nw.iterrows():

            if (student["studentExtId"].strip() in ids_stripped): #Check to see if student is already in tracker
                index = ids_stripped.index(student["studentExtId"].strip())
                row = index + base_row #If student is in tracker, label the row in tracker.
                update_student_row(row,since_date,student,s)
                change = True
                book.save(starfish_tracker_edits)

            elif (str(student["category"]) in ["FLAG" , "TO_DO"]) and (student["name"] not in ["Attendance Concern","Lack of Class Participation","Electronic Distraction"]) :
                AddStudentToTopRow(student,s)
                base_row += 1
                change = True
                book.save(starfish_tracker_edits)

        tracker = prep_tracking_sheet(tracker0,subj)
        for index,student in tracker.iterrows():
            if (str(student["ID#"]) in starfish_w["studentExtId"].values):
                row = index + 4

                if s["M" + str(row)].value != "Withdrew":
                    s["M" + str(row)] = "Withdrew"
                    change = True
            
        book.save(starfish_tracker_edits)
        add_sheet = yes_no("Would you like to add another sheet?")

        
    if change:
        s["A2"] = dt.datetime.today()
        book.save(starfish_tracker_edits)
    
    
    update_visit_log(subj,starfish_tracker_edits)


if __name__ == "__main__":
    starfish_tracker = get_file("The main starfish tracker",initial_folder=".\Casey's Starfish")
    shutil.copyfile(starfish_tracker,starfish_tracker_edits)
    if yes_no("Would you like to update math?") == "Yes":
        addupdate_students(starfish_tracker_edits,"MATH")
    if yes_no("Would you like to update English?") == "Yes":
        addupdate_students(starfish_tracker_edits,"ENGL")
