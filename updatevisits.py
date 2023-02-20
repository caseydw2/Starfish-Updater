# -*- coding: utf-8 -*-
"""
Created on Tue Nov 22 12:07:18 2022

@author: E1448105
"""

import pandas as pd
from ListStudentsCOPYv2 import subjdict
from openpyxl import load_workbook

from GUI.GUI_Utilities import get_file, which_subj

def update_visit_log(subj:str,stfs_trkr_edit:str):
    acc_path = get_file("Path to visit log from accudemia")
    vl0 = pd.read_csv(acc_path, header=2)
    vl0 = vl0.drop_duplicates(["Student2","SubjectArea"],keep = "last")
    vl0["Student2"] = vl0["Student2"].str[1:]
    

    sheet = subjdict[subj]["ssheet"]

    refs0 = pd.read_excel(stfs_trkr_edit, sheet, header = 2).astype({"ID#" : "str"})
    refs = refs0["ID#"].to_numpy()
    refs = list(refs)
    vl0 = vl0[vl0["Student2"].isin(refs)]
    vl = vl0[["Student2","Date"]].to_numpy()
    book = load_workbook(stfs_trkr_edit)
    s = book[sheet]

    for i,stud in enumerate(vl):
        snum = stud[0]
        date = stud[1]
        index = refs.index(snum)
        row = str(index + 4)
        col = "K"
        cell = s[col + row].value
        if cell != None and s[col + row] != date:
            s["L" + row] = "Y"
        s[col + row] = date

    book.save(stfs_trkr_edit)


if __name__ == "__main__":
    update_visit_log(which_subj("Which subject would you like to update the visit log?").upper())



