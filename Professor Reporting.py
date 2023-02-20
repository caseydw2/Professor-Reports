# -*- coding: utf-8 -*-
"""
Created on Mon Dec  5 11:11:11 2022

To Use:

Download the most current Writing Studio Smartsheet and put it in professor reports folder (Check name. Most current name is Fall 2022 Writing Studio)
Go to accudemia, select the week you would like to compile. Check the box that says Allowed report visit. Set Subject area to ENGL. Download as csv.
Put the accudemia report in the professor reports folder and rename it "Accudemia Report"
Run the Debugger


@author: E1448105
"""
import pandas as pd
import datetime as dt
from dateutil.relativedelta import relativedelta, FR
from os import makedirs
import _pickle as cPickle
from GUI.GUI_Utilities import get_file

#We can get the files from a folder
smart_sheet = pd.read_excel(get_file("Writing Studio Smartsheet"))
emails = smart_sheet.loc[smart_sheet["Type of Session"] == "Email"]
emails_to_report = emails.loc[emails["Report Visit?"] == "Yes"]
# emails_to_report = emails_to_report[["Student's Name","Tutoring:Writing",]]
accudemia_sheet = pd.read_csv(get_file("Accudemia Report"))
acc_report = accudemia_sheet[["FullName","Services","Course","Task","SignInTime","SignOutTime","Period","Instructor","Comments","Place"]]
acc_report["Instructor"] = acc_report["Instructor"].fillna("Professor Unknown")
acc_no_prof = acc_report[acc_report["Instructor"].isin([" ","None","N/A","Professor Unknown","NaN",'nan'])]
acc_prof_avail = acc_report[~acc_report["Instructor"].isin([" ","None","N/A","BLANK","Professor Unknown","NaN","nan"])]


professor_name_inputs = smart_sheet["Instructor's Name"].unique()


def make_professor_dir(professor_name):
    today = dt.datetime.today()
    this_friday = today + relativedelta(weekday=FR(-1))
    this_friday = this_friday.strftime("%m-%d-%y")
    #get file from folder
    path = r"C:/Users/E1448105/OneDrive - Metropolitan Community College - Kansas City/Programs/Professor Reports Documents/Professor Reports for Leslie"
    dir_path = path+f"/{this_friday}/{professor_name}"

    try:
        makedirs(dir_path)
    except FileExistsError:
        pass

def make_professor_excel(df,professor_name):
    if type(professor_name) != float:
        last_name = professor_name.split(" ")[-1]
        today = dt.datetime.today()
        this_friday = today + relativedelta(weekday=FR(-1))
        this_friday = this_friday.strftime("%m-%d-%y")
        last_friday = today + relativedelta(weekday=FR(-2))
        last_friday = last_friday.strftime("%m.%d")
        today = today.strftime("%m.%d") 
        file_name = str(last_name)+ " " + str(last_friday) +"-"+str(today) +".xlsx"
        path = f"./Professor Reports for Leslie/{this_friday}/{professor_name}/{file_name}"
        df.to_excel(path,sheet_name = f"{professor_name}'s {today} Report",index = False)
        print(f"{file_name} has been created")


def isin_acc_prof_avail(fullname):
    student_name_list = acc_prof_avail["FullName"].unique()
    return fullname in student_name_list

def findprof_in_acc_prof_avail(fullname):
    prof = acc_prof_avail.loc[acc_prof_avail["FullName"] == fullname]["Instructor"].unique()[0]
    return prof

def make_name_var(student):
    student["First Name"] = student["FullName"].split(" ")[0]
    student["Last Name"] = student["FullName"].split(" ")[-1]
    student["FNLN"] = student["First Name"] + " " + student["Last Name"]
    if len(student["FullName"].split(" ")) >= 3:
        student["First Name"] = student["FullName"].split(" ")[0]
        student["Last Name"] = student["FullName"].split(" ")[-1]
        student["Middle Name"] = student["FullName"].split(" ")[1]
        student["FNMNLN"] = student["First Name"] + " " +student["Middle Name"] + " " + student["Last Name"]
    else:
        student["FNMNLN"] = student["FNLN"]
    return [student["FullName"], student["FNLN"], student["FNMNLN"]]

def isin_smart_sheets(namelist):
    with open("nicknames.pickle","rb") as nn_file_rd:
        nickname_dict = cPickle.load(nn_file_rd)
    try:
        namelist = nickname_dict[namelist[0]]
    except KeyError:
        nickname_dict[namelist[0]] = namelist
    for name in namelist:
        if name in smart_sheet["Student's Name"].unique():
            return name
    else:
        user_input_name = input(f"Is {namelist} in starfish? If so, please input how their name appears in starfish. If not, press Enter. ")
        if user_input_name:
            namelist.append(user_input_name)
            nickname_dict[namelist[0]] = namelist
            with open("nicknames.pickle","wb") as nn_file_wd:
                cPickle.dump(nickname_dict,nn_file_wd)
            return user_input_name
        else:
            return False


#I can break up the below into seperate functions
def findprof_in_smart_sheets(name):
    with open("prof_dict.pickle","rb") as prof_dict_rb:
        prof_dict = cPickle.load(prof_dict_rb)
    inputs = smart_sheet.loc[smart_sheet["Student's Name"] == name]["Instructor's Name"].unique()
    bool_list = [prof_input in prof_dict.keys() for prof_input in inputs]
    trues = [i for i,x in enumerate(bool_list) if x]
    if any(bool_list):
        prof_input = inputs[trues[0]]
        prof = prof_dict[prof_input]
        dummy_dict = dict.fromkeys(inputs,prof)
        prof_dict = {**prof_dict,**dummy_dict}
        with open("prof_dict.pickle","wb") as prof_file_wd:
            prof_file_wd.seek(0)
            cPickle.dump(prof_dict,prof_file_wd)
            prof_file_wd.close()
        return prof
    else:
        prof = "next"
        i = 0
        while prof == "next" and i < len(inputs):
            prof = input(f"Who is this instructor: {inputs[i]} (if unknown, type 'next'). ")
            i+=1
        if prof == "next":
            return "Professor Unknown"
        else:
            dummy_dict = dict.fromkeys(inputs,prof)
            prof_dict = {**prof_dict,**dummy_dict}
            with open("prof_dict.pickle","wb") as prof_file_wd:
                prof_file_wd.seek(0)
                cPickle.dump(prof_dict,prof_file_wd)
                prof_file_wd.close()
            return prof


def add_prof(student,prof):
    student["Instructor"] = prof

def create_reports(df):
    instructors = df["Instructor"].unique()
    for prof in instructors:
        if type(prof) == float:
            prof = "Unknown Professor"
        make_professor_dir(prof)
        prof_df = df.loc[df["Instructor"] == prof]
        make_professor_excel(prof_df,prof)


def professor_reports():
    acc_prof_avail0 = acc_prof_avail
    acc_no_prof0 = acc_no_prof
    for index, student in acc_no_prof0.iterrows():
        if student["FullName"] == "Professor Unknown":
            pass
        else:
            fn = student["FullName"]
            namelist = make_name_var(student)
        
        if isin_acc_prof_avail(fn):
            prof = findprof_in_acc_prof_avail(fn)
            acc_no_prof0.loc[index,"Instructor"] = prof
                    
        elif isin_smart_sheets(namelist):
            name = isin_smart_sheets(namelist)
            prof = findprof_in_smart_sheets(name)
            acc_no_prof0.loc[index,"Instructor"] = prof

        else:
            prof = input(f"Who is {fn}'s instructor? If unknown, press Enter. ")
            if type(prof) != None:
                acc_no_prof0.loc[index,"Instructor"] = prof
            else:
                acc_no_prof0.loc[index,"Instructor"] = "Professor Unknown"
                print(f"Could not find {fn}'s instructor.")
    acc_no_prof01 = acc_no_prof0.loc[acc_no_prof0["Instructor"] == "Professor Unknown"]
    acc_prof_found = acc_no_prof0[acc_no_prof0["Instructor"] != "Professor Unknown"]
    x = pd.concat([acc_prof_avail0,acc_prof_found], ignore_index= True)
    create_reports(x)
    create_reports(acc_no_prof01)


if __name__ == "__main__":
    professor_reports()