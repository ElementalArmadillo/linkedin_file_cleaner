import pandas
import csv
import openpyxl
import re
import os
from pathlib import Path
from tkinter import *
from tkinter import filedialog
import glob2
from datetime import datetime
#import database_classes

#db_path=Path(os.getcwd(),'profiles.db')

#methods for cleaning algorithm
def read_line(line):
    if is_name_line(line):
        return "name_line_found"
    if is_name_line2(line):
        return "name_line2_found"
    if is_job_line(line):
        return "job_line_found"
    if is_job_line2(line):
        return "job_line2_found"
    if is_location_line(line):
        return "location_line_found"
    if is_end_of_profile(line):
        return "end_of_profile_found"
    return ""

def is_name_line(line):
    if re.search("^Profile result -", line[0], re.IGNORECASE):
        return True
    return False

def is_name_line2(line):
    if re.search(" degree contact$", line[0], re.IGNORECASE):
        return True
    return False

def get_name(line):
    profile_result_index = re.search("^Profile result -", line[0], re.IGNORECASE)
    name = line[0][profile_result_index.start() + 16:]
    name = name.strip()
    return name

def get_name2(line):
    if line[0].strip() == "":
        if len(line) > 1:
            return line[1].strip()
    return line[0].strip()

def is_job_line(line):
    if re.search("at ", line[0]) and re.search(" Go to ", line[0], re.IGNORECASE):
        if re.search("at ", line[0]).start() < line[0].lower().rfind(" go to "):
            return True
    return False

def is_job_line2(line):
    if len(line) > 1:
        if re.search("at ", line[1], re.IGNORECASE) and re.search(" Go to ", line[1], re.IGNORECASE):
            if re.search("at ", line[1], re.IGNORECASE).start() < line[1].lower().rfind(" go to "):
                return True
    return False

def get_job(line):
    at_index = re.search("at ", line[0])
    job = line[0][: at_index.start()]
    return job

def get_company(line):
    at_index = re.search("at ", line[0])
    go_to_index = line[0].lower().rfind(" go to ")
    company = line[0][at_index.start() + 3: go_to_index]
    return company

def get_job2(line):
    return line[0].strip()

def get_company2(line):
    at_index = re.search("at ", line[1], re.IGNORECASE)
    go_to_index = line[1].lower().rfind(" go to ")
    company = line[1][at_index.start() + 3: go_to_index]
    return company

def is_location_line(line):
    if re.search(" Area$", line[0]) or re.search(" Area,", line[0]):
        if re.search(" Area", line[0]).start() > 0:
            return True
    return False

def get_location(line):
    return line[0]

def is_end_of_profile(line):
    if  line[0].strip() == "" and line[1].strip() == "":
        return True
    if re.search("^Add tag", line[0], re.IGNORECASE) or re.search("^View profile", line[0], re.IGNORECASE) or re.search("^Show more", line[0], re.IGNORECASE) or re.search("^Profile result context -", line[0], re.IGNORECASE) or is_location_line(line):
        return True
    return False


#Simple GUI to select working folder
def select_folder():
    global dir_path
    root.withdraw()
    directory = filedialog.askdirectory(initialdir=os.getcwd(),title='Please select a directory')
    root.update()
    if len(directory) > 0:
        print ("You chose %s \n" % directory)
        dir_path = directory
        root.destroy()
    else:
        print ("No directory selected. Please try again\n")
        root.deiconify()

root = Tk()
root.geometry('350x50')
root.title("Select folder:")

l1 = Label(root, text="Please select the folder with the extracted files:", width = 40)
l1.grid(column=1, row=0)

b1 = Button(root, text="Select folder", width=15, highlightbackground="light blue", activebackground = "blue", command=select_folder)
b1.grid(column=1, row=1)

root.mainloop()

start_time = datetime.now()

try:
    os.chdir(dir_path)
except NameError:
    print("No directory selected. Using current directory instead.\n")

file_list = glob2.glob("*.xlsx")
print("Reading excel files from " + os.getcwd() + " ...\n")

temp_df_dir = "extracted_csv_files"
temp_df_path = Path(os.getcwd(), temp_df_dir)
cleaned_df_dir = "final_output_files"
cleaned_df_path = Path(os.getcwd(), cleaned_df_dir)
other_files_dir="extra_files"
other_files_path=Path(cleaned_df_path, other_files_dir)

num_files_read = 0
incomplete_profiles_dict_list = []
company_dict_list=[]
profiles_found_num = 0

if file_list:
    try:
        os.mkdir(temp_df_path)
    except FileExistsError:
        pass

    try:
        os.mkdir(cleaned_df_path)
    except FileExistsError:
        pass

    try:
        os.mkdir(other_files_path)
    except FileExistsError:
        pass

    program_result_save_location = other_files_path / "result.txt"
    program_result = open(program_result_save_location, "w")
    program_result.write("Results of the program: \n\n")
    program_result.write("Directory chosen: " + os.getcwd() + "\n\n")
    program_result.write("Files found: " + str(file_list) + "\n\n")

    #profiles_db=database_classes.ProfileDatabase(db_path)
    #companies_db=database_classes.CompanyDatabase(db_path)


#Looping through each file in folder
for file in file_list:

    file_name = str(file)[:-5]
    num_sheets = len(pandas.ExcelFile(file_name + ".xlsx").sheet_names)
    dict_list = []

    try:
        #cleaning algorithm
        for i in range(num_sheets):

            #saving each excel sheet as a csv file
            sheet_num = i + 1
            csv_file_name = file_name + "_sheet" + str(sheet_num) + ".csv"
            temp_df = pandas.read_excel(file_name + ".xlsx", sheet_name=i)
            temp_df_save_location = temp_df_path / csv_file_name
            temp_df.to_csv(temp_df_save_location, index=False)

            #reading each csv file
            with open(temp_df_save_location, encoding = 'utf8') as csvfile:
                csv_reader = csv.reader(csvfile, delimiter=',')
                d = {}
                previous_line = ""
                line_num = 0
                name_found = False
                work_found = False
                location_found = False

                #goes through each line and sorts out the expected data into a dict_list
                for line in csv_reader:
                    line_num += 1
                    if not line:
                        continue

                    result = read_line(line)

                    #act on result
                    if result == "":
                        previous_line = line
                        continue

                    if result == "name_line_found":
                        d["Name"] = get_name(line)
                        name_found = True
                    elif result == "name_line2_found":
                        if name_found == False:
                            d["Name"] = get_name2(previous_line)
                            name_found = True
                    elif result == "job_line_found":
                        d["Designation"] = get_job(line)
                        d["Company"] = get_company(line)
                        work_found = True
                    elif result == "job_line2_found":
                        if work_found == False:
                            d["Designation"] = get_job2(line)
                            d["Company"] = get_company2(line)
                            work_found = True
                    elif result == "location_line_found":
                        d["Location"] = get_location(line)
                        location_found = True

                    if not name_found and (result == "job_line_found" or result == "job_line2_found"):
                        if previous_line[0].strip() != "":
                            d["Name"] = get_name2(previous_line)
                            name_found = True

                    end_of_profile_found = is_end_of_profile(line)

                    if end_of_profile_found:
                        if name_found or work_found or location_found:
                            if not name_found:
                                d["Name"] = "NOT_FOUND"
                            if not work_found:
                                d["Designation"] = "NOT_FOUND"
                                d["Company"] = "NOT_FOUND"
                            if not location_found:
                                d["Location"] = "NOT_FOUND"

                            if not (name_found and work_found and location_found):
                                d["File"] = file_name + "_sheet" + str(sheet_num) + "_line " + str(line_num)
                                incomplete_profiles_dict_list.append(d)
                            else:
                                if d not in dict_list:
                                    dict_list.append(d)
                                    #profiles_db.add_to_db(d['Name'], d['Designation'], d['Company'], d['Location'])
                                company_d={"Company Name":d["Company"], "Location":d['Location']}
                                if company_d not in company_dict_list:
                                    company_dict_list.append(company_d)
                                    #companies_db.add_to_db(d['Company'], d['Location'])
                        d = {}
                        name_found = False
                        work_found = False
                        location_found = False

                    previous_line = line


        #creates DataFrame from dict_list and stores it in an excel file as the final output
        if dict_list:
            cleaned_df = pandas.DataFrame(dict_list)
            cleaned_df = cleaned_df[['Name', 'Designation', 'Company', 'Location']]

            cleaned_df_name = file_name + "_final_output.csv"
            cleaned_df_save_location = cleaned_df_path / cleaned_df_name
            cleaned_df.to_csv(cleaned_df_save_location, index = False)
            num_files_read += 1
            profiles_found_num += len(dict_list)
            print(file_name + ".xlsx read successfully")
            program_result.write(file_name + ".xlsx read successfully\n")
        else:
            print("No profiles found in " + file_name + ".xlsx")
            program_result.write("No profiles found in " + file_name + ".xlsx\n")

    except Exception as e:
        print("Error reading file " + file_name + ".xlsx: " + str(e))
        program_result.write("Error reading file " + file_name + ".xlsx: " + str(e))
        continue

if incomplete_profiles_dict_list:
    incomplete_profiles_df = pandas.DataFrame(incomplete_profiles_dict_list)
    incomplete_profiles_df = incomplete_profiles_df[['Name', 'Designation', 'Company', 'Location', 'File']]
    incomplete_profiles_df_name = "incomplete_profiles.csv"
    incomplete_profiles_df_save_location = other_files_path / incomplete_profiles_df_name
    incomplete_profiles_df.to_csv(incomplete_profiles_df_save_location, index = False)

if company_dict_list:
    company_df=pandas.DataFrame(company_dict_list)
    company_df=company_df[['Company Name','Location']]
    company_df_name="companies.csv"
    company_df_save_location=other_files_path / company_df_name
    company_df.to_csv(company_df_save_location, index = False)

end_time = datetime.now()
run_time = end_time - start_time
run_time_minutes = 0
run_time_seconds = run_time.seconds

if run_time_seconds >= 60:
    run_time_minutes = int(run_time_seconds / 60)
    run_time_seconds = run_time_seconds % 60

if num_files_read == 0:
    print("\nNo files found")
    print("Run time: " + str(run_time_minutes) + ":%02d" % run_time_seconds)
else:
    print("\n" + str(num_files_read) + " out of " + str(len(file_list)) + " files found and read successfully.")
    print("\n" + str(profiles_found_num) + " profiles found.")
    print(str(len(company_dict_list)) + " companies found.")
    print("\nRun time: " + str(run_time_minutes) + ":%02d" % run_time_seconds)
    #print("Profiles in profile table: "+str(len(profiles_db.view())))
    #print("Companies in company table: "+str(len(companies_db.view())))

    program_result.write("\n" + str(num_files_read) + " out of " + str(len(file_list)) + " files found and read successfully.")
    program_result.write("\n" + str(profiles_found_num) + " profiles found.")
    program_result.write("\n" + str(len(company_dict_list)) + " companies found.")
    program_result.write("\n" + str(len(incomplete_profiles_dict_list)) + " incomplete profiles found.")
    program_result.write("\nRun time: " + str(run_time_minutes) + ":%02d" % run_time_seconds)

    program_result.close()
