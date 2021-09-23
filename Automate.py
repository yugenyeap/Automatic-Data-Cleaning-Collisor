import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz
import re
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfile 
import time


import math



#Read the Category keyword files

#Accounting
with open('category_words/Accounting.txt','r') as file:
    AccountingStr = file.read()
#print(AccountingStr)
with open('category_words/Agriculture.txt','r') as file:
    AgricultureStr = file.read()
with open('category_words/Architecture.txt','r') as file:
    ArchitectureStr = file.read()
with open('category_words/Built.txt','r') as file:
    BuiltStr = file.read()
with open('category_words/Business.txt','r') as file:
    BusinessStr = file.read()
with open('category_words/Communications.txt','r') as file:
    CommunicationsStr = file.read()
with open('category_words/Computing.txt','r') as file:
    ComputingStr = file.read()
with open('category_words/CreativeArts.txt','r') as file:
    CreativeStr = file.read()
with open('category_words/EnvironmentalStudies.txt','r') as file:
    EnvironmentalStr = file.read()
with open('category_words/Dentistry.txt','r') as file:
    DentistryStr = file.read()
with open('category_words/Economics.txt','r') as file:
    EconomicsStr = file.read()
with open('category_words/Education.txt','r') as file:
    EducationStr = file.read()
with open('category_words/Engineering.txt','r') as file:
    EngineeringStr = file.read()
with open('category_words/Rehab.txt','r') as file:
    RehabStr = file.read()
with open('category_words/Health.txt','r') as file:
    HealthStr = file.read()
with open('category_words/Tourism.txt','r') as file:
    TourismStr = file.read()
with open('category_words/Humanities.txt','r') as file:
    HumanitiesStr = file.read()
with open('category_words/Languages.txt','r') as file:
    LanguagesStr = file.read()
with open('category_words/Law.txt','r') as file:
    LawStr = file.read()
with open('category_words/Para-legal.txt','r') as file:
    ParalegalStr = file.read()
with open('category_words/Pharmacy.txt','r') as file:
    PharmacyStr = file.read()
with open('category_words/Sport.txt','r') as file:
    SportStr = file.read()
with open('category_words/Mathematics.txt','r') as file:
    MathStr = file.read()
with open('category_words/Medicine.txt','r') as file:
    MedicineStr = file.read()
with open('category_words/Nursing.txt','r') as file:
    NursingStr = file.read()
with open('category_words/Psychology.txt','r') as file:
    PsychoStr = file.read()
with open('category_words/Sciences.txt','r') as file:
    ScienceStr = file.read()
with open('category_words/Social.txt','r') as file:
    SocialStr = file.read()
with open('category_words/Surveying.txt','r') as file:
    SurveyingStr = file.read()
with open('category_words/Vet.txt','r') as file:
    VetStr = file.read()

#GUI Variables
ws = Tk()
ws.title('Collisor - Automatic Data cleaner')
ws.geometry('700x800')

c1_val=IntVar()
c2_val=IntVar()
c3_val=IntVar()
c4_val=IntVar()
c5_val=IntVar()
c6_val=IntVar()
c7_val=IntVar()
c8_val=IntVar()
c9_val=IntVar()
c10_val=IntVar()
c11_val=IntVar()
c12_val=IntVar()
c13_val=IntVar()
c14_val=IntVar()
c15_val=IntVar()
c16_val=IntVar()
c17_val=IntVar()
c18_val=IntVar()
c19_val=IntVar()
c20_val=IntVar()
c21_val=IntVar()
c22_val=IntVar()
c23_val=IntVar()
c24_val=IntVar()
c25_val=IntVar()
c26_val=IntVar()
c27_val=IntVar()
c28_val=IntVar()
c29_val=IntVar()
c30_val=IntVar()


#Inputs for checkboxes checked and keyword entered
def checkboxSelected(*args):
    cat_checked = False
    cat_list = []
    cat_list.append(c1_val.get())
    cat_list.append(c2_val.get())
    cat_list.append(c3_val.get())
    cat_list.append(c4_val.get())
    cat_list.append(c5_val.get())
    cat_list.append(c6_val.get())
    cat_list.append(c7_val.get())
    cat_list.append(c8_val.get())
    cat_list.append(c9_val.get())
    cat_list.append(c10_val.get())
    cat_list.append(c11_val.get())
    cat_list.append(c12_val.get())
    cat_list.append(c13_val.get())
    cat_list.append(c14_val.get())
    cat_list.append(c15_val.get())
    cat_list.append(c16_val.get())
    cat_list.append(c17_val.get())
    cat_list.append(c18_val.get())
    cat_list.append(c19_val.get())
    cat_list.append(c20_val.get())
    cat_list.append(c21_val.get())
    cat_list.append(c22_val.get())
    cat_list.append(c23_val.get())
    cat_list.append(c24_val.get())
    cat_list.append(c25_val.get())
    cat_list.append(c26_val.get())
    cat_list.append(c27_val.get())
    cat_list.append(c28_val.get())
    cat_list.append(c29_val.get())
    cat_list.append(c30_val.get())

    for cat in cat_list:
        if cat==1:
            cat_checked = True
            break

    if not cat_checked:
        confirm_button.config(state="disabled")
    else:
        confirm_button.config(state="normal")



c1 = tk.Checkbutton(ws, text='Accounting',variable=c1_val, command=checkboxSelected)
c1.grid(row=20, columnspan=3, )
c2 = tk.Checkbutton(ws, text='Agriculture',variable=c2_val, command=checkboxSelected)
c2.grid(row=21, columnspan=3, )
c2 = tk.Checkbutton(ws, text='Architecture',variable=c3_val, command=checkboxSelected)
c2.grid(row=22, columnspan=3, )
c4 = tk.Checkbutton(ws, text='Built Environment',variable=c4_val, command=checkboxSelected)
c4.grid(row=23, columnspan=3, )
c5 = tk.Checkbutton(ws, text='Business and Management',variable=c5_val, command=checkboxSelected)
c5.grid(row=24, columnspan=3, )
c6 = tk.Checkbutton(ws, text='Communications',variable=c6_val, command=checkboxSelected)
c6.grid(row=25, columnspan=3, )
c7 = tk.Checkbutton(ws, text='Computing and Information Technology',variable=c7_val, command=checkboxSelected)
c7.grid(row=26, columnspan=3, )
c8 = tk.Checkbutton(ws, text='Creative Arts',variable=c8_val, command=checkboxSelected)
c8.grid(row=27, columnspan=3, )
c9 = tk.Checkbutton(ws, text='Environmental Studies',variable=c9_val, command=checkboxSelected)
c9.grid(row=28, columnspan=3, )
c10 = tk.Checkbutton(ws, text='Dentistry',variable=c10_val, command=checkboxSelected)
c10.grid(row=29, columnspan=3, )
c11 = tk.Checkbutton(ws, text='Economics',variable=c11_val, command=checkboxSelected)
c11.grid(row=30, columnspan=3, )
c12 = tk.Checkbutton(ws, text='Education and Training',variable=c12_val, command=checkboxSelected)
c12.grid(row=31, columnspan=3, )
c13 = tk.Checkbutton(ws, text='Engineering and Technology',variable=c13_val, command=checkboxSelected)
c13.grid(row=32, columnspan=3, )
c14 = tk.Checkbutton(ws, text='Rehabilitation',variable=c14_val, command=checkboxSelected)
c14.grid(row=33, columnspan=3, )
c15 = tk.Checkbutton(ws, text='Health Services and Support',variable=c15_val, command=checkboxSelected)
c15.grid(row=34, columnspan=3, )
c16 = tk.Checkbutton(ws, text='Tourism and Hospitality',variable=c16_val, command=checkboxSelected)
c16.grid(row=20, column=4, columnspan=3, )
c17 = tk.Checkbutton(ws, text='Humanities and Social Sciences',variable=c17_val, command=checkboxSelected)
c17.grid(row=21, column=4, columnspan=3, )
c18 = tk.Checkbutton(ws, text='Languages',variable=c18_val, command=checkboxSelected)
c18.grid(row=22, column=4, columnspan=3, )
c19 = tk.Checkbutton(ws, text='Law',variable=c19_val, command=checkboxSelected)
c19.grid(row=23, column=4, columnspan=3, )
c20 = tk.Checkbutton(ws, text='Para-legal Studies',variable=c20_val, command=checkboxSelected)
c20.grid(row=24, column=4, columnspan=3, )
c21 = tk.Checkbutton(ws, text='Pharmacy',variable=c21_val, command=checkboxSelected)
c21.grid(row=25, column=4,  columnspan=3, )
c22 = tk.Checkbutton(ws, text='Sport and Leisure',variable=c22_val, command=checkboxSelected)
c22.grid(row=26, column=4, columnspan=3, )
c23 = tk.Checkbutton(ws, text='Mathematics',variable=c23_val, command=checkboxSelected)
c23.grid(row=27, column=4, columnspan=3, )
c24 = tk.Checkbutton(ws, text='Medicine',variable=c24_val, command=checkboxSelected)
c24.grid(row=28, column=4, columnspan=3, )
c25 = tk.Checkbutton(ws, text='Nursing',variable=c25_val, command=checkboxSelected)
c25.grid(row=29, column=4, columnspan=3, )
c26 = tk.Checkbutton(ws, text='Psychology',variable=c26_val, command=checkboxSelected)
c26.grid(row=30, column=4, columnspan=3, )
c27 = tk.Checkbutton(ws, text='Sciences',variable=c27_val, command=checkboxSelected)
c27.grid(row=31, column=4, columnspan=3, )
c28 = tk.Checkbutton(ws, text='Social Work',variable=c28_val, command=checkboxSelected)
c28.grid(row=32, column=4, columnspan=3, )
c29 = tk.Checkbutton(ws, text='Surveying',variable=c29_val, command=checkboxSelected)
c29.grid(row=33, column=4, columnspan=3, )
c30 = tk.Checkbutton(ws, text='Veterinary Science',variable=c30_val, command=checkboxSelected)
c30.grid(row=34, column=4, columnspan=3, )

var = IntVar()

confirm_button = Button(ws, text="Confirm", command=lambda: [OnConfirmClick(),var.set(1)])
confirm_button.grid(row=17, columnspan=6, pady=30)
confirm_button.config(state="disabled")

cat_selected = []

def OnConfirmClick():
    global cat_selected
    cat_selected = []
    cat_list = []
    cat_list.append(c1_val.get())
    cat_list.append(c2_val.get())
    cat_list.append(c3_val.get())
    cat_list.append(c4_val.get())
    cat_list.append(c5_val.get())
    cat_list.append(c6_val.get())
    cat_list.append(c7_val.get())
    cat_list.append(c8_val.get())
    cat_list.append(c9_val.get())
    cat_list.append(c10_val.get())
    cat_list.append(c11_val.get())
    cat_list.append(c12_val.get())
    cat_list.append(c13_val.get())
    cat_list.append(c14_val.get())
    cat_list.append(c15_val.get())
    cat_list.append(c16_val.get())
    cat_list.append(c17_val.get())
    cat_list.append(c18_val.get())
    cat_list.append(c19_val.get())
    cat_list.append(c20_val.get())
    cat_list.append(c21_val.get())
    cat_list.append(c22_val.get())
    cat_list.append(c23_val.get())
    cat_list.append(c24_val.get())
    cat_list.append(c25_val.get())
    cat_list.append(c26_val.get())
    cat_list.append(c27_val.get())
    cat_list.append(c28_val.get())
    cat_list.append(c29_val.get())
    cat_list.append(c30_val.get())

    iter = 0
    for cat in cat_list:
        iter = iter + 1
        if cat==1:
            cat_selected.append(iter)
            print(cat_selected)
    
    c1_val.set(False)
    c2_val.set(False)
    c3_val.set(False)
    c4_val.set(False)
    c5_val.set(False)
    c6_val.set(False)
    c7_val.set(False)
    c8_val.set(False)
    c9_val.set(False)
    c10_val.set(False)
    c11_val.set(False)
    c12_val.set(False)
    c13_val.set(False)
    c14_val.set(False)
    c15_val.set(False)
    c16_val.set(False)
    c17_val.set(False)
    c18_val.set(False)
    c19_val.set(False)
    c20_val.set(False)
    c21_val.set(False)
    c22_val.set(False)
    c23_val.set(False)
    c24_val.set(False)
    c25_val.set(False)
    c26_val.set(False)
    c27_val.set(False)
    c28_val.set(False)
    c29_val.set(False)
    c30_val.set(False)

def prompt_for_category(coursename):
    #Set checked category variable to false
    global cat_checked
    cat_checked = False

    checkboxSelected()
    #Enable comfirm button
    
    #Create entry label to display coursename
    #Label(ws, text='Unrecognised coursename: ').grid(row=7)
    coursenamelabel = Label(ws, text=coursename, width=80)
    coursenamelabel.grid(row=10, columnspan=9)
    coursenamelabel.config(anchor=CENTER)
    coursenamelabel.configure(state="readonly")
    
    def entryInputted(*args):
        keyword = stringvar1.get()
        if not keyword:
            confirm_button.config(state="disabled")
            checkboxSelected()
        else:
            confirm_button.config(state="normal")
            checkboxSelected()

    stopwords = ['Bachelor Degree (Honours) ',
                 'Bachelor of ', 'Bachelor Degree ', 
                 'Diploma of ', 'Diploma in ',
                 'Graduate Certificate in ',
                 'Graduate Certificate of ',
                 'Undergraduate Certificate in ', 'Undergraduate Certificate',
                 'Undergraduate Certificate of ',
                 'Associate Degree in ','Associate Degree of ',
                 'Master of ','Master in ',
                 'Doctor of ','Doctor in ',
                 'Doctorate by ', 'Doctorate of '
                 'Doctorate by coursework of ', 'Doctorate of coursework of ',
                 'Certificate II in ','Certificate III in ',
                 'Certificate IV in ','Certificate V in '
                 ,'Certificate VI in ','Certificate VII in ',
                 'Advanced Diploma of ','Advanced Diploma in ',
                 'Graduate Diploma of ','Graduate Diploma in ' 
                 ]

    for word in stopwords:
        if word in coursename:
            coursename = coursename.replace(str(word),"")
            
    stringvar1 = tk.StringVar(ws,value=coursename)
    stringvar1.trace("w", entryInputted)

    #Create input for category entry
    #Label(ws, text='Enter Keyword: ').grid(row=8)
    new_cat = Entry(ws, width=80, justify=CENTER, textvariable=stringvar1)
    new_cat.grid(row=11, columnspan=9)
  
    #Wait for button to be pressed
    confirm_button.wait_variable(var)

    # while not cat_checked and not keyword_entered:
    #     warning = Label(ws, text='Enter keyword/category', foreground='red')
    #     warning.grid(row=4, columnspan=3, pady=10)
    # warning.destroy()

    #Get entry to be entered
    keyword = new_cat.get()
    print(keyword)   

    #Get chosen categories from category list
    categories = ""
    var.set(0)
    for cat in cat_selected:
        if cat==1:
            categories = "Accounting"
            with open('category_words/Accounting.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
        if cat==2:
            with open('category_words/Agriculture.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Agriculture"
            else:
                categories = categories + ",Agriculture"
        if cat==3:
            with open('category_words/Architecture.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Architecutre"
            else:
                categories = categories + ",Architecture"
        if cat==4:
            with open('category_words/Built.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Built Environment"
            else:
                categories = categories + ",Built Environment"
        if cat==5:
            with open('category_words/Business.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Business and Management"
            else:
                categories = categories + ",Business and Management"
        if cat==6:
            with open('category_words/Communications.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Communications"
            else:
                categories = categories + ",Communications"
        if cat==7:
            with open('category_words/Computing.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Computing and Information Technology"
            else:
                categories = categories + ",Computing and Information Technology"
        if cat==8:
            with open('category_words/CreativeArts.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Creative Arts"
            else:
                categories = categories + ",Creative Arts"
        if cat==9:
            with open('category_words/EnvironmentalStudies.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Environmental Studies"
            else:
                categories = categories + ",Environmental Studies"
        if cat==10:
            with open('category_words/Dentistry.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Dentistry"
            else:
                categories = categories + ",Dentistry"

        if cat==11:
            with open('category_words/Economics.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Economics"
            else:
                categories = categories + ",Economics"

        if cat==12:
            with open('category_words/Education.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Education and Training"
            else:
                categories = categories + ",Education and Training"
        if cat==13:
            with open('category_words/Engineering.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Engineering and Technology"
            else:
                categories = categories + ",Engineering and Technology"

        if cat==14:
            with open('category_words/Rehab.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Rehabilitation"
            else:
                categories = categories + ",Rehabilitation"

        if cat==15:
            with open('category_words/Health.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Health Services and Support"
            else:
                categories = categories + ",Health Services and Support"

        if cat==16:
            with open('category_words/Tourism.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Tourism and Hospitality"
            else:
                categories = categories + ",Tourism and Hospitality"
            
        if cat==17:
            with open('category_words/Humanities.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Humanities and Social Sciences"
            else:
                categories = categories + ",Humanities and Social Sciences"
        if cat==18:
            with open('category_words/Languages.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Languages"
            else:
                categories = categories + ",Languages"
        if cat==19:
            with open('category_words/Law.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Law"
            else:
                categories = categories + ",Law"
        if cat==20:
            with open('category_words/Para-legal.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Para-legal Studies"
            else:
                categories = categories + ",Para-legal Studies"
        if cat==21:
            with open('category_words/Pharmacy.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Pharmacy"
            else:
                categories = categories + ",Pharmacy"
        if cat==22:
            with open('category_words/Sport.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Sport and Leisure"
            else:
                categories = categories + ",Sport and Leisure"
        if cat==23:
            with open('category_words/Mathematics.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Mathematics"
            else:
                categories = categories + ",Mathematics"

        if cat==24:
            with open('category_words/Medicine.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Medicine"
            else:
                categories = categories + ",Medicine"
        if cat==25:
            with open('category_words/Nursing.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Nursing"
            else:
                categories = categories + ",Nursing"

        if cat==26:
            with open('category_words/Psychology.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Psychology"
            else:
                categories = categories + ",Psychology"
        if cat==27:
            with open('category_words/Sciences.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Sciences"
            else:
                categories = categories + ",Sciences"
        if cat==28:
            with open('category_words/Social.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Social Work"
            else:
                categories = categories + ",Social Work"
        if cat==29:
            with open('category_words/Surveying.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Surveying"
            else:
                categories = categories + ",Surveying"
        if cat==30:
            with open('category_words/Vet.txt','a') as keyword_list:
                keyword_list.write("|" + keyword)
                keyword_list.close()
            if not categories:
                categories = "Veterinary Science"
            else:
                categories = categories + ",Veterinary Science"
        
    new_cat.destroy()
    coursenamelabel.destroy()

    return categories
    #pass

hygiened_filename=''


def Hygiene(unhygiened,courses,providers,course_percentage,provider_percentage):
    #Read the CSV file
    #df = pd.read_csv('UQ-good-universities-data.csv')
    df = pd.read_csv(unhygiened)

    #remove extension from filename
    filename = hygiened_filename.split(".")
    filenameraw = filename[0]

    complete = Label(ws, text='Hygiene Complete!', foreground='green')
    hygienedfilenotice = Label(ws, text='Hygiened file saved as ' + hygiened_filename + '-hygiened.csv')
    datamatchingnotice = Label(ws, text='Fuzzy Matching data saved as ' + hygiened_filename + '-fuzzy-match-data.csv')
    complete.destroy()
    hygienedfilenotice.destroy()
    datamatchingnotice.destroy()
    complete = Label(ws, text='Hygiene Complete!', foreground='green')
    hygienedfilenotice = Label(ws, text='Hygiened file saved as ' + filenameraw + '-hygiened.csv')
    datamatchingnotice = Label(ws, text='Fuzzy Matching data saved as ' + filenameraw + '-fuzzy-match-data.csv')

    Unhygiened_filebtn.config(state="disabled")
    mastercoursebtn.config(state="disabled")
    masterproviderbtn.config(state="disabled")
    providerMatchSpinbox.config(state="disabled")
    courseMatchSpinbox.config(state="disabled")
    upld.config(state="disabled")
    status = Label(ws, text='Removing trailing spaces...', foreground='orange')
    status.grid(row=8, columnspan=6, pady=5)
    ws.update()

    #duplicates
    #print(df.duplicated(subset=['Course Name', 'School Name']).sum())

    #Read master providers and courses database
    Providers_Master = pd.read_csv(providers)
    Courses_Master = pd.read_csv(courses)
    
    #Providers_Master = pd.read_csv("Provider Database MASTER.csv")
    #Courses_Master = pd.read_csv("scrap_courseDetails.csv")

    #Clear all trailing spaces
    ##def clean_spaces(data):
    ##    if isinstance(data, float):
    ##        data = str(data)
    ##    data = re.sub(r'\s+',' ',data)
    ##    data = re.sub(r'\s+',' ',data)
    ##    data = re.sub('\n+',' ',data)
    ##    data = re.sub('\t+',' ',data)
    ##    data = re.sub('  ',' ',data)
    ##    return data

    ##df =  df.applymap(clean_spaces)

    

    #Add new words if not recognised
    #keyword_list.write("|Accounting")
    #keyword_list.close()
    #print(keyword_list)

    rows = df.shape[0] 
    #print(rows)


    #Course fee, extract just the numbers
    fee_num = df['Course Fee']
    #print(fee_num)
    df['Course Fee'] = fee_num.str.replace(',','') #remove comma
    df['Course Fee'] = fee_num.str.extract('([0-9]+)', expand=False) #extract number
    #print(fee_num)

    #Categorise course based on course name or inputted category
    course_name = df['Course Name']
    category_name = df['Category Name']

    status['text'] = 'Categorising courses...'
    ws.update()

    def categorise():
        #Try and categorise by category name given, if fails (because cell is empty) categorise by coursename
        #Course Name Boolean
        Accounting = course_name.str.contains(AccountingStr)
        Agriculture = course_name.str.contains(AgricultureStr)
        Architecture = course_name.str.contains(ArchitectureStr)
        Built_Environment = course_name.str.contains(BuiltStr)
        Business = course_name.str.contains(BusinessStr)
        Communications = course_name.str.contains(CommunicationsStr)
        Computing = course_name.str.contains(ComputingStr)
        Creative_Arts = course_name.str.contains(CreativeStr)
        Environment = course_name.str.contains(EnvironmentalStr)
        Dentistry = course_name.str.contains(DentistryStr)
        Economics = course_name.str.contains(EconomicsStr)
        Education = course_name.str.contains(EducationStr)
        Engineering = course_name.str.contains(EngineeringStr)
        Rehabilitation = course_name.str.contains(RehabStr)
        Health_Services_and_Support = course_name.str.contains(HealthStr)
        Tourism_and_Hospitality = course_name.str.contains(TourismStr)
        Humanities_and_Social_Sciences = course_name.str.contains(HumanitiesStr)
        Languages = course_name.str.contains(LanguagesStr)
        Law = course_name.str.contains(LawStr)
        Para_legal_Studies = course_name.str.contains(ParalegalStr)
        Pharmacy = course_name.str.contains(PharmacyStr)
        Sport_and_Leisure = course_name.str.contains(SportStr)
        Mathematics = course_name.str.contains(MathStr)
        Medicine = course_name.str.contains(MedicineStr)
        Nursing = course_name.str.contains(NursingStr)
        Psychology = course_name.str.contains(PsychoStr)
        Sciences = course_name.str.contains(ScienceStr)
        Social_Work = course_name.str.contains(SocialStr)
        Surveying = course_name.str.contains(SurveyingStr)
        Veterinary_Science = course_name.str.contains(VetStr)
        try:

            Accounting_category = category_name.str.contains(AccountingStr)
            Agriculture_category = category_name.str.contains(AgricultureStr)
            Architecture_category = category_name.str.contains(ArchitectureStr)
            Built_Environment_category = category_name.str.contains(BuiltStr)
            Business_category = category_name.str.contains(BusinessStr)
            Communications_category = category_name.str.contains(CommunicationsStr)
            Computing_category = category_name.str.contains(ComputingStr)
            Creative_Arts_category = category_name.str.contains(CreativeStr)
            Environment_category = category_name.str.contains(EnvironmentalStr)
            Dentistry_category = category_name.str.contains(DentistryStr)
            Economics_category = category_name.str.contains(EconomicsStr)
            Education_category = category_name.str.contains(EducationStr)
            Engineering_category = category_name.str.contains(EngineeringStr)
            Rehabilitation_category = category_name.str.contains(RehabStr)
            Health_Services_and_Support_category = category_name.str.contains(HealthStr)
            Tourism_and_Hospitality_category = category_name.str.contains(TourismStr)
            Humanities_and_Social_Sciences_category = category_name.str.contains(HumanitiesStr)
            Languages_category = category_name.str.contains(LanguagesStr)
            Law_category = category_name.str.contains(LawStr)
            Para_legal_Studies_category = category_name.str.contains(ParalegalStr)
            Pharmacy_category = category_name.str.contains(PharmacyStr)
            Sport_and_Leisure_category = category_name.str.contains(SportStr)
            Mathematics_category = category_name.str.contains(MathStr)
            Medicine_category = category_name.str.contains(MedicineStr)
            Nursing_category = category_name.str.contains(NursingStr)
            Psychology_category = category_name.str.contains(PsychoStr)
            Sciences_category = category_name.str.contains(ScienceStr)
            Social_Work_category = category_name.str.contains(SocialStr)
            Surveying_category = category_name.str.contains(SurveyingStr)
            Veterinary_Science_category = category_name.str.contains(VetStr)

            #First Category
            df['Category Name'] = np.where(Accounting_category, 'Accounting',
                                np.where(Agriculture_category, 'Agriculture',
                                np.where(Architecture_category, 'Architecture',
                                np.where(Built_Environment_category, 'Built Environment',
                                np.where(Business_category, 'Business and Management',
                                np.where(Communications_category, 'Communications',
                                np.where(Computing_category, 'Computing and Information Technology',
                                np.where(Creative_Arts_category, 'Creative Arts',
                                np.where(Environment_category, 'Environmental Studies',
                                np.where(Dentistry_category, 'Dentistry',
                                np.where(Economics_category, 'Economics',
                                np.where(Education_category, 'Education and Training',
                                np.where(Engineering_category, 'Engineering and Technology',
                                np.where(Rehabilitation_category, 'Rehabilitation',
                                np.where(Health_Services_and_Support_category, 'Health Services and Support',
                                np.where(Tourism_and_Hospitality_category, 'Tourism and Hospitality',
                                np.where(Humanities_and_Social_Sciences_category, 'Humanities and Social Sciences',
                                np.where(Languages_category, 'Languages',
                                np.where(Law_category, 'Law',
                                np.where(Para_legal_Studies_category, 'Para-legal Studies',
                                np.where(Pharmacy_category, 'Pharmacy',
                                np.where(Sport_and_Leisure_category, 'Sport and Leisure',
                                np.where(Mathematics_category, 'Mathematics',
                                np.where(Medicine_category, 'Medicine',
                                np.where(Nursing_category, 'Nursing',
                                np.where(Psychology_category, 'Psychology',
                                np.where(Sciences_category, 'Sciences',
                                np.where(Social_Work_category, 'Social Work',
                                np.where(Surveying_category, 'Surveying',
                                np.where(Veterinary_Science_category, 'Veterinary Science',
                                np.where(Accounting, 'Accounting',
                                np.where(Agriculture, 'Agriculture',
                                np.where(Architecture, 'Architecture',
                                np.where(Built_Environment, 'Built Environment',
                                np.where(Business, 'Business and Management',
                                np.where(Communications, 'Communications',
                                np.where(Computing, 'Computing and Information Technology',
                                np.where(Creative_Arts, 'Creative Arts',
                                np.where(Environment, 'Environmental Studies',
                                np.where(Dentistry, 'Dentistry',
                                np.where(Economics, 'Economics',
                                np.where(Education, 'Education and Training',
                                np.where(Engineering, 'Engineering and Technology',
                                np.where(Rehabilitation, 'Rehabilitation',
                                np.where(Health_Services_and_Support, 'Health Services and Support',
                                np.where(Tourism_and_Hospitality, 'Tourism and Hospitality',
                                np.where(Humanities_and_Social_Sciences, 'Humanities and Social Sciences',
                                np.where(Languages, 'Languages',
                                np.where(Law, 'Law',
                                np.where(Para_legal_Studies, 'Para-legal Studies',
                                np.where(Pharmacy, 'Pharmacy',
                                np.where(Sport_and_Leisure, 'Sport and Leisure',
                                np.where(Mathematics, 'Mathematics',
                                np.where(Medicine, 'Medicine',
                                np.where(Nursing, 'Nursing',
                                np.where(Psychology, 'Psychology',
                                np.where(Sciences, 'Sciences',
                                np.where(Social_Work, 'Social Work',
                                np.where(Surveying, 'Surveying',
                                np.where(Veterinary_Science, 'Veterinary Science',
                                        ''))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))
            
        except:
            #First Category
            df['Category Name'] = np.where(Accounting, 'Accounting',
                            np.where(Agriculture, 'Agriculture',
                            np.where(Architecture, 'Architecture',
                            np.where(Built_Environment, 'Built Environment',
                            np.where(Business, 'Business and Management',
                            np.where(Communications, 'Communications',
                            np.where(Computing, 'Computing and Information Technology',
                            np.where(Creative_Arts, 'Creative Arts',
                            np.where(Environment, 'Environmental Studies',
                            np.where(Dentistry, 'Dentistry',
                            np.where(Economics, 'Economics',
                            np.where(Education, 'Education and Training',
                            np.where(Engineering, 'Engineering and Technology',
                            np.where(Rehabilitation, 'Rehabilitation',
                            np.where(Health_Services_and_Support, 'Health Services and Support',
                            np.where(Tourism_and_Hospitality, 'Tourism and Hospitality',
                            np.where(Humanities_and_Social_Sciences, 'Humanities and Social Sciences',
                            np.where(Languages, 'Languages',
                            np.where(Law, 'Law',
                            np.where(Para_legal_Studies, 'Para-legal Studies',
                            np.where(Pharmacy, 'Pharmacy',
                            np.where(Sport_and_Leisure, 'Sport and Leisure',
                            np.where(Mathematics, 'Mathematics',
                            np.where(Medicine, 'Medicine',
                            np.where(Nursing, 'Nursing',
                            np.where(Psychology, 'Psychology',
                            np.where(Sciences, 'Sciences',
                            np.where(Social_Work, 'Social Work',
                            np.where(Surveying, 'Surveying',
                            np.where(Veterinary_Science, 'Veterinary Science',
                                    ''))))))))))))))))))))))))))))))

        #Any Additional Categories
        df['Category Name'] = np.where(Accounting & (df['Course Category']!='Accounting'),df['Category Name'] + ',Accounting',df['Category Name'])
        df['Category Name'] = np.where(Agriculture & (df['Category Name']!='Agriculture'),df['Category Name'] + ',Agriculture',df['Category Name'])
        df['Category Name'] = np.where(Architecture & (df['Category Name']!='Architecture'),df['Category Name'] + ',Architecture',df['Category Name'])
        df['Category Name'] = np.where(Built_Environment & (df['Category Name']!='Built Environment'),df['Category Name'] + ',Built Environment',df['Category Name'])
        df['Category Name'] = np.where(Business & (df['Category Name']!='Business and Management'),df['Category Name'] + ',Business and Management',df['Category Name'])
        df['Category Name'] = np.where(Communications & (df['Category Name']!='Communications'),df['Category Name'] + ',Communications',df['Category Name'])
        df['Category Name'] = np.where(Computing & (df['Category Name']!='Computing and Information Technology'),df['Category Name'] + ',Computing and Information Technology',df['Category Name'])
        df['Category Name'] = np.where(Creative_Arts & (df['Category Name']!='Creative Arts'),df['Category Name'] + ',Creative Arts',df['Category Name'])
        df['Category Name'] = np.where(Environment & (df['Category Name']!='Environmental Studies'),df['Category Name'] + ',Environmental Studies',df['Category Name'])
        df['Category Name'] = np.where(Dentistry & (df['Category Name']!='Dentistry'),df['Category Name'] + ',Dentistry',df['Category Name'])
        df['Category Name'] = np.where(Economics & (df['Category Name']!='Economics'),df['Category Name'] + ',Economics',df['Category Name'])
        df['Category Name'] = np.where(Education & (df['Category Name']!='Education and Training'),df['Category Name'] + ',Education and Training',df['Category Name'])
        df['Category Name'] = np.where(Engineering & (df['Category Name']!='Engineering and Technology'),df['Category Name'] + ',Engineering and Technology',df['Category Name'])
        df['Category Name'] = np.where(Rehabilitation & (df['Category Name']!='Rehabilitation'),df['Category Name'] + ',Rehabilitation',df['Category Name'])
        df['Category Name'] = np.where(Health_Services_and_Support & (df['Category Name']!='Health Services and Support'),df['Category Name'] + ',Health Services and Support',df['Category Name'])
        df['Category Name'] = np.where(Tourism_and_Hospitality & (df['Category Name']!='Tourism and Hospitality'),df['Category Name'] + ',Tourism and Hospitality',df['Category Name'])
        df['Category Name'] = np.where(Humanities_and_Social_Sciences & (df['Category Name']!='Humanities and Social Sciences'),df['Category Name'] + ',Humanities and Social Sciences',df['Category Name'])
        df['Category Name'] = np.where(Languages & (df['Category Name']!='Languages'),df['Category Name'] + ',Languages',df['Category Name'])
        df['Category Name'] = np.where(Law & (df['Category Name']!='Law'),df['Category Name'] + ',Law',df['Category Name'])
        df['Category Name'] = np.where(Para_legal_Studies & (df['Category Name']!='Para-legal Studies'),df['Category Name'] + ',Para-legal Studies',df['Category Name'])
        df['Category Name'] = np.where(Pharmacy & (df['Category Name']!='Pharmacy'),df['Category Name'] + ',Pharmacy',df['Category Name'])
        df['Category Name'] = np.where(Sport_and_Leisure & (df['Category Name']!='Sport and Leisure'),df['Category Name'] + ',Sport and Leisure',df['Category Name'])
        df['Category Name'] = np.where(Mathematics & (df['Category Name']!='Mathematics'),df['Category Name'] + ',Mathematics',df['Category Name'])
        df['Category Name'] = np.where(Medicine & (df['Category Name']!='Medicine'),df['Category Name'] + ',Medicine',df['Category Name'])
        df['Category Name'] = np.where(Nursing & (df['Category Name']!='Nursing'),df['Category Name'] + ',Nursing',df['Category Name'])
        df['Category Name'] = np.where(Psychology & (df['Category Name']!='Psychology'),df['Category Name'] + ',Psychology',df['Category Name'])
        df['Category Name'] = np.where(Sciences & (df['Category Name']!='Sciences'),df['Category Name'] + ',Sciences',df['Category Name'])
        df['Category Name'] = np.where(Social_Work & (df['Category Name']!='Social Work'),df['Category Name'] + ',Social Work',df['Category Name'])
        df['Category Name'] = np.where(Surveying & (df['Category Name']!='Surveying'),df['Category Name'] + ',Surveying',df['Category Name'])
        df['Category Name'] = np.where(Veterinary_Science & (df['Category Name']!='Veterinary Science'),df['Category Name'] + ',Veterinary Science',df['Category Name'])

    #Read the Category keyword files
    def readcategorynames():
        #Accounting
        with open('category_words/Accounting.txt','r') as file:
            global AccountingStr
            AccountingStr = file.read()
        #print(AccountingStr)
        with open('category_words/Agriculture.txt','r') as file:
            global AgricultureStr
            AgricultureStr = file.read()
        with open('category_words/Architecture.txt','r') as file:
            global ArchitectureStr
            ArchitectureStr = file.read()
        with open('category_words/Built.txt','r') as file:
            global BuiltStr
            BuiltStr = file.read()
        with open('category_words/Business.txt','r') as file:
            global BusinessStr
            BusinessStr = file.read()
        with open('category_words/Communications.txt','r') as file:
            global CommunicationsStr
            CommunicationsStr = file.read()
        with open('category_words/Computing.txt','r') as file:
            global ComputingStr
            ComputingStr = file.read()
        with open('category_words/CreativeArts.txt','r') as file:
            global CreativeStr
            CreativeStr = file.read()
        with open('category_words/EnvironmentalStudies.txt','r') as file:
            global EnvironmentalStr
            EnvironmentalStr = file.read()
        with open('category_words/Dentistry.txt','r') as file:
            global DentistryStr
            DentistryStr = file.read()
        with open('category_words/Economics.txt','r') as file:
            global conomicsStr
            conomicsStr = file.read()
        with open('category_words/Education.txt','r') as file:
            global EducationStr
            EducationStr = file.read()
        with open('category_words/Engineering.txt','r') as file:
            global EngineeringStr
            EngineeringStr = file.read()
        with open('category_words/Rehab.txt','r') as file:
            global RehabStr
            RehabStr = file.read()
        with open('category_words/Health.txt','r') as file:
            global HealthStr
            HealthStr = file.read()
        with open('category_words/Tourism.txt','r') as file:
            global TourismStr
            TourismStr = file.read()
        with open('category_words/Humanities.txt','r') as file:
            global HumanitiesStr
            HumanitiesStr = file.read()
        with open('category_words/Languages.txt','r') as file:
            global LanguagesStr
            LanguagesStr = file.read()
        with open('category_words/Law.txt','r') as file:
            global LawStr
            LawStr = file.read()
        with open('category_words/Para-legal.txt','r') as file:
            global ParalegalStr
            ParalegalStr = file.read()
        with open('category_words/Pharmacy.txt','r') as file:
            global PharmacyStr
            PharmacyStr = file.read()
        with open('category_words/Sport.txt','r') as file:
            global SportStr
            SportStr = file.read()
        with open('category_words/Mathematics.txt','r') as file:
            global MathStr
            MathStr = file.read()
        with open('category_words/Medicine.txt','r') as file:
            global MedicineStr
            MedicineStr = file.read()
        with open('category_words/Nursing.txt','r') as file:
            global NursingStr
            NursingStr = file.read()
        with open('category_words/Psychology.txt','r') as file:
            global PsychoStr
            PsychoStr = file.read()
        with open('category_words/Sciences.txt','r') as file:
            global ScienceStr
            ScienceStr = file.read()
        with open('category_words/Social.txt','r') as file:
            global SocialStr
            SocialStr = file.read()
        with open('category_words/Surveying.txt','r') as file:
            global SurveyingStr
            SurveyingStr = file.read()
        with open('category_words/Vet.txt','r') as file:
            global VetStr
            VetStr = file.read()
    readcategorynames()
    categorise()
    
    #Interrupt Program to ask for category if category is still empty
    #df['Category Name'] = np.where(df['Category Name']=="",prompt_for_category(),df['Category Name'])
    cat_column = df['Category Name']
    temp_cat=[]
    column_row = 0

    #Count unrecognised columns
    def get_unrec_courses():
        cat_column = df['Category Name']
        total_unrecognised_names = 0
        for data in cat_column:
            if data=="":
                total_unrecognised_names = total_unrecognised_names + 1
        return total_unrecognised_names

    #Prompt for category in unrecognised columns WHILE LOOP ATTEMPT
    while column_row < rows:
        unrecognised = Label(ws, text='Unrecognised Courses left: '+ str(get_unrec_courses()), foreground='orange')
        if df['Category Name'][column_row]=="":
            ws.update()
            unrecognised.grid(row=9, columnspan=6, pady=5)
            #print(PsychoStr)
            df['Category Name'][column_row] = prompt_for_category(df['Course Name'][column_row])
            unrecognised.destroy()
            readcategorynames()
            categorise()
        unrecognised.destroy()
        value = {}
        value.update({"cat name" : df['Category Name'][column_row]})
        temp_cat.append(value)
        column_row = column_row + 1
    
    #print(df['Course Name'][34])
    cat_data_conversion = pd.DataFrame(temp_cat)
    df['Category Name'] = cat_data_conversion['cat name']
    #print(df['Category Name'])

   


    #Set category number based on category name
    df['Course Category'] = df['Category Name']
    df['Course Category'] = df['Course Category'].str.replace('Accounting', '1')
    df['Course Category'] = df['Course Category'].str.replace('Agriculture', '2')
    df['Course Category'] = df['Course Category'].str.replace('Architecture', '3')
    df['Course Category'] = df['Course Category'].str.replace('Built Environment', '4')
    df['Course Category'] = df['Course Category'].str.replace('Business and Management', '5')
    df['Course Category'] = df['Course Category'].str.replace('Communications', '6')
    df['Course Category'] = df['Course Category'].str.replace('Computing and Information Technology', '7')
    df['Course Category'] = df['Course Category'].str.replace('Creative Arts', '8')
    df['Course Category'] = df['Course Category'].str.replace('Environmental Studies', '9')
    df['Course Category'] = df['Course Category'].str.replace('Dentistry', '10')
    df['Course Category'] = df['Course Category'].str.replace('Economics', '11')
    df['Course Category'] = df['Course Category'].str.replace('Education and Training', '12')
    df['Course Category'] = df['Course Category'].str.replace('Engineering and Technology', '13')
    df['Course Category'] = df['Course Category'].str.replace('Rehabilitation', '14')
    df['Course Category'] = df['Course Category'].str.replace('Health Services and Support', '15')
    df['Course Category'] = df['Course Category'].str.replace('Tourism and Hospitality', '16')
    df['Course Category'] = df['Course Category'].str.replace('Humanities and Social Sciences', '17')
    df['Course Category'] = df['Course Category'].str.replace('Languages', '18')
    df['Course Category'] = df['Course Category'].str.replace('Law', '19')
    df['Course Category'] = df['Course Category'].str.replace('Para-legal Studies', '20')
    df['Course Category'] = df['Course Category'].str.replace('Pharmacy', '21')
    df['Course Category'] = df['Course Category'].str.replace('Sport and Leisure', '22')
    df['Course Category'] = df['Course Category'].str.replace('Mathematics', '23')
    df['Course Category'] = df['Course Category'].str.replace('Medicine', '24')
    df['Course Category'] = df['Course Category'].str.replace('Nursing', '25')
    df['Course Category'] = df['Course Category'].str.replace('Psychology', '26')
    df['Course Category'] = df['Course Category'].str.replace('Sciences', '27')
    df['Course Category'] = df['Course Category'].str.replace('Social Work', '28')
    df['Course Category'] = df['Course Category'].str.replace('Surveying', '29')
    df['Course Category'] = df['Course Category'].str.replace('Veterinary Science', '30')

    #Get study level based on course name
    study_level=df['Study Level']
    Undergraduate = study_level.str.contains('Bachelor|Associate|Undergraduate|AssocDeg')
    Postgraduate = study_level.str.contains('Postgraduate|Doctor|Doctorate|Graduate|Graduate Diploma')
    Vocational = study_level.str.contains('Certificate|Diploma|Vocational')
    Microcredential = study_level.str.contains('RSA|First Aid|Short Course')

    df['Study Level'] = np.where(Undergraduate, 'Undergraduate',
                                          np.where(Postgraduate, 'Postgraduate',
                                                   np.where(Vocational, 'Vocational',
                                                            np.where(Microcredential, 'Microcredential','Microcredential'))))



    #Study Mode
    full_time_str = df['Study Mode'].str.contains("Full-time|full-time|Full Time|Full time|full time|Full-Time")
    part_time_str = df['Study Mode'].str.contains("Part-time|part-time|Part Time|Part time|part time|Part-Time")
    online_str = df['Study Mode'].str.contains("Online|online")

    #First Mode
    df['Study Mode'] = np.where(full_time_str, 'Full Time',
                            np.where(part_time_str, 'Part Time',
                                     np.where(df['Study Mode'].str.contains('Flexible Delivery'),'Full Time, Part Time, Online',
                                              np.where(df['Study Mode'].str.contains('On-campus'),'Full Time, Part Time',
                                                    np.where(online_str, 'Online', 'Full Time')))))

    #Second Mode if exists
    df['Study Mode'] = np.where(part_time_str,df['Study Mode'] + ',Part Time',df['Study Mode'])
    df['Study Mode'] = np.where(online_str,df['Study Mode'] + ',Online',df['Study Mode'])



    #Course Duration
    duration = df['Course Duration']



    #Extract the duration
    #Remove the mode (full-time partime etc.)
    df['Course Duration'] = duration.str.replace('Full-time|full-time|Full Time|Full time|full time|Full-Time','')
    df['Course Duration'] = duration.str.replace('Part-time|part-time|Part Time|Part time|part time|Part-Time','')
    df['Course Duration'] = duration.str.replace('Online|online','')
    df['Course Duration'] = duration.str.replace('Flexible Delivery','')
    df['Course Duration'] = duration.str.replace('Duration','')

    df['Course Duration'] = duration.str.replace(':','',1)
    df['Course Duration'] = duration.str.replace(':',',')

    #Remove Trailing Spaces
    df['Course Duration'] = duration.str.replace(r'\s+',' ')
    df['Course Duration'] = duration.str.replace('\n+',' ')
    df['Course Duration'] = duration.str.replace('\t+',' ')
    df['Course Duration'] = duration.str.replace('  ',' ')

    ###########################################
    ##Fuzzy Match for course and provider ids##
    ###########################################

    status['text'] = 'Fuzzy Matching...'

    def match_name(name, list_names, list_ids, min_score=0):
        # -1 score incase we don't get any matches
        max_score = -1
        # Returning empty name for no match as well
        max_name = ""
        id_=""
        # Iternating over all names in the other
        for idx, name2 in enumerate(list_names):
            #Finding fuzzy match score
            score = fuzz.ratio(str(name), str(name2))
            # Checking if we are above our threshold and have a better score
            if ((score > min_score) & (score > max_score))or score==100:
                max_name = name2
                max_score = score
                id_= list_ids.iloc[idx]
        return (max_name, max_score, id_)

    def match_provider(name, list_providers, list_ids, min_score=0):
        # -1 score incase we don't get any matches
        max_score = -1
        # Returning empty name for no match as well
        max_provider = ""
        id_=""
        # Iternating over all names in the other
        for idx, name2 in enumerate(list_providers):
            #Finding fuzzy match score
            score = fuzz.ratio(str(name), str(name2))
            # Checking if we are above our threshold and have a better score
            if ((score > min_score) & (score > max_score))or score==100:
                max_provider = name2
                max_score = score
                id_= list_ids.iloc[idx]
        return (max_provider, max_score, id_)

    # List for dicts for easy dataframe creation
    dict_list = []

    #counting
    row = 1

    #GUI things
    # pb1 = Progressbar(
    #     ws, 
    #     orient=HORIZONTAL, 
    #     length=rows, 
    #     mode='determinate'
    #     )
    # pb1.grid(row=4, columnspan=3, pady=20)
        
    

    #iterating Scrapped database
    for data in df.iloc[ : , [5, 11] ].iterrows():
        #Gui things
        ws.update()
        #print(pb1['value'])
        #pb1['value'] += round(row/rows)
        matchingcoursecount = Label(ws, text='Course: '+ str(row) +'/' + str(rows), foreground='orange')
        matchingcoursecount.grid(row=9, columnspan=6, pady=10)
        
        # Get provider ID
        Provider_ID = ""
        #print(data[1][1]) #Business name

        #Compare scrapped business name with business name in provider database. If score above 90, use new name
        Provider_match_check = match_provider((data[1][1]), Providers_Master['Business Name'],Providers_Master['Institution Id'], provider_percentage)
        #print(Provider_match_check)
        #Get Provider ID
        if Provider_match_check[1] == -1: #if match not found above threshold
            #Provider_ID = ""
            business_name = data[1][1]
        else:
            business_name = Provider_match_check[0]
            Provider_ID = Provider_match_check[2]

        #Get courses with provider ID
        CoursesUnderProviderIndex = Courses_Master[Courses_Master['Provider_ID']==Provider_ID].index.values #Get index of row where business names match
        #print(CoursesUnderProviderIndex) #indexes of rows that have matching provider ID
        CoursesUnderProvider = (Courses_Master.loc[CoursesUnderProviderIndex, ['Course_Name', 'Course_ID']])
        #print(CoursesUnderProvider['Course_ID'])
                
        # Use our method to find best match, we can set a threshold here
        match = match_name((data[1][0]), CoursesUnderProvider['Course_Name'],CoursesUnderProvider['Course_ID'], course_percentage)

        #Get CourseID
        if match[1] == -1: #if match not found above threshold
            Course_ID = ""
            final_match_name = data[1][0]
        else:
            Course_ID = CoursesUnderProvider.at[CoursesUnderProviderIndex[0],'Course_ID']
            final_match_name = match[0]
        
        # New dict for storing data
        dict_ = {}
        dict_.update({"Provider_id" : Provider_ID})
        dict_.update({"Business_name" : data[1][1]})
        dict_.update({"Matched_Business_name" : Provider_match_check[0]})
        dict_.update({"Final_Business_name" : business_name})
        dict_.update({"Matched_Business_score" : Provider_match_check[1]})
        dict_.update({"course_name" : data[1][0]})
        dict_.update({"match_name" : match[0]})
        dict_.update({"final_match_name" : final_match_name})
        dict_.update({"Existing_course_id" : match[2]})
        dict_.update({"score" : match[1]})
        dict_list.append(dict_)

        row = row + 1
        
    

    merge_table = pd.DataFrame(dict_list)

    

    merge_table.to_csv(filenameraw+'-fuzzy-match-data.csv', index=False)

    df['Provider Id'] = merge_table['Provider_id']
    df['Course Id'] = merge_table['Existing_course_id']

    #Display results
    df.to_csv(filenameraw+'-hygiened.csv', index=False,mode='w+' )  

    #More gui things
    matchingcoursecount.destroy()
    status.destroy()
    Unhygiened_filebtn.config(state="normal")
    mastercoursebtn.config(state="normal")
    masterproviderbtn.config(state="normal")
    providerMatchSpinbox.config(state="normal")
    courseMatchSpinbox.config(state="normal")
    upld.config(state="normal")
    complete.grid(row=10, columnspan=6, pady=10)
    hygienedfilenotice.grid(row=11, columnspan=6)
    datamatchingnotice.grid(row=12, columnspan=6)

#####END FUZZY MATCH#####

########START GUI########

#Integer spinboxes for fuzzy matching percentage
coursePercent= IntVar(ws)
coursePercent.set("90")

courseMatchPercentageLabel = Label(ws, text='Select course name matching accuracy')
courseMatchPercentageLabel.grid(row=4, padx=10, pady=10)

courseMatchSpinbox = Spinbox(ws,from_= 50, to = 100,textvariable=coursePercent)
courseMatchSpinbox.grid(row=5)

providerPercent= IntVar(ws)
providerPercent.set("90")

providerMatchPercentageLabel = Label(ws, text='Select provider name matching accuracy')
providerMatchPercentageLabel.grid(row=4, column=4, padx=10, pady=10)

providerMatchSpinbox = Spinbox(ws,from_= 50, to = 100,textvariable=providerPercent)
providerMatchSpinbox.grid(row=5, column=4)

CPercentage = coursePercent.get()
PPercentage = providerPercent.get()

hygiened_file=''
courses_file=''
providers_file=''



def open_unhygiened():
    hygiened_file_path = fd.askopenfilename(filetypes=[('CSV Files', '*csv')])
    if hygiened_file_path is not None:
        filename = hygiened_file_path.split("/")
        Unhygiened_file['text'] = "Unhygiened file : " + filename[-1]
        global hygiened_file
        hygiened_file = hygiened_file_path
        global hygiened_filename
        hygiened_filename = filename[-1]
        print(hygiened_filename)
        pass
    
def open_courses():
    courses_file_path = fd.askopenfilename(filetypes=[('CSV Files', '*csv')])
    if courses_file_path is not None:
        filename = courses_file_path.split("/")
        mastercourse['text'] ="Master Course file : " + filename[-1]
        global courses_file
        courses_file = courses_file_path
        pass
    
def open_providers():
    providers_file_path = fd.askopenfilename(filetypes=[('CSV Files', '*csv')])
    if providers_file_path is not None:
        filename = providers_file_path.split("/")
        masterprovider['text'] ="Master Provider file : " + filename[-1]
        global providers_file
        providers_file = providers_file_path
        pass

def File_Chosen():
    UploadAssert = Label(ws, text='Upload files!', foreground='red')
    if hygiened_file!='' and courses_file!='' and providers_file!='':
        #print(hygiened_file)
        #print(courses_file)
        #print(providers_file)
        UploadAssert.destroy()

        Hygiene(hygiened_file,courses_file,providers_file,CPercentage,PPercentage)
        #Hygiene('Courseseeker-unhygiened-JUNE.csv',"scrap_courseDetails.csv","Provider Database MASTER.csv",CPercentage,PPercentage)
    else:
        UploadAssert.grid(row=6, columnspan=6)
        
    
    
Unhygiened_file = Label(
    ws, 
    text='Select Unhygiened file'
    )
Unhygiened_file.grid(row=0, column=0, padx=10)

Unhygiened_filebtn = Button(
    ws, 
    text ='Choose File', 
    command = lambda:open_unhygiened()
    ) 
Unhygiened_filebtn.grid(row=0, column=1)

mastercourse = Label(
    ws, 
    text='Select Master Course file '
    )
mastercourse.grid(row=1, column=0, padx=10)

mastercoursebtn = Button(
    ws, 
    text ='Choose File ', 
    command = lambda:open_courses()
    ) 
mastercoursebtn.grid(row=1, column=1)

masterprovider = Label(
    ws, 
    text='Select Master Provider file '
    )
masterprovider.grid(row=2, column=0, padx=10)

masterproviderbtn = Button(
    ws, 
    text ='Choose File', 
    command = lambda:open_providers()
    ) 
masterproviderbtn.grid(row=2, column=1)

upld = Button(
    ws, 
    text='Start Hygiene', 
    command=File_Chosen
    #command=Hygiene('UQ-good-universities-data.csv',"scrap_courseDetails.csv","Provider Database MASTER.csv")
    )
upld.grid(row=7, columnspan=6, pady=10)

copyrightLabel = Label(ws, text='© Created by Yu Gen Yeap for Collisor Pty Ltd')
copyrightLabel.grid(row=40, padx=5, pady=30)




# new_cat = Entry(ws)
# new_cat.grid(row=8, columnspan=3, pady=30)

# c1_val=IntVar()
# c2_val=IntVar()
# c1 = tk.Checkbutton(ws, text='Accounting',variable=c1_val)
# c1.pack()
# c2 = tk.Checkbutton(ws, text='Agrictulture',variable=c2_val)
# c2.pack()

# cat_selected = ""

# def OnButtonClick():
#     for c in (c1,c2):
#         print(c.get())

# confirm_button = Button(ws, text="Confirm", command=OnButtonClick)
# confirm_button.grid(row=17, columnspan=3, pady=30)
    


#var = tk.IntVar()
#button = tk.Button(root, text="Click Me", command=lambda: var.set(1))


ws.mainloop()


##END GUI##











##CODE CEMETARY##

#Prompt for category in unrecognised columns
    # for data in df['Category Name']:
    #     unrecognised = Label(ws, text='Unrecognised Courses left: '+ str(get_unrec_courses()), foreground='orange')
    #     if data=="":
    #         ws.update()
    #         unrecognised.grid(row=9, columnspan=6, pady=5)
    #         #print(PsychoStr)
    #         data = prompt_for_category(df['Course Name'][column_row])
    #         readcategorynames()
    #         categorise()
    #     unrecognised.destroy()
    #     value = {}
    #     value.update({"cat name" : data})
    #     temp_cat.append(value)
    #     column_row = column_row + 1

##def Hygiene():
##    pb1 = Progressbar(
##        ws, 
##        orient=HORIZONTAL, 
##        length=300, 
##        mode='determinate'
##        )
##    pb1.grid(row=4, columnspan=3, pady=20)
##    for i in range(5):
##        ws.update_idletasks()
##        pb1['value'] += 20
##        time.sleep(1)
##        Label(ws, text='Hygiening: '+ str(i) +'/5', foreground='orange').grid(row=5, columnspan=3, pady=10)
##    pb1.destroy()
##    Label(ws, text='Hygiene Complete!', foreground='green').grid(row=4, columnspan=3, pady=10)
                                                                                                                                                                                                                                                                                                  
#for data in df.iloc[ : , [5, 6, 7, 8] ].iterrows():
    #print(data)
    
##for x in range(rows):
##    #print('---------course-----------')
##    #print(x)
##    for data in df.iloc[ x , ]:
##        #Remove all trailing white spaces
##        if isinstance(data, float):
##            data = str(data)
##        data = re.sub(r'\s+',' ',data)
##        data = re.sub(r'\s+',' ',data)
##        data = re.sub('\n+',' ',data)
##        data = re.sub('\t+',' ',data)
##        data = re.sub('  ',' ',data)
##        #print(data)
      

#iterate through each row in the csv file
#for data in df.iloc[ : , [5, 11] ].iterrows():
    #print(row)

##for data in df.iloc[ 7 , ]:
##    print(data)
##print("----------------------------------")
##

##def remove_spaces(item):
##    if isinstance(item, float):
##        item = str(item)
##    item = re.sub(r'\s+',' ',item)
##    item = re.sub('\n',' ',item)
##    item = re.sub('\t',' ',item)
##    item = re.sub('  ',' ',item)
##    return item

#df.applymap(remove_spaces)
#pd.set_option("display.max_rows", None, "display.max_columns", None)
#print(df)
    

#5 = course name
#11 = School Name
#print(df.iloc[:,[5, 11]])

#Extract fields in a column with regex
#extr = df['Date of Publication'].str.extract(r'^(\d{4})', expand=False)
#extr.head()


