# -*- coding: utf-8 -*-
"""
Created on Thu Jan 12 15:21:24 2023

@author: fedig
"""

import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog



#from tkinter import *
'''
from tkinter import ttk
from tkinter import messagebox
from tkinter import PhotoImage
'''

import script





def textBoxIsEmpty(textbox,btn):
    try:
        if str(textbox.get(1.0,tk.END)).isspace() :
            btn['state']=tk.DISABLED
            tk.messagebox.showerror('Error', 'Number of files is missing !')
            #print('The Widget Is Empty')
        elif int(textbox.get(1.0,tk.END))==0 :
            btn['state']=tk.DISABLED
            tk.messagebox.showerror('Error', '0 ???!!!')
        else:
            #print('The Widget Is Not Empty')
            int(textbox.get(1.0,tk.END))
            btn['state']=tk.NORMAL
    except Exception:
        btn['state']=tk.DISABLED
        tk.messagebox.showerror('Error', 'Enter an Integer !')

def proceedPaymentWindow():
    tk.messagebox.showwarning('Warning !!!', 'Use it on your own responsibility.')
    newWindow = tk.Tk()
    
    years_options = [str(i) for i in range (2023,2050)]
    months_options= ["January","February","March","April","May","June","July","August","September","October","November","December"]
    value_inside_year = tk.StringVar(newWindow)
    value_inside_month = tk.StringVar(newWindow)
    value_inside_year.set("Select the expiration year")
    value_inside_month.set("Select the expiration month")

    newWindow.title("Referring member info & Payment Info")
 
    newWindow.geometry("300x350")
    

    
    name = tk.Label(newWindow, text = "Referring Member Full Name")
    name.pack()
    
    memberName = tk.Text(newWindow, height = 1, width = 20)
    memberName.pack()
    
    memberIDD = tk.Label(newWindow, text = "Referring Member IEEE ID")
    memberIDD.pack()
    
    memberID = tk.Text(newWindow, height = 1, width = 20)
    memberID.pack()
    
    cardNumberr = tk.Label(newWindow, text = "Credit Card Number")
    cardNumberr.pack()

    cardNumber = tk.Text(newWindow, height = 1, width = 20)
    cardNumber.pack()
    
    
    
    cardYearr = tk.Label(newWindow, text = "Credit Card Expiration Year")
    cardYearr.pack()
    
    cardYear = tk.OptionMenu(newWindow, value_inside_year, *years_options)
    cardYear.pack()
    
    cardMonthh = tk.Label(newWindow, text = "Credit Card Expiration Month")
    cardMonthh.pack()

    cardMonth = tk.OptionMenu(newWindow, value_inside_month, *months_options)
    cardMonth.pack()
    
    
    cardOwnerr = tk.Label(newWindow, text = "Credit Card Owner Full Name")
    cardOwnerr.pack()

    cardOwner = tk.Text(newWindow, height = 1, width = 20)
    cardOwner.pack()
    
    
    

    
    excelFile = filedialog.askopenfile(mode='rb',title="Choose an excel file !")
    startBtn = tk.Button(newWindow,text="START",state = tk.DISABLED,
                         command=lambda:[script.mainWithPayment(excelFile.name,memberName.get(1.0,tk.END),memberID.get(1.0,tk.END),
                                                                cardNumber.get(1.0,tk.END), value_inside_year.get() ,
                                                                value_inside_month.get(), cardOwner.get(1.0,tk.END))])
    
    if excelFile!=None:
        startBtn["state"]= tk.NORMAL
    else:
        startBtn["state"]= tk.DISABLED
    
    startBtn.pack()
    
    

def notProceedPaymentWindow():

    newWindow = tk.Tk()
 

    newWindow.title("Referring member info")
 
    newWindow.geometry("250x160")
    name = tk.Label(newWindow, text = "Referring Member Full Name")
    name.pack()
    
    memberName = tk.Text(newWindow, height = 1, width = 20)
    memberName.pack()
    
    memberID = tk.Label(newWindow, text = "Referring Member IEEE ID")
    memberID.pack()
    
    memberID = tk.Text(newWindow, height = 1, width = 20)
    memberID.pack()
    
    '''
    #th.Thread(target=).start()
    
    progressBar = ttk.Progressbar(newWindow, orient = HORIZONTAL,length = 100, mode = 'determinate') 
    progressBar.pack(pady=10)
    '''
    
    startBtn = tk.Button(newWindow,text="START", state = tk.DISABLED,
                         command=lambda:[script.mainNoPayment(excelFile.name,memberName.get(1.0,tk.END),memberID.get(1.0,tk.END))])
    
    startBtn.pack()


    excelFile = filedialog.askopenfile(mode='rb',title="Choose an excel file !")
    if excelFile:
        startBtn["state"]= tk.NORMAL
    else:
        startBtn["state"]= tk.DISABLED
    
    
    
  


    
'''
    selectFileLabel = tk.Label(newWindow, text = "Browse Excel File")
    selectFileLabel.pack()
    

    
    splitBtn = tk.Button(newWindow,text="split",state=tk.DISABLED,
                         command=lambda:[script.splitExcel(file.name, int(nbr)),msgSplit(),newWindow.destroy()])
    splitBtn.pack()
    
'''




def checkBoxIsChecked(btn):
    if cb.get() == 1:
        #btn['state'] = tk.NORMAL
        #btn.configure(text='Awake!')
        proceedPaymentWindow()
    elif cb.get() == 0:
        #btn['state'] = tk.DISABLED
        #btn.configure(text='Sleeping!')
        notProceedPaymentWindow()
    else:
        tk.messagebox.showerror('Error', 'Something went wrong!')    


def get_input(textbox):
    global nbr
    nbr = textbox.get(1.0, "end-1c")
    
def browse_file(window):
    global file
    file = filedialog.askopenfile(mode='rb',title="Choose an excel file !")
    
def msgSplit():
    tk.messagebox.showinfo("Information", "- Your excel file is splited into "+nbr+" files !")


def splitExcelWindow():
    newWindow = tk.Tk()
 

    newWindow.title("Split Excel File")
 
    newWindow.geometry("250x120")
    l = tk.Label(newWindow, text = "Numbre of generated Excel Files")
    l.pack()
    nbrExcelFiles = tk.Text(newWindow, height = 1, width = 10)
    nbrExcelFiles.pack()

    selectFileLabel = tk.Label(newWindow, text = "Browse Excel File")
    selectFileLabel.pack()
    selectFileBtn = tk.Button(newWindow,text="Browse",
                              command=lambda:[browse_file(newWindow),get_input(nbrExcelFiles),
                                             textBoxIsEmpty(nbrExcelFiles,splitBtn)])
    splitBtn = tk.Button(newWindow,text="split",state=tk.DISABLED,
                         command=lambda:[script.splitExcel(file.name, int(nbr)),msgSplit(),newWindow.destroy()])
    selectFileBtn.pack()

    
    
    splitBtn.pack()
    








mainWindow = tk.Tk()

tk.messagebox.showinfo("Information",
                       "- This tool is developed by Fedi GALFAT. \n- If you need any support, contact me at fedi.galfat@ieee.org")
mainWindow.geometry("400x200")
mainWindow.title("IEEE Membership Automation")

image = Image.open("sb.png")
imageSB = image.resize((250, 83))
imgSB = ImageTk.PhotoImage(imageSB)

imgLabel = tk.Label(image=imgSB)
imgLabel.image = imgSB
imgLabel.place(x=70 ,y=10)




splitExcelBtn = tk.Button(mainWindow,text ="Split Excel File",command=splitExcelWindow)
splitExcelBtn.place(x=200 ,y=120)

startBtn = tk.Button(mainWindow,text ="Start Script !",command=lambda:[checkBoxIsChecked(splitExcelBtn)])
startBtn.place(x=100 ,y=120)

cb = tk.IntVar()

paymentCheckbox = tk.Checkbutton(mainWindow,
                                 text="I want the script to proceed to the payment phase !",variable=cb, onvalue=1, offvalue=0)
paymentCheckbox.place(x=50,y=160)




mainWindow.mainloop()
