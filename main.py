from tkinter import BooleanVar
from tkinter.constants import CENTER, TRUE
from PySimpleGUI.PySimpleGUI import WINDOW_CLOSED, WINDOW_CLOSE_ATTEMPTED_EVENT
import openpyxl
from openpyxl import Workbook
import PySimpleGUI as sg
import matplotlib.pyplot as mpl
import numpy as np
import tracking_sheet


sheet1 = tracking_sheet.Sheet()

sg.theme('Dark Blue 3') 

balance = sheet1.getBalance()
layout = [[sg.Text('Finance Tracker', font="Times", size = (45, 0), justification = 'center')],
          [sg.Text('Month', font="Times"), sg.Input(key='-M-')],
          [sg.Text('Day', font="Times"), sg.Input(key='-D-')],
          [sg.Text('Title', font="Times"), sg.Input(key='-T-')],
          [sg.Text('Category', font="Times"), sg.Input(key='-C-')],
          [sg.Text('Amount', font="Times"), sg.Input(key = '-A-')],
          [sg.Button('Add', font="Times"), sg.Button('Monthly Summary', font="Times"), sg.Button('Income/Expense Summary', font="Times")], 
          [sg.Text('Balance From ' + sheet1.getStartDate() + ":", font="Times"), sg.Text(balance, size=(9,1), key='-B-', font="Times"), sg.Button('Reset Tracker', font="Times")]]

layout = [[sg.Column(layout, element_justification='c')]]
window = sg.Window('Finance Tracker', layout)
while True:  # Event Loop
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    # Must change the prints to display on the GUI
    if event == 'Add':
        error = False
        if (values['-M-'] == ''):
            print("Missing Month")
            error = True
        if (values['-D-'] == ''):
            print("Missing Day")
            error = True
        if (values['-T-'] == ''):
            print("Missing Title")
            error = True
        if (values['-C-'] == ''):
            print("Missing Category")
            error = True
        if (values['-A-'] == ''):
            print("Missing Amount")
            error = True
        try:
            val = int(values['-A-'])
            print("Amount is a valid number")
        except ValueError:
            try:
                val = float(values['-A-'])
                print("Amount is a valid fp number")
            except:
                print("Enter a valid number")
                error = True
        # If there are no problems with input
        if (error == False):
            income = False 
            expense = False
            if (val >= 0):
                income = True
                expense = False
            elif (val < 0): # If the amount is negative
                expense = True
                income = False
                val = val * -1
            # addExpense Format: (month, date, description, category, income, expense, amount)
            # Month and Category names stripped of surrounding whitespace and made capitalized in order to avoid redundant months and categories made from simple user mistakes 
            sheet1.addExpense(values['-M-'].strip().upper(), values['-D-'], values['-T-'], values['-C-'].strip().upper(), income, expense, val)
        window['-B-'].update(sheet1.getBalance())
    
    if event == 'Reset Tracker':
        layout2 = [[sg.Text('Are you sure?', pad = (20, 0), font="Times")], [sg.Button('Yes', pad = (15, 0), font="Times"), sg.Button('No', pad = (15, 0), font="Times")]]
        window2 = sg.Window(' ', layout2)
        while True:
            event2, values2 = window2.read()
            if event2 == sg.WIN_CLOSED:
                break
            if event2 == 'Yes':
                sheet1.resetSheet()
                window['-B-'].update(sheet1.getBalance())
                window2.close()
                break
            if event2 == 'No':
                window2.close()
                break
        
    if event == 'Income/Expense Summary':
        catSums = sheet1.getCategorySums(income=True)
        fig, (ax1, ax2) = mpl.subplots(1, 2)
        fig.set_size_inches(15, 7)
        fig.suptitle("Income and Expense Summary Since " + sheet1.getStartDate())
        ax1.pie(catSums.values(), labels=catSums.keys(), autopct='%1.1f%%')
        ax1.title.set_text("Income Summary")
        catSums = sheet1.getCategorySums(income=False)
        ax2.pie(catSums.values(), labels=catSums.keys(), autopct='%1.1f%%')
        ax2.title.set_text("Expenses")
        mpl.show()
    
    if event == 'Monthly Summary':
        monthTotals = sheet1.getMonthlyTotals()
        keysList = list(monthTotals.keys())
        ypos = np.arange(len(keysList))
        mpl.figure(figsize=(15, 7))
        mpl.xticks(ypos, keysList)
        incomeList = []
        expenseList = []
        for value in monthTotals.items():
            incomeList.append(value[1][0])
            expenseList.append(value[1][1])
        mpl.bar(ypos-0.2, incomeList, width=0.4, label="Income")
        mpl.bar(ypos+0.2, expenseList, width=0.4, label="Expenses")
        mpl.ylabel("Amount ($)")
        mpl.title("Monthly Income and Expense Summary")
        mpl.legend()
        mpl.show()

window.close()