import numpy as np
import pandas as pd
import openpyxl
import os
from time import sleep
from settings import *


pd.set_option("display.colheader_justify", "center") # dataframe formatting


class New_Entry: # creates a new entry
    def req(self): # Requesting user input
        name = input("Please enter your name: ")
        date = input("Please enter the date: ")
        timein = input("Please enter the time you entered: ")
        timeout = input("Please enter the time you left: ")
        weaponid = input("Please enter the weapon ID: ")
        pelletcount = input("Please enter the pellet count: ")
        list1 = [name, date, timein, timeout, weaponid, pelletcount]
        return list1
    
    def read(self):
        df = pd.DataFrame(self.list1, columns = ["Name", "Date", "Time In", "Time Out", "Weapon ID", "Pellet Count"], index = [''])


class Todays_entries: # holds/manages current session's entries
    def __init__(self, list1): # requires compilation of New_Entries from compile_entries
        self.list1 = list1

    def savedf(self): # saves the dataframe to xlsx file
        wb = openpyxl.load_workbook("E:/CEP-WA3/attendance.xlsx")
        ws = wb.active
        for i in range(len(self.list1)):
            ws.append(self.list1[i])
        wb.save("E:/CEP-WA3/attendance.xlsx")

    def read(self): # reads the current session's entries
        df = pd.DataFrame(
            self.list1,
            columns=[
                "Name",
                "Date",
                "Time In",
                "Time Out",
                "Weapon ID",
                "Pellet Count",
            ],
            index=[""] * len(self.list1),
        )
        return df


class All_entries: # facilitates interaction with the main database
    def read(self): # reads the entire database
        df = pd.read_excel("E:/CEP-WA3/attendance.xlsx")
        print(df)

    def search(self, column, value): # searches the database for a specific value in a specific column
        df = pd.read_excel("E:/CEP-WA3/attendance.xlsx")
        df = df.loc[df[column] == value]
        print(df)


def compile_entries():  # compiles entries into a list
    list1 = []
    entry1 = New_Entry()
    yes = True
    
    while yes == True:
        list1.append(entry1.req())
        choice = input("Do you want to add another entry? (y/n): ")
        if choice == "y":
            pass
        elif choice == "n":
            yes = False
        else:
            print("Please enter a valid choice.")
    
    return list1


def start(): # to be called later in order to start the program
    print("Welcome to the Shooting Club Attendance Tracker.")
    sleep(0.5)
    main()


def main(): # main process
    while True:
        print("You can do the following things:")
        sleep(0.5)
        print("1. Add a new entry")
        print("2. View today's entries")
        print("3. Search entries")
        print("4. Exit")
        sleep(0.5)
        
        choice = input("Please enter your choice: ")
        os.system("cls")
        
        if choice == "1": # add a new entry
            list1 = compile_entries()
            entry1 = Todays_entries(list1)
            entry1.savedf()
            input()
            os.system("cls")
        
        elif choice == "2": # view current session's entries
            try:
                entry1.read()
            except:
                print("There are no entries for today.")
            input()
            os.system("cls")
        
        elif choice == "3": # search entries
            print('Column 1 is "Name", Column 2 is "Date", Column 3 is "Time In",\nColumn 4 is "Time Out", Column 5 is "Weapon ID"and Column 6 is "Pellet Count."')
            sleep(0.5)
            entry2 = All_entries()
            column = input(
                "Please enter the column name: "
            )
            value = input("Enter your search term: ")
            entry2.search(column, value)
            input()
            os.system("cls")
        
        elif choice == "4": # exits the program
            if input("Are you sure you want to exit? (y/n): ") == "y":
                exit()
        
        else: # failsafe
            print("Please enter a valid choice.")
            os.system("cls")


start()
