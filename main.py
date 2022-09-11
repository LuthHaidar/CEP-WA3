import pandas as pd
import openpyxl as xl
import os
import time
import msvcrt
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from settings import *


pd.set_option("display.colheader_justify", "center")  # dataframe formatting
attendance_sheet = os.path.join(os.path.dirname(__file__), "attendance.xlsx")
attendance_sheet_sorted = os.path.join(os.path.dirname(__file__), "attendance_sorted.xlsx")


class New_Entry:  # creates a new entry
    def req(self):  # Requesting user input
        while True: # name
            name = input("Please enter your name: ")
            if name == "":
                print("Please enter a valid name.")
                input()
                os.system("cls")
                continue
            else:
                name = name.lower()
                break

        while True: # Date
            date = input("Please enter the date(e.g. 17/9/2022): ")
            try:  # checks if date is valid
                time.strptime(date, "%d/%m/%Y")
                break
            except:
                print("Please enter a valid date.")
                input()
                os.system("cls")
                continue

        while True: # Time In
            timein = input("Please enter the time you entered(e.g. 1600): ")
            try:  # checks if time is valid
                time.strptime(timein, "%H%M")
                timein = int(timein)
                break
            except:
                print("Please enter a valid time.")
                input()
                os.system("cls")
                continue

        while True: # Time Out
            timeout = input("Please enter the time you left(e.g. 1800): ")
            try:  # checks if time is valid
                time.strptime(timeout, "%H%M")
                timeout = int(timeout)
                break
            except:
                print("Please enter a valid time.")
                input()
                os.system("cls")
                continue

        while True: # Weapon ID
            weaponid = input("Please enter the weapon ID: ")
            if weaponid == "": # checks if weapon ID is empty
                print("Please enter a valid weapon ID.")
                input()
                os.system("cls")
                continue
            else:
                break

        while True: # Pellet Count
            pelletcount = input("Please enter the pellet count: ")
            try: # ensures pellet count is an integer
                pelletcount = int(pelletcount)
                break
            except:
                print("Please enter a valid pellet count.")
                input()
                os.system("cls")
                continue
        
        list1 = [name, date, timein, timeout, weaponid, pelletcount]
        self.list1 = list1
        return list1

    def read(self): # reads the new entry
        df = pd.DataFrame(
            [self.list1],
            columns=[
                "Name",
                "Date",
                "Time In",
                "Time Out",
                "Weapon ID",
                "Pellet Count",
            ],
            index=[""],
        )
        print(df)


class Todays_entries:  # holds/manages current session's entries
    def __init__(self, list1):  # requires compilation of New_Entries from compile_entries
        self.list1 = list1

    def savedf(self):  # appends dataframe to xlsx file
        wb = xl.load_workbook(attendance_sheet)
        ws = wb.active
        for i in range(len(self.list1)):
            ws.append(self.list1[i])
        wb.save(attendance_sheet)

    def read(self):  # reads the current session's entries
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
        print(df)

    def delete_entry(self, number): # deletes an entry
        df = pd.read_excel(attendance_sheet)
        df.drop(df.index[number], inplace=True)
        print(df)


class All_entries:  # facilitates interaction with the main database
    def read(self):  # reads the entire database
        df = pd.read_excel(attendance_sheet)
        print(df)

    def search(
        self, column, value
    ):  # searches the database for a specific value in a specific column
        if column == "Name":
            value = value.lower()
        df = pd.read_excel(attendance_sheet)
        df = df[df[column] == value]
        if df.empty:
            print("No entries found.")
        else:
            print(df)

    def sort_by(self, column):
        df = pd.read_excel(attendance_sheet)
        df = df.sort_values(column)
        df.to_excel(attendance_sheet_sorted, index=False)
        print(df)

    def send_mail(self):  # sends attendance.xlsx to the email address given
        youremail = input(
            "Please enter the email address you want to send a copy of the attendance sheet to: "
        )
        msg = MIMEMultipart()
        msg["From"] = "Yourself"
        msg["To"] = youremail
        msg["Subject"] = "Attendance"
        body = "Here's a copy of attendance.xlsx as valid for today."
        msg.attach(MIMEText(body, "plain"))
        filename = attendance_sheet
        attachment = open(attendance_sheet, "rb")
        part = MIMEApplication(attachment.read(), Name=basename(filename))
        part["Content-Disposition"] = 'attachment; filename="%s"' % basename(filename)
        msg.attach(part)
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("cepwa3@gmail.com", "pqomoupmwibsujwr")
        text = msg.as_string()
        server.sendmail("cepwa3@gmail.com", youremail, text)
        server.quit()
        print("Email sent successfully to ", youremail, ".")


def compile_entries():  # compiles entries into a list
    list1 = []
    entry1 = New_Entry()

    while True:  # loops until user enters "no"
        list1.append(entry1.req())
        print(entry1.read())
        time.sleep(0.5)
        while msvcrt.kbhit():
            flush = input()

        correct = input("Is this correct? (y/n): ")
        if correct == "y":
            pass
        elif correct == "n":
            list1.pop()
            os.system("cls")
            continue
        
        choice = input("Do you want to add another entry? (y/n): ")
        if choice == "y":
            os.system("cls")
            continue
        elif choice == "n":
            break
        else:
            print("Please enter a valid choice.")
    return list1


def start():  # to be called later in order to start the program
    print("Welcome to the Shooting Club Attendance Tracker.\n")
    time.sleep(0.5)
    while msvcrt.kbhit():
        flush = input()
    main()


def main():  # main process
    while True:
        print("You can do the following things:")

        time.sleep(0.5)
        while msvcrt.kbhit():
            flush = input()

        print("1. Add a new entry")
        time.sleep(0.3)
        while msvcrt.kbhit():
            flush = input()

        print("2. View today's entries")
        time.sleep(0.3)
        while msvcrt.kbhit():
            flush = input()

        print("3. Search entries")
        time.sleep(0.3)
        while msvcrt.kbhit():
            flush = input()

        print("4. Sort entries")
        time.sleep(0.3)
        while msvcrt.kbhit():
            flush = input()

        print("5. Email a copy of the attendance sheet")
        time.sleep(0.3)
        while msvcrt.kbhit():
            flush = input()

        print("6. Exit\n")
        time.sleep(0.3)
        while msvcrt.kbhit():
            flush = input()

        choice = input("Please enter your choice: ")
        os.system("cls")

        if choice == "1":  # add a new entry
            list1 = compile_entries()
            entry1 = Todays_entries(list1)
            entry1.savedf()
            input()
            os.system("cls")

        elif choice == "2":  # view current session's entries
            try:
                entry1.read()
                ask = input("Do you want to delete any entries? (y/n): ")
                if ask == "y":
                    while True:
                        which = input(
                            "Which entry do you want to delete? (1, 2, 3, etc.): "
                        )
                        if len(which) == 1:
                            try:
                                entry1.delete_entry(int(which) - 1)
                                break
                            except:
                                print("Please enter a single, valid number.")
                                input()
                                continue
                        else:
                            print("Please enter a single, valid number.")
                            input()
                            os.system("cls")
                        another = input("Do you want to delete any entries? (y/n): ")
                        if another == "y":
                            continue
                        elif another == "n":
                            break
                    entry1.savedf()
                elif ask == "n":
                    pass
                else:
                    print("Please enter a valid choice.")
                    input()
            except:
                print("There are no entries for today.")
            input()
            os.system("cls")

        elif choice == "3":  # search entries
            print(
                'The following columns are available for searching: "Name", "Date", "Time In", "Time Out", "Weapon ID" or "Pellet Count".'
            )
            time.sleep(0.5)
            while msvcrt.kbhit():
                flush = input()
            entry1 = All_entries()
            while True:
                column = input("Please enter the column name: ")
                value = input("Enter your search term: ")
                try:
                    entry1.search(column, value)
                    break
                except:
                    print("Please enter a valid column name.")
                    continue
            input()
            os.system("cls")

        elif choice == "4":  # sort entries
            entry1 = All_entries()
            print(
                'You can sort by: "Name", "Date", "Time In", "Time Out", "Weapon ID" or "Pellet Count".'
            )
            time.sleep(0.5)
            while msvcrt.kbhit():
                flush = input()
            while True:
                column = input("Please enter the column name by which you want to sort: ")
                try:    
                    entry1.sort_by(column) # sorts by column and prints the sorted dataframe
                    print(
                    "You can also find the sorted entries in the file 'attendance_sorted.xlsx'."
                    )
                    break
                except:
                    print("Please enter a valid column name.")
                    continue
                
            input()
            os.system("cls")

        elif choice == "5":  # email attendance sheet
            entry1 = All_entries()
            entry1.send_mail()
            input()
            os.system("cls")

        elif choice == "6":  # exits the program
            while True:
                exit_query = input("Are you sure you want to exit? (y/n): ")
                if exit_query == "y":
                    exit()
                elif exit_query == "n":
                    os.system("cls")
                    break
                else:
                    print("Please enter a valid choice.")
                    input()
                    os.system("cls")
                    continue

        else:  # failsafe
            print("Please enter a valid choice.")
            input()
            os.system("cls")


start()
