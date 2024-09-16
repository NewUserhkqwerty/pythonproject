from tkinter import *
from tkinter.font import Font
from tkinter import messagebox
import csv
import os
import pyodbc

m = Tk()
m.title('Akash Technologies')
m.geometry("1000x700")
m.configure(bg="lightblue")


def sumbit():
    data = [date.get(), name.get(), mobile.get(), Alternate.get(), email.get(), address.get(), Course.get(),
            batch.get(),
            batch_know.get(), fresher.get(), Besant.get(), counselor.get(), fees.get(), comment.get()]

    file_path = r'D:\kani\student.csv'
    file_exists = os.path.isfile(file_path)

    with open(file_path, 'a', newline='') as f:
        w = csv.writer(f)

        if not file_exists:
            headers = ["Date", "Name", "Mobile No", "Alternate No", "Email id", "Address", "Course Interested",
                       "Batch Preferred", "How You Came To Know Us", "Are You Experienced or Fresher",
                       "Contact Person From Besant Technologies", "Counselor", "Fees", "Comment"]
            w.writerow(headers)

        w.writerow(data)

    print("Success", "Form submitted successfully!")

    try:

        connection = pyodbc.connect(
            'DRIVER={SQL Server};' + 'Server=DESKTOP-51TEJF4;' + 'Database=DP1;' + 'Trusted_Connection=True')
        connection.autocommit = True
        cursor = connection.cursor()

        cursor.execute('USE DP10')
        print('Connected successfully and using database DP10')

        create_table_query = '''
                IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Emp4' AND xtype='U')
                CREATE TABLE Emp4 (
                    [Date] DATE,
                    [Name] VARCHAR(100),
                    [Mobile No] VARCHAR(15),
                    [Alternate No] VARCHAR(15),
                    [Email id] VARCHAR(100),
                    [Address] VARCHAR(255),
                    [Course Interested] VARCHAR(100),
                    [Batch Preferred] INT,
                    [How You Came To Know Us] VARCHAR(100),
                    [Are You Experienced or Fresher] VARCHAR(50),
                    [Contact Person From Besant Technologies] VARCHAR(100),
                    [Counselor] VARCHAR(100),
                    [Fees] INT,
                    [Comment] VARCHAR(255)
                )
                '''
        cursor.execute(create_table_query)

        insert_query = '''INSERT INTO Emp4 ([Date], [Name], [Mobile No], [Alternate No], [Email id], [Address],
        [Course Interested], [Batch Preferred], [How You Came To Know Us],
        [Are You Experienced or Fresher], [Contact Person From Besant Technologies],
        [Counselor], [Fees], [Comment]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''

        data1 = (
        date.get(), name.get(), mobile.get(), Alternate.get(), email.get(), address.get(), Course.get(), batch.get(),
        batch_know.get(), fresher.get(), Besant.get(), counselor.get(), fees.get(), comment.get())

        cursor.execute(insert_query, data1)
        print("Data  successfully insert in the database")


    except pyodbc.Error as msg:
        print('Error:', msg)

    finally:
        if connection:
            connection.close()
            print("Database connection close ")


myfont = Font(family="times", size=13)
myfont1 = Font(family="times", size=16)

title_label = Label(m, text="Akash Technologies \nEnquiry Form", fg="red", font=myfont, justify='center',
                    anchor='center')
title_label.grid(row=0, column=1, ipady=15, pady=9, ipadx=50)

label = Label(m, text="Date :", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left", width=34)
label.grid(row=1, column=0, pady=2)
date = Entry(m, width=40, font=myfont1)
date.grid(row=1, column=1, padx=10, pady=2)

labe2 = Label(m, text="Name :", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left", width=34)
labe2.grid(row=2, column=0)
name = Entry(m, width=40, font=myfont1)
name.grid(row=2, column=1, padx=10, pady=2)

labe3 = Label(m, text="Mobile No :", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left", width=34)
labe3.grid(row=3, column=0)
mobile = Entry(m, width=40, font=myfont1)
mobile.grid(row=3, column=1, padx=10, pady=2)

labe4 = Label(m, text="Alternate No:", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left", width=34)
labe4.grid(row=4, column=0)
Alternate = Entry(m, width=40, font=myfont1)
Alternate.grid(row=4, column=1, padx=10, pady=2)

labe5 = Label(m, text="Email id:", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left", width=34)
labe5.grid(row=5, column=0)
email = Entry(m, width=40, font=myfont1)
email.grid(row=5, column=1, padx=10, pady=2)

labe6 = Label(m, text="Address:", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left", width=34)
labe6.grid(row=6, column=0)
address = Entry(m, width=40, font=myfont1)
address.grid(row=6, column=1, padx=10, pady=2)

labe7 = Label(m, text="Course Interested:", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left",
              width=34)
labe7.grid(row=7, column=0)
Course = Entry(m, width=40, font=myfont1)
Course.grid(row=7, column=1, padx=10, pady=2)

labe8 = Label(m, text="Batch Pregerred:", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left", width=34)
labe8.grid(row=8, column=0)
batch = Entry(m, width=40, font=myfont1)
batch.grid(row=8, column=1, padx=10, pady=2)

labe8 = Label(m, text="How You Came To Know US:", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left",
              width=34)
labe8.grid(row=9, column=0)
batch_know = Entry(m, width=40, font=myfont1)
batch_know.grid(row=9, column=1, padx=10, pady=2)

labe8 = Label(m, text="Are You Experience or Fresher:", font=myfont, bg="lightblue", fg="black", anchor="w",
              justify="left", width=34)
labe8.grid(row=10, column=0)
fresher = Entry(m, width=40, font=myfont1)
fresher.grid(row=10, column=1, padx=10, pady=2)

labe9 = Label(m, text="Contact Person From Besant Technologies:", font=myfont, bg="lightblue", fg="black", anchor="w",
              justify="left", width=34)
labe9.grid(row=11, column=0)
Besant = Entry(m, width=40, font=myfont1)
Besant.grid(row=11, column=1, padx=10, pady=2)

labe10 = Label(m, text="counselor:", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left", width=34)
labe10.grid(row=12, column=0)
counselor = Entry(m, width=40, font=myfont1)
counselor.grid(row=12, column=1, padx=10, pady=2)

labe11 = Label(m, text="Fess:", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left", width=34)
labe11.grid(row=13, column=0)
fees = Entry(m, width=40, font=myfont1)
fees.grid(row=13, column=1, padx=10, )

labe12 = Label(m, text="Comment:", font=myfont, bg="lightblue", fg="black", anchor="w", justify="left", width=34)
labe12.grid(row=14, column=0)
comment = Entry(m, width=40, font=myfont1)
comment.grid(row=14, column=1, padx=10, pady=2)

enquiry = Checkbutton(m, text="Enquiry")
enquiry.grid(row=15, column=2, padx=10, pady=10, sticky=W)

registration = Checkbutton(m, text="Registration")
registration.grid(row=15, column=1, padx=10, pady=10, sticky=W)

sumbit = Button(m, text="Sumbit", bg="green", fg="white", padx=20, pady=5, width=5, command=sumbit)
sumbit.grid(row=16, column=1, sticky=W)

clear = Button(m, text="Quit", bg="red", fg="white", padx=20, pady=5, width=5)
clear.grid(row=16, column=2, sticky=W)

m.mainloop()
