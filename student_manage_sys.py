import pandas as pd 
import matplotlib. pyplot as plt 
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows

def main_menu():
    print("\n------- Student Management System -------\n")
    print("1. Create/Import New Data frame")
    print("2. Student Data Analysis")
    print("3. Student Data Visualization")
    print("4. Exit")
	print("5. Export to Excel with Charts")

def create_dataframe_menu():
    print("\n------- Create Data frame -------\n")
    print("1. Import Data frame from csv file")
    print("2. Add/Modify Custom Index")
    print("3. Add/Modify Custom Column Head")
    print("4. Return to main menu")

def analysis_menu():
    print("\n------- Data Analysis using Python -------\n")
    print("1.  Display All records")
    print("2.  Print first nth records")
    print("3.  Print last nth records")
    print("4.  Display student with maximum marks")
    print("5.  Display student with minimum marks")
    print("6.  Display students who have secured passing marks")
    print("7. Delete a row from Data frame")
    print("8. Return to main menu")

def visualisation_menu():
    print("\n------- Visualization using Matplotlib -------\n")
    print("1. Plot Line graph (Subject wise marks)")
    print("2. Plot Bar graph (Students, Marks)")
    print("3. Return to main menu")
	
cols = ['admn','name','dob','class','maths','english','science','marks']
df = pd. DataFrame([],columns = cols) # Create an Empty Data Frame
while True:
    main_menu()
    ch = int(input("Select Option: "))
    if ch == 1:
        # Create New Data frame
        create_dataframe_menu()
        ch = int(input("Select Option: "))
        if ch == 1:
            file = input("File name: ")
            df = pd.read_csv(file)
        elif ch == 2:
            index_list = input("Index List: ").split(",")
            df.index = index_list
        elif ch == 3:
            column_list= input("Column List: ").split(",")
            df.columns = column_list
        print(df)

    elif ch == 2:
        while True:
            # Student  Data Analysis
            analysis_menu()
            ch = int(input("Select Option: "))
            if ch == 1:
                print(df)
            elif ch == 2:
                nth = int(input("Enter no of rows to display: "))
                print(df.head(nth))
            elif ch == 3:
                nth = int(input("Enter number of rows to display: "))
                print(df.tail(nth))
            elif ch == 4:
                print(df[df['Total'] == df['Total'].max()])
            elif ch == 5:
                print(df[df.Total== df["Total"].min()])
            elif ch == 6:
                print(df[df['Total']*100/240 >= 33])
            elif ch == 7:
                print("1. Delete Row by Index")
                print("2. Delete Row by Admn No.")
                ch = int(input("Select Option: "))
                if ch == 1:
                    idx = int(input("Index to delete: "))
                    df = df.drop(index = idx)
                elif ch == 2:
                    admn = int(input("Admn no to delete: "))
                    df = df.drop(df[df["admission number"] == admn].index)
                else:
                    print("Wrong Option Selected! ")
            else:
                print("Returning to main menu")
                break
    elif ch == 3:
        while True:
            # Student Data Visualisation
            visualisation_menu()
            ch = int(input("Select Option: "))
            if ch == 1:
                plt.plot(df['First Name'], df['Maths'], label='Maths', color ="blue", marker="*")
                plt.plot(df['First Name'], df['Physics'], label='Physics', color = "green", marker="*")
                plt.plot(df['First Name'], df['Chemistry'], label='Chemistry', color = "purple", marker="*")
                plt.xlabel("Student", fontsize=12)
                plt.ylabel("Marks", fontsize=12)
                plt.title("Subject Wise Marks of Students", fontsize=16)
                plt.legend()
                plt.show()
            elif ch == 2:
                x_values = df["First Name"]
                y_values = df['Total']
                plt.bar(x_values, y_values, color = 'orange')
                plt.xlabel("Students", fontsize=12)
                plt.ylabel("Marks", fontsize=12)
                plt.title("Students - Marks Visualisation", fontsize=14)
                plt.show()
            elif ch == 3:
                print("Returning to main menu")
                break
            else:
                print("Wrong Option Selected! ")
    elif ch == 4:
        # Exit
        print("Bye ...")
        exit()
	elif ch == 5:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.utils.dataframe import dataframe_to_rows

    file_name = input("Enter Excel file name (e.g. students.xlsx): ")

    wb = Workbook()
    ws = wb.active
    ws.title = "Student Data"

    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    chart = BarChart()
    chart.title = "Student Marks"
    chart.x_axis.title = "Students"
    chart.y_axis.title = "Marks"

    data = Reference(ws, min_col=8, min_row=1, max_row=len(df)+1)
    cats = Reference(ws, min_col=2, min_row=2, max_row=len(df)+1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    ws.add_chart(chart, "J2")

    wb.save(file_name)

    print("Excel file with chart created successfully!")
    else:
        # Error Display and Exit
        print("Error! Wrong option selected. ")
        break
