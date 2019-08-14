from tkinter import *
from tkinter import filedialog
import csv
import xlsxwriter
import datetime


stroot = Tk()
stroot.geometry("+200+200")
stroot.title('Daily Patrolman Point Checker')
stroot.configure(background='white', padx=40, pady=40)


def openFile():
    global str
    str = filedialog.askopenfilename(initialdir="/Users/unclechopper/Desktop", title='Select a file',
                                                 filetypes=(("csv files", "*.csv"), ("All files", "*.*")))
    stroot.destroy()
    runChecker()


def masterList(filename, new_name):
    with open(filename, ) as master:
        lines = csv.reader(master)
        for row in lines:
            new_name.append(row[2])


def rowStrip(setDiff, id, libox):
    setDiff = list(setDiff)
    setDiff.sort()
    if bool(setDiff):
        for row in setDiff:
            if '&#163;' in row and '%' in row and '&amp;' in row:
                row = row[:-15]
            elif '&#163;' in row and '%' in row:
                row = row[:-9]
            elif '&#163;' in row and '&amp;' in row:
                row = row[:-12]
            else:
                row = row[:-7]
            libox.insert("end", f'{id} - {row[35:]}')


def diff_lists_export(dif):
    # Create new list for the three patrol difference sets to go to
    global report_dif_list
    report_dif_list = []

    # Strip the rows and append to new string
    for row in list(dif):
        if '&#163;' in row and '%' in row and '&amp;' in row:
            row = row[35:-15]
            report_dif_list.append(row)

        elif '&#163;' in row and '%' in row:
            row = row[35:-9]
            report_dif_list.append(row)

        elif '&#163;' in row and '&amp;' in row:
            row = row[35:-12]
            report_dif_list.append(row)

        elif '&#163;' in row:
            row = row[35:-7]
            report_dif_list.append(row)

    return report_dif_list


def resultLabel(list, label, list2, list3, list4, list5):
    if list > 0:
        label.set(f"{len(list2)} point(s) checked out of {list3} total points. \n"
                     f"\nPoints checked percentage: {round(list4)}% "
                     f" \n\n!!! {list5} point(s) missed !!!")
    elif list == 0:
        label.set(f"{len(list2)} point(s) checked out of {list3} total points. \n"
                     f"\nPoints checked percentage: {round(list4)}% "
                     f" \n\n!!! Patrol Completed !!!")


def runChecker():

    # Create GUI window
    root = Tk()
    root.geometry("+200+200")
    root.title('Daily Patrolman Point Checker')
    root.configure(background='white', padx=30, pady=30)

    # Declare GUI variables
    labText = StringVar()
    labText2 = StringVar()
    labText3 = StringVar()
    labText4 = StringVar()
    labCheckData = StringVar()

    labCheckData.set(str)
    # Create a label to show parsed document
    Label(root, textvariable=labCheckData, borderwidth=2, relief="groove", background="white", padx=5,
              pady=10, font=("Helvetica", 9)).grid(row=2, column=1, columnspan=1, padx=10, pady=10)


    def checkPoints():

        # Clear listbox contents
        lbox.delete(0, END)

        # Declare master lists
        p1master_list = []
        p2master_list = []
        p3master_list = []

        # Create master patrol lists
        masterList('master_p1.csv', p1master_list)
        masterList('master_p2.csv', p2master_list)
        masterList('master_p3.csv', p3master_list)

        # Output the name in labText label
        labCheckData.set(str)

        # Global declaration of the report saved as a list
        global lines_list
        lines_list = []

        # Open the csv file
        with open(str) as file:
            # Declare list variable to store lines of the file
            lines = csv.reader(file, delimiter=',')
            for row in lines:
                lines_list.append(row)

        # Number of points using the list
        global startDate
        global endDate
        no_points = (len(lines_list)-2)
        endDate = lines_list[2][1]
        startDate = lines_list[no_points][1]

        # Output the dates start and end
        labStartDate.set(startDate)
        labEndDate.set(endDate)

        # Declare counters for patrols - using new sets for unique values
        global count
        global p1list
        global p2list
        global p3list
        count = 0
        p1list = set()
        p2list = set()
        p3list = set()

        for row in lines_list:
        # Filter P1 Patrol using £
            count += 1
            time = row[1][-5:-3]
            if ('21' <= time < '24') or ('0' <= time < '04'):
                if row[2] in p1master_list:
                    p1list.add(row[2])

        # Filter P2 Patrol using %
            if ('10' <= time <= '13'):
                if row[2] in p2master_list:
                    p2list.add(row[2])

        # Filter P3 Patrol &
            if ('15' <= time <= '17'):
                if row[2] in p3master_list:
                    p3list.add(row[2])

        count = count - 2
        pat1_comp = ""
        pat2_comp = ""
        pat3_comp = ""
        if len(p1list) == len(p1master_list):
            pat1_comp = 'Fully Complete'
        else:
            pat1_comp = 'Partially Complete'

        if len(p2list) == len(p2master_list):
            pat2_comp = 'Fully Complete'
        else:
            pat2_comp = 'Partially Complete'

        if len(p3list) == len(p3master_list):
            pat3_comp = 'Fully Complete'
        else:
            pat3_comp = 'Partially Complete'

        global p1diff
        global p2diff
        global p3diff

        # Check the difference between report lists and master lists
        p1diff = set(p1master_list).difference(p1list)
        p2diff = set(p2master_list).difference(p2list)
        p3diff = set(p3master_list).difference(p3list)

        # Display missed points for each patrol in the labelbox and strip symbols
        rowStrip(p1diff, 'P1', lbox)
        rowStrip(p2diff, 'P2', lbox)
        rowStrip(p3diff, 'P3', lbox)

        global lenP1
        global lenP2
        global lenP3
        lenP1 = len(p1master_list)
        lenP2 = len(p2master_list)
        lenP3 = len(p3master_list)

        # Declare length of checked list as integer
        global lenP1Ch
        global lenP2Ch
        global lenP3Ch
        lenP1Ch = len(p1list)
        lenP2Ch = len(p2list)
        lenP3Ch = len(p3list)

        # Calculate percentage of hits
        global p1_hit_rate
        global p2_hit_rate
        global p3_hit_rate
        p1_hit_rate = (lenP1Ch / lenP1) * 100
        p2_hit_rate = (lenP2Ch / lenP2) * 100
        p3_hit_rate = (lenP3Ch / lenP3) * 100

        # Calculate the missed points
        p1_missed = lenP1 - lenP1Ch
        p2_missed = lenP2 - lenP2Ch
        p3_missed = lenP3 - lenP3Ch

        global p_missed_total
        global dup_points
        p_missed_total = p1_missed + p2_missed + p3_missed
        dup_points = count - lenP1Ch - lenP2Ch - lenP3Ch

        # Output the points missed and percentages
        labText.set(f"Patrol 1: {pat1_comp}!\n"
                    f"Patrol 2: {pat2_comp}!\n"
                    f"Patrol 3: {pat3_comp}!\n"
                    f"\n{count} point(s) checked overall. \n"
                    f"\nAdditional Point(s) Checked: {dup_points}"
                    f" \n\n!!! Total Point(s) Missed: {p_missed_total} !!!")

        # Display Patrol results
        resultLabel(p1_missed, labText2, p1list, lenP1, p1_hit_rate, p1_missed)
        resultLabel(p2_missed, labText3, p2list, lenP2, p2_hit_rate, p2_missed)
        resultLabel(p3_missed, labText4, p3list, lenP3, p3_hit_rate, p3_missed)

        # Create the labels to house the points data
        Label(root, textvariable=labText, borderwidth=2, relief="groove", pady=10, background="white",
              font=("Helvetica", 10)).grid(row=6, column=2, padx=10, pady=3, sticky=W + E + N)
        Label(root, textvariable=labText2, borderwidth=2, relief="groove", pady=10, background="white",
              font=("Helvetica", 10)).grid(row=4, column=0, padx=10, pady=10, sticky=W + E + N + S)
        Label(root, textvariable=labText3, borderwidth=2, relief="groove", pady=10, background="white",
              font=("Helvetica", 10)).grid(row=4, column=1, padx=10, pady=10, sticky=W + E + N + S)
        Label(root, textvariable=labText4, borderwidth=2, relief="groove", pady=10, background="white",
              font=("Helvetica", 10)).grid(row=4, column=2, padx=10, pady=10, sticky=W + E + N + S)


    def openFile1():
        global str
        str = filedialog.askopenfilename(initialdir="/Users/unclechopper/Desktop", title='Select a file',
                                         filetypes=(("csv files", "*.csv"), ("All files", "*.*")))
        checkPoints()

    # Declare the variables for the I/O
    labStartDate = StringVar()
    labEndDate = StringVar()
    labTotal = StringVar()
    labP1Total = StringVar()
    labP2Total = StringVar()
    labP3Total = StringVar()

    # Create the Patrol headings
    labTotal.set('Summary Report \n'
                 'All Points (06:00 - 06:00)')
    labP1Total.set('P1 Patrol (21:00 - 04:00)\n'
                   'East & West & External')
    labP2Total.set('P2 Patrol (10:30 - 12:30)\n'
                   'Ground, Basement & External')
    labP3Total.set('P3 Patrol (15:00 - 17:00)\n'
                   'Vulnerable Areas')

    startLF = LabelFrame(root, text="Start Date/Time", fg='white', background='gray')
    startLF.grid(row=0, column=0, sticky=E+W)

    endLF = LabelFrame(root, text="End Date/Time", fg='white', background='gray')
    endLF.grid(row=0, column=2, sticky=E+W)

    # Create the start and end dates and output at the top
    Label(startLF, textvariable=labStartDate, fg='white', background='gray', font=("Helvetica", 12)).grid(row=0,
                                                                                                          column=0, padx=10, pady=10)

    logo = PhotoImage(file="3HSMlogo.gif")

    Label(root, image=logo, bd=0).grid(row=0,
                                                                                                   column=1, padx=10, pady=10)
    Label(endLF, textvariable=labEndDate, fg='white', background="gray", font=("Helvetica", 12)).grid(row=0,
                                                                                                      column=2, padx=10, pady=10)
    # Create button to open file browser
    b3 = Button(root, text='Open Another CSV File', command=openFile1, font=("Helvetica", 12))
    b3.grid(row=2, column=0, padx=10, pady=10, sticky=W + E + N + S)

    # Create button to run the report
    b2 = Button(root, text="Run Excel Report", command=runReport, font=("Helvetica", 12))
    b2.grid(row=2, column=2, padx=10, pady=10, sticky=W+E+N+S)

    # Create the labels to house the points titles
    Label(root, textvariable=labTotal, fg='white', borderwidth=2, relief="groove", bg='grey', width=25, bd=2,
          font=("Helvetica", 12)).grid(row=5, column=2, padx=10, pady=10, sticky=W+E+N)
    Label(root, textvariable=labP1Total, fg='white', borderwidth=2, relief="groove", bg='grey', width=25,
          font=("Helvetica", 12)).grid(row=3, column=0,padx=10, pady=10, sticky=W+E+N+S)
    Label(root, textvariable=labP2Total, fg='white', borderwidth=2, relief="groove", bg='grey', width=25,
          font=("Helvetica", 12)).grid(row=3, column=1,padx=10, pady=10, sticky=W+E+N+S)
    Label(root, textvariable=labP3Total, fg='white', borderwidth=2, relief="groove", bg='grey', width=25,
          font=("Helvetica", 12)).grid(row=3, column=2,padx=10, pady=10, sticky=W+E+N+S)

    Label(root, text=":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
                     ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::",
            font=("Helvetica", 12)).grid(row=1, column=0,columnspan=4, padx=10,
                                                             pady=10, sticky=W+E+N+S)

    # Create the listbox to house the missed points
    lbox = Listbox(root, fg='white', background='grey', font=("Helvetica", 9))
    lbox.grid(row=5, columnspan=2, rowspan=3, padx=10, pady=10, sticky=W+E+N+S)

    # Create button to quit program
    b4 = Button(root, text="Quit", command=root.destroy, activebackground="red", fg='red', font=("Helvetica", 12))
    b4.grid(row=7, column=2, padx=10, pady=10, sticky=W+E+N+S)

    # Run points checker
    checkPoints()

    root.mainloop()


def runReport():
    # Check the date format
    if '/' in startDate:
        sD = startDate[:2] + startDate[3:5] + startDate[8:10]
        eD = endDate[:2] + endDate[3:5] + endDate[8:10]
    else:
        months = {
            'Jan': '01',
            'Feb': '02',
            'Mar': '03',
            'Apr': '04',
            'May': '05',
            'Jun': '06',
            'Jul': '07',
            'Aug': '08',
            'Sep': '09',
            'Oct': '10',
            'Nov': '11',
            'Dec': '12'
        }

        sD = startDate[:2] + months[startDate[3:6]] + startDate[9:11]
        eD = endDate[:2] + months[endDate[3:6]] + endDate[9:11]

    # Format report names
    report_name = f"{sD} to {eD} Daily Patrol Check.xlsx"
    workbook = xlsxwriter.Workbook(f"../{report_name}")
    worksheet = workbook.add_worksheet('Summary')
    worksheet1 = workbook.add_worksheet('Missed Points')

# C:\\Users\\capta\\OneDrive\\Desktop\\
    # Format the cells
    bold = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': 'gray', 'font_color': 'white'})
    normal = workbook.add_format({'align': 'center', 'border': 1})
    normalData = workbook.add_format({'align': 'left', 'border': 1})
    normalTitle = workbook.add_format({'align': 'center', 'bg_color': 'gray', 'bold': True, 'border': 1, 'font_color': 'white'})

    worksheet.set_column(0,  10, width=20)
    worksheet1.set_column(0,  0, width=10)
    worksheet1.set_column(7,  7, width=60)
    worksheet1.set_column(9,  9, width=25)


    # Set the titles for the columns
    worksheet.write('A1', 'Start Date', bold)
    worksheet.write('A2', 'End Date', bold)
    worksheet.write('D1', 'P1 Patrol Points', bold)
    worksheet.write('D2', 'P2 Patrol Points', bold)
    worksheet.write('D3', 'P3 Patrol Points', bold)
    worksheet.write('F1', 'Missed Points', bold)
    worksheet.write('F2', 'Additional Points', bold)
    worksheet.write('A3', 'Total Points Checked', bold)
    worksheet.write('F3', 'Total Points (%)', bold)

    worksheet1.write('A1', 'Patrol Type', normalTitle)
    worksheet1.merge_range(0, 1, 0, 6, 'Missed Points', normalTitle)
    worksheet1.write('H1', 'Comments', normalTitle)
    worksheet1.write('J1', 'Checked by:', normalTitle)
    worksheet1.write('J2', ' ', normal)
    worksheet1.write('J3', 'Date Checked:', normalTitle)
    worksheet1.write('J4', ' ', normal)



    total_ave = ((lenP1Ch + lenP2Ch + lenP3Ch) / (lenP1 + lenP2 + lenP3)) * 100
    total_ave = f"{round(total_ave)}%"

    # Populate the cells with calculated data
    worksheet.write('B1', startDate, normal)
    worksheet.write('B2', endDate, normal)
    worksheet.write('C1', f'{round(p1_hit_rate)}%', normal)
    worksheet.write('C2', f'{round(p2_hit_rate)}%', normal)
    worksheet.write('C3', f'{round(p3_hit_rate)}%', normal)
    worksheet.write('E1', lenP1Ch, normal)
    worksheet.write('E2', lenP2Ch, normal)
    worksheet.write('E3', lenP3Ch, normal)
    worksheet.write('G1', p_missed_total, normal)
    worksheet.write('G2', dup_points, normal)
    worksheet.write('B3', count, normal)
    worksheet.write('G3', total_ave, normal)





    # Declare rows for the below loop
    rS = 5
    rE = 5

    # Calculate which patrol and strip symbols
    for row in range(2, len(lines_list)):
        if '&#163;' in lines_list[row][2] and '%' in lines_list[row][2] and '&amp;' in lines_list[row][2]:
            lines_list[row][2] = lines_list[row][2][:-15]
        elif '&#163;' in lines_list[row][2] and '%' in lines_list[row][2]:
            lines_list[row][2] = lines_list[row][2][:-9]
        elif '&#163;' in lines_list[row][2] and '&amp;' in lines_list[row][2]:
            lines_list[row][2] = lines_list[row][2][:-12]
        else:
            lines_list[row][2] = lines_list[row][2][:-7]

    a = 5
    b = 0
    # For loop to populate the data body in the report
    for i in range(2, len(lines_list)):
        worksheet.write(a, b, lines_list[i][1], normal)
        worksheet.merge_range(rS, 1, rE, 6, lines_list[i][2], normalData)
        a += 1
        rS += 1
        rE += 1

    worksheet.write(4, 0, 'Time Scanned', normalTitle)
    worksheet.merge_range(4, 1, 4, 6, 'Points Checked', normalTitle)

    # Call function to strip dif lists
    p1listdif = diff_lists_export(list(p1diff))
    p2listdif = diff_lists_export(list(p2diff))
    p3listdif = diff_lists_export(list(p3diff))
    p1listdif.sort()
    p2listdif.sort()
    p3listdif.sort()
    diftotallist = p1listdif + p2listdif + p3listdif

    r2S = 1
    r2E = 1
    e = 1
    f = 0
    g = 1
    h = 7

    for i in range(len(p1listdif)):
        worksheet1.write(e, f, 'P1', normalData)
        e += 1
    for i in range(len(p2listdif)):
        worksheet1.write(e, f, 'P2', normalData)
        e += 1
    for i in range(len(p3listdif)):
        worksheet1.write(e, f, 'P3', normalData)
        e += 1

    # For loop to populate the missed points tab
    for row in diftotallist:
        worksheet1.merge_range(r2S, 1, r2E, 6, row, normalData)
        worksheet1.write(g, h, ' ', normalData)
        r2S += 1
        r2E += 1
        g += 1

    worksheet.hide_gridlines()
    worksheet1.hide_gridlines()

    workbook.set_properties({
        'title': report_name,
        'company': 'Knight Frank - Ward Security',
        'category': 'Patrol Points',
        'manager': 'John Hughes',
        'created': datetime.datetime.now()
    })
    workbook.close()


l1 = Label(stroot, text="Click on button and load a CSV file.", justify=LEFT, borderwidth=2, bg='lightgrey', bd=2, font=("Helvetica", 12), pady=20, padx=20)
l1.grid(row=0, column=0, columnspan=2, padx=20, pady=20, sticky=W+E+N)

# Create button to open file browser
b4 = Button(stroot, text='Open CSV File', command=openFile, font=("Helvetica", 12), bg='gray')
b4.grid(row=1, column=0, columnspan=2, padx=20, pady=30, sticky=W+E)

stroot.mainloop()
