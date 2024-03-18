import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import pandas as pd


class App(tk.Tk):
    def __init__(self, title, size):

        # main settings
        super().__init__()
        self.title(title)
        self.geometry(f"{size[0]}x{size[1]}")
        self.minsize(size[0], size[1])

        # widgets
        self.header = Header(self)
        self.menu = Menu(self)
        self.lastRow = LastRow(self)

        # run
        self.mainloop()


class Header(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(anchor=NW, pady=(0, 20))

        # main title
        labelTitle = Label(
            self, text="AttendanceKeeper v1.0", font=("Helvetica", "20", "bold")
        )
        # second row title
        labelImport = Label(
            self, text="Select student list Excel file:", font=("Helvetica", "14")
        )
        # import button

        buttonImport = Button(
            self,
            text="Import List",
            width=9,
            height=1,
            font=("Helvetica", "10"),
            command=lambda: importStudent(parent),
        )

        # Add Widgets
        labelTitle.pack(ipadx=160)
        labelImport.pack(side=LEFT)
        buttonImport.pack(side=LEFT, padx=20)


class Menu(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(anchor=W)

        self.leftMenu = LeftMenu(self)
        self.middleMenu = MiddleMenu(self)
        self.rightMenu = RightMenu(self)


class LeftMenu(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(anchor=NW, padx=20, side=LEFT)

        # Listbox title
        labelTitle = Label(
            self, text="Select a student", font=("Helvetica", "10", "bold")
        )
        # Student Listbox
        listboxStudents = Listbox(self, width=30, selectmode="multiple")

        # Add Widgets
        labelTitle.pack()
        listboxStudents.pack()


class MiddleMenu(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(anchor=NW, side=LEFT)

        # Combobox setup
        items = [
            "AP 01",
            "AP 02",
            "AP 03",
            "AP 04",
            "AP 05",
            "AP 06",
            "AP 07",
            "AP 08",
            "AP 09",
            "AP 10",
            "AP 11",
            "AP 12",
            "AP 13",
            "AP 14",
            "AP 15",
            "AP 16",
            "AP 17",
            "AP 18",
            "AP 19",
            "AP 20",
        ]

        # Combobox title
        labelTitle = Label(self, text="Section:", font=("Helvetica", "10", "bold"))

        comboBoxSection = ttk.Combobox(
            self,
            values=items,
            width=10,
            postcommand=lambda: changeSelectedAP(self.children["!combobox"], parent),
        )
        comboBoxSection.set(items[0])

        buttonAdd = Button(
            self,
            text="Add =>",
            width=9,
            height=1,
            command=lambda: addStudent(parent),
            font=("Helvetica", "10"),
        )

        buttonRemove = Button(
            self,
            text="<= Remove",
            width=9,
            height=1,
            command=lambda: removeStudent(parent),
            font=("Helvetica", "10"),
        )

        # Add Widgets
        labelTitle.pack()
        comboBoxSection.pack(ipadx=60)
        buttonAdd.pack(ipadx=60, pady=5)
        buttonRemove.pack(ipadx=60, pady=5)


class RightMenu(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(anchor=NW, side=LEFT, padx=(20, 0))

        # Listbox title
        labelTitle = Label(self, text="Attended Students", font=("Helvetica", "10"))

        listboxStudents = Listbox(self, width=30, selectmode="multiple")

        labelTitle.pack()
        listboxStudents.pack()


class LastRow(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(anchor=W, pady=(10, 0))

        labelWeek = Label(
            self, text="Please enter week: ", font=("Helvetica", "10", "bold")
        )

        textField = Entry(self)

        labelType = Label(
            self, text="Please select file type: ", font=("Helvetica", "10", "bold")
        )

        # menu items
        items = [".txt", ".xlsx", ".csv"]
        combo = ttk.Combobox(self, values=items, width=5)
        combo.set(items[0])

        exportButton = Button(
            self,
            text="export as File",
            width=10,
            height=1,
            font=("Helvetica", "10"),
            command=lambda: exportFile(parent),
        )

        # Add Widgets
        labelType.pack(side=LEFT, padx=(5, 0))
        combo.pack(side=LEFT)
        labelWeek.pack(side=LEFT, padx=(20, 0))
        textField.pack(side=LEFT)
        exportButton.pack(side=LEFT, padx=(30, 0))


class Student:

    def __init__(self, name, id, section, dept):
        self.id = id
        self.dept = dept
        self.name = name
        self.section = section


selectedAP = "AP 01"
students: list[Student] = []
attendedStudents: list[Student] = []


def findStudentListbox(parent: tk.Frame) -> Listbox:
    try:
        lb = parent.children["!menu"].children["!leftmenu"].children["!listbox"]
        return lb
    except:
        lb = parent.children["!leftmenu"].children["!listbox"]
        return lb


def findAttendedStudentListbox(parent: tk.Frame) -> Listbox:
    try:
        lb = parent.children["!menu"].children["!rightmenu"].children["!listbox"]
        return lb
    except:
        lb = parent.children["!rightmenu"].children["!listbox"]
        return lb


def importStudent(parent: tk.Frame):
    filetypes = (("Excel Files", "*.xlsx *.xls"), ("All files", "*.*"))

    # select file
    filename = askopenfilename(title="Open a file", initialdir="/", filetypes=filetypes)

    # import file
    sheet = pd.read_excel(filename)

    # find students
    for index, item in sheet.iterrows():

        student = Student(
            id=str(item["Id"]),
            name=str(item["Name"]),
            section=str(item["Section"]),
            dept=str(item["Dept."]),
        )

        students.append(student)

    listStudent(parent)


def listStudent(parent: tk.Frame):
    lb = findStudentListbox(parent)

    lb.delete(0, END)
    _list = []
    for item in students:
        if item.section == selectedAP:

            name = (
                item.name.split(" ")[-1]
                + " , "
                + item.name.split(" ")[-0]
                + " , "
                + item.id
            )
            _list.append(name)

    sortAndAddList(lb, _list)


def changeSelectedAP(combobox: ttk.Combobox, parent: tk.Frame):
    global selectedAP

    lb = findAttendedStudentListbox(parent)
    lb.delete(0, END)
    attendedStudents.clear()

    selected_option = combobox.get()
    selectedAP = selected_option

    listStudent(parent)


def sortAndAddList(lb: Listbox, items: list):
    _items = sorted(items)

    count = 1
    for item in _items:
        lb.insert(count, item)
        count += 1


def addStudent(parent: tk.Frame):
    lb = findStudentListbox(parent)
    lbA = findAttendedStudentListbox(parent)

    allLbItems = lb.get(0, END)

    for sIndex in lb.curselection()[::-1]:
        sItem = allLbItems[sIndex]
        sId = sItem.split(" , ")[2]

        for s in students:
            if s.id == sId:
                attendedStudents.append(s)
                students.remove(s)

        lb.delete(sIndex)

    # Last Step
    lbA.delete(0, END)

    _list = []

    for item in attendedStudents:
        name = (
            item.name.split(" ")[-1]
            + " , "
            + item.name.split(" ")[-0]
            + " , "
            + item.id
        )
        _list.append(name)

    sortAndAddList(lbA, _list)


def removeStudent(parent: tk.Frame):
    lb: Listbox = findStudentListbox(parent)
    lbA: Listbox = findAttendedStudentListbox(parent)

    allLbaItems = lbA.get(0, END)

    for sIndex in lbA.curselection()[::-1]:
        sItem = allLbaItems[sIndex]
        sId = sItem.split(" , ")[2]

        for s in attendedStudents:
            if s.id == sId:
                students.append(s)
                attendedStudents.remove(s)

        lbA.delete(sIndex)

    # Last Step
    lb.delete(0, END)

    _list = []
    for item in students:
        if item.section == selectedAP:
            name = (
                item.name.split(" ")[-1]
                + " , "
                + item.name.split(" ")[-0]
                + " , "
                + item.id
            )
            _list.append(name)

    sortAndAddList(lb, _list)


def exportFile(parent: tk.Frame):
    combo: ttk.Combobox = parent.children["!lastrow"].children["!combobox"]
    entry: Entry = parent.children["!lastrow"].children["!entry"]

    fileType = combo.get()
    week = entry.get()

    fileName = f"{selectedAP} Week {week}{fileType}"

    if fileType == ".txt":
        txtExport(fileName)
    elif fileType == ".xlsx":
        excelExport(fileName)
    else:
        csvExport(fileName)


def txtExport(fileName: str):
    data = {"Id": [], "Name": [], "Departmant": []}

    for student in attendedStudents:
        data["Id"].append(student.id)
        data["Name"].append(student.name)
        data["Departmant"].append(student.dept)

    df = pd.DataFrame(data)
    df.to_string(fileName, index=False)


def excelExport(fileName: str):
    data = {"Id": [], "Name": [], "Departmant": []}

    for student in attendedStudents:
        data["Id"].append(student.id)
        data["Name"].append(student.name)
        data["Departmant"].append(student.dept)

    df = pd.DataFrame(data)
    df.to_excel(fileName, index=False)


def csvExport(fileName: str):
    data = {"Id": [], "Name": [], "Departmant": []}

    for student in attendedStudents:
        data["Id"].append(student.id)
        data["Name"].append(student.name)
        data["Departmant"].append(student.dept)

    df = pd.DataFrame(data)
    df.to_csv(fileName, index=False)


App("tk", (650, 380))
