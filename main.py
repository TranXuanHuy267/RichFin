import tkinter as tk
from tkinter import ttk
from tkinter import *
from datetime import date

import xlwings as xw
import pandas as pd

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        # Max geometry: 1430 * 860
        self.height = 800
        self.width = 650
        self.geometry(str(self.width) + "x" + str(self.height) + "+" + str(int((1430-self.width)/2)) + "+" + str(int((860 - self.height)/2)))
        self.title("Rich Finance")
        # self.configure(background="pink")
        self.menu = tk.StringVar()
        self.menu2 = tk.StringVar()
        self.input1 = tk.IntVar()
        self.input2 = tk.IntVar()
        self.input3 = tk.IntVar()
        self.input4 = tk.StringVar()
        self.input5 = tk.IntVar()
        self.input6 = tk.IntVar()
        self.input7 = tk.IntVar()
        self.input8 = tk.IntVar()
        self.input9 = tk.IntVar()
        self.type1 = ("What?", "Income", "Expense", "Assets", "Liabilities")
        self.type2 = ("What more?","Earned", "Dividend", "Passive")
        self.remember = []
        self.row_update = []
        self.create_widgets()

    def create_widgets(self):
        ttk.Entry(self, width = 2, textvariable=self.input1, justify='center').grid(row = 2, column = 4, sticky=tk.W)
        ttk.Entry(self, width = 4, textvariable=self.input2, justify='center').grid(row = 2, column = 5)
        ttk.Entry(self, width = 4, textvariable=self.input3, justify='center').grid(row = 2, column = 6, sticky=tk.E)
        ttk.Entry(self, width = 20, textvariable=self.input4, justify='center').grid(row = 5, column = 3)
        ttk.Entry(self, width = 13, textvariable=self.input5, justify='center').grid(row = 5, column = 4, columnspan=3)
        
        tk.Button(self, width = 7, text='Update', fg='black', command=self.Update_add).grid(row = 5, column = 10)
        tk.Button(self, width = 7, text='Submit', fg='black', command=self.Submit).grid(row = 10, column = 1, columnspan=10)
        tk.Button(self, width = 7, text='Remove', fg='black', command=self.Update_del).grid(row = 6, column= 10)
        self.track = tk.Button(self, width = 7, text=['Extract ⇨' if self.width == 650 else '⇦ Rolling'][0], fg='black', command=self.Extract).grid(row = 7, column = 10)
        
        ttk.Label(self, text="Welcome to the private secret!").grid(row = 0, column = 1, columnspan = 10, pady=30)
        ttk.Label(self, text="Day   Month     Year  ").grid(row = 1, column = 4, columnspan=3)
        Label(self, text="", width=88, height=10).grid(row = 11, column = 1, columnspan=11)
        
        self.extension_frame = Label(self, text = "", width=16, height=55)
        self.extension_frame['bg'] = 'aliceblue'
        self.extension_frame.grid(column=11, row=0, rowspan=12)

        # ttk.Label(self, text = 'leuleu').grid(row = 8, column = 8)
        # ttk.Label(self, text = 'leuleu').grid(row = 8, column = 8)

        today = date.today()
        self.input1.set(today.day)
        self.input2.set(today.month)
        self.input3.set(today.year)
        if self.input4.get() == "": self.input4.set("What more and more...?")
        self.input5.set(0)

        men = ttk.OptionMenu(self, self.menu, "What?", *self.type1, command=self.Yourchoice)
        men.config(width=7)
        men["menu"].config(bg="black")
        men.grid(row = 5, column = 1, sticky=tk.W)
        self.set_menu2()

        self.Frame_Update()

    def set_menu2(self):
        men2 = ttk.OptionMenu(self, self.menu2, self.type2[0], *self.type2)
        men2.config(width=8)
        men2["menu"].config(bg="black")
        men2.grid(row = 5, column = 2)

    
    def Yourchoice(self, *args):
        lst = {
            "Income": ("What more?","Earned", "Dividend", "Passive"),
            "Expense": ("What more?", "Rent", "Technology", "Invest", "Meal", "Transport"),
            "Assets": ("What more?", "Real Estate", "Stock", "Debt"),
            "Liabilities": ("What more?", "Room", "Motorbike", "Loan")
        }
        self.type2 = lst[self.menu.get()]
        self.set_menu2()

    def Frame_Update(self):
        self.frame = tk.Frame(self)
        self.frame.grid(row = 9, column = 1, columnspan=10, pady=10)
        self.tree = ttk.Treeview(self.frame, height=20)
        self.tree["columns"] = ['','Year', 'Month', 'Day', 'Type', 'Detail', 'Money']
        self.tree['show'] = "headings"
        self.tree.column('', anchor=CENTER, width=30)
        self.tree.column('Year', anchor=CENTER, width=40)
        self.tree.column('Month', anchor=CENTER, width=40)
        self.tree.column('Day', anchor=CENTER, width=30)
        self.tree.column('Type', anchor=CENTER, width=100)
        self.tree.column('Detail', anchor=CENTER, width=200)
        self.tree.column('Money', anchor=CENTER, width=80)
        for col in self.tree["column"]:
            self.tree.heading(col, text=col, anchor=CENTER)
        self.tree.grid(row = 9, column = 1)

    def Update_add(self):
        new_row = [len(self.row_update)+1,self.input3.get(), self.input2.get(), self.input1.get(), self.menu2.get(), self.input4.get(), self.input5.get()]
        for row in self.row_update:
            if new_row[1:-1] == row[1:-1]:
                break

        self.row_update.append(new_row)
        self.tree.insert("", "end", values=self.row_update[-1])

    
    def Update_del(self):
        selected_item = self.tree.selection()
        for item in selected_item[::-1]:
            item_index = list(self.tree.get_children()).index(item)
            self.tree.delete(item)
            self.row_update.pop(item_index)

        for i in range(len(self.row_update)):
            self.row_update[i][0] = i+1

        self.Frame_Update()
        for row in self.row_update:
            self.tree.insert("", "end", values=row)


    def Submit(self):
        filename = 'richfin.xlsx'
        sheetname = str(self.menu.get())
        workbook = xw.Book(filename)
        sheet = workbook.sheets(sheetname)
        data = pd.read_excel(filename, sheet_name=sheetname)
        next = 'A' + str(len(data)+2)
        sheet.range(next).value = [row[1:] for row in self.row_update]
        workbook.save()

    def Extract(self):
        if self.width == 780:
            self.width = 650
            self.geometry(str(self.width) + "x" + str(self.height) + "+" + str(int((1430-self.width)/2)) + "+" + str(int((860 - self.height)/2)))
            list = self.grid_slaves()
            for l in list:
                l.destroy()
            self.create_widgets()
            self.row_update = []
        else:
            self.width = 780
            self.geometry(str(self.width) + "x" + str(self.height) + "+" + str(int((1430-self.width)/2)) + "+" + str(int((860 - self.height)/2)))
            self.Extensions()
            self.row_update = []

    def Extensions(self):      
        ttk.Entry(self, width = 2, textvariable=self.input6, justify='center').grid(row = 2, column = 7, sticky=tk.W)
        ttk.Entry(self, width = 4, textvariable=self.input7, justify='center').grid(row = 2, column = 8)
        ttk.Entry(self, width = 4, textvariable=self.input8, justify='center').grid(row = 2, column = 9, sticky=tk.E)
        ttk.Entry(self, width = 13, textvariable=self.input9, justify='center').grid(row = 5, column = 7, columnspan=3)

        ttk.Label(self, text="Day   Month     Year  ").grid(row = 1, column = 7, columnspan=3)
        tk.Button(self, width = 7, text=['Extract ⇨' if self.width == 650 else '⇦ Rolling'][0], fg='black', command=self.Extract).grid(row = 7, column = 10)
        tk.Button(self, width = 7, text='Find', fg='black', command=self.Get_Data).grid(row = 8, column = 10)
        today = date.today()
        self.input6.set(today.day)
        self.input7.set(today.month)
        self.input8.set(today.year)
        self.input9.set(0)
        self.Frame_Update()

    def Get_Data(self):
        filename = 'richfin.xlsx'
        Find = [self.menu.get(), self.menu2.get(), self.input1.get(), self.input2.get(), self.input3.get(), self.input4.get(), self.input5.get(), self.input6.get(), self.input7.get(), self.input8.get(), self.input9.get()]
        Full_Data = []
        workbook = xw.Book(filename)
        if Find[0] == "What?":
            sheet_name_list = list(self.type1[1:])
            for sheetname in sheet_name_list:
                sheet = workbook.sheets(sheetname)
                data = pd.read_excel(filename, sheet_name=sheetname)
                Full_Data += data.to_dict(orient='records')
        else:
            sheetname = Find[0]
            sheet = workbook.sheets(sheetname)
            data = pd.read_excel(filename, sheet_name=sheetname)
            Full_Data = data.to_dict(orient='records')
        Output = []
        index = 0
        sum = 0
        for line in Full_Data:
            if Find[5] != "What more and more...?" and Find[5]!="" and Find[5] != line['Detail']: continue
            if Find[1] != "What more?" and Find[1] != line['Type']: continue
            if int(line['Year']) < int(Find[4]) or int(line['Year']) > int(Find[9]): continue
            if int(Find[4]) == int(Find[9]) and (int(line['Month']) < int(Find[3]) or int(line['Month']) > int(Find[8])): continue
            if int(Find[4]) == int(Find[9]) and int(Find[3]) == int(Find[8]) and (int(line['Day']) < int(Find[2]) or int(line['Day']) > int(Find[7])): continue
            if ((Find[10] != 0 and Find[10]!= "") or (Find[6] != 0 and Find[6]!= "")) and (int(Find[10]) < int(line['Money']) or int(Find[6]) > int(line['Money'])): continue
            index += 1
            sum += int(line['Money'])
            Output.append([index, line['Year'], line['Month'], line['Day'], line['Type'], line['Detail'], line['Money']])
        


        if len(Output) == 0:
            self.row_update = [[index, 00, 00, 0000, Find[1], "Nothing", 0]]
            self.Frame_Update()
            for row in self.row_update:
                self.tree.insert("", "end", values=row)
        else:
            Output.append(["", "", "", "", "", "", sum])
            self.row_update = Output
            self.Frame_Update()
            for row in self.row_update:
                self.tree.insert("", "end", values=row)
            


if __name__ == '__main__':
    app = App()
    app.mainloop()