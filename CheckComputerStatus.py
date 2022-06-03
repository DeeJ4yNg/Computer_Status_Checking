import tkinter as tk
import pandas as pd
from tkinter import ttk
import tkinter.messagebox as messagebox

def Read_CSV(Computer_Obj):
    #read computer info csv
    csv = pd.read_csv('./Computer_Status.csv')
    Computer_Obj = Computer_Obj.upper()
    Computer = csv[csv['Machine'] == Computer_Obj]
    Computer = Computer.fillna('-')
    #read patch status csv
    csv_patch_info = pd.read_csv('./Compliance_Status.csv')
    Computer_patch = csv_patch_info[csv_patch_info['ComputerName'] == Computer_Obj]
    failed_patches = []
    for i in Computer_patch.values:
        if i[1] == '2' or i[1] == '0':
            failed_patches.append(i[3])

    if Computer['Machine'].empty:
        messagebox.showerror('Computer Info Checking Tool',
                			'Error, Computer was not found.')
    else:
        Computer_Name = Computer['Machine'].values[0]
        Computer_Account_Status = 'True' if Computer['Enabled'].values[0] == 1 else 'False'
        Model = Computer['Model'].values[0]
        User = Computer['User_name'].values[0]
        User_IID = Computer['User_IID'].values[0]
        Main_OU = Computer['OU'].values[0]
        Office = Computer['ADOU'].values[0]
        Grade = Computer['Grade'].values[0]
        CMDB_Status = Computer['Status'].values[0]
        GD = Computer['System'].values[0]
        Office_365_Version = Computer['Version0'].values[0]
        Acrobat_Version = Computer['Version1'].values[0]

        if Office != 'S01' and Office != 'S02':
            Restriction = '-'
        elif Office == 'S01':
            Restriction = 'Restriction level 1'
        elif Office == 'S02':
            Restriction = 'Restriction level 2'
       
        AD_Description = Computer['Description'].values[0]
        CMDB_Remark = Computer['remarks'].values[0]
        Computer_Patch_Status = failed_patches

    return Computer_Name,Computer_Account_Status,str(Model).replace(" ", "\ "),str(User).replace(" ", "\ "),ID,OU,Office,Grade,Status,System,Version0,Version1,str(Restriction).replace(" ","\ "),str(Description).replace(" ","\ "), str(Remark).replace(" ","\ "), Compliance_Status

class POPUP():
    def __init__(self):
        pass

    def TKWindow(self):
        window = tk.Tk()
        window.title('Computer Info Checking Tool')
        window.geometry('450x150')
        var = tk.StringVar()
        l = tk.Label(window, text='This is a tool for checking computer information.', bg='white', font=('Arial', 12),
                 width=50, height=2)
        l.pack()
    #l2 = tk.Label(window, text='Please enter a computer name below', bg='white', font=('Arial', 12),
    #             width=50, height=2)
    #l2.pack()
        l3 = tk.Label(window, text='Please type a computer name:')
        l3.pack()
        E1 = tk.Entry(window, bd=10)
        E1.pack()
        def hitbutton():
            global on_hit
            on_hit = False
            if on_hit == False:
                on_hit = True
                self.Result(Read_CSV(E1.get()))
            else:
                on_hit = False

        b = tk.Button(window, text='Check', font=('Arial', 12), width=10, height=1, command=hitbutton)
        b.pack()
        window.mainloop()

    def Result(self,info):
        window_Table = tk.Tk()
        window_Table.title('Computer Info Checking Tool')
        window_Table.geometry('750x450')
        tree = ttk.Treeview(window_Table, height=50)
        tree['columns'] = ("Computer Info", "Value")
        tree.column("Computer Info", width=750)
        #tree.column("Value", width=300)
        tree.heading("Computer Info", text="Computer Info", anchor='w')
        tree.tag_configure('tag', foreground='red', font=('Arial Bold', 10))
        tree.tag_configure('tag_blue', foreground='blue')
        #tree.heading("Value", text="Value")
        tree.insert("", 0, text="Computer_Name", values=info[0])
        tree.insert("", 1, text="Computer_Account_Status", values=info[1])
        tree.insert("", 2, text="Model", values=info[2])
        tree.insert("", 3, text="OU", values=info[5])
        tree.insert("", 4, text="Office", values=info[6])
        tree.insert("", 5, text="Status", values=info[8])
        tree.insert("", 6, text="User", values=info[3])
        tree.insert("", 7, text="ID", values=info[4])
        tree.insert("", 8, text="Grade", values=info[7])
        tree.insert("", 9, text="System", values=info[9], tags='tag_blue')
        tree.insert("", 10, text="Version0", values=info[10], tags='tag_blue')
        tree.insert("", 11, text="Version1", values=info[11], tags='tag_blue')
        tree.insert("", 12, text="Restriction", values=info[12])
        tree.insert("", 13, text="Description", values=info[13])
        tree.insert("", 14, text="Remark", values=info[14])
        Missing_Patches_Count = 14
        for j in info[15]:
            Missing_Patches_Count += 1
            tree.insert("", Missing_Patches_Count, text="Missing_Security_Patch", values=str(j).replace(" ", "\ "))
        if info[6] == 'S01' or info[6] == 'S02':
            items = list(tree.get_children())
            tree.item(items[12], tag='tag')
        tree.pack()
        window_Table.mainloop()


a = POPUP()
a.TKWindow()
