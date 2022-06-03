import tkinter as tk
import pandas as pd
from tkinter import ttk
import tkinter.messagebox as messagebox

def Read_CSV(Computer_Obj):
    #read computer info csv
    csv = pd.read_csv(r'\\cnfuosms02\SecurityUpdateSelfCheck\AD_CMDB_OPP_Acrobat.csv')
    Computer_Obj = Computer_Obj.upper()
    Computer = csv[csv['AD_Machine'] == Computer_Obj]
    Computer = Computer.fillna('-')
    #read patch status csv
    csv_patch_info = pd.read_csv(r'\\cnfuosms02\SecurityUpdateSelfCheck\Patch_Deployment_Status.csv')
    Computer_patch = csv_patch_info[csv_patch_info['ComputerName'] == Computer_Obj]
    failed_patches = []
    for i in Computer_patch.values:
        if i[1] == '2' or i[1] == '0':
            failed_patches.append(i[3])

    if Computer['AD_Machine'].empty:
        messagebox.showerror('Computer Info Checking Tool – KPMG SCCM',
                			'Error, Computer was not found.')
    else:
        Computer_Name = Computer['AD_Machine'].values[0]
        Computer_Account_Status = 'True' if Computer['Enabled'].values[0] == 1 else 'False'
        Model = Computer['CMDB_Model_Desc'].values[0]
        User = Computer['CMDB_display_name'].values[0]
        User_IID = Computer['User_IID'].values[0]
        Main_OU = Computer['OU'].values[0]
        Office = Computer['SubOU'].values[0]
        Grade = Computer['Grade'].values[0]
        CMDB_Status = Computer['CMDB_Status'].values[0]
        GD = Computer['GD'].values[0]
        Office_365_Version = Computer['OPPVersion'].values[0]
        Acrobat_Version = Computer['AcrobatVersion'].values[0]

        if Office != 'Servicing01' and Office != 'Servicing02' and Office != 'Servicing03':
            Restriction = '-'
        elif Office == 'Servicing01':
            Restriction = 'Restriction level 1, enabled watermark and reminder'
        elif Office == 'Servicing02':
            Restriction = 'Restriction level 2, enabled watermark and reminder, restricted browsers'
        else:
            Restriction = 'Restriction level 3, enabled watermark and reminder, restricted browsers, Office and Acrobat applications'

        AD_Description = Computer['AD_Description'].values[0]
        CMDB_Remark = Computer['ci_remarks'].values[0]
        Computer_Patch_Status = failed_patches

    return Computer_Name,Computer_Account_Status,str(Model).replace(" ", "\ "),str(User).replace(" ", "\ "),User_IID,Main_OU,Office,Grade,CMDB_Status,GD,Office_365_Version,Acrobat_Version,str(Restriction).replace(" ","\ "),str(AD_Description).replace(" ","\ "), str(CMDB_Remark).replace(" ","\ "), Computer_Patch_Status

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
        tree.insert("", 3, text="Main_OU", values=info[5])
        tree.insert("", 4, text="Sub_OU", values=info[6])
        tree.insert("", 5, text="CMDB_Status", values=info[8])
        tree.insert("", 6, text="User", values=info[3])
        tree.insert("", 7, text="User_IID", values=info[4])
        tree.insert("", 8, text="Grade", values=info[7])
        tree.insert("", 9, text="GD", values=info[9], tags='tag_blue')
        tree.insert("", 10, text="Office_365_Version", values=info[10], tags='tag_blue')
        tree.insert("", 11, text="Acrobat_Version", values=info[11], tags='tag_blue')
        tree.insert("", 12, text="Restriction", values=info[12])
        tree.insert("", 13, text="AD_Description", values=info[13])
        tree.insert("", 14, text="CMDB_Remark", values=info[14])
        #tree.insert("", 13, text="Computer_Missing_Patches", values=info[13])
        Missing_Patches_Count = 14
        for j in info[15]:
            Missing_Patches_Count += 1
            tree.insert("", Missing_Patches_Count, text="Missing_Security_Patch", values=str(j).replace(" ", "\ "))
        if info[6] == 'Servicing01' or info[6] == 'Servicing02' or info[6] == 'Servicing03':
            items = list(tree.get_children())
            tree.item(items[12], tag='tag')
        tree.pack()
        window_Table.mainloop()


a = POPUP()
a.TKWindow()
#print(Read_CSV('cnpc0gh2ve'))
#CNPC1XF068