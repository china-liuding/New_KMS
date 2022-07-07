import tkinter
import os
import tkinter.ttk

Window = tkinter.Tk()
Window.title("KMSon --version 2.0.0")
Window.geometry("500x500+100+100")


class KMS_ON:
    
    def __init__(self):
        self.office_bit_label = None
        self.office_bit_choose = None
        self.office_active_button = None
        self.office_version_choose = None
        self.office_choose_label = None
        self.office_label = None

        self.windows_button = None
        self.windows_label = None

        self.Active_Office_Var = tkinter.StringVar()

    @staticmethod
    def Active_Windows():
        os.system("slmgr /skms kms.luody.info")
        os.system("slmgr /ato")

    def Active_Office(self):
        if self.Active_Office_Var == "64":
            os.system("MondoVL.bat")
        else:
            os.system("MondoVL86.bat")

    def Active_Office_Tools(self, windows):
        self.office_label = tkinter.Label(windows, text="******激活Office******")
        self.office_choose_label = tkinter.Label(windows, text="请输入您想要激活的Office版本")
        self.office_version_choose = tkinter.ttk.Combobox(windows, width=22)
        self.office_version_choose["values"] = (
            "Microsoft 365(Office 365)",
            "Office 2021",
            "Office 2019",
            "Office 2016",
            "Office 2013",
            "Office 2010",
            "Office 2007",
        )
        self.office_version_choose.current(0)
        self.office_active_button = tkinter.Button(windows, text="激活", command=self.Active_Office)
        self.office_bit_choose = tkinter.ttk.Combobox(windows)
        self.office_bit_choose["values"] = (
            "64位",
            "32位",
        )
        self.office_bit_choose.current(0)
        self.office_bit_label = tkinter.Label(windows, text="请输入您安装的office的位数：")

    def Active_Windows_Tools(self, windows):
        self.windows_label = tkinter.Label(windows, text="******激活Windows******")
        self.windows_button = tkinter.Button(windows, text="直接激活Windows", command=self.Active_Windows)

    @staticmethod
    def Menus(windows):
        menu_father = tkinter.Menu(windows)

        tools_menu = tkinter.Menu(menu_father, tearoff=False)
        tools_menu.add_command(label="自动激活")
        tools_menu.add_command(label="指定Office安装位置激活")
        tools_menu.add_command(label="指定Office激活服务器激活")
        menu_father.add_cascade(label="工具", menu=tools_menu)

        windows.config(menu=menu_father)

    def Windows_Tools_remove(self):
        self.windows_label.grid_remove()
        self.windows_button.grid_remove()

    def Office_Tools_remove(self):
        self.office_label.grid_remove()
        self.office_choose_label.grid_remove()
        self.office_version_choose.grid_remove()
        self.office_active_button.grid_remove()
        self.office_bit_choose.grid_remove()
        self.office_bit_label.grid_remove()

    def Windows_Tools_grid(self):
        self.windows_label.grid(column=1, row=0)
        self.windows_button.grid(column=1, row=1)

    def Office_Tools_grid(self):
        self.office_label.grid(column=1, row=0)
        self.office_choose_label.grid(column=1, row=1)
        self.office_version_choose.grid(column=2, row=1)
        self.office_active_button.grid(column=3, row=2)
        self.office_bit_choose.grid(column=2, row=2)
        self.office_bit_label.grid(column=1, row=2)

    def Office_Button_(self):
        self.Windows_Tools_remove()
        self.Office_Tools_grid()

    def Windows_Button_(self):
        self.Office_Tools_remove()
        self.Windows_Tools_grid()

    def Buttons_Indexes(self, windows):
        button_1 = tkinter.Button(windows, text="激活Windows", width=15, command=self.Windows_Button_)
        button_2 = tkinter.Button(windows, text="激活Office", width=15, command=self.Office_Button_)

        button_1.grid(column=0, row=0)
        button_2.grid(column=0, row=1)

    def Run_Project(self, windows):
        self.Buttons_Indexes(windows)
        self.Menus(windows)
        self.Active_Office_Tools(windows)
        self.Active_Windows_Tools(windows)

        windows.mainloop()


if __name__ == '__main__':
    KMS = KMS_ON()
    KMS.Run_Project(Window)
