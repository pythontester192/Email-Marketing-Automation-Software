from tkinter import *
from tkinter import messagebox,filedialog
import os
import pandas as pd # pandas and pyxl for reading the content from the excel file
from PIL.ImageTk import PhotoImage

import code_email
import time

from PIL import ImageTk

root = Tk()


class Email:


    def __init__(self, root):
        self.root = root
        self.root.title("My_Email_sender:")
        self.root.geometry("1000x550+200+50")  ### 200 and 50 are x and y axizs from windo
        self.root.resizable(False, False)  ###SIZE CANT BERESIZE first false is width and second is height
        self.root.config(bg="skyblue")




        #########################  ICON #######################################
        self.Email_icon = ImageTk.PhotoImage(file="email01.png")
        self.Setting_icon=ImageTk.PhotoImage(file="setting.png")

        ##################### BUTTON #######################

        self.var_choice = StringVar()

        single = Radiobutton(root, text="Single", value="single", command=self.check_single_OR_bulk,
                             activebackground="skyblue", variable=self.var_choice, font=("times new roman", 30, "bold"),
                             bg="skyblue", fg="black").place(x=50, y=120)
        multiple = Radiobutton(root, text="Multiple", value="multiple", command=self.check_single_OR_bulk,
                               variable=self.var_choice, activebackground="skyblue",
                               font=("times new roman", 30, "bold"), bg="skyblue", fg="black").place(x=200, y=120)
        self.var_choice.set("single")




        ##################### TITLE ###########################

        title = Label(self.root, text="Bulk Email Sender Panel",font=("Goudy Old Style", 48, "bold"), bg="blue",
                      fg="white").place(x=0, y=0, relwidth=1)

        REF = Label(self.root, text="Use Excel File For Sending Bulk Email At Once", font=("Calibri (body)", 14,),
                    bg="yellow", fg="black").place(x=0, y=80, relwidth=1)

        btn_set1 = Button(self.root, image=self.Email_icon,
                          bg="blue", bd='0', command=LEFT, height=65, width=100,activebackground="blue").place(x=1, y=4)

        btn_set = Button(self.root,image=self.Setting_icon,
                         bg="blue", bd='0', command=self.setting_window,height= 65, width=100,activebackground="blue",).place(x=890, y=4)

        #########################   text   #####################################

        To = Label(self.root, text="To (Email Address)", font=("times new roman", 18, "bold"), bg="skyblue",
                   fg="black").place(x=50, y=200)
        Subject = Label(self.root, text="SUBJECT", font=("times new roman", 18, "bold"), bg="skyblue",
                        fg="black").place(x=50, y=250)
        Message = Label(self.root, text="MESSAGE", font=("times new roman", 18, "bold"), bg="skyblue",
                        fg="black").place(x=50, y=300)

        ############################################## STATUS ##################################################################

        self.Total = Label(self.root, font=("times new roman", 18, "bold"), bg="skyblue", fg="black")
        self.Total.place(x=50, y=500)

        self.Sent = Label(self.root, font=("times new roman", 18, "bold"), bg="skyblue", fg="darkgreen")
        self.Sent.place(x=350, y=500)

        self.Left = Label(self.root, font=("times new roman", 18, "bold"), bg="skyblue", fg="orange")
        self.Left.place(x=450, y=500)

        self.Failed = Label(self.root, font=("times new roman", 18, "bold"), bg="skyblue", fg="red")
        self.Failed.place(x=550, y=500)






        ############### ENTERY LABELS ##########################################

        self.to_entry = Entry(self.root, font=("times new roman", 18), bg="lightgrey")
        self.to_entry.place(x=280, y=200, width=350, height=30)

        self.sub_entry = Entry(self.root, font=("times new roman", 18), bg="lightgrey")
        self.sub_entry.place(x=280, y=250, width=450, height=30)

        self.message_entry = Text(self.root, font=("times new roman", 18), bg="lightgrey")
        self.message_entry.place(x=280, y=300, width=700, height=190)

        btn1 = Button(root, activebackground="skyblue", command=self.send_email, text="SEND",
                      font=("times new roman", 20, "bold"), bg="black",
                      fg="white").place(x=700, y=500, width=130, height=30)
        btn2 = Button(root, activebackground="skyblue", command=self.clear1, text="CLEAR",
                      font=("times new roman", 20, "bold"), bg="#ffcccb",
                      fg="black").place(x=850, y=500, width=130, height=30)
        self.btn3 = Button(root, activebackground="skyblue", text="BROWSE", font=("times new roman", 20, "bold"),
                           bg="lightblue", command=self.Browse_button, cursor="hand2", state=DISABLED, fg="black")
        self.btn3.place(x=650, y=200, width=150, height=30)


        ################################################ Browse ########################################################################

        self.check_file_exist()


    def Browse_button(self):
        op = filedialog.askopenfile(initialdir='/', title="Select Excel File for Emails",
                                    filetypes=(("All Files", "*.*"), ("Excel Files", ".xlsx")))
        if op != None:

            data = pd.read_excel(op.name)

            if 'Email' in data.columns:
                self.EMAIL = list(data['Email'])
                # print(EMAIL)
                c = []
                for i in self.EMAIL:
                    # print(i)
                    if (pd.isnull(i)) == False:
                        ## givs data only which is filled
                        # print(i)
                        c.append(i)
                self.EMAIL = c
                if len(self.EMAIL) > 0:
                    self.to_entry.config(state=NORMAL)
                    self.to_entry.delete(0, END)
                    self.to_entry.insert(0, str(op.name.split("/")[-1]))
                    self.to_entry.config(state='readonly')
                    self.Total.config(text="Total: " + str(len(self.EMAIL)) )
                    self.Sent.config(text="Sent: ")
                    self.Left.config(text="Left: ")
                    self.Failed.config(text="Failed: ")
                # print(EMAIL)


            else:
                messagebox.showinfo("Error", "Please Select A File Which Has Emails", parent=self.root)



    #################################### SEND EMAIL #############################

    def send_email(self):
        x = len(self.message_entry.get('1.0', END))
        if self.to_entry.get() == "" or self.sub_entry.get() == "" or x == 1:
            messagebox.showerror("ERROR", "All feilds are required", parent=self.root)
        else:
            if self.var_choice.get() == "single":
                status=code_email.Email_send_function(self.to_entry.get(),self.sub_entry.get(),self.message_entry.get('1.0',END),self.uname,self.pasw)
                if status=="s":
                    messagebox.showinfo("SUCCESS", "Email Has Been Sent", parent=self.root)
                if status=="f":
                    messagebox.showerror("Failed", "Email Not Sent", parent=self.root)

            if self.var_choice.get()=="multiple":
                self.failed = []
                self.s_count=0
                self.f_count = 0
                for x in self.EMAIL:
                    status=code_email.Email_send_function(x,self.sub_entry.get(),self.message_entry.get('1.0',END),self.uname,self.pasw)

                    if status=="s":
                       self.s_count+=1
                    if status=="f":
                       self.f_count+=1
                    self.status_bar()
                    time.sleep(1)


                messagebox.showinfo("Success", "Email Has Been Sent,Please Check Status....", parent=self.root)






    def clear1(self):
        self.to_entry.config(state=NORMAL)
        self.to_entry.delete(0, END)
        self.sub_entry.delete(0, END)
        self.message_entry.delete('1.0', END)
        self.var_choice.set("single")
        self.btn3.config(state=DISABLED)
        self.Total.config(text="")
        self.Sent.config(text="")
        self.Left.config(text="")
        self.Failed.config(text="")



    def status_bar(self):
        self.Total.config(text="Status " + str(len(self.EMAIL))+":-")
        self.Sent.config(text="Sent: "+ str(self.s_count))
        self.Left.config(text="Left: "+ str(len(self.EMAIL)-(self.f_count+self.s_count)))
        self.Failed.config(text="Failed: "+ str(self.f_count))
        self.Total.update()
        self.Sent.update()
        self.Left.update()
        self.Failed.update()





    def check_single_OR_bulk(self):
        if self.var_choice.get() == "single":
            messagebox.showinfo("single", "Setted To Single", parent=self.root)
            self.btn3.config(state=DISABLED)
            self.to_entry.config(state=NORMAL)
            self.to_entry.delete(0, END)
            self.clear1()

        if self.var_choice.get() == "multiple":
            messagebox.showinfo("multiple", "Setted To Bulk", parent=self.root)

            self.btn3.config(state=NORMAL)
            self.to_entry.delete(0, END)
            self.to_entry.config(state='readonly')



        ####################################### SETTING FUNCTION ########################

    def setting_clear(self):
        self.uname_entry.delete(0, END)
        self.pasw_entry.delete(0, END)

    def setting_window(self):
        self.check_file_exist()
        self.root2 = Toplevel()
        self.root2.title("Setting")
        self.root2.resizable(False, False)
        self.root2.geometry("700x450+350+90")
        self.root2.focus_force()
        self.root2.grab_set()
        self.root2.config(bg="lightgrey")
        title2 = Label(self.root2, text="Bulk Email Sender", padx=10, compound=LEFT,
                       font=("Goudy Old Style", 48, "bold"), bg="black",
                       fg="white").place(x=0, y=0, relwidth=1)
        REF2 = Label(self.root2, text="Enter your valid Email Id and Password", font=("Calibri (body)", 14,),
                     bg="yellow", fg="black").place(x=0, y=80, relwidth=1)

        uname = Label(self.root2, text="Email Address", font=("times new roman", 18, "bold"), bg="lightgrey",
                      fg="black").place(x=50, y=150)

        pasw = Label(self.root2, text="Password", font=("times new roman", 18, "bold"), bg="lightgrey",
                     fg="black").place(x=50, y=200)

        self.uname_entry = Entry(self.root2, font=("times new roman", 18), bg="lightyellow")
        self.uname_entry.place(x=250, y=150, width=330, height=30)

        self.pasw_entry = Entry(self.root2, font=("times new roman", 18), bg="lightyellow", show="*")
        self.pasw_entry.place(x=250, y=200, width=330, height=30)



        ################################################# BUTTON OF SETTING ############################################

        btn1 = Button(self.root2, activebackground="skyblue", text="SEND", font=("times new roman", 20, "bold"),
                      bg="black",
                      fg="white", command=self.save_setting).place(x=250, y=250, width=130, height=30)
        btn2 = Button(self.root2, activebackground="skyblue", text="CLEAR", font=("times new roman", 20, "bold"),
                      bg="#ffcccb", command=self.setting_clear,
                      fg="black").place(x=400, y=250, width=130, height=30)

        self.uname_entry.insert(0, self.uname)
        self.pasw_entry.insert(0, self.pasw)




    #######################################  FOR EMAIL AND PASS IN SETTING ##############################################

    def check_file_exist(self):
        if os.path.exists("important.txt") == False:
            f = open('important.txt', 'w')
            f.write(",")
            f.close()
        f2 = open('important.txt', 'r')
        self.credentials = []
        for i in f2:
            self.credentials.append([i.split(",")[0], i.split(",")[1]])
        # print(self.credentials)
        self.uname = self.credentials[0][0]
        self.pasw = self.credentials[0][1]
        # print(self.uname,self.pasw)

    def save_setting(self):
        if self.uname_entry.get() == "" or self.pasw_entry.get() == "":
            messagebox.showinfo("ERROR", "All feilds are required", parent=self.root2)

        else:
            f = open('important.txt', 'w')
            f.write(self.uname_entry.get() + "," + self.pasw_entry.get())
            f.close()
            messagebox.showinfo("Sent", "Email and password are saved Successfully", parent=self.root2)
            self.check_file_exist()

obj = Email(root)
root.mainloop()
