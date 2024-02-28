import tkinter,os,openpyxl
from tkcalendar import DateEntry
from openpyxl.styles import Font


class data_entry:
    def __init__(self):
        self.window = tkinter.Tk()
        self.window.title("Data Entry Form")
        self.frame = tkinter.Frame(self.window)
        self.frame.pack()


    def enter_data(self):
        self.accepted = self.accept_var.get()
        
        if self.accepted=="Accepted":
            # User info
            self.sr_no = self.serial_no_entry.get()
            self.firstname = self.first_name_entry.get()
            self.lastname = self.last_name_entry.get()
            self.date = self.cal.get()
        
            if self.firstname and self.lastname and self.sr_no:
                # Course info
                self.registration_status = self.reg_status_var.get()
                self.starthrs = self.starthrs_spinbox.get()
                self.stophrs = self.stophrs_spinbox.get()
                total_hr = int(self.stophrs) - int(self.starthrs)
                self.bucket = self.buc.get()
                self.dis = self.diesel.get()
                self.rt = self.rate.get()
                print("Sr No.: ", self.sr_no, "Date: ", self.date, "First name: ", self.firstname, "Last name: ", self.lastname)
                print("Start Hrs: ", self.starthrs, "Stop Hrs: ", self.stophrs)
                print("Total Hrs: ",total_hr, "Bucket: ",self.bucket, "Diesel: ",self.dis, "Rate: ",self.rt)
                print("Working Status", self.registration_status)
                print("------------------------------------------")
                
                filepath = "data.xlsx"
                
                if not os.path.exists(filepath):
                    workbook = openpyxl.Workbook()
                    sheet = workbook.active
                    heading = ["Sr No", "Date", "First Name", "Last Name",
                            "Start Hrs", "Stop Hrs", "Total Hrs", "Bucket", "Diesel", "Working status"]
                    sheet.append(heading)
                    workbook.save(filepath)
                workbook = openpyxl.load_workbook(filepath)
                sheet = workbook.active
                sheet.append([self.sr_no, self.date, self.firstname, self.lastname, self.starthrs,
                            self.stophrs, total_hr, self.bucket, self.dis, self.registration_status])
                workbook.save(filepath)      
            else:
                tkinter.messagebox.showwarning(title="Error", message="First name and last name are required.")
        else:
            tkinter.messagebox.showwarning(title= "Error", message="You have not accepted the terms")

    def clearfunc(self):
        self.serial_no_entry.delete(0, tkinter.END)
        self.first_name_entry.delete(0, tkinter.END)
        self.last_name_entry.delete(0, tkinter.END)
        self.reg_status_var.set(0)
        self.starthrs_spinbox.delete(0, tkinter.END)
        self.stophrs_spinbox.delete(0, tkinter.END)
        self.buc.delete(0, tkinter.END)
        self.diesel.delete(0, tkinter.END)
        self.rate.delete(0, tkinter.END)
        
    def main(self):
        # Saving User Info
        self.user_info_frame =tkinter.LabelFrame(self.frame, text="User Information")
        self.user_info_frame.grid(row= 0, column=0, padx=20, pady=10)

        self.serial_no_label = tkinter.Label(self.user_info_frame, text="Sr No.")
        self.serial_no_label.grid(row=0, column=0)
        self.first_name_label = tkinter.Label(self.user_info_frame, text="First Name")
        self.first_name_label.grid(row=0, column=2)
        self.last_name_label = tkinter.Label(self.user_info_frame, text="Last Name")
        self.last_name_label.grid(row=0, column=3)

        self.serial_no_entry = tkinter.Entry(self.user_info_frame)
        self.first_name_entry = tkinter.Entry(self.user_info_frame)
        self.last_name_entry = tkinter.Entry(self.user_info_frame)
        self.serial_no_entry.grid(row=1, column=0)
        self.first_name_entry.grid(row=1, column=2)
        self.last_name_entry.grid(row=1, column=3)

        self.cal_label = tkinter.Label(self.user_info_frame, text="Date")
        self.cal = DateEntry(self.user_info_frame,selectmode='day')
        self.cal_label.grid(row=0, column=1)
        self.cal.grid(row=1,column=1,padx=15)

        for widget in self.user_info_frame.winfo_children():
            widget.grid_configure(padx=10, pady=5)

        # Saving Course Info
        self.work_frame = tkinter.LabelFrame(self.frame)
        self.work_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

        self.working_label = tkinter.Label(self.work_frame, text="Working Status")

        self.reg_status_var = tkinter.StringVar(value="Not Work Done")
        self.registered_check = tkinter.Checkbutton(self.work_frame, text="Work Done",
                                            variable=self.reg_status_var, onvalue="Work Done", offvalue="Not Work Done")

        self.working_label.grid(row=0, column=5)
        self.registered_check.grid(row=1, column=5)

        self.starthrs_label = tkinter.Label(self.work_frame, text= "Start Hrs")
        self.starthrs_spinbox = tkinter.Entry(self.work_frame, width=20)
        self.starthrs_label.grid(row=0, column=0)
        self.starthrs_spinbox.grid(row=1, column=0)

        self.stophrs_label = tkinter.Label(self.work_frame, text="Stop Hrs")
        self.stophrs_spinbox = tkinter.Entry(self.work_frame, width=20)
        self.stophrs_label.grid(row=0, column=1)
        self.stophrs_spinbox.grid(row=1, column=1)

        self.buc_label = tkinter.Label(self.work_frame, text="Bucket")
        self.buc_label.grid(row=0, column=2)
        self.buc = tkinter.Entry(self.work_frame, width=20)
        self.buc.grid(row=1, column=2)
        
        self.diesel_label = tkinter.Label(self.work_frame, text="Diesel")
        self.diesel_label.grid(row=0, column=3)
        self.diesel = tkinter.Entry(self.work_frame, width=20)
        self.diesel.grid(row=1, column=3)
        
        self.rate_label = tkinter.Label(self.work_frame, text="Rate")
        self.rate_label.grid(row=0, column=4)
        self.rate = tkinter.Entry(self.work_frame, width=20)
        self.rate.grid(row=1, column=4)

        for widget in self.work_frame.winfo_children():
            widget.grid_configure(padx=10, pady=5)

        # Accept terms
        self.terms_frame = tkinter.LabelFrame(self.frame, text="Terms & Conditions")
        self.terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

        self.accept_var = tkinter.StringVar(value="Not Accepted")
        self.terms_check = tkinter.Checkbutton(self.terms_frame, text= "I accept the terms and conditions.",
                                        variable=self.accept_var, onvalue="Accepted", offvalue="Not Accepted")
        self.terms_check.grid(row=0, column=0)
        
        self.button1 = tkinter.Button(self.terms_frame, text="Reset", command= self.clearfunc)
        self.button1.grid(row=0, column=2, sticky="news", padx=20, pady=10)

        # Button
        self.button = tkinter.Button(self.frame, text="Enter data", bg='green', command= self.enter_data)
        self.button.grid(row=3, column=0, sticky="news", padx=20, pady=10)
        
        self.window.mainloop()
        
if __name__ == "__main__":
    obj = data_entry()
    obj.main()
