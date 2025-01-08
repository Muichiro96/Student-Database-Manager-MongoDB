#Importing all the modules required for the app to function
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from pymongo import MongoClient,errors
import csv

#Install the following package " MongoDB-clientector-python" using "pip install"

class StudentManagementSystem: #Starting the App 
    def __init__(self, root): #Declaration of the App
        self.root = root 
        self.root.title("Student Management System")
        self.root.geometry("1250x625")
        self.root.resizable(False, False)
        self.root.tk.call("source", "azure.tcl")
        self.root.tk.call("set_theme", "light")
        
        #Style configuration
        style = ttk.Style()
        style.configure(
                "Treeview", #Widget Nme 
                background="black", 
                foreground="white",
                fieldbackground="black",
                font=("Helvetica", 10),
                rowheight=25
                )
        style.configure(
                "Accent.TButton", 
                foreground="white", 
                background="blue",
                borderwidth=0,
                relief=tk.RAISED,
                padding=10
                )

        #All Variables 
        self.students = []
        self.Student_Id_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.grade_section_var = tk.StringVar()
        self.email_var = tk.StringVar()
        self.gender_var = tk.StringVar()
        self.contact_var = tk.StringVar()
        self.search_by = tk.StringVar()
        self.search_txt = tk.StringVar()


        #Main Title of the App 
        self.title = ttk.Label(root, text="Student Management System", font=("bold", 16), style="Title.TLabel")
        self.title.pack(pady=20)


        
        #The frame which contains all the Field Entries like "Name, Email" etc.
        self.manage_frame = ttk.Frame(root, borderwidth=0, relief="flat", style="Frame.TFrame")
        self.manage_frame.pack(side=tk.LEFT, padx=20, pady=20, expand=True, fill=tk.BOTH)
        
        self.manage_title = ttk.Label(self.manage_frame, text="Manage Students", font=("bold", 16), style="Title.TLabel")
        self.manage_title.pack(side=tk.TOP, pady=10)

        self.roll_label = ttk.Label(self.manage_frame, text="Student ID", font=("bold", 11), style="Label.TLabel")
        self.roll_label.pack()
        self.roll_entry = ttk.Entry(self.manage_frame, textvariable=self.Student_Id_var, font=("bold", 9), style="Entry.TEntry")
        self.roll_entry.pack()

        self.name_label = ttk.Label(self.manage_frame, text="Name", font=("bold", 11), style="Label.TLabel")
        self.name_label.pack()
        self.name_entry = ttk.Entry(self.manage_frame, textvariable=self.name_var, font=("bold", 9), style="Entry.TEntry")
        self.name_entry.pack()

        self.grade_section_label = ttk.Label(self.manage_frame, text="Grade and Section", font=("bold", 11), style="Label.TLabel")
        self.grade_section_label.pack()
        self.grade_section_entry = ttk.Entry(self.manage_frame, textvariable=self.grade_section_var, font=("bold", 9), style="Entry.TEntry")
        self.grade_section_entry.pack()

        self.email_label = ttk.Label(self.manage_frame, text="Email", font=("bold", 11), style="Label.TLabel")
        self.email_label.pack()
        self.email_entry = ttk.Entry(self.manage_frame, textvariable=self.email_var, font=("bold", 9), style="Entry.TEntry")
        self.email_entry.pack()

        self.gender_label = ttk.Label(self.manage_frame, text="Gender", font=("bold", 11), style="Label.TLabel")
        self.gender_label.pack()
        self.gender_combo = ttk.Combobox(self.manage_frame, textvariable=self.gender_var, state='readonly', font=("bold", 9), style="Combobox.TCombobox")
        self.gender_combo['values'] = ('Male', 'Female')
        self.gender_combo.pack()

        self.contact_label = ttk.Label(self.manage_frame, text="Contact", font=("bold", 11), style="Label.TLabel")
        self.contact_label.pack()
        self.contact_entry = ttk.Entry(self.manage_frame, textvariable=self.contact_var, font=("bold", 9), style="Entry.TEntry")
        self.contact_entry.pack()

        

        self.button_frame = ttk.Frame(self.manage_frame, relief="flat",borderwidth=0, style="Frame.TFrame")
        self.button_frame.pack(pady=13)

        self.add_button = ttk.Button(self.button_frame, text="Add", style="Accent.TButton", command=self.add_student)
        self.add_button.grid(row=0, column=0, padx=5, pady=10)
        self.update_button = ttk.Button(self.button_frame, text="Update", style="Accent.TButton", command=self.update_student)
        self.update_button.grid(row=0, column=1, padx=5, pady=10)
        self.delete_button = ttk.Button(self.button_frame, text="Delete", style="Accent.TButton", command=self.delete_student)
        self.delete_button.grid(row=0, column=2, padx=5, pady=10)
        self.clear_button = ttk.Button(self.button_frame, text="Clear", style="Accent.TButton", command=self.clear_entries)
        self.clear_button.grid(row=0, column=3, padx=5, pady=10)


        self.details_frame = ttk.Frame(root, borderwidth=0, relief="flat", style="Frame.TFrame")
        self.details_frame.pack(side=tk.TOP, padx=20, pady=20)

        self.data_frame = ttk.Frame(root, borderwidth=0, relief="flat", style="Frame.TFrame")
        self.data_frame.pack(side=tk.TOP, padx=20, pady=5)

        # Add buttons for data export and import to the data frame
        self.export_button = ttk.Button(self.data_frame, text="Export Data", command=self.export_data,style="Accent.TButton",width=20)
        self.export_button.grid(row=0, column=1, padx=10, pady=5)

        self.import_button = ttk.Button(self.data_frame, text="Import Data", command=self.import_data,style="Accent.TButton",width=20)
        self.import_button.grid(row=0, column=0, padx=10, pady=5)

        self.clear_button = ttk.Button(self.data_frame, text="Clear Data", command=self.delete_table,style="Accent.TButton",width=20)
        self.clear_button.grid(row=0, column=2, padx=10, pady=5)

        self.search_frame = ttk.Frame(self.details_frame, borderwidth=0, relief="flat", style="Frame.TFrame")
        self.search_frame.pack(side=tk.TOP, pady=20)

        self.search_label = ttk.Label(self.search_frame, text="Search By", font=("bold", 14), style="Label.TLabel")
        self.search_label.grid(row=0, column=0, padx=5)
        self.search_combo = ttk.Combobox(self.search_frame, textvariable=self.search_by, state='readonly', style="Combobox.TCombobox")
        self.search_combo['values'] = ('Student_Id', 'Name', 'Grade_Section')
        self.search_combo.grid(row=0, column=1, padx=5)
        self.search_entry = ttk.Entry(self.search_frame, textvariable=self.search_txt, style="Entry.TEntry")
        self.search_entry.grid(row=0, column=2, padx=5)
        self.search_button = ttk.Button(self.search_frame, text="Search", style="Accent.TButton", command=self.search_students)
        self.search_button.grid(row=0, column=3, padx=5)
        self.left_button = ttk.Button(self.search_frame, text="<", command=self.move_left,style="Accent.TButton",width=1)
        self.left_button.grid(row=0, column=5, padx=5)
        self.right_button = ttk.Button(self.search_frame, text=">", command=self.move_right,style="Accent.TButton",width=1)
        self.right_button.grid(row=0, column=6, padx=5)
        self.show_all_button = ttk.Button(self.search_frame, text="Show All", style="Accent.TButton", command=self.display_students)
        self.show_all_button.grid(row=0, column=4, padx=5)

        self.students_tree = ttk.Treeview(self.details_frame, columns=(
            'Student_Id', 'Name', 'Grade_Section', 'Email', 'Gender', 'Contact'))

        self.students_tree.heading('Student_Id', text='Student ID', anchor='w')
        self.students_tree.heading('Name', text='Name', anchor='w')
        self.students_tree.heading('Grade_Section', text='Grade and Section', anchor='w')
        self.students_tree.heading('Email', text='Email', anchor='w')
        self.students_tree.heading('Gender', text='Gender', anchor='w')
        self.students_tree.heading('Contact', text='Contact', anchor='w')
        


        self.students_tree['show'] = 'headings'
        self.students_tree.column('Student_Id', width=100)
        self.students_tree.column('Name', width=150)
        self.students_tree.column('Grade_Section', width=120)
        self.students_tree.column('Email', width=200)
        self.students_tree.column('Gender', width=100)
        self.students_tree.column('Contact', width=120)
        

        self.students_tree.pack(fill=tk.BOTH, expand=1)
        self.students_tree.bind('<ButtonRelease-1>', self.get_selected_row)

        self.display_students()

    
#helps to move the Treeviw to the Left 
    def move_left(self):
        self.students_tree.xview_scroll(-30, "units")
#helps to move the Treeview to the Right 
    def move_right(self):
        self.students_tree.xview_scroll(30, "units")
#Helps to export the data from the System Manager 
    def export_data(self):
        search_by = self.search_by.get()
        search_text = self.search_txt.get()

       

        try:
            client = self.connect_to_database()
            students_collection = client["Student-cluster"]["students"]
            if search_by != "" and search_text != "":
                students=students_collection.find({search_by : search_text},{"_id":0})
                count=students_collection.count_documents({search_by : search_text})
                
            else:
                students=students_collection.find({},{"_id":0})
                count=students_collection.count_documents({})

            if count == 0:
                messagebox.showwarning("No Records Found", "No records matching the search criteria.")
                return
            file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
            if not file_path:
                return

            with open(file_path, mode='w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(["Student_Id", "Name", "Grade_Section", "Email", "Gender", "Contact"])
                for student in students:
                    writer.writerow([
                    student.get("Student_Id", ""),  # Use .get() to avoid KeyError
                    student.get("Name", ""),
                    student.get("Grade_Section", ""),
                    student.get("Email", ""),
                    student.get("Gender", ""),
                    student.get("Contact", "")
                    ])

            messagebox.showinfo("Export Successful", "Data exported to CSV file successfully.")
            self.students_tree.delete(*self.students_tree.get_children())
            if search_by != "" and search_text != "":
                students_collection.delete_many({search_by : search_text})
            else :
                students_collection.delete_many({})
                
            


        finally:
            # Close the database clientection
            client.close()


#Helps to import data from a CSV file -- ( Note :-- It will accept only the data file it has exported ..)        
    def import_data(self):
    # Open a file dialog to choose the import file
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return

        try:
            # Clear existing student data
            self.students_tree.delete(*self.students_tree.get_children())
            self.students.clear()

            # Read the CSV file
            with open(file_path, mode='r') as file:
                reader = csv.reader(file)
                headings = next(reader)  # Get the column headings

                # Check if the file has the expected columns
                if headings != ["Student_Id", "Name", "Grade_Section", "Email", "Gender", "Contact"]:
                    messagebox.showerror("Invalid File", "The selected file does not have the expected columns.")
                    return

                # clientect to the MongoDB database
                client = self.connect_to_database()
                students_collection = client["Student-cluster"]["students"]

                
                

                # Process each row of data
                for row in reader:
                    student_data = dict(zip(headings, row))
                    self.students.append(student_data)
                    self.students_tree.insert('', tk.END, values=list(student_data.values()))

                    # Insert the data into the "students" table
                    students_collection.insert_one({'Student_Id':
                        student_data['Student_Id'],'Name':
                        student_data['Name'],'Grade_Section':
                        student_data['Grade_Section'],'Email':
                        student_data['Email'],'Gender':
                        student_data['Gender'],'Contact':
                        student_data['Contact']
                        })
                    
                    

                # Commit the changes and close the database clientection
                
                client.close()

                messagebox.showinfo("Import Successful", "Data imported from CSV file successfully.")
        except Exception as e:
                messagebox.showerror("Import Error", f"An error occurred while importing data: {str(e)}")


#This function is used to Add the student to the Treeview and Database 
    def add_student(self):
        if self.Student_Id_var.get() == '' or self.name_var.get() == '' or self.grade_section_var.get() == '':
            messagebox.showerror("Error", "Please fill in all required fields")
        else:
            try:
                client = self.connect_to_database()
                students_collection = client["Student-cluster"]['students']
               
                students_collection.insert_one({'Student_Id':self.Student_Id_var.get(),'Name':self.name_var.get(), 'Grade_Section':self.grade_section_var.get(), 'Email':self.email_var.get(), 'Gender':self.gender_var.get(), 'Contact':self.contact_var.get()})
                
                
                client.close()
                self.clear_entries()
                self.display_students()
                messagebox.showinfo("Success", "Student added successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Error while clientecting to the database: {e}")
                
#This function is used to update the student's Record 
    def update_student(self):
        if self.Student_Id_var.get() == '' or self.name_var.get() == '' or self.grade_section_var.get() == '':
            messagebox.showerror("Error", "Please select a student")
        else:
            
            try:
                client = self.connect_to_database()
                students_collection = client["Student-cluster"]["students"]
                students_collection.update_one({'Student_Id':self.Student_Id_var.get()},{ '$set' :{'Name':self.name_var.get(), 'Grade_Section':self.grade_section_var.get(), 'Email':self.email_var.get(), 'Gender':self.gender_var.get(), 'Contact':self.contact_var.get()}})
                
                
                
                self.clear_entries()
                self.display_students()
                client.close()
                messagebox.showinfo("Success", "Student updated successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Error while clientecting to the database: {e}")
            

    
#This helps to delete the student's record from the database 
    def delete_student(self):
        if self.Student_Id_var.get() == '':
            messagebox.showerror("Error", "Please select a student")
        else:
            confirmation = messagebox.askyesno("Confirmation", "Are you sure you want to delete this student?")
            if confirmation == 1:
                try:
                    client = self.connect_to_database()
                    students_collection = client["Student-cluster"]['students']
                    students_collection.delete_one({"Student_Id":self.Student_Id_var.get()})
                
                    self.clear_entries()
                    self.display_students()
                    messagebox.showinfo("Success", "Student deleted successfully")
                    client.close()
                except Exception as e:
                    messagebox.showerror("Error", f"Error while clientecting to the database: {e}")
        
    #This function helps to search the student in the database 
    def search_students(self):
        try:
            client = self.connect_to_database()
            students_collection = client["Student-cluster"]['students']
            
                
            students = {}
            if self.search_by.get() == 'Student_Id':
                students =students_collection.find({"Student_Id": self.search_txt.get()})
            elif self.search_by.get() == 'Name':
                students =students_collection.find({"Name": self.search_txt.get()})
            elif self.search_by.get() == 'Grade_Section':
                students =students_collection.find({"Grade_Section": self.search_txt.get()})
            
            
            self.students_tree.delete(*self.students_tree.get_children())
            for student in students:
                self.students_tree.insert('', tk.END, values=(student['Student_Id'],student['Name'],student['Grade_Section'],student['Email'],student['Gender'],student['Contact']))
            client.close()
        except Exception as e:
            messagebox.showerror("Error", f"Error while clientecting to the database: {e}")

    
#This is used to display the student records on the Treeview 
    def display_students(self):
        
        try:
            client = self.connect_to_database()
            students_collection = client["Student-cluster"]['students']
            students =students_collection.find()
            self.students_tree.delete(*self.students_tree.get_children())
            for student in students:
                self.students_tree.insert('', tk.END, values=(student['Student_Id'],student['Name'],student['Grade_Section'],student['Email'],student['Gender'],student['Contact']))
            client.close()  
        except errors.ConnectionFailure as e:
            messagebox.showerror("Error", f"Error while connecting to the database: {e}")
         
#This helps to delete the Table from the Database 
    def delete_table(self):
        
        try:
            client = self.connect_to_database()
            client["Student-cluster"]['students'].drop()
            messagebox.showinfo("Table Deleted", f"The table Students has been deleted successfully.")
            self.students_tree.delete(*self.students_tree.get_children())
            client.close()
            

        except Exception as e:
            print(f"Error deleting table: {e}")
        
#This selects the row you select on the Treeview 
    def get_selected_row(self, event):
        try:
            selected_row = self.students_tree.focus()
            values = self.students_tree.item(selected_row, 'values')
            self.Student_Id_var.set(values[0])
            self.name_var.set(values[1])
            self.grade_section_var.set(values[2])
            self.email_var.set(values[3])
            self.gender_var.set(values[4])
            self.contact_var.set(values[5])
            
        except IndexError:
            pass
#This Clears all the entries on the fields 
    def clear_entries(self):
        self.Student_Id_var.set('')
        self.name_var.set('')
        self.grade_section_var.set('')
        self.email_var.set('')
        self.gender_var.set('')
        self.contact_var.set('')
        

#This Helps to clientect the App to the "MongoDB Server" in your PC/Laptop 
    
    def connect_to_database(self):
        
        try:
            connection_string =  "mongodb+srv://ouss12fr:ouss2002@student-cluster.gh1qe.mongodb.net/Student-cluster"
            client= MongoClient(connection_string)
            

            return client

        except errors.ConnectionFailure as e:
            print("Error connecting to MongoDb:", e)
            messagebox.showerror("Error", "Failed to connect to MongoDb")
            return None
        
#A function declared to open the app 
def open_student_system(): 
    root = tk.Tk()
    sms = StudentManagementSystem(root)
    root.mainloop()
#open_student_system()
