from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
from tkinter.filedialog import askopenfile
from score import *
from createpdf import *
import os.path
import time
import pandas as pd

root = Tk()
w = 550 # width for the Tk root
h = 650 # height for the Tk root
wh = '{w}x{h}'

ws = root.winfo_screenwidth() # width of the screen
hs = root.winfo_screenheight() # height of the screen

# calculate x and y coordinates for the Tk root window
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)

root.geometry('%dx%d+%d+%d' % (w, h, x, y/2))
root.resizable(width=True, height=True)
root.title('PyRChicken')
root.iconbitmap('reportcard/icon.ico')
root.grid_columnconfigure([0], weight=1)
root.grid_rowconfigure(3, weight=1)

#variables
class_var = StringVar()
semester_var = StringVar()
year_var = StringVar()
subjdir_var = StringVar()
ekskuldir_var = StringVar()
outputdir_var = StringVar()
absenfile_var = StringVar()
descfile_var = StringVar()

#teacher
class teacher():
    def __init__ (self):
        self.teacher_name = StringVar()
        self.teacher_nik = StringVar()
        
    def popup(self):
        global top
        button_change.configure(state = DISABLED)
        top = Toplevel(root)
        w = 300 # width for the Tk root
        h = 200 # height for the Tk root
        ws = root.winfo_screenwidth() # width of the screen
        hs = root.winfo_screenheight() # height of the screen
        # calculate x and y coordinates for the Tk root window
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)

        top.geometry('%dx%d+%d+%d' % (w, h, x, y/2))
        top.resizable(width=False, height=False)
        top.title('Teacher\'s Information')
        top.iconbitmap('reportcard/icon.ico')
        def disable_event():
            pass
        top.protocol("WM_DELETE_WINDOW", disable_event)
        label_name = Label(top, text = 'Teacher\'s Name:')
        label_name.pack(side = TOP, pady = 2, padx = 10, fill = X)
        entry_name = Entry(top, textvariable = self.teacher_name)
        entry_name.pack(side = TOP, pady = 2, padx = 30, fill = X)
        label_nik = Label(top, text = 'Teacher\'s NIK:')
        label_nik.pack(side = TOP, pady = 2, padx = 10, fill = X)
        entry_nik = Entry(top, textvariable = self.teacher_nik)
        entry_nik.pack(side = TOP, pady = 2, padx = 30, fill = X)
        btn = Button (top, command = self.confirm, state = NORMAL, text = 'Confirm')
        btn.pack(side = BOTTOM, pady = 10, ipadx = 50)
        self.label_err = Label(top, text = '')
        self.label_err.pack(side = TOP, pady = 2, padx = 10)
        
    def confirm(self):
        button_change.configure(state = NORMAL)
        self.name = self.teacher_name.get()
        self.teacher_name.set(self.name)
        self.nik = self.teacher_nik.get()
        self.teacher_nik.set(self.nik)
        if self.name != '' and self.nik != '':
            label_welcomename.configure(text = f'Welcome, {self.name}')
            label_welcomeNIK.configure(text = f'NIK. {self.nik}')
            print(f'''Welcome, {self.name}
NIK:{self.nik}
========================
*You can either use your own files or the dummy files provided in 'Dummy files'.
 It is suggested to read the 'Instruction.pdf' file.
 
 The 'Create PDF' is available once everything is correctly filled.

-) Details
    • Class     : Fill with either 10.1, 10.2, 11.1, 11.2, 12.1, or 12.2
    • Semester  : Fill with either 1 or 2
    • Year      : Format is "(academic year start)-(academic year end)"
                  (Example: "2022-2023")
    
-) Files
    • Absen     : Template is available in 'Template' folder
    • Subject   : Template is available in 'Template' folder
    
-) Output       : Any folder is allowed, no specific criteria.

-) Score
    • Subjects  : Naming format is "(subject)_(class)_semester (semester).xlsx"
    • Ekskul    : Naming format is "(ekskul)_(academic start year)-(academic end year)"
                  ''')
            top.destroy()
        else:
            self.label_err.configure(text = 'Please fill out the fields', fg = 'Red')

groupA = ['Agama', 'Bahasa Indonesia', 'Bahasa Inggris', 'Matematika (Wajib)', 'PKN', 'Sejarah Indonesia'] #Kelompok A (Umum)
groupB = ['Seni Budaya', 'PJOK', 'PKWU'] #Kelompok B (Umum)
mutlok = ['ICT'] #Muatan Lokal
groupC_IPA = ['Matematika (Peminatan)', 'Biologi', 'Kimia', 'Fisika'] #Kelompok C (Peminatan) - IPA
groupC_IPS = ['Ekonomi', 'Sosiologi', 'Geografi', 'Sejarah Internasional'] #Kelompok C (Peminatan) - IPS
subj = groupA + groupB + mutlok + groupC_IPA + groupC_IPS + ['Literature']
ekskul_reg = ['Archery', 'Badminton', 'Band', 'Basketball', 'Content Creator', 'Cooking', 'Dance', 'Graphic Design', 'Illustration', 'Knitting', 'Mini Soccer', 'Public Speaking', 'Robotic']
ekskul_prem = ['Bowling', 'Coding', 'Swimming', 'Tae Kwon Do']
ekskul = ekskul_reg + ekskul_prem

#command   
class file_func():
    def upload_file(self, label, var):
        self.file = filedialog.askopenfilename(
            title = 'Select File',
            initialdir = '/Subjects',
            filetypes = [("Excel files", ".xlsx .xls")]
        )
        label.configure(text = self.file, fg = 'blue')
        var.set(self.file)
        
    def ask_directory(self, label, var):
        self.dir = filedialog.askdirectory(
            title = 'Select Folder'
        )
        label.configure(text = self.dir, fg = 'blue')
        var.set(self.dir)
        
class available():
    def __init__(self, master, files, var):
        self.master = master
        self.files = files
        self.var = var
        self.dir = ''
    
    def set(self):
        kelas_list = ['10.1', '10.2', '11.1', '11.2', '12.1', '12.2']
        kelas = class_var.get()
        class_var.set(kelas)
        semester = semester_var.get()
        semester_var.set(semester)
        year = year_var.get()
        year_var.set(year)
        
        if kelas == '':
            label_error_class.configure(text = '*Error: Please enter a class', fg = 'Red')
        elif kelas not in kelas_list:
            label_error_class.configure(text = f'*Error: The class \'{kelas}\' doesn\'t exist', fg = 'Red')
        elif kelas in kelas_list:
            label_error_class.configure(text = f'Class is set to \'{kelas}\'', fg = 'Blue')
            
        if semester == '':
            label_error_semester.configure(text = '*Error: Please enter a semester', fg = 'Red')
        elif semester not in ['1', '2']:
            label_error_semester.configure(text = f'*Error: The semester \'{semester}\' doesn\'t exist', fg = 'Red')
        elif semester in ['1', '2']:
            label_error_semester.configure(text = f'Semester is set to \'{semester}\'', fg = 'Blue')
        
        if year == '':
            label_error_year.configure(text = '*Error: Please enter a year', fg = 'Red')
        else:
            label_error_year.configure(text = f'Year is set to \'{year}\'', fg = 'Blue')
        
        if semester in ['1', '2'] and kelas in kelas_list and year != '':
            if self.var == subjdir_var:   
                    self.files = ([s + f'_{kelas}_semester {semester}.xlsx' for s in subj])
            elif self.var == ekskuldir_var:
                self.files = ([s + f'_{year}.xlsx' for s in ekskul])
            self.btn_folder.configure(state = NORMAL)
            self.check_available(self.files)
              
    def create_frame(self):
        self.details = Label(master = self.master, font = ('Calibri', 8), wraplength = 250, fg = 'red', text = 'Please set directory!')
        self.details.pack(fill = X)
        
        self.avail = Frame(master = self.master, borderwidth = 2, relief = GROOVE, bg = 'white')
        self.avail.pack(side = TOP, fill = 'both', expand = True)
        self.label_avail = Label(self.avail, text = 'Available Files (✔)', font = ('Calibri', 8), fg = 'white', bg = 'green').pack(fill = X, anchor = NW)
        self.unavail = Frame(master = self.master, borderwidth = 2, relief = GROOVE, bg = 'white')
        self.unavail.pack(side = TOP, fill = 'both', expand = True)
        self.label_unavail = Label(self.unavail, text = 'Unavailable Files (✘)', font = ('Calibri', 8), fg = 'white', bg = 'red').pack(fill = X, anchor = NW)
        
        self.v_avail = Scrollbar(self.avail, orient = 'vertical')
        self.v_avail.pack(side = RIGHT, fill = Y)
        self.v_unavail = Scrollbar(self.unavail, orient = 'vertical')
        self.v_unavail.pack(side = RIGHT, fill = Y)
        self.t_avail = Text(master = self.avail, width = 40, height = 1, wrap = NONE, font = ('Calibri', 8), yscrollcommand = self.v_avail.set, state = DISABLED)
        self.t_unavail = Text(master = self.unavail, width = 40, height = 1, wrap = NONE, font = ('Calibri', 8), yscrollcommand = self.v_unavail.set, state = DISABLED)
        self.t_avail.pack(anchor = NW, fill = 'both', expand = True) 
        self.t_unavail.pack(anchor = NW, fill = 'both', expand = True)
        
        self.btn_section = Frame(master = self.master)
        self.btn_section.pack(fill = X)
        self.btn_folder = Button(self.btn_section, text = 'Choose Folder', font = ('Calibri', 8), command = lambda: [self.check_directory(self.details), pdf_trigger.change_state()], state = DISABLED)
        self.btn_folder.pack(expand = True, anchor = SE)
        
    def check_directory(self, label):
        self.dir = filedialog.askdirectory(
            title = 'Select Folder'
            ) #ask directory
        #self.var.set(self.dir)
        self.check_available(self.files)
        label.configure(text = f'Checking files in \'{self.dir}\'', fg = 'blue')#configure a label to show 'selected path'
    
    def check_available(self, files):   
        self.t_avail.configure(state = NORMAL)
        self.t_avail.delete(1.0, END)
        self.t_unavail.configure(state = NORMAL)
        self.t_unavail.delete(1.0, END)
        
        self.a_files = files
        self.a_exist = [f for f in self.a_files if os.path.isfile(f'{self.dir}/{str(f)}')]
        self.a_non_exist = [f for f in self.a_files if not os.path.isfile(f'{self.dir}/{str(f)}')]

        for f in self.a_exist:
           self.t_avail.insert (END, f'• {f}\n')
        
        for f in self.a_non_exist:
           self.t_unavail.insert (END, f'• {f}\n')
           
        if self.a_exist == self.files:
            self.var.set(self.dir)

        self.v_avail.configure(command = self.t_avail.yview)
        self.v_unavail.configure(command = self.t_unavail.yview)
        self.t_avail.configure(state = DISABLED)
        self.t_unavail.configure(state = DISABLED)
        
class pdf_trigger():
        
    def change_state():
        if class_var.get() != '' and semester_var.get() != '' and year_var.get() != '' and subjdir_var.get() != '' and ekskuldir_var.get() != '' and outputdir_var.get() != '' and absenfile_var.get() != '' and descfile_var.get() != '':
            pdf_btn.configure(state = NORMAL)
        else:
            pdf_btn.configure(state = DISABLED)
        
    def generate_pdf():
        messagebox.showinfo("Generating pdf", "The Application is going to not respond, this is normal.\nPlease confirm to continue the process.")
        print('Printing in progress... please wait until another message box shows up')
        st = time.time()
        collect_data(absenfile_var.get(), class_var.get(), semester_var.get(), subjdir_var.get())
        WaliKelas = mainteacher.teacher_name.get()
        Nik_WaliKelas = mainteacher.teacher_nik.get()
        Kelas = class_var.get()
        Semester = semester_var.get()
        Thn = year_var.get()
        Description = descfile_var.get()
        Ekskul = ekskuldir_var.get()
        absen_file = absenfile_var.get()
        output_dir = outputdir_var.get()
        createpdf(WaliKelas, Nik_WaliKelas, Kelas, Semester, Thn, Description, Ekskul, absen_file, output_dir)
        et = time.time()
        df_absen = pd.read_excel(f'{absen_file}', header = 1, sheet_name = 'Profile')
        print(f'Thank you for using PyRChicken\nby Christian Raphael Heryanto\nAverage time per report card: {round((et - st)/len(df_absen.index))} s\nDon\'t forget to Screenshot and fill the Questionnaire!')
        messagebox.showinfo('Generating pdf', f'All Done!\nPlease SCREENSHOT the whole screen, this will act as evidence that you have tried the application\nAverage time per report card: {round((et - st)/len(df_absen.index))} s\n Thank you, {WaliKelas}!')
        
#frames
welcomeframe = Frame(root, width = 100, height = 100)
welcomeframe.grid(row = 0, column = 0, padx = 5, pady = 2.5, ipadx = 5, ipady = 2, sticky = NSEW)

entryframe = LabelFrame(root, width = 100, height = 100, text = 'Details')
entryframe.grid(row = 1, column = 0, padx = 5, pady = 2.5, ipadx = 5, ipady = 2, sticky = NSEW)
entryframe.grid_columnconfigure((0,1), weight=1)
entryframe.grid_rowconfigure(5, weight=1)

filesframe = LabelFrame(root, width = 200, height = 150, text = 'Files')
filesframe.grid(row = 2, column = 0, padx = 5, pady = 2.5, ipadx = 5, ipady = 2, sticky = NSEW)
filesframe.grid_columnconfigure(1, weight=1)

outputframe = LabelFrame(root, width = 200, height = 150, text = 'Output')
outputframe.grid(row = 3, column = 0, padx = 5, pady = 2.5, ipadx = 5, ipady = 2, sticky = NSEW)
outputframe.grid_columnconfigure(0, weight=1)
outputframe.grid_rowconfigure(1, weight=1)

final_frame = Frame(root)
final_frame.grid(row = 4, column = 0, columnspan = 2, padx = 5, pady = 5, sticky = NSEW)
#pdf button
pdf_btn = Button(final_frame, text = 'Create PDF', state = DISABLED, command = pdf_trigger.generate_pdf)

#welcome frame
mainteacher = teacher()
label_welcomename = Label(welcomeframe, text = 'Hi Teachers,')
label_welcomename.grid(row = 0, column = 0, padx = 5, pady = 2, sticky = W)
label_welcomeNIK = Label(welcomeframe, text = f'Welcome to my software application!')
label_welcomeNIK.grid(row = 1, column = 0, padx = 5, pady = 2, sticky = W)
button_change = Button(welcomeframe, text = 'Change', state = NORMAL, font = ('Calibri', 8), command = lambda: mainteacher.popup())
button_change.grid(row = 2, column = 0, ipadx = 10, padx = 5, pady = 2, sticky = SW)
mainteacher.popup()

#set details
label_class = Label(entryframe, text = 'Class').grid(row = 0, column = 0, padx = 2, pady = 2, sticky = W)
entry_class = Entry(entryframe, textvariable = class_var).grid(row = 0, column = 1, padx = 5, pady = 2, sticky = E)
label_semester = Label(entryframe, text = 'Semester').grid(row = 1, column = 0, padx = 2, pady = 2, sticky = W)
entry_semester = Entry(entryframe, textvariable = semester_var).grid(row = 1, column = 1, padx = 5, pady = 2, sticky = E)
label_year = Label(entryframe, text = 'Year').grid(row = 2, column = 0, padx = 2, pady = 2, sticky = W)
entry_year = Entry(entryframe, textvariable = year_var).grid(row = 2, column = 1, padx = 5, pady = 2, sticky = E)
label_error_class = Label(entryframe, text = '*Class is required', fg = 'red', font = ('Calibri', 8), wraplength = 150, justify = LEFT)
label_error_class.grid(row = 3, column = 0, columnspan = 2, padx = 2, pady = 2, sticky = W)
label_error_semester = Label(entryframe, text = '*Semester is required', fg = 'red', font = ('Calibri', 8), wraplength = 150, justify = LEFT)
label_error_semester.grid(row = 4, column = 0, columnspan = 2, padx = 2, pady = 2, sticky = W)
label_error_year = Label(entryframe, text = '*Year is required', fg = 'red', font = ('Calibri', 8), wraplength = 150, justify = LEFT)
label_error_year.grid(row = 5, column = 0, columnspan = 2, padx = 2, pady = 2, sticky = W)

set_btn = Button(entryframe, text = 'Set', font = ('Calibri', 8), command = lambda:[subjavail.set(), ekskulavail.set(), pdf_trigger.change_state()])
set_btn.grid(row = 6, column = 1, ipadx = 10, padx = 5, pady = 5, sticky = SE)



#absen
absen_label = Label(filesframe, text = 'Absen').grid(row = 0, column = 0, padx = 2, pady = 2, sticky = W)
absen_btn = Button(filesframe, text = 'Choose File', font = ('Calibri', 8), command = lambda: [absen_func.upload_file(absen_path, absenfile_var), pdf_trigger.change_state()])
absen_btn.grid(row = 0, column = 1, padx = 5, pady = 2, sticky = E)
absen_func = file_func()
absen_path = Label(filesframe, text = 'File is not found!', fg = 'red', wraplength = 190, font = ('Calibri', 8), justify = CENTER)
absen_path.grid(row = 1, column = 0, padx = 2, pady = 2, columnspan = 2, sticky = EW)


#description
desc_label = Label(filesframe, text = 'Description').grid(row = 2, column = 0, padx = 2, pady = 2, sticky = W)
desc_btn = Button(filesframe, text = 'Choose File', font = ('Calibri', 8), command = lambda: [desc_func.upload_file(desc_path, descfile_var), pdf_trigger.change_state()])
desc_btn.grid(row = 2, column = 1, padx = 5, pady = 2, sticky = E)
desc_func = file_func()
desc_path = Label(filesframe, text = 'File is not found!', fg = 'red', wraplength = 190, font = ('Calibri', 8), justify = LEFT)
desc_path.grid(row = 3, column = 0, padx = 2, pady = 2, columnspan = 2, sticky = EW)

#output
output_label = Label(outputframe, text = 'Select output directory!').grid(row = 0, column = 0, padx = 2, pady = 2, sticky = W)
output = file_func()
output_path = Label(outputframe, text = 'Folder is not found!', fg = 'red', wraplength = 190, font = ('Calibri', 8), justify = LEFT)
output_path.grid(row = 1, column = 0, padx = 2, pady = 2, columnspan = 2, sticky = EW)
output_button = Button(outputframe, text = 'Choose Folder', font = ('Calibri', 8), command = lambda: [output.ask_directory(output_path, outputdir_var), pdf_trigger.change_state()])
output_button.grid(row = 2, column = 0, padx = 5, pady = 2, sticky = SE)

#set the required files tab
tabrequired = ttk.Notebook(root)

           
#subject window 
subject_win = Frame(tabrequired, borderwidth = 1, relief = GROOVE, background = 'white')
subjavail = available(subject_win, '', subjdir_var)
subjavail.create_frame()

#ekskul window
ekskul_win = Frame(tabrequired, width = 100, height = 200, borderwidth = 1, relief = GROOVE, background = 'white')
ekskulavail = available(ekskul_win, ekskul, ekskuldir_var)
ekskulavail.create_frame()

tabrequired.add(subject_win, text='Subjects')
tabrequired.add(ekskul_win, text='Ekskul')
tabrequired.grid(row = 0, column = 1, rowspan = 4, padx = 5, pady = 5, sticky = NSEW)


separator = ttk.Separator(final_frame, orient='horizontal')
separator.pack(fill='x')
final_label = Label(final_frame, text = 'Please fill the required details!', fg = 'Red')
final_label.pack(fill = X, expand = True)
separator = ttk.Separator(final_frame, orient='horizontal')
separator.pack(fill='x')

pdf_btn.pack(fill = X, expand = True, ipadx = 100, pady = 5)

root.mainloop()