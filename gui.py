from tkinter import *
from tkinter.ttk import Frame, Button, Style, Progressbar
from tkinter import filedialog
from tkhtmlview import HTMLLabel
from os.path import basename

class GUI(Frame):

    def __init__(self, root):
        super().__init__()
        self.controller = None
        self.default_title = "Bulk Email Sender"
        self.current_selection = None
        self.details_window = None
        self.job_completion_window = None
        self.root = root

    def load(self):
        try:
            filename = filedialog.askopenfilename(initialdir = "",
                                                title = "Select a File",
                                                filetypes = (("Yaml files",
                                                                "*.yml*"),
                                                            ("all files",
                                                                "*.*")))
        except:
            return None
        self.controller.load_yaml(filename)

    def display_details_window(self):
        self.details_window = DetailsWindow(self)

    def display_job_completion_widow(self):
        self.job_completion_window = JobCompletionWindow(self)

    def update_html_rederer(self):
        if(self.current_selection != None):
            current_email = self.controller.generated_email_list[self.current_selection]
            to_list = ", ".join(current_email.to)
            cc_list = ", ".join(current_email.cc)

            html = "<b>To:</b> " + to_list + "<br>"
            html +="<b>CC:</b> " + cc_list + "<br>"
            html +="<b>Subject:</b> " + current_email.subject+"<br>"
            html +="<b>Attachments:</b> " + ", ".join([basename(attachment) for attachment in current_email.attachments]) +"<br><br>"

            current_email.body_raw = self.html_editor.get("1.0",'end-1c')
            current_email.parse_body(self.controller.worksheet,self.controller.email_config["DATA"])

            self.email_generated_display.set_html("<html>"+html+self.controller.generated_email_list[self.current_selection].body_generated+"</html>")
        else:
            self.email_generated_display.set_html("<html></html>")
        self.root.update_idletasks()
    def listbox_listener(self, event):
        if(self.email_list_box.curselection()):
            self.current_selection = int(self.email_list_box.curselection()[0])
        self.update_html_rederer()
    def html_editor_listener(self, event):
        if(self.email_list_box.curselection()):
            self.current_selection = int(self.email_list_box.curselection()[0])
        self.update_html_rederer()
    def initialise(self):
        self.master.title(self.default_title)
        self.style = Style()
        self.style.theme_use("default")
        #Main Menu
        
        self.mainmenu = Menu(self.master)
        self.mainmenu.add_command(label = "Load", command= self.load)
        self.mainmenu.add_command(label = "Exit", command= self.master.destroy) 
        self.mainmenu.add_command(label = "Generate Report", command = self.controller.logging.output_to_file, state='disabled') 
        self.master.config(menu = self.mainmenu)

        self.frame = Frame(self, relief=RAISED, borderwidth=1)
        self.frame.pack(fill=BOTH, expand=True)

        self.pack(fill=BOTH, expand=True)

        self.email_list_box = Listbox(self.frame)  
        self.email_list_box.bind("<<ListboxSelect>>", self.listbox_listener)
        self.email_list_box.pack(side = LEFT, fill=Y)   

        self.html_editor = HTMLLabel(self.frame,relief=RIDGE, background="white")
        self.html_editor.bind('<KeyRelease>',  self.html_editor_listener)
        self.html_editor.pack(fill=BOTH, expand=True)
        self.html_editor.fit_height()

        self.email_generated_display = HTMLLabel(self.frame, relief=RIDGE, state='disabled')
        self.email_generated_display.pack(fill=BOTH, expand=True)
        self.email_generated_display.fit_height()

        self.send_emails_button = Button(self, text="SEND", command=self.controller.init_sending_emails_authentication)
        self.send_emails_button.pack(side=RIGHT, padx=5, pady=5)

        self.progress_bar = Progressbar(self, mode="determinate", value=0, maximum=100)
        self.progress_bar.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
    def update_list(self, generated_email_list):
        self.email_list_box.delete(0, END) 
        for email in generated_email_list: 
            if(len(email.to) > 0):
                self.email_list_box.insert(END, email.to[0]) 
            else:
                self.email_list_box.insert(END, "NULL") 
        if(len(generated_email_list) > 0):
            self.email_list_box.select_clear(0, END)
            self.email_list_box.select_set(0)
            if(self.email_list_box.curselection()):
                self.current_selection = int(self.email_list_box.curselection()[0])
                self.update_html_rederer()
    def update_progress_bar(self, amount):
        self.progress_bar.step(amount)
        self.root.update_idletasks()
    def reset(self):
        self.current_selection = None
        self.progress_bar['value'] = 0
        self.update_html_rederer()
        self.update_list(self.controller.generated_email_list)
        self.mainmenu.entryconfig(3, state="disabled" )
        
##REFACTOR TO INHERIT FROM Toplevel
class DetailsWindow:
    def __init__(self,parent):
        self.master = parent
        if self.master != "":
            self.window = Toplevel(self.master,width=200,height=100,padx=40,pady=40)

        #window settings
        self.window.title("Login")
        self.window.protocol("WM_DELETE_WINDOW", self.disable_event)
        self.window.resizable(False, False)

        #window widgets
        self.details_canvas = Canvas(self.window, width = 200, height = 100)
        self.details_canvas.pack()

        self.username_input = Entry(self.window)
        self.details_canvas.create_window(100, 30, window=self.username_input)
        self.username_label = Label(self.window, text = "username")
        self.details_canvas.create_window(10, 30, window=self.username_label)

        self.password_input = Entry (self.window,show="*") 
        self.details_canvas.create_window(100, 50, window=self.password_input)
        self.password_label = Label(self.window, text = "password")
        self.details_canvas.create_window(10, 50, window=self.password_label)

        self.cancel_button = Button(self.window,text="Cancel",command=self.window.destroy)
        self.details_canvas.create_window(60, 80, window=self.cancel_button)

        self.login_button = Button(self.window,text="Login",command=self.login_button_onclick)
        self.details_canvas.create_window(135, 80, window=self.login_button)
        
        self.authentication_status_text = StringVar()
        self.authentication_status_label = Label(self.window, textvariable=self.authentication_status_text)
        self.details_canvas.create_window(100, 105, window=self.authentication_status_label)
    def login_button_onclick(self):
        self.master.controller.bulk_send_emails(self.username_input.get(), self.password_input.get())
    
    def reset(self):
        self.authentication_status_text.set("")
    def disable_event(self):
        pass
    def update_authentication_status_label(self, text):
        self.authentication_status_text.set(text)
        self.master.root.update_idletasks()

class JobCompletionWindow:    
    def __init__(self,parent):
        self.master = parent
        if self.master != "":
            self.window = Toplevel(self.master,width=200,height=100,padx=40,pady=40)

        #window settings
        self.window.title("Job Completed")
        self.window.resizable(False, False)

        #window widgets
        self.details_canvas = Canvas(self.window, width = 200, height = 100)
        self.details_canvas.pack()
        
        self.job_completion_status = StringVar()
        self.job_completion_status_label = Label(self.window, textvariable=self.job_completion_status)
        self.details_canvas.create_window(100, 10, window=self.job_completion_status_label)

        self.cancel_button = Button(self.window,text="Close",command=self.window.destroy)
        self.details_canvas.create_window(60, 80, window=self.cancel_button)

        self.login_button = Button(self.window,text="Generate Report",command=self.master.controller.logging.output_to_file)
        self.details_canvas.create_window(150, 80, window=self.login_button)
    
    def update_job_completion_status(self):
        text = "Emails Sent Success: " + str(len(self.master.controller.logging.emails_success)) + "\n"
        text += "Emails Sent Failure: " + str(len(self.master.controller.logging.emails_failed))
        self.job_completion_status.set(text)
        self.master.root.update_idletasks()

def main():

    root = Tk()
    root.geometry("960x720+150+150")
    gui = GUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
