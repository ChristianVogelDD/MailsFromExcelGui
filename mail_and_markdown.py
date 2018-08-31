#pyinstaller to exe wont handle second level imports
#need to add them manually
####
import pandas._libs.skiplist
import pandas._libs.tslibs.nattype
import pandas._libs.tslibs.np_datetime
####
import tkinter as tk
from tkinter.filedialog import askopenfilename
import pandas as pd
import yagmail as ym
import markdown as md

class MarkdownEditor(tk.Frame):
    """Class with Textfield and Buttons to realize simple Markdown Editor"""

    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.markdown_button_texts = ["Bold", "Heading",
                "Picture", "Link"] 
        self.markdown_buttons = []
        self.column_buttons = []
        for i in range(len(self.markdown_button_texts)):
            self.markdown_buttons.append(tk.Button(self,
                text = self.markdown_button_texts[i],
                command = lambda idx = i: self.markdown_button_pressed(idx)))
            self.markdown_buttons[i].grid(row = 0, column = i,
                    sticky = "we")
        self.mail_text = tk.Text(self, width = 80,font=("Arial",20))
        self.subject_entry = tk.Entry(self)
        self.subject_entry.grid(row = 2, column = 0, columnspan = 4,
                sticky = "we")
        self.mail_text.grid(row = 3, column = 0, columnspan = 10)

    def get_mail_text(self):
        return(self.mail_text.get("1.0", "end-1c"))
    
    def get_subject_entry(self):
        return self.subject_entry.get()  

    def set_dataframe(self, data_frame):
        self.data_frame = data_frame

    def remove_old_buttons(self):
        for button in self.column_buttons:
            button.grid_remove()
        del self.column_buttons[:]

    def update_buttons(self, column_names):
        self.remove_old_buttons()
        self.column_names = column_names
        for i in range(len(column_names)):
            self.column_buttons.append(tk.Button(self, text = column_names[i],
                command = lambda idx = i: self.column_button_pressed(idx)))
            self.column_buttons[i].grid(row = 1, column = i, sticky = "we") 
    
    def markdown_button_pressed(self, idx):
        button_pressed = self.markdown_buttons[idx].cget("text")
        if button_pressed == "Bold":
            self.mail_text.insert(tk.END, "****")
            pos =self.mail_text.index("end-3c")
            self.mail_text.mark_set("insert", pos) 
        elif button_pressed == "Heading":
            self.mail_text.insert("insert linestart", "#")
        elif button_pressed == "Picture":
            self.mail_text.insert(tk.INSERT, " ![ImageName](link to image) ")
        elif button_pressed == "Link":
            self.mail_text.insert(tk.INSERT, " [LinkName](link to website) ")

    def column_button_pressed(self, idx):
        button_pressed = self.column_buttons[idx].cget("text")
        self.mail_text.insert(tk.INSERT, " _"+button_pressed+"_ ")

    def put_in_attributes(self, row):
        self.replaced_markdown_text = self.get_mail_text()
        for column_name in self.column_names:
            search = "_"+column_name+"_"
            self.replaced_markdown_text =self.replaced_markdown_text.replace(
                    "_"+column_name+"_",self.data_frame.loc[row][column_name])

    def generate_html(self, row):
        self.get_mail_text()
        self.put_in_attributes(row)
        html = md.markdown(self.replaced_markdown_text)
        return html

class FarantoGUI(tk.Frame):
    """Frame Class consisting of Top-Frame, Bottom-Frame and Markdown-Editor"""
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.master = master
        self.master.title("faranto mail gui")
        self.initialize_top_frame()
        self.initialize_bot_frame()
        self.mark = MarkdownEditor(self)

        self.top_frame.pack(side = "top", fill = "x")
        self.bot_frame.pack(side = "bottom", fill = "x")
        self.mark.pack(side = "bottom", fill = "x")
               
    def initialize_bot_frame(self):
        """Bottom Frame with Buttons for Mail Sending and adding attachments"""
        self.bot_frame = tk.Frame(self)
        self.add_attachment_button = tk.Button(self.bot_frame,
                text = "Add attachment", command = self.add_attachment)
        self.attachment_entry = tk.Entry(self.bot_frame, width = 45)
                
        self.send_mails_button = tk.Button(self.bot_frame, 
                text = "Send Mails", command = self.send_mails)
        self.send_test_button = tk.Button(self.bot_frame, 
                text = "Send Test Mail", command = self.send_test_mail)

        self.add_attachment_button.grid(row = 0,column = 0, sticky="w")
        self.attachment_entry.grid(row = 0, column = 1,
                sticky = "w")
        self.send_mails_button.grid(row = 1, column = 1, sticky = "w")
        self.send_test_button.grid(row = 1, column = 0, sticky = "we")

    def send_test_mail(self):
        yag = self.login_gmail() 
        subject = "Test Email"
        to = self.username_entry.get()
        try:
            body = self.mark.generate_html(0)
            if not self.attachment_entry.get():
                yag.send(to = to, subject = subject, contents = body )
            else:
                yag.send(to = to, subject = subject, contents = [body,
                    self.attachment_entry.get()])
        except Exception:
            self.print_error("Test Mail Failed.")

    def login_gmail(self):
        try:
            return ym.SMTP(self.username_entry.get(),
                    self.password_entry.get())
        except Exception:
            self.print_error("Wrong Username or Password.")
            self.file_entry.delete(0, "end")


    def send_mails(self):
        yag = self.login_gmail()
        subject = self.mark.get_subject_entry()
        try:
            for index, row in self.data_frame.iterrows():
                to = row[self.to_column]
                body = self.mark.generate_html(index)
                try:
                    if not self.attachment_entry.get():
                        yag.send(to = to, subject = subject, contents = body )
                    else:
                        yag.send(to = to, subject = subject, contents = [body,
                            self.attachment_entry.get()])
                except Exception:
                    self.print_error(("Unable to send Mails.Check Username "
                    "Password or Mail To."))
                    self.username_entry.delete(0, "end")
                    self.password_entry.delete(0, "end")
                    break;
        except AttributeError:
            self.print_error("No valid Excel Sheet selected.")

    def add_attachment(self):
        self.attachment_file_name = askopenfilename()
        self.attachment_entry.delete(0, "end")
        self.attachment_entry.insert(0, self.attachment_file_name)
        
    def initialize_top_frame(self):
        """Top Frame with Labels and Entries for G-Mail Login"""
        self.top_frame = tk.Frame(self)
        self.username_label = tk.Label(self.top_frame, text = "Username G-Mail:")
        self.username_entry = tk.Entry(self.top_frame, width = 80)
        self.username_entry.insert(0, "Enter your G-Mail Username here..")
        self.username_entry.bind("<FocusIn>",
                lambda event: self.click_in_entry(event, self.username_entry))

        self.password_label = tk.Label(self.top_frame, text = "Password G-Mail:")
        self.password_entry = tk.Entry(self.top_frame, show = "*", width = 80)
        self.password_entry.insert(0, "Enter your G-Mail Password here..")
        self.password_entry.bind("<FocusIn>",
                lambda event: self.click_in_entry(event, self.password_entry))

        self.file_button = tk.Button(self.top_frame,
                text = "Press to load .xlsx File", command = self.open_file)
        self.file_entry = tk.Entry(self.top_frame, width = 80)
        self.file_entry.insert(0, "Path/to/the/xlsxFile")
        self.first_click = 0

        #Set Layout of widgets
        self.username_label.grid(row = 0, column = 0, sticky="w")
        self.username_entry.grid(row = 0, column = 1, sticky = "w")
        self.password_label.grid(row = 1, column = 0, sticky = "w")
        self.password_entry.grid(row = 1, column = 1, sticky ="w")
        self.file_button.grid(row = 2, column = 0, sticky = "w")
        self.file_entry.grid(row = 2, column = 1, sticky = "w")

        #self.top_frame.pack(fill = "x")

    def open_file(self):
        """Open File Dialog. Accepts .xlsx File and transforms it to pandas data frame.
        Updates Mail-To-Button and Markdown-Frame Buttons"""
        self.file_name = askopenfilename()
        self.file_entry.delete(0, "end")
        self.file_entry.insert(0, self.file_name)
        try:
            self.excel_sheet = pd.ExcelFile(self.file_name)
            self.data_frame = self.excel_sheet.parse(
                self.excel_sheet.sheet_names[0])
            self.column_names = list(self.data_frame.columns.values)
            self.mark.set_dataframe(self.data_frame)
            self.mark.update_buttons(self.column_names)
            self.add_mail_to_option()
        except Exception:
            self.print_error("File format should be xlsx.")
            self.file_entry.delete(0, "end")
                
    def add_mail_to_option(self):
        """Adds Option menu with column names. Column Values are Recipients of the mails"""
        self.drop_menu_label = tk.Label(self.top_frame, text = "Mail to:")
        self.dropVar = tk.StringVar()
        self.dropVar.set("Choose Recipient here..")
        self.drop_menu = tk.OptionMenu(self.top_frame, self.dropVar, 
                *self.column_names, command = self.mail_to_listener)
        self.drop_menu.config(width = 80)

        self.drop_menu_label.grid(row = 3, column = 0, sticky = "w")
        self.drop_menu.grid(row = 3, column = 1, sticky = "w")

    def mail_to_listener(self, value):
        self.to_column = value

    def click_in_entry(self,event, entry):
        """Called to clear Entry when User clicks into it, just design"""
        if self.first_click < 2:
            entry.delete(0, "end")
            self.first_click += 1

    def print_error(self, error):
        """ Pops up window with error message """
        window = tk.Toplevel(root)
        ok = tk.Button(window, text = "Ok!", command = window.destroy)
        label = tk.Label(window, text = error)
        label.pack()
        ok.pack()

if __name__ == "__main__":
    root = tk.Tk()
    FarantoGUI(root).pack(side="top", fill="both", expand = True)
    root.mainloop()
#todo
#updater für mail to mit optionmenu <-
#markdown section von markdowngui_final ziehen <-
#makdown section updater schreiben für column buttons <-
#button logic für text area schreiben todo <-
#parse +columnname+ to variable todo <-
#generate and view html schreiben <-
#send mails logic schreiben todo <-
#testing 
#design
#refactoring
#py to exe?
