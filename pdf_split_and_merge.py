"""
Program to split and merge pdf files

Author : Evan Giavina
Date : 29.09.2020
"""

from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
from pathlib import Path
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from functools import partial
from PIL import Image,ImageTk
from pdf2image import convert_from_path
import os

class App:
    """A class to create the pdf split and merge app"""
    def __init__(self, master):
        """Initialize the GUI"""
        self.file_field_counter = 0
        self.dir = []
        self.dirSplit = ''
        
        #Create tabs
        self.master = master
        self.notebk = ttk.Notebook(self.master)
        self.notebk.grid(column=0, row=0)
        self.frame1 = ttk.Frame(self.notebk, width = 400, height = 400)
        self.frame2 = ttk.Frame(self.notebk, width = 400, height = 400)
        self.notebk.add(self.frame1, text = 'Split')
        self.notebk.add(self.frame2, text = 'Merge')
        
        
        #Split
        self.groupFilePath = LabelFrame(self.frame1, text="Select file", padx=5, pady=5,labelanchor=N)
        self.groupFilePath.grid(column=0,row=0,columnspan=3, sticky="WE")
        
        self.L_filepathSplit = Label(self.groupFilePath, width=60,text="Filepath", pady=10)
        self.L_filepathSplit.grid(column=0, row=0)
        
        self.B_chooseFileSplit = Button(self.groupFilePath, text="Select file to split", width=30, command=self.openFileSplit)
        self.B_chooseFileSplit.grid(column=1, row=0)  
        
        
        
        self.groupCut = LabelFrame(self.frame1, text="Cuts", padx=5, pady=5,labelanchor=N)
        self.groupCut.grid(column=0,row=1,columnspan=3, sticky="WE")
        
        self.E_cuts = Entry(self.groupCut, width=60)
        self.E_cuts.insert(END, '')
        self.E_cuts.grid(column=0,row=1)
        
        self.L_example = Label(self.groupCut, width=60,text="Use visual tool or enter pages after which it should cut, ex:'1,4,5'", pady=10)
        self.L_example.grid(column=1, row=1)
        
        
        
        self.groupSave = LabelFrame(self.frame1, text="Split", padx=5, pady=5,labelanchor=N)
        self.groupSave.grid(column=0,row=2,columnspan=3, sticky="WE")
        
        self.C_openVar = IntVar()
        self.C_open = Checkbutton(self.groupSave, text="Open created files",variable=self.C_openVar, pady=10, width=60)
        self.C_open.grid(column=0, row=2)
        self.C_open.select()     
        
        self.B_FileSplit = Button(self.groupSave, text="Split the file", width=30, command=self.fileSplit)
        self.B_FileSplit.grid(column=1, row=2)  
        

        
        #Merge
        frame = Frame(self.frame2)
        frame.grid(column=2,row=2)

        self.B_merge = Button(frame, text="Merge files", command=self.start_merge)
        self.B_merge.grid(column=0,row=0)
        
        #Load file
        self.groupFile = LabelFrame(frame, text="Files", padx=5, pady=5,labelanchor=N)
        self.groupFile.grid(column=0,row=1,columnspan=3, sticky="WE")
        
        self.L_filepath = []
        self.B_File = []
        self.dir = []
        self.add_file_field()
 

        #Console
        self.console = Text(root, height=30, width=60)
        self.console.grid(column=80, row=0, columnspan=20)
        self.console.tag_configure("stderr", foreground="#b22222")

        sys.stdout = TextRedirector(self.console, "stdout")
        sys.stderr = TextRedirector(self.console, "stderr")
       
    def fileSplit(self):
        """Split the file, asks user where to save"""
        pages = self.E_cuts.get()
        if(pages == ''):
            p = self.pagesToCut
        else:
            for char in pages:
                if not((char.isnumeric()) or (char == ",")):
                    print("The cuts inputted are incorrect")
            p = []
            for i in pages.split(","):
                if i != '':
                    p.append(int(i))
           
        if self.dirSplit == '':
            print("Select a file")
        else:
            dirSplitSave = filedialog.askdirectory(parent=root, title="Select save folder")
            if dirSplitSave != '':
                pdf = PdfFileReader(self.dirSplit, strict=False) 
                p.append(pdf.getNumPages())
                p.sort()
                
                
                pdf_writer = PdfFileWriter()
                for page in range(pdf.getNumPages()):
                    pdf_writer.addPage(pdf.getPage(page))
                    
                    if (page+1) == p[0]:
                        output_filename = dirSplitSave + '\\' + 'page_{}.pdf'.format(page+1)
                        with open(output_filename, 'wb') as out:
                            pdf_writer.write(out)
                            
                        print("File created: {}".format(output_filename))
                        if self.C_openVar.get():
                            os.system(output_filename)
                            
                        pdf_writer = PdfFileWriter()
                        p.pop(0)
                

        
    def openFileSplit(self):
        """Open the file to split, then display the visual cut guide"""
        
        self.dirSplit = filedialog.askopenfilename(parent=root, title="Select file")
        
        if self.dirSplit != '':
            notSelected = "grey"
            highlightNotSelected = "DarkSlateGray4"
            highlightSelected = "DarkSlateGray2"
            selected = "DarkSlateGray3"
            
            #Split: pdf view
            self.pagesToCut = []
            self.pdfWindow = Toplevel(self.master)
            self.pdfWindow.title("Pdf view")
            self.pdfWindow.geometry("600x849+0+0")
            self.scroll = Scrollbar(self.pdfWindow)
            self.pdf = Text(self.pdfWindow, yscrollcommand=self.scroll.set, bg=notSelected, insertborderwidth=15)
            self.scroll.pack(side=RIGHT, fill=Y)
            self.scroll.config(command=self.pdf.yview)
            self.pdf.pack(fill=BOTH, expand=1)
            
            def leaveTag(event, tagNum):
                tag = "tag" + str(tagNum)
                self.pdf.config(cursor="arrow")
                if(self.pdf.tag_cget(tag, 'background') == highlightNotSelected):
                    self.pdf.tag_config(tag, background=notSelected)
                else:
                    self.pdf.tag_config(tag, background=selected)
        
            def enterTag(event, tagNum):
                tag = "tag" + str(tagNum)
                self.pdf.config(cursor="sb_h_double_arrow")
                
                if(self.pdf.tag_cget(tag, 'background') == notSelected):
                    self.pdf.tag_config(tag, background=highlightNotSelected)
                else:
                    self.pdf.tag_config(tag, background=highlightSelected)
                
            def cut(event, tagNum):
                tag = "tag" + str(tagNum)
                if(self.pdf.tag_cget(tag, 'background') == highlightSelected):
                    self.pdf.tag_config(tag, background=highlightNotSelected)
                    self.pagesToCut.remove(tagNum)
                else:
                    self.pdf.tag_config(tag, background=highlightSelected)
                    self.pagesToCut.append(tagNum)
                

            self.pdf.config(cursor="arrow")
            
            #get pdf images
            self.master.update()
            width = 600 - self.scroll.winfo_width()
            height = width * 2**0.5;
            self.pages = convert_from_path(self.dirSplit, size=(width, height))
            self.pageScreenshots = []
            for i in range(len(self.pages)):
                self.pageScreenshots.append(ImageTk.PhotoImage(self.pages[i]))
                
            first = True
            i = 0
            tagNumList = []
            for screenshot in self.pageScreenshots:
                if first:
                    first = False
                else:
                    tag = "tag" + str(i)
                    tagNum = int(i)
                    self.pdf.insert(END, '\n\n')
                    self.pdf.insert(END, '\n\n', tag)
                    self.pdf.tag_config(tag, background=notSelected)
                    self.pdf.tag_bind(tag, "<Enter>", partial(enterTag, tagNum=i))
                    self.pdf.tag_bind(tag, "<Leave>", partial(leaveTag, tagNum=i))
                    self.pdf.tag_bind(tag, "<Button-1>", partial(cut, tagNum=i))
                    self.pdf.insert(END, '\n')
                self.pdf.image_create(END, image=screenshot)
                tagNumList.append(i)
                i += 1
            
            self.pdf.pack(fill=BOTH, expand=1)
            self.pdf.config(state=DISABLED)
            
        self.L_filepathSplit.configure(text=self.dirSplit)
    
    def add_file_field(self):        
        """Add a 'add file' field"""
        self.L_filepath.append(Label(self.groupFile, width=60,text="Filepath"))
        self.L_filepath[self.file_field_counter].grid(column=0, row=self.file_field_counter)
        
        self.B_File.append(Button(self.groupFile, text="Select file", width=30, command=partial(self.open_file, self.file_field_counter)))
        self.B_File[self.file_field_counter].grid(column=1, row=self.file_field_counter)   
        
        self.file_field_counter = self.file_field_counter + 1
        
            
    def start_merge(self):
        """Takes the inputted files, merge them and asks user where to save result"""
        pdf_merger = PdfFileMerger(strict=False)

        files_path = self.dir


        for path in files_path:
            for path2 in path:
                pdf_merger.append(str(path2))
        
        filename = filedialog.asksaveasfilename(parent=root,title = "Save as",defaultextension=".pdf", filetypes =(("pdf files","*.pdf"),("all files","*.*")))
        with Path(filename).open(mode="wb") as output_file:
            pdf_merger.write(output_file)
            print("File created: " + str(filename))
            
        os.system(filename)

    def quit_program(self):
        """Quit the app"""
        root.destroy()

    def open_file(self, id, event = None):
        """Open file and add new 'add file' field if needed"""
        if id >= (len(self.B_File)-1):
            self.add_file_field()
            self.dir.append("")
        
        self.dir[id] = filedialog.askopenfilename(parent=root, title="Select file", multiple=True, filetypes =(("pdf files","*.pdf"),("all files","*.*")))
        self.L_filepath[id].configure(text=self.dir[id])

        
class TextRedirector(object):
    """A class that lets you direct standard and error output to tkinter text widget"""
    ## Initialization method
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    ## Method used to write to the text widget
    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.configure(state="disabled")
        self.widget.see(END)

if __name__ == '__main__':
    root = Tk()

    app = App(root)

    root.title("Pdf split & merge, created by Evan Giavina, 29.09.2020")
    root.protocol("WM_DELETE_WINDOW", app.quit_program)
    root.mainloop()
