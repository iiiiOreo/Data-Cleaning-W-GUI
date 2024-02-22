import tkinter as tk
import aspose.words as aw
import os
from tkinter import filedialog
from nltk.stem import PorterStemmer , WordNetLemmatizer
from nltk.tokenize import word_tokenize




class cleaner:
    def __init__(self):
        self.window = tk.Tk()
        self.window.geometry("600x300")
        self.window.resizable(0, 0)
        self.window.configure(bg='#222831')
        self.window.title("AIO 101 Project")
        self.create_fun_buttons()
        self.word_count = {}
        self.listy = []
        self.result = ""
        self.find_entry = tk.StringVar()
        
    def open_file(self):
        self.text = filedialog.askopenfile(mode='r', filetypes=[('Text Files', '*.txt'),('pdf Files', '*.pdf')])
        if self.text:
            filepath = os.path.abspath(self.text.name)
            tk.Label(self.window, text="\nThe File is located at : " + str(filepath), bg = "#222831" ,fg = "#D52484" ,font=('Aerial 11')).pack()
        x, self.y = os.path.splitext(self.text.name)
        if self.y == ".pdf":
            self.txt()
            self.text = open("C:/Users/youse/Desktop/Untitled.txt","r",encoding='utf-8')
            
    def txt(self):
        txt = aw.Document(self.text.name)
        txt.save("C:/Users/youse/Desktop/Untitled.txt")

    def uppercase(self):
        self.result = self.text.read().upper()
        self.before_run()

    def lowercase(self):
        self.result = self.text.read().lower()
        self.before_run()
    
    def lemmatization(self):
        lemmatizer = WordNetLemmatizer()
        words = word_tokenize(self.text.read())
        for w in words:
            self.listy.append(str(w)+" : "+str(lemmatizer.lemmatize(w)))
        for item in self.listy:
            self.result = self.result + item + "\n"
        self.before_run()

    def stemming(self):
        ps = PorterStemmer()
        words = word_tokenize(self.text.read())
        for w in words:
            self.listy.append(str(w)+" : "+str(ps.stem(w)))
        for item in self.listy:
            self.result = self.result + item + "\n"
        self.before_run()

    def edit(self):
        toplevel = tk.Toplevel()
        toplevel.title("Edit Window")
        toplevel.geometry("700x510")
        toplevel.resizable(0, 0)
        toplevel.configure(bg='#222831')
        label = tk.Label(toplevel, text="Edit what you want : ", font=("Arial", 18), bg = "#222831" ,fg = "#D52484")
        label.pack()
        toplevel.focus_set()
        my_frame = tk.Frame(toplevel)
        my_frame.pack(pady=5)
        my_text = tk.Text(my_frame, bg = "#222831" ,fg = "#CD4DCC" , width=95, height=25, undo=True)
        my_text.insert(tk.INSERT, self.text.read())
        my_text.pack()
        edit_btn = tk.Button(toplevel, bg = "#CD4DCC" , text='Save', command=lambda: self.save())
        edit_btn.pack(pady=20)
        self.changeOnHover(edit_btn, "#D62196", "#CD4DCC")

    def find_cont(self):
        word_finded = self.find_entry.get()
        i = 0
        w = 0
        flag = False
        for line in self.text:
            i = i + 1
            words = line.split()
            for word in words:
                word = word.lower()
                w = w + 1
                if word_finded == word:
                    resultt = i
                    resulttt = w
                    flag = True
                    self.result += "\nThe word you want to find is in the line number : " + str(
                        resultt) + " and the word number : " + str(resulttt)
            if not flag:
                self.result = "can't be found in the file"
        self.before_run()

    def find(self):
        toplevel = tk.Toplevel()
        toplevel.title("Find Window")
        toplevel.geometry("200x200")
        toplevel.configure(bg='#222831')
        label = tk.Label(toplevel, text="\n enter the text you want to find : \n", font=("Arial", 10), bg = "#222831" ,fg = "#D52484")
        self.find_entry = tk.Entry(toplevel, font=('calibre', 10, 'normal'))
        label.pack()
        self.find_entry.pack()
        find_btn = tk.Button(toplevel, text="Find ", bg = "#CD4DCC", command=lambda: self.find_cont())
        find_btn.pack(pady=10)
        toplevel.focus_set()
        self.changeOnHover(find_btn, "#D62196", "#CD4DCC")

    def pdf(self):
        doc = aw.Document(self.text.name)
        doc.save("THE PDF.pdf", aw.SaveFormat.PDF)
        self.before_run()

    def count(self):
        for line in self.text:
            words = line.split()
            for word in words:
                word = word.lower()
                if word in self.word_count:
                    self.word_count[word] += 1
                else:
                    self.word_count[word] = 1
        self.result = (self.word_count)
        self.before_run()

    def create_fun_buttons(self):
        self.b1 = tk.Button(self.window, width=10, height=1, bg = "#BB86FC" ,text="Uppercase", command=lambda: self.uppercase())
        self.b1.place(x=55, y=160)
        self.b2 = tk.Button(self.window, width=10, height=1, bg = "#BB86FC" ,text="Lowercase", command=lambda: self.lowercase())
        self.b2.place(x=185, y=160)
        self.b3 = tk.Button(self.window, width=10, height=1, bg = "#BB86FC" ,text="lemmatize", command=lambda: self.lemmatization())
        self.b3.place(x=335, y=160)
        self.b4 = tk.Button(self.window, width=10, height=1, bg = "#BB86FC" ,text="stemming", command=lambda: self.stemming())
        self.b4.place(x=470, y=160)
        self.b5 = tk.Button(self.window, width=10, height=1, bg = "#7E0CF5" ,text="Edit", command=lambda: self.edit())
        self.b5.place(x=55, y=230)
        self.b6 = tk.Button(self.window, width=10, height=1, bg = "#7E0CF5" ,text="Find", command=lambda: self.find())
        self.b6.place(x=185, y=230)
        self.b7 = tk.Button(self.window, width=10, height=1, bg = "#7E0CF5" ,text="PDF Converter", command=lambda: self.pdf())
        self.b7.place(x=335, y=230)
        self.b8 = tk.Button(self.window, width=10, height=1, bg = "#7E0CF5" ,text="Counter", command=lambda: self.count())
        self.b8.place(x=470, y=230)
        self.browse = tk.Button(self.window, width=10, height=1, bg = "#CD4DCC" ,text="Browse", command=lambda: self.open_file())
        self.browse.place(x=260, y=80)
        
    def changeOnHover(self, button, colorOnHover, colorOnLeave):
        button.bind("<Enter>", func=lambda e: button.config(
            background=colorOnHover))
        button.bind("<Leave>", func=lambda e: button.config(
            background=colorOnLeave))

    def before_run(self):
        toplevel = tk.Toplevel()
        toplevel.title("Save Window")
        toplevel.geometry("700x510")
        toplevel.resizable(0, 0)
        toplevel.configure(bg='#222831')
        label = tk.Label(toplevel, text="Done Successfully", font=("Arial", 18), bg = "#222831" ,fg = "#D52484")
        label.pack()
        toplevel.focus_set()
        my_frame = tk.Frame(toplevel)
        my_frame.pack(pady=5)
        my_text = tk.Text(my_frame, bg = "#222831" ,fg = "#CD4DCC", width=95, height=25, undo=True)
        my_text.insert(tk.INSERT, self.result)
        my_text.configure(state='disabled')
        my_text.pack()
        save_btn = tk.Button(toplevel, text='Save',bg = "#CD4DCC", command=lambda: self.save())
        save_btn.pack(pady=20)
        self.changeOnHover(save_btn, "#D62196", "#CD4DCC")

    def save(self):
        file = open("C:/Users/youse/Desktop/result.txt", 'w')
        file.write(str(self.result))
        file.close()
        self.text.close()

    def run(self):
        self.changeOnHover(self.b1, "#CD4DCC", "#BB86FC")
        self.changeOnHover(self.b2, "#CD4DCC", "#BB86FC")
        self.changeOnHover(self.b3, "#CD4DCC", "#BB86FC")
        self.changeOnHover(self.b4, "#CD4DCC", "#BB86FC")
        self.changeOnHover(self.b5, "#D62196", "#7E0CF5")
        self.changeOnHover(self.b6, "#D62196", "#7E0CF5")
        self.changeOnHover(self.b7, "#D62196", "#7E0CF5")
        self.changeOnHover(self.b8, "#D62196", "#7E0CF5")
        self.changeOnHover(self.browse, "#D62196", "#CD4DCC")
        self.window.mainloop()



AIO = cleaner()
AIO.run()