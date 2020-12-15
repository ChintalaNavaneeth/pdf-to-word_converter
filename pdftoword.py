import win32com.client
import os
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, asksaveasfile

# =================open file method======================
def openFile():
    global FILE_PATH
    FILE_PATH= askopenfilename(defaultextension=".pdf",
                           filetypes=[("Pdf files", "*.pdf")])
    if FILE_PATH == "":
        FILE_PATH = None
    else:
        fileEntry.delete(0, END)
        fileEntry.config(fg="black")
        fileEntry.insert(0, FILE_PATH)
# =========== File conversion ==========================
def convert():
    word = win32com.client.Dispatch("word.Application")
    word.visible = 0
    doc_pdf = FILE_PATH
    input_file = os.path.abspath(doc_pdf)
    wb = word.Documents.Open(input_file)
    output_file = os.path.abspath(doc_pdf[0:-4] + "".format())
    wb.SaveAs2(output_file, FileFormat=16)
    messagebox.showinfo(None,"Converted sucessfully")
    wb.Close()
    word.Quit()
# =================== Front End Design =====================
root = Tk()
root.geometry("600x200")
root.config(bg="light blue")
root.title("PDF Converter [Designed by : Chintala Navaneeth]")
root.resizable(0, 0)
FILE_PATH = ""
# ==============App Name==============================================================>>
appName = Label(root, text="PDF to WORD Converter ", font=('arial', 20, 'bold'),
                bg="light blue", fg='maroon')
appName.place(x=150, y=5)
# Select pdf file
labelFile = Label(root, text="Select Pdf File", font=('arial', 12, 'bold'))
labelFile.place(x=30, y=50)
fileEntry = Entry(root, font=('calibri', 12), width=80)
fileEntry.pack(ipadx=250, pady=50, padx=150)
# ===========button to access openFile method=================================
openFileButton = Button(root, text=" Open ", font=('arial', 12, 'bold'), width=30,
                        bg="sky blue", fg='green', command=openFile)
openFileButton.place(x=150, y=80)
#==============================Button convert method=================
save2Word = Button(root, text=" Convert ", font=('arial', 12, 'bold'),width = 30,
                   bg="light green", fg='black', command=convert)
save2Word.place(x=150, y=120)
# ===================halt window=============================>>
if __name__ == "__main__":
    root.mainloop()