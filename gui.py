import tkinter, os
from config.definitions import ROOT_DIR
from mainScript import mainClass

class guiClass(mainClass):
    
    def __init__(self): 
        self.root = tkinter.Tk()
        self.root.configure(background='light green')
        self.root.title("Proces Verbal")
        self.root.geometry("420x210+500+300")
        
        self.heading = tkinter.Label(self.root, text="Formular", bg="light green")
        self.heading.grid(row=0, column=1)
        
        self.denumire = tkinter.Label(self.root, text="Denumire", bg="light green")
        self.denumire.grid(row=1, column=0)
        
        self.lot = tkinter.Label(self.root, text="Lot", bg="light green")
        self.lot.grid(row=2, column=0)
        
        self.concentratie = tkinter.Label(self.root, text="Concentratie", bg="light green")
        self.concentratie.grid(row=3, column=0)
        
        self.bbd = tkinter.Label(self.root, text="Bbd", bg="light green")
        self.bbd.grid(row=4, column=0)
        
        self.serie = tkinter.Label(self.root, text="Serie", bg="light green")
        self.serie.grid(row=5, column=0)
        
        self.email_pr = tkinter.Label(self.root, text="Email pr", bg="light green")
        self.email_pr.grid(row=6, column=0)
        
        self.adresa = tkinter.Label(self.root, text="Adresa", bg="light green")
        self.adresa.grid(row=7, column=0)
        
        #excel()
        
        self.denumire_field = tkinter.Entry(self.root)
        self.lot_field = tkinter.Entry(self.root)
        self.concentratie_field = tkinter.Entry(self.root)
        self.bbd_field = tkinter.Entry(self.root)
        self.serie_field = tkinter.Entry(self.root)
        self.email_pr_field = tkinter.Entry(self.root)
        self.adresa_field = tkinter.Entry(self.root)
        
        self.denumire_field.grid(row=1, column=1, ipadx="100")
        self.lot_field.grid(row=2, column=1, ipadx="100")
        self.concentratie_field.grid(row=3, column=1, ipadx="100")
        self.bbd_field.grid(row=4, column=1, ipadx="100")
        self.serie_field.grid(row=5, column=1, ipadx="100")
        self.email_pr_field.grid(row=6, column=1, ipadx="100")
        self.adresa_field.grid(row=7, column=1, ipadx="100")
        
        self.submit = tkinter.Button(self.root, text="Submit", fg="Black",
							bg="light blue", command=self.insert)
        self.submit.place(x=150,y=170)
        
        self.exitButton = tkinter.Button(self.root, text = "Exit", bg = "Red", command = self.exit)
        self.exitButton.place(x=270,y=170)
        
        self.root.mainloop()
        
obj = guiClass()