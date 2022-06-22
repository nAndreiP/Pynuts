import os, xlwt, xlrd, xlutils.copy
from config.definitions import ROOT_DIR
from tkinter import messagebox

class mainClass:

    def exit(self):
        self.root.destroy()

    def clear(self):
        self.denumire_field.delete(0, "end")
        self.lot_field.delete(0, "end")
        self.concentratie_field.delete(0, "end")
        self.bbd_field.delete(0, "end")
        self.serie_field.delete(0, "end")
        self.email_pr_field.delete(0, "end")
        self.adresa_field.delete(0, "end")

    def excel(self):
        if os.path.exists(os.path.join(ROOT_DIR, "output/0.xls")) is False:
            rb = xlwt.Workbook(os.path.join(ROOT_DIR,"output/0.xls"))
            sheet = rb.add_sheet("PV")
            
            top_style = xlwt.XFStyle()
            top_border = xlwt.Borders()
            top_border.top = top_border.right = top_border.bottom = top_border.left = 1
            top_font = xlwt.Font()
            top_font.bold = True
            top_font.height = 280
            top_align = xlwt.Alignment()
            top_align.horz = 0x02
            top_align.wrap = 1
            top_style.borders = top_border
            top_style.font = top_font
            top_style.alignment = top_align
           
            for n in range(1,8):
                sheet.col(n).width = 5000
            
            sheet.write(0,1,"Denumire", style = top_style)
            sheet.write(0,2,"Lot", style = top_style)
            sheet.write(0,3,"Concentratie", style = top_style)
            sheet.write(0,4,"BBD", style = top_style)
            sheet.write(0,5,"Serie", style = top_style)
            sheet.write(0,6,"Email pr", style = top_style)
            sheet.write(0,7,"Adresa", style = top_style)
            rb.save(os.path.join(ROOT_DIR,"output/0.xls"))
    
    
        rb = xlrd.open_workbook(os.path.join(ROOT_DIR,"output/0.xls"), formatting_info = True)
        sheet = rb.sheet_by_index(0)
        
        self.wb = xlutils.copy.copy(rb)
        self.wb_sheet = self.wb.get_sheet(0)
        
        try:
        
            last_row = sheet.nrows
        
        except IndexError:
            last_row = 0
            
        top_style = xlwt.XFStyle()
        top_border = xlwt.Borders()
        top_border.top = top_border.right = top_border.bottom = top_border.left = 1
        top_font = xlwt.Font()
        top_font.bold = True
        top_font.height = 280
        top_align = xlwt.Alignment()
        top_align.horz = 0x02
        top_align.wrap = 1
        top_style.borders = top_border
        top_style.font = top_font
        top_style.alignment = top_align    
        
        
        self.wb_sheet.write(0,1,"Denumire", style = top_style)
        self.wb_sheet.write(0,2,"Lot", style = top_style)
        self.wb_sheet.write(0,3,"Concentratie", style = top_style)
        self.wb_sheet.write(0,4,"BBD", style = top_style)
        self.wb_sheet.write(0,5,"Serie", style = top_style)
        self.wb_sheet.write(0,6,"Email pr", style = top_style)
        self.wb_sheet.write(0,7,"Adresa", style = top_style)
        
        style1 = xlwt.easyxf("align: horiz center, wrap yes; border: left thin, right thin, top thin, bottom thin; font: bold off, colour red")
        
        
        self.wb_sheet.write(last_row, 1, self.denumire_field.get(), style = style1)
        self.wb_sheet.write(last_row, 2, self.lot_field.get(), style = style1)
        self.wb_sheet.write(last_row, 3, self.concentratie_field.get(), style = style1)
        self.wb_sheet.write(last_row, 4, self.bbd_field.get(), style = style1)
        self.wb_sheet.write(last_row, 5, self.serie_field.get(), style = style1)
        self.wb_sheet.write(last_row, 6, self.email_pr_field.get(), style = style1)
        self.wb_sheet.write(last_row, 7, self.adresa_field.get(), style = style1)
        
        self.wb.save(os.path.join(ROOT_DIR, "output/0.xls"))
    
    def insert(self):
        if (self.denumire_field.get() == "" and
            self.lot_field.get() == "" and
            self.concentratie_field.get() == "" and
            self.bbd_field.get() == "" and
            self.serie_field.get() == "" and
            self.email_pr_field.get() == "" and
            self.adresa_field.get() == ""):
                
            messagebox.showinfo("Atentie!", "Fields empty")
            
        else:
            
            self.excel()
            self.clear()
            
            