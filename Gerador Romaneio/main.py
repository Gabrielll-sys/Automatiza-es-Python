from tkinter import *
from tkinter import ttk
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import os


def Salvar():
    wb= Workbook()
    ws = wb.active
    ws.title="Testes"
    # for i in range(int(first_name_entry.get()),int(last_name_entry.get())):
        
    #     ws['A'+str(i)]=i
    #     ws['B'+str(i)]=i
    #     ws['C'+str(i)]=i
    #     ws['D'+str(i)]=i
    #     ws['E'+str(i)]=i
    #     ws['F'+str(i)]=i
    #     ws['G'+str(i)]=i
    #     ws['H'+str(i)]=i
    #     ws['I'+str(i)]=i
    #     ws['J'+str(i)]=i
    #     ws['K'+str(i)]=i
    ws['A4']= ferramenta_entry.get()

    wb.save("Teste.xlsx")
window = Tk()
window.title("Automatização Planilhas Master")

frame = ttk.Frame(window)
frame.pack()

# Saving User Info
user_info_frame =ttk.LabelFrame(frame, text="Gerador romaneio")
user_info_frame.grid(row= 0, column=0, padx=60, pady=10)

ferramenta_label = ttk.Label(user_info_frame, text="Ferramenta")
ferramenta_label.grid(row=0, column=0)


ferramenta_entry = ttk.Combobox(user_info_frame,value=["Alicate de bico"])
ferramenta_entry.grid(row=1, column=0)


title_label = ttk.Label(user_info_frame, text="Polegadas")
unidade_combobox = ttk.Combobox(user_info_frame, values=['1/8"', '1/4"', '1/2"', '1"'])
title_label.grid(row=0, column=1)
unidade_combobox.grid(row=1, column=1)

quantidade_label = ttk.Label(user_info_frame, text="Quantidade")
quantidade_spinbox = ttk.Spinbox(user_info_frame, from_=1, to="infinity")
quantidade_label.grid(row=0, column=3)
quantidade_spinbox.grid(row=1, column=3)


button = ttk.Button(user_info_frame, text="Adicionar", command= Salvar)
button.grid(row=3, column=1, sticky="news", padx=20, pady=10)


for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Saving Course Info
courses_frame = ttk.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

registered_label = ttk.Label(courses_frame, text="Registration Status")



registered_label.grid(row=0, column=0)


numcourses_label = ttk.Label(courses_frame, text= "# Completed Courses")
numcourses_spinbox = ttk.Spinbox(courses_frame, from_=0, to='infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)

numsemesters_label = ttk.Label(courses_frame, text="# Semesters")
numsemesters_spinbox = ttk.Spinbox(courses_frame, from_=0, to="infinity")
numsemesters_label.grid(row=0, column=2)
numsemesters_spinbox.grid(row=1, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)





# Button
button = ttk.Button(frame, text="Enter data", command= Salvar)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)
 
window.mainloop()

# wb= Workbook()
# ws = wb.active
# ws.title="Testes"
# img=Image('')
# for i in range(1,20):
    
#     ws['A'+str(i)]=i
#     ws['B'+str(i)]=i
#     ws['C'+str(i)]=i
#     ws['D'+str(i)]=i
#     ws['E'+str(i)]=i
#     ws['F'+str(i)]=i
#     ws['G'+str(i)]=i
#     ws['H'+str(i)]=i
#     ws['I'+str(i)]=i
#     ws['J'+str(i)]=i
#     ws['K'+str(i)]=i
# ws.add_image(img,'A21')
# wb.save("Teste.xlsx")

