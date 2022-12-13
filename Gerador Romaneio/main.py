from tkinter import *
from tkinter import ttk
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import *
import datetime
import os
ferramentas=['Alicate ','Alicate','Serra','Alicate']

data=datetime.date.today()
def Criar():
    # ferramentas.append(ferramenta_entry)
    nova_label=ttk.Label(ferramentas[len(ferramentas)-1])
    
    print(ferramentas[1])
def Salvar():
    valores = [ferramenta_entry.get(),quantidade_spinbox.get(),unidade_combobox.get()]
    # ferramentas.sort()
    print(ferramentas)
    wb= Workbook()
    ws = wb.active
    ws.title="Testes"
    for i in range(1,20):
        ws.append([valores[0],valores[1],valores[2]])
       
        
    
    # # img=Image("Logo.jpg")
    # # # img.anchor = 'A10'
    # # ws.add_image(img,'B1')
    wb.save("Teste.xlsx")
print(data)


window = Tk()
window.title("Automatização Planilhas Master")

frame = ttk.Frame(window)
frame.pack()

# Saving User Info
user_info_frame =ttk.LabelFrame(frame, text="Gerador romaneio")
user_info_frame.grid(row= 0, column=0, padx=60, pady=10)

ferramenta_label = ttk.Label(user_info_frame, text="Ferramenta")
ferramenta_label.grid(row=0, column=0)


ferramenta_entry = ttk.Combobox(user_info_frame,value=ferramentas)
ferramenta_entry.grid(row=1, column=0)

itens_adicionados_label=ttk.Label(user_info_frame,text=ferramentas)
itens_adicionados_label.grid(row=4,column=1)
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








# Button
button = ttk.Button(frame, text="Enter data", command= Salvar)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)
 
window.mainloop()
