# -*- coding: utf-8 -*-
"""
Created on Mon Feb 24 18:44:24 2020

@author: Alex
"""

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import datetime
import xlsxwriter
from tkinter import filedialog
import tkinter as tk
from tkinter import ttk
import os

def generate(url,media):

    print(url)

    df = pd.read_excel(url,header=0,converters={'Codigos':str})
        
    df.drop(df.columns[[0]], axis = 1, inplace = True)
    df.drop(df.tail(1).index,inplace=True)
    df.drop(columns=['P.v.p.','SMín','Lote ','Tot. ','A.An '],inplace=True)
    current_date = datetime.datetime.now()
    current_year = current_date.year
    current_month = current_date.month
    print(df.columns)
    if current_month == 1:
        df.rename(columns = {'Ene. ':datetime.date(current_year-1, 1, 1),
                             'Feb. ':datetime.date(current_year-1, 2, 1),
                             'Mar. ':datetime.date(current_year-1, 3, 1),
                             'Abr. ':datetime.date(current_year-1, 4, 1),
                             'May. ':datetime.date(current_year-1, 5, 1),
                             'Jun. ':datetime.date(current_year-1, 6, 1),
                             'Jul. ':datetime.date(current_year-1, 7, 1),
                             'Ago. ':datetime.date(current_year-1, 8, 1),
                             'Sep. ':datetime.date(current_year-1, 9, 1),
                             'Oct. ':datetime.date(current_year-1, 10, 1),
                             'Nov. ':datetime.date(current_year-1, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'Ene. .1':datetime.date(current_year, current_month, 1)},inplace = True)
    
    elif current_month == 2:
        df.rename(columns = {'Ene. ':datetime.date(current_year, 1, 1),
                             'Feb. ':datetime.date(current_year-1, 2, 1),
                             'Mar. ':datetime.date(current_year-1, 3, 1),
                             'Abr. ':datetime.date(current_year-1, 4, 1),
                             'May. ':datetime.date(current_year-1, 5, 1),
                             'Jun. ':datetime.date(current_year-1, 6, 1),
                             'Jul. ':datetime.date(current_year-1, 7, 1),
                             'Ago. ':datetime.date(current_year-1, 8, 1),
                             'Sep. ':datetime.date(current_year-1, 9, 1),
                             'Oct. ':datetime.date(current_year-1, 10, 1),
                             'Nov. ':datetime.date(current_year-1, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'Feb. .1':datetime.date(current_year, current_month, 1)},inplace = True)
    elif current_month == 3:
        df.rename(columns = {'Ene. ':datetime.date(current_year, 1, 1),
                             'Feb. ':datetime.date(current_year, 2, 1),
                             'Mar. ':datetime.date(current_year-1, 3, 1),
                             'Abr. ':datetime.date(current_year-1, 4, 1),
                             'May. ':datetime.date(current_year-1, 5, 1),
                             'Jun. ':datetime.date(current_year-1, 6, 1),
                             'Jul. ':datetime.date(current_year-1, 7, 1),
                             'Ago. ':datetime.date(current_year-1, 8, 1),
                             'Sep. ':datetime.date(current_year-1, 9, 1),
                             'Oct. ':datetime.date(current_year-1, 10, 1),
                             'Nov. ':datetime.date(current_year-1, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'Mar. .1':datetime.date(current_year, current_month, 1)},inplace = True)
    elif current_month == 4:
        df.rename(columns = {'Ene. ':datetime.date(current_year, 1, 1),
                             'Feb. ':datetime.date(current_year, 2, 1),
                             'Mar. ':datetime.date(current_year, 3, 1),
                             'Abr. ':datetime.date(current_year-1, 4, 1),
                             'May. ':datetime.date(current_year-1, 5, 1),
                             'Jun. ':datetime.date(current_year-1, 6, 1),
                             'Jul. ':datetime.date(current_year-1, 7, 1),
                             'Ago. ':datetime.date(current_year-1, 8, 1),
                             'Sep. ':datetime.date(current_year-1, 9, 1),
                             'Oct. ':datetime.date(current_year-1, 10, 1),
                             'Nov. ':datetime.date(current_year-1, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'Abr. .1':datetime.date(current_year, current_month, 1)},inplace = True)  
    elif current_month == 5:
        df.rename(columns = {'Ene. ':datetime.date(current_year, 1, 1),
                             'Feb. ':datetime.date(current_year, 2, 1),
                             'Mar. ':datetime.date(current_year, 3, 1),
                             'Abr. ':datetime.date(current_year, 4, 1),
                             'May. ':datetime.date(current_year-1, 5, 1),
                             'Jun. ':datetime.date(current_year-1, 6, 1),
                             'Jul. ':datetime.date(current_year-1, 7, 1),
                             'Ago. ':datetime.date(current_year-1, 8, 1),
                             'Sep. ':datetime.date(current_year-1, 9, 1),
                             'Oct. ':datetime.date(current_year-1, 10, 1),
                             'Nov. ':datetime.date(current_year-1, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'May. .1':datetime.date(current_year, current_month, 1)},inplace = True) 
    elif current_month == 6:
        df.rename(columns = {'Ene. ':datetime.date(current_year, 1, 1),
                             'Feb. ':datetime.date(current_year, 2, 1),
                             'Mar. ':datetime.date(current_year, 3, 1),
                             'Abr. ':datetime.date(current_year, 4, 1),
                             'May. ':datetime.date(current_year, 5, 1),
                             'Jun. ':datetime.date(current_year-1, 6, 1),
                             'Jul. ':datetime.date(current_year-1, 7, 1),
                             'Ago. ':datetime.date(current_year-1, 8, 1),
                             'Sep. ':datetime.date(current_year-1, 9, 1),
                             'Oct. ':datetime.date(current_year-1, 10, 1),
                             'Nov. ':datetime.date(current_year-1, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'Jun. .1':datetime.date(current_year, current_month, 1)},inplace = True)  
    elif current_month == 7:
        df.rename(columns = {'Ene. ':datetime.date(current_year, 1, 1),
                             'Feb. ':datetime.date(current_year, 2, 1),
                             'Mar. ':datetime.date(current_year, 3, 1),
                             'Abr. ':datetime.date(current_year, 4, 1),
                             'May. ':datetime.date(current_year, 5, 1),
                             'Jun. ':datetime.date(current_year, 6, 1),
                             'Jul. ':datetime.date(current_year-1, 7, 1),
                             'Ago. ':datetime.date(current_year-1, 8, 1),
                             'Sep. ':datetime.date(current_year-1, 9, 1),
                             'Oct. ':datetime.date(current_year-1, 10, 1),
                             'Nov. ':datetime.date(current_year-1, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'Jul. .1':datetime.date(current_year, current_month, 1)},inplace = True)  
    elif current_month == 8:
        df.rename(columns = {'Ene. ':datetime.date(current_year, 1, 1),
                             'Feb. ':datetime.date(current_year, 2, 1),
                             'Mar. ':datetime.date(current_year, 3, 1),
                             'Abr. ':datetime.date(current_year, 4, 1),
                             'May. ':datetime.date(current_year, 5, 1),
                             'Jun. ':datetime.date(current_year, 6, 1),
                             'Jul. ':datetime.date(current_year, 7, 1),
                             'Ago. ':datetime.date(current_year-1, 8, 1),
                             'Sep. ':datetime.date(current_year-1, 9, 1),
                             'Oct. ':datetime.date(current_year-1, 10, 1),
                             'Nov. ':datetime.date(current_year-1, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'Ago. .1':datetime.date(current_year, current_month, 1)},inplace = True)  
    elif current_month == 9:
        df.rename(columns = {'Ene. ':datetime.date(current_year, 1, 1),
                             'Feb. ':datetime.date(current_year, 2, 1),
                             'Mar. ':datetime.date(current_year, 3, 1),
                             'Abr. ':datetime.date(current_year, 4, 1),
                             'May. ':datetime.date(current_year, 5, 1),
                             'Jun. ':datetime.date(current_year, 6, 1),
                             'Jul. ':datetime.date(current_year, 7, 1),
                             'Ago. ':datetime.date(current_year, 8, 1),
                             'Sep. ':datetime.date(current_year-1, 9, 1),
                             'Oct. ':datetime.date(current_year-1, 10, 1),
                             'Nov. ':datetime.date(current_year-1, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'Sep. .1':datetime.date(current_year, current_month, 1)},inplace = True)  
    elif current_month == 10:
        df.rename(columns = {'Ene. ':datetime.date(current_year, 1, 1),
                             'Feb. ':datetime.date(current_year, 2, 1),
                             'Mar. ':datetime.date(current_year, 3, 1),
                             'Abr. ':datetime.date(current_year, 4, 1),
                             'May. ':datetime.date(current_year, 5, 1),
                             'Jun. ':datetime.date(current_year, 6, 1),
                             'Jul. ':datetime.date(current_year, 7, 1),
                             'Ago. ':datetime.date(current_year, 8, 1),
                             'Sep. ':datetime.date(current_year, 9, 1),
                             'Oct. ':datetime.date(current_year-1, 10, 1),
                             'Nov. ':datetime.date(current_year-1, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'Oct. .1':datetime.date(current_year, current_month, 1)},inplace = True)  
    elif current_month == 11:
        df.rename(columns = {'Ene. ':datetime.date(current_year, 1, 1),
                             'Feb. ':datetime.date(current_year, 2, 1),
                             'Mar. ':datetime.date(current_year, 3, 1),
                             'Abr. ':datetime.date(current_year, 4, 1),
                             'May. ':datetime.date(current_year, 5, 1),
                             'Jun. ':datetime.date(current_year, 6, 1),
                             'Jul. ':datetime.date(current_year, 7, 1),
                             'Ago. ':datetime.date(current_year, 8, 1),
                             'Sep. ':datetime.date(current_year, 9, 1),
                             'Oct. ':datetime.date(current_year, 10, 1),
                             'Nov. ':datetime.date(current_year-1, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'Nov. .1':datetime.date(current_year, current_month, 1)},inplace = True)
    elif current_month == 12:
        df.rename(columns = {'Ene. ':datetime.date(current_year, 1, 1),
                             'Feb. ':datetime.date(current_year, 2, 1),
                             'Mar. ':datetime.date(current_year, 3, 1),
                             'Abr. ':datetime.date(current_year, 4, 1),
                             'May. ':datetime.date(current_year, 5, 1),
                             'Jun. ':datetime.date(current_year, 6, 1),
                             'Jul. ':datetime.date(current_year, 7, 1),
                             'Ago. ':datetime.date(current_year, 8, 1),
                             'Sep. ':datetime.date(current_year, 9, 1),
                             'Oct. ':datetime.date(current_year, 10, 1),
                             'Nov. ':datetime.date(current_year, 11, 1),
                             'Dic. ':datetime.date(current_year-1, 12, 1),                     
                             'Dic. .1':datetime.date(current_year, current_month, 1)},inplace = True) 
    
    df.rename(columns = {'Denominación':'Denominacion'},inplace = True)
    df = df.set_index('Codigos')
    print(df.columns)
    #current_month = df.columns[len(df.columns)-1].month
    #current_date = datetime.date(current_year, current_month, 1)
    meses_rot = media
    meses_rot_name = 'Media-'+str(meses_rot)
    df[meses_rot_name] = 0
    
    for col in df.columns[(df.shape[1]-(meses_rot+1)):df.shape[1]-1]:
        df[meses_rot_name] += df[col]
    
    df[meses_rot_name] = (df[meses_rot_name]/meses_rot).astype(int)
    
    
    print(df.dtypes)
    print(df.head())
    print(df.describe())
    
    #df[meses_rot_name] = ((df.iloc[:,df.shape[1]-3] + df.iloc[:,df.shape[1]-4] + df.iloc[:,df.shape[1]-5]) /3 ).astype(int)
    df = df[df[meses_rot_name] > df['Exist']]
    df['Pedido'] = df[meses_rot_name] - df['Exist']
    df.sort_values(by=['Pedido'], inplace=True, ascending=False)
    
    
    print(df)
    
    # return the transpose 
    df_transpose = df.drop(columns=['Denominacion','Exist','Pedido',meses_rot_name]).transpose()
    df_transpose.drop(df_transpose.tail(1).index,inplace=True)
    
    # Print the result 
    print(df_transpose.dtypes) 
    print(df_transpose.head())
    print(df_transpose.describe())
    
    
    #df.to_excel (r'Pedido.xlsx', index =True, header=True)
    
    #♦df['Nombre-Corto'] = df.Denominacion.str.replace('/', ' ').str.split().str.get(0)
    
    # Use seaborn style defaults and set the default figure size
    #sns.set(rc={'figure.figsize':(11, 4)})
    
    from pathlib import Path
    Path("./fig").mkdir(parents=True, exist_ok=True)
    
    
    df['url'] =''
    df['trend'] = 0
    #df['Google'] = 0
    
    
    def trendline(index,data, order=1):
        coeffs = np.polyfit(index, list(data), order)
        slope = coeffs[-2]
        return float(slope)
    
    #pytrends = TrendReq(hl='es-ES', tz=1)
    #timeframe = df_transpose.index[0].strftime("%Y-%m-%d") + ' ' + df_transpose.index[len(df_transpose)-1].strftime("%Y-%m-%d")
    
    len_index = len(df.index)
    i = 1
    for ind in df.index:
        
        name = df.loc[ind, 'Denominacion']
        df_transpose_rolling = df_transpose[ind].rolling(4, center=True).mean().dropna()
        
        #plt.figure()
        plt.plot(df_transpose_rolling,marker='.', linestyle='-', linewidth=0.5, label=name)    
        plt.xlabel('Fecha')
        plt.ylabel('Rotación')    
        plt.legend([df.loc[ind, 'Denominacion']])
        
        
        #fig, axes = plt.subplots(2, 1, figsize=(11, 10), sharex=True)
        #axes[0].plot(df_transpose_rolling,marker='.', linestyle='-', linewidth=0.5, label=name)    
        #axes[0].set_xlabel('Fecha')
        #axes[0].set_ylabel('Rotación')    
        #axes[0].legend([df['Nombre'][ind]])
            
        df_trend = df_transpose_rolling.iloc[(len(df_transpose_rolling)-meses_rot):len(df_transpose_rolling)]
        print(df_trend)
        df.loc[ind, 'url'] = '=HYPERLINK("'+'./fig/'+ind+'.svg","'+name+'")'
        df.loc[ind,'trend'] = trendline(list(range(0, len(df_trend))),df_trend)
    
        #pytrends.build_payload(name, cat=0, timeframe=timeframe, geo='ES', gprop='')
        #df_trends = pytrends.interest_over_time()
        #df.loc[ind,'Google'] = df_trends[name].mean()    
        #axes[1].plot(df_trends[df['Nombre-Corto'][ind]],marker='.', linestyle='-', linewidth=0.5, label=df['Nombre-Corto'][ind])    
        #axes[1].set_xlabel('Fecha')
        #axes[1].set_ylabel('Búsqueda')    
        #axes[1].legend([df['Nombre'][ind]])    
        plt.savefig('./fig/'+ind+'.svg')
        plt.clf()
        progress_bar['value'] = (i*100)/len_index
        window.update_idletasks() 
        i+=1
    
    cols = list(df)
    # move the column to head of list using index, pop and insert
    cols.insert(len(cols)-3, cols.pop(cols.index('Exist')))
    print(cols)


            
    try:
        df[cols].to_excel (r'Pedido.xlsx', index =True, header=True)    
    except xlsxwriter.exceptions.FileCreateError:            
        tk.messagebox.showinfo("Permiso denegado", "No se puede escribir el fichero de pedido.")
    btn_open_file['state'] = 'normal'
        
        
def open_file():
    """Open a file for editing."""
    filepath = filedialog.askopenfilename(initialdir = "./",title = "Seleccionar archivo",filetypes = (("xls files","*.xls"),("xlsx files","*.xlsx")))
    if not filepath:
        return

    window.title(filepath)
    lbl_url["text"] = filepath
    btn_generate['state'] = 'normal'
    

def callback():
    os.system("start ./Pedido.xlsx")

window = tk.Tk()
window.title("Generador de Pedidos de Unycop")

window.rowconfigure(0, minsize=200, weight=1)
window.columnconfigure(0, minsize=100, weight=1)

#♣window.resizable(width=False, height=False)
 
#url =  filedialog.askopenfilename(initialdir = "/",title = "Seleccionar archivo",filetypes = (("xls files","*.xls"),("xlsx files","*.xlsx")))

frame = tk.Frame(master=window, borderwidth=1)

lbl_url = tk.Label(master=frame, text='')
lbl_url.grid(row=0, column=0, pady=2, sticky="nsew")

btn_open = tk.Button(master=frame, text="Abrir listado", command=open_file)
btn_open.grid(row=1, column=0, pady=2, sticky="nsew")

scale_entry = tk.Scale(master=frame,orient='horizontal', from_=2, to=12,resolution=1,width=25,label='Media de meses:')
scale_entry.set(3)
scale_entry.grid(row=2, column=0, pady=2, sticky="nsew")

#btn_generate = tk.Button(master=frame, text="Generar Pedido", command= lambda:generate(url,scale_entry.get()))
btn_generate = tk.Button(master=frame, text="Generar Pedido", command= lambda:generate(window.title(),scale_entry.get()))
btn_generate.grid(row=3, column=0, pady=2, sticky="nsew")
btn_generate['state'] = 'disable'

btn_open_file = tk.Button(master=frame, text="Abrir Pedido", command= lambda:callback())
btn_open_file.grid(row=4, column=0, pady=2, sticky="nsew")
btn_open_file['state'] = 'disable'

# Progress bar widget 
progress_bar = ttk.Progressbar(master=frame, orient = 'horizontal', length = 100, mode = 'determinate')
progress_bar.grid(row=5, column=0, pady=2, sticky="nsew")

frame.grid(row=0, column=0, padx=5, pady=5)

# Run the application
window.mainloop()