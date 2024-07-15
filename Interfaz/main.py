import tkinter as tk
from tkinter import ttk, messagebox
import os
import openpyxl
from tkcalendar import DateEntry
import customtkinter
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Patch
import pandas as pd
import networkx as nx
from gantt import Gantt_button_clicked
from cpm import PERT_CPM_button_clicked
from PIL import Image
import gantt
import cpm
from PIL import Image, ImageTk
import customtkinter


def Ventana_Principal():
    #Crear la figura y los ejes 
    global fig, ax
    fig, ax = plt.subplots(figsize=(10, 6))

    # Frame principal
    principal = tk.Tk()
    principal.title("")
    principal.resizable(False,False)

    #Añadir frame principal a customtkinter
    frameprincipal = customtkinter.CTkFrame(principal,fg_color="#E5E5E5")
    frameprincipal.pack()
    principal.eval('tk::PlaceWindow . center')

    #Frame de contenido 
    contenido_frame = customtkinter.CTkFrame(frameprincipal,fg_color="#D9D9D9")
    contenido_frame.grid(row=0, column=0,padx=20, pady=10)

    #Etiqueta del título
    task_label1 = customtkinter.CTkLabel(contenido_frame, text="Bienvenido a GPC ", font=('Helvetica', 40, 'bold'),text_color="black")
    task_label1.grid(row=0, column=0, padx=20, pady=10)

    #Etiqueta Gantt
    task_label = customtkinter.CTkLabel(contenido_frame, text="Selecciona el tipo de metodo", font=('Helvetica', 30),text_color="black")
    task_label.grid(row=1, column=0, padx=20, pady=10) # Centrar la etiqueta

    #Frame de botones
    botones_frame = customtkinter.CTkFrame(frameprincipal,fg_color="#D9D9D9")
    botones_frame.grid(row=1, column=0, pady=(20, 20), padx=20)  # Modificado pady

    #Botón para Gantt
    buttonG = customtkinter.CTkButton(botones_frame, text="Gantt", height=40, font=('Helvetica', 20, 'bold'),fg_color='#3A7EBF', command=lambda: [principal.destroy(), Gantt_button_clicked()])  # Se llama función Gantt_button_clicked
    buttonG.grid(row=0, column=0, padx=20, pady=(10, 10))  # Modificado pady

    #Botón para PERT
    buttonP = customtkinter.CTkButton(botones_frame, text="PERT- CPM", height=40, font=('Helvetica', 20, 'bold'), fg_color='#3A7EBF',command=lambda: [principal.destroy(), PERT_CPM_button_clicked()])  # Se llama función PERT_CPM_button_clicked
    buttonP.grid(row=0, column=2, padx=20, pady=(10, 10))  # Modificado pady

    #Botón para salir
    buttonSalir = customtkinter.CTkButton(botones_frame, text="Salir", height=40, font=('Helvetica', 20, 'bold'), fg_color='red', hover_color='#B80507', command=principal.destroy)
    buttonSalir.grid(row=1, column=1, padx=20, pady=(10, 10))  # Modificado pady


    principal.mainloop()

if __name__ == "__main__":
    # Mostrar la interfaz principal y no la de editar
    Ventana_Principal()
