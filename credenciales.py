#------------------------------------------------------------
#Creado por Cristhian Guzmán                                |
#cgp2409@gmail.com | 318 689 9502 | linkedin.com/in/cgp2409 |
#------------------------------------------------------------

#librerias
import tkinter as tk
from tkinter import messagebox
import sys

#Mensaje de advertencia
root=tk.Tk()
root.geometry("300x100")

titulo = "ADVERTENCIA"
mensaje = """Antes de continuar no deben existir archivos 
de excel abiertos, estos se pueden cerrar
perdiendo toda la información sin guardar
                 ¿Desea continuar?     """
#variable continuar es booleano 
continuar = messagebox.askyesno(message=mensaje, title=titulo)
root.destroy()  
if continuar == False:
    sys.exit()



#Preguntar por credenciales necesarias para las macros de ventas
root=tk.Tk()

root.geometry("300x100") #tamaño de la ventana
usuario = tk.StringVar()
contraseña = tk.StringVar()
 
  
# definir funcion del boton
def submit():
 
    global usuario
    global contraseña
    usuario = usuario.get()
    contraseña = contraseña.get()
    root.destroy()
     
#Hacer el texto del primer cuadro de entrada
texto_usuario = tk.Label(root, text = 'Usuario', 
                      font=('calibre',10, 'bold'))  
#Hacer el primer cuadro de entrada
texto = tk.StringVar()
texto.set("cobguzman")
usuario = tk.Entry(root, font=('calibre',10,'normal'),
                   textvariable = texto)
  
#Hacer el texto del segundo cuadro de entrada
texto_contraseña = tk.Label(root, text = 'Contraseña', 
                       font = ('calibre',10,'bold'))
#Hacer el segundo cuadro de entrada
texto = tk.StringVar()
texto.set("$Febrero23")
contraseña = tk.Entry(root, font = ('calibre',10,'normal'), 
                      textvariable = texto, show = '*')
  
# Crear el boton
boton = tk.Button(root,text = 'Aceptar', command = submit)

# ubicación de cada texto y cuadro de entrada
texto_usuario.grid(row=0,column=0)
usuario.grid(row=0,column=1)
texto_contraseña.grid(row=1,column=0)
contraseña.grid(row=1,column=1)
boton.grid(row=2,column=1)
  
#loop infinito pra mostrar la ventana
root.mainloop()



