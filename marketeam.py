#------------------------------------------------------------
#Creado por Cristhian Guzmán                                |
#cgp2409@gmail.com | 318 689 9502 | linkedin.com/in/cgp2409 |
#------------------------------------------------------------

#Importación de librerias
import pandas as pd
import easygui
import numpy as np
import xlwings as xw
import ctypes  
from tqdm import tqdm

#crear una barra de progreso
pbar = tqdm(total=100)




#Mensajes y datos iniciales
#Aquí se utiliza otro scrip llamado credenciales
import credenciales as cs
usuario = cs.usuario
contraseña = cs.contraseña

#---CARGAR MAESTRA DE CLIENTES---

ctypes.windll.user32.MessageBoxW(0, "Seleccione la maestra de clientes manualmente", "", 0)
ruta = easygui.fileopenbox(msg="Seleccione la maestra de clientes") #seleccionar manualmente
df_maestra_clientes = pd.read_excel(ruta,engine='openpyxl',
                                    skiprows=1, usecols=list(range(1,48)),
                                    dtype = str)
df_maestra_clientes.sort_values("Cliente") #ordenar por código de clientes
df_maestra_clientes['Observacion'] = df_maestra_clientes['Observacion'].str.upper()


#---CARGAR ARCHIVO INFO---
df_info = pd.read_excel( 'info.xlsx', sheet_name='codigos',
                        dtype = str)
df_zapatocas = pd.read_excel( 'info.xlsx', sheet_name='zapatocas',
                        dtype = str)

pbar.update(10) #Actualizar barra de progreso

#Generar lista y excluir clientes por bloqueo
bloqueados_excluir = pd.DataFrame(df_info["Clientes a excluir por bloqueo"].values)
bloqueados_excluir = bloqueados_excluir.dropna()
bloqueados_excluir = bloqueados_excluir[0].values.tolist()
df_maestra_clientes = df_maestra_clientes.loc[~df_maestra_clientes["Bloqueo Clientes Pedido"].isin(bloqueados_excluir)]

#Generar lista de códigos de población a excluir 
poblacion_excluir = pd.DataFrame(df_info["Poblaciones a excluir"].values)
poblacion_excluir = poblacion_excluir.dropna()
poblacion_excluir = poblacion_excluir[0].values.tolist() 

#Generar lista de códigos de Nit a excluir 
nit_excluir = pd.DataFrame(df_info["Nit a excluir"].values)
nit_excluir = nit_excluir.dropna()
nit_excluir = nit_excluir[0].values.tolist() 

#Generar lista de códigos de tipologias seleccionadas
tip_selec = pd.DataFrame(df_info["Tipologias seleccionadas"].values)
tip_selec = tip_selec.dropna()
tip_selec = tip_selec[0].values.tolist() 

#Generar lista de códigos de of ventas a excluir 
ofventas_excluir = pd.DataFrame(df_info["of ventas a excluir"].values)
ofventas_excluir = ofventas_excluir.dropna()
ofventas_excluir = ofventas_excluir[0].values.tolist() 


pbar.update(10) #Actualizar barra de progreso

#lista de drogerias a excluir
drog_excluir = ['DC','DD']

 
df_final = df_maestra_clientes.copy(deep=True) #crear maestra final
df_final['Codigo Postal'] = df_final['Codigo Postal'].str[7:] #seleccionar solo una parte de código postal
df_final['Num Ident Fiscal1'] = df_final['Num Ident Fiscal1'].astype("string")
df_final["Cliente"] = pd.to_numeric(df_final["Cliente"])
lista_columnas = df_final.columns.tolist()
for i in lista_columnas:
    df_final[i]=df_final[i].astype('string')

#Renombrar columnas
dict = {'Cliente': 'Nºcliente','Nombre1 Cliente': 'Nombre 1',
        'Observacion': 'Coment referentes di','Nombre1 Cliente': 'Nombre 1',
        'Num Ident Fiscal1': 'N ident f','Poblacion': 'Población',
        'Num Calle': 'Calle y nº','Telefono1': 'Núm teléf',
        'Canal': 'Canal dist','Subcanal': 'Grup clien',
        'Oficina Ventas': 'Ofic venta','Num Z1': 'Z1 Núm person',
        'Nomb Z1': 'Z1 Nombre Vendedor 1','Cedula Z1': 'Z1 Cédula Vendedor 1',
        'Num ZA': 'ZA Núm person','Nomb ZA': 'ZA Nombre Vendedor A',
        'Cedula ZA': 'ZA Cédula Vendedor A','Num Z6': 'Z6 Núm person',
        'Nomb Z6': 'Z6 Nombre Jefe de Zona','Cedula Z6': 'Z6 Cédula Jefe de Zona',
        'Num Y9': 'Y9 Núm person','Nomb Y9': 'Y9 Nombre Coord. Pto. Vta.',
        'Cedula Y9': 'Y9 Cédula Coord. Pto. Vta.','Num Y3': 'Y3 Núm person',
        'Nomb Y3': 'Y3 Nombre Mercaderista 1','Cedula Y3': 'Y3 Cédula Mercaderista 1',
        'Num Y3 - 2': 'Y3 Núm person 1','Nomb Mercaderista 2': 'Y3 Nombre Mercaderista 2',
        'Cedula Y3 - 2': 'Y3 Cédula Mercaderista 2','Num Y3 - 3': 'Y3 Núm person 2',
        'Nomb Mercaderista 3': 'Y3 Nombre Mercaderista 3','Cedula Y3 - 3': 'Y3 Cédula Mercaderista 3',
        'Num Y3 - 4': 'Y3 Núm person 3','Nomb Mercaderista 4': 'Y3 Nombre Mercaderista 4',
        'Cedula Y3 - 4': 'Y3 Cédula Mercaderista 4','Num Y8': 'Y8 Núm person 4',
        'Nomb Y8': 'Y8 Nombre Coord. Trade','Cedula Y8': 'Y8 Cédula Coord. Trade',
        'Coordenadas Lugar': 'Cód. Loc.'}    
df_final.rename(columns=dict,inplace=True)
 
pbar.update(10) #Actualizar barra de progreso

#Agregar información sobe oxxos
df_oxxo = df_info[['Num cliente','Clientes Oxxo']]
df_oxxo = df_oxxo.dropna()
df_oxxo.rename(columns = {'Num cliente':'Nºcliente'}, inplace = True)
df_final = pd.merge(left=df_final,right=df_oxxo, how='left')

#Agregar nueva columnas
nuevas_col = ['Portafolio Clave','Can_Estruct','Subc_Estruc',
              'Socio N.','Modelo Atención']
for i in nuevas_col: #llenar las columnas nuevas con NA
    df_final[i] = np.nan 
lista_columnas = df_final.columns.tolist()
for i in lista_columnas:
    df_final[i]=df_final[i].astype('string')

#llenar columna Modelo atención con la palabra "Directa"
df_final['Modelo Atención'] = "Directa"

#Excluir registros
df_final = df_final.loc[df_final.Tipologia.isin(tip_selec)]

df_final = df_final.loc[~df_final["N ident f"].isin(nit_excluir)]
df_final = df_final.loc[~df_final["Codigo Postal"].isin(poblacion_excluir)]
df_final = df_final.loc[~df_final["Grup clien"].isin(drog_excluir)]
df_final = df_final.loc[~df_final["Of Ventas"].isin(ofventas_excluir)]

#llenar columna segmento
df_segmento = pd.DataFrame(df_info[['Tipologia','Segmento']])
df_segmento = df_segmento.dropna()
df_final = pd.merge(left=df_final,right=df_segmento, how='left')


#llenar columna segmento vital
df_segmento_vital = pd.DataFrame(df_info[["Grp4 Cliente","Segmento Vital"]])
df_segmento_vital = df_segmento_vital.dropna()
df_final = pd.merge(left=df_final,right=df_segmento_vital, how='left')


pbar.update(10) #Actualizar barra de progreso


#Remover columnas inncesarias y dar el orden requerido
df_final = df_final.drop(['Grp Cta Deudor','Tipologia','Grp4 Cliente', 'Of Ventas'], axis=1)

orden = ['Nºcliente', 'Cód. Loc.', 'Nombre 1', 'Coment referentes di',
                     'N ident f','Codigo Postal', 'Población', 'Calle y nº', 'Núm teléf', 'Distrito',
                     'Canal dist', 'Grup clien', 'Ofic venta', 'Portafolio Clave',
                     'Can_Estruct', 'Subc_Estruc','Segmento',  'Socio N.','Z1 Núm person',
                     'Z1 Nombre Vendedor 1', 'Z1 Cédula Vendedor 1', 'ZA Núm person',
                     'ZA Nombre Vendedor A', 'ZA Cédula Vendedor A', 'Z6 Núm person',
                     'Z6 Nombre Jefe de Zona', 'Z6 Cédula Jefe de Zona', 'Y9 Núm person',
                     'Y9 Nombre Coord. Pto. Vta.','Y9 Cédula Coord. Pto. Vta.',
                     'Y3 Núm person', 'Y3 Nombre Mercaderista 1','Y3 Cédula Mercaderista 1',
                     'Y3 Núm person 1', 'Y3 Nombre Mercaderista 2',
                     'Y3 Cédula Mercaderista 2', 'Y3 Núm person 2',
                     'Y3 Nombre Mercaderista 3', 'Y3 Cédula Mercaderista 3',
                     'Y3 Núm person 3', 'Y3 Nombre Mercaderista 4',
                     'Y3 Cédula Mercaderista 4', 'Y8 Núm person 4', 'Y8 Nombre Coord. Trade',
                     'Y8 Cédula Coord. Trade', 'Coordenada X', 'Coordenada Y', 'Modelo Atención',
                     'Clientes Oxxo', 'Segmento Vital']

df_final = df_final[orden]

#Agregar zapatocas como nuevos registros
df_final = pd.concat([df_final, df_zapatocas])


#Generar lista de los números de clientes de los exito express a excluir
df_exitos = df_final[df_final["N ident f"].str.contains("8909006089") == True]
df_exitos = df_exitos.reset_index()
indices = df_exitos.index.values.tolist()
df_exitos = df_exitos.drop(['index'], axis=1)
exitos_excluir = []
for i in indices:
    a ="EXPRESS" in df_exitos.iloc[i,3]
    b = "EXITO" in df_exitos.iloc[i,3]
    c = "ÉXITO" in df_exitos.iloc[i,3]
    d = df_exitos.iloc[i,1]
    if (b or c) and a:
        exitos_excluir.append(d)
 
pbar.update(10) #Actualizar barra de progreso    
        
#Generar lista de los números de clientes de los turbo carulla a excluir
df_turbo = df_final[df_final["N ident f"].str.contains("8909006089") == True]
df_turbo = df_turbo.reset_index()
indices = df_turbo.index.values.tolist()
turbo_excluir = []
for i in indices:
    a ="TURBO" in df_turbo.iloc[i,3]
    b = "CARULLA" in df_turbo.iloc[i,3]
    d = df_turbo.iloc[i,1]
    if b and a:
        turbo_excluir.append(d)


#Generar lista de vendedores a excluir
vendedores_excluir = pd.DataFrame(df_info["Vendedores a excluir"].values)
vendedores_excluir = vendedores_excluir.dropna()
vendedores_excluir = vendedores_excluir[0].values.tolist() 

#Generar lista de clientes a excluir
clientes_excluir = pd.DataFrame(df_info["Clientes a excluir"].values)
clientes_excluir = clientes_excluir.dropna()
clientes_excluir = clientes_excluir[0].values.tolist()
       
#Excluir registros
df_final = df_final.loc[~df_final["Nºcliente"].isin(exitos_excluir)]
df_final = df_final.loc[~df_final["Nºcliente"].isin(turbo_excluir)]
df_final = df_final.loc[~df_final["ZA Núm person"].isin(vendedores_excluir)]
df_final = df_final.loc[~df_final["Z1 Núm person"].isin(vendedores_excluir)]
df_final = df_final.loc[~df_final["Nºcliente"].isin(clientes_excluir)]

pbar.update(10) #Actualizar barra de progreso

#Generar DataFrame para llenar el portafolio clave
df_portafolio = pd.DataFrame(df_info[["Concatenado","Portafolio Clave"]].values)
df_portafolio = df_portafolio.dropna()

#Rellenar columna de Portafolio Clave !!!!!REVISAR NO SE COMPLETAN TODOS!!!
df_final = df_final.reset_index()
df_final = df_final.drop(['index'], axis=1)
indices = df_final.index.values.tolist()
for i in indices:
    OficVenta = df_final.iloc[i,12]
    Poblacion = df_final.iloc[i,6]
    concat_final = OficVenta + Poblacion
    df = df_portafolio.loc[df_portafolio[0] == concat_final]
    if df.empty==False:
        portafolio = str(df[1].iloc[0])
        df_final.iloc[i,13] = portafolio
    else:
        df_final.iloc[i,13] = "Error raro"

#Generar lista de clientes para estructura de canal y subcanal
vendedores_estructura = pd.DataFrame(df_info["ZA-Z1 Núm person"].values)
vendedores_estructura = vendedores_estructura.dropna()
vendedores_estructura = vendedores_estructura[0].values.tolist() 

#Rellenar columna de can_struc u subc_struc !!!!!REVISAR NO SE COMPLETAN TODOS!!!
df_estructura = pd.DataFrame(df_info[["ZA-Z1 Núm person","Can_Estruct","Subc_Estruc"]].values)
for i in indices:
    ZAvalor = df_final.iloc[i,21]
    Z1valor = df_final.iloc[i,18]
    if pd.isna(ZAvalor):
        ZAexiste = False
    else:
        ZAexiste = df_final.iloc[i,21] in vendedores_estructura
    if pd.isna(Z1valor):
        Z1existe = False
    else:
        Z1existe = df_final.iloc[i,18] in vendedores_estructura 

    if ZAexiste:
        df = df_estructura.loc[df_estructura[0] == ZAvalor] 
        if df.empty==False:
            can_est = str(df[1].iloc[0])
            sub_est = str(df[2].iloc[0])
            df_final.iloc[i,14] = can_est
            df_final.iloc[i,15] = sub_est 
        else:
            df_final.iloc[i,14] = "Error raro df vacio ZA"
            df_final.iloc[i,15] = "Error raro df vacio ZA" 
    elif Z1existe:  
        df = df_estructura.loc[df_estructura[0] == Z1valor] 
        if df.empty==False:
            can_est = str(df[1].iloc[0])
            sub_est = str(df[2].iloc[0])
            df_final.iloc[i,14] = can_est      
            df_final.iloc[i,15] = sub_est 
        else:
            df_final.iloc[i,14] = "Error raro df vacio Z1"
            df_final.iloc[i,15] = "Error raro df vacio Z1" 
    else:
        df_final.iloc[i,14] = "ERORR RARO no hay ZA ni Z1" 
        df_final.iloc[i,15] = "ERORR RARO no hay ZA ni Z1" 

pbar.update(10) #Actualizar barra de progreso

#socios no plus
df_socios_Noplus = pd.DataFrame(df_info[["Socios No Plus","Segmento socios"]].values) #llamar dataframe de info
df_socios_Noplus.rename(columns = {0:'Nºcliente',1:'Socio N.'}, inplace = True) #cambiar nombres de columnas
df_socios_Noplus.loc[df_socios_Noplus['Socio N.'] == 'DR', 'Socio N.'] = 'TD' #reemplazar todo DR pot TD
sociosAU = df_socios_Noplus[df_socios_Noplus['Socio N.'].str.contains("AU", case=False)] #df que de todo lo que tenga AU
socios_noplus = df_socios_Noplus['Nºcliente'].values.tolist() #lista de clientes socios no plus 
sociosAU = sociosAU['Nºcliente'].values.tolist() #lista de socios que tienen AU
df_socios_Noplus.loc[df_socios_Noplus['Nºcliente'].isin(sociosAU), 'Socio N.'] = "AU" #remplazar los que tienen AU + otras letras, solo por palara AU

for i in indices: #poner SocioN en el df final
    cliente = df_final.iloc[i,0]
    if pd.isna(cliente):
        cliente_existe = False
    else:
        cliente_existe = df_final.iloc[i,0] in socios_noplus
    if cliente_existe:
        df = df_socios_Noplus.loc[df_socios_Noplus['Nºcliente'] == cliente] 
        if df.empty==False:
            tipo = str(df['Socio N.'].iloc[0])
            df_final.iloc[i,17] = tipo
        else:
            df_final.iloc[i,17] = "Error raro"   
    else:
        df_final.iloc[i,17] = "No"

#se sacan todos los turbo que no esten en la población turbo 
#eliminar ktronix por nombre
df_final[df_final.columns[3]] = df_final[df_final.columns[3]].str.upper()
df_final[df_final.columns[6]] = df_final[df_final.columns[6]].str.upper()
lista_eliminar =[]
for i in indices:
    empresa = df_final.iloc[i,3]
    lugar = df_final.iloc[i,6]
    empresaktronix = "KTRONIX" in empresa
    empresa = "TURBO" in empresa
    lugar = "TURBO" in lugar
    if empresa and (lugar is False):
        lista_eliminar.append(df_final.iloc[i,0])
    if empresaktronix:
        lista_eliminar.append(df_final.iloc[i,0])
df_final = df_final.loc[~df_final["Nºcliente"].isin(lista_eliminar)]


#eliminar todo lo que sea express o drogeria y que pertenescan a cadenas (valor S en clumna "dist")
df_final = df_final.reset_index()
indices = df_final.index.values.tolist()
df_final = df_final.drop(['index'], axis=1)
lista_eliminar =[] 
df_final[df_final.columns[10]] = df_final[df_final.columns[10]].str.upper()
for i in indices:
    es_cadenas = df_final.iloc[i,10] == "S"
    palabraclave = df_final.iloc[i,3]
    es_palabra = (("DROG" in palabraclave) or ("EXPRESS" in palabraclave))
    if es_cadenas and es_palabra:
        lista_eliminar.append(df_final.iloc[i,0])
df_final = df_final.loc[~df_final["Nºcliente"].isin(lista_eliminar)]

#-----Ejecutar macro y cargar información de ventas directa-----

#Generar input de meses de ventas a evaluar para las macros
meses_ventas = pd.DataFrame(df_info["meses ventas"].values)
meses_ventas = meses_ventas.dropna()
meses_ventas = meses_ventas[0].values.tolist() 
meses = ""
for i in meses_ventas:
    meses = meses + i + ";"
meses = meses[:-1]    

#Información de ventas Directa
LibroVentasI = xw.Book("ventas Directa.xlsm") #abrir libro
macro = LibroVentasI.macro("principal.ActivarAnalysis") #referenciar macro 
xl = xw.apps.active.api
macro(meses,usuario,contraseña) #ejecutar macro referenciada
LibroVentasI.save() #guardar
LibroVentasI.close() #cerrar 
xl.Quit() #salir de excel


pbar.update(10) #Actualizar barra de progreso
    
#cargar información de la consulta de analysis
columnas = ["oficina de ventas","cod cliente","nom cliente"] + meses_ventas
ventasD_cliente = pd.read_excel( 'ventas Directa.xlsm', names=columnas,
                             sheet_name='consulta',skiprows=2)


#Crear columna promedio de ventas, ignorando valores nulos
ventasD_cliente["promedio"] = ventasD_cliente.loc[:, meses_ventas].mean(axis = 1, skipna = True) 
#Unir el promedio al df final
ventasD_cliente.rename(columns = {'cod cliente':'Nºcliente'}, inplace = True)
df_final = pd.merge(left=df_final,right=ventasD_cliente.loc[:,['Nºcliente','promedio']], how='left')

#Crear columna clientes desarrolados
df_desarrollados = pd.DataFrame(df_info['Clientes desarrollados'])
df_desarrollados = df_desarrollados.dropna()
df_desarrollados.rename(columns = {'Clientes desarrollados':'Nºcliente'}, inplace = True)
df_desarrollados['Desarrollados'] = "Si"
df_final = pd.merge(left=df_final,right=df_desarrollados, how='left')
df_final['Desarrollados'] =  df_final['Desarrollados'].fillna('No')


#---Armar df ordenado por socios y ventas---
#listas de vendedores a los cuales se les encuentra los clientes aptos
codigosZA = df_final['ZA Núm person'].values.tolist() 
codigosZA = [x for x in codigosZA if pd.isnull(x) == False]
codigosZA = list(set(codigosZA)) #eliminar duplicados de lista
#en caso de no tener vendedor ZA se busca el Z1
indicesz1 = df_final[df_final['ZA Núm person'].isnull()].index.tolist()
codigosZ1 = df_final.iloc[indicesz1] 
codigosZ1 = codigosZ1['Z1 Núm person'].values.tolist() 
codigosZ1 = [x for x in codigosZ1 if pd.isnull(x) == False]
codigosZ1 = list(set(codigosZ1))
df_directa = pd.DataFrame(columns=df_final.columns)
topclientes = int(df_info['Top de clientes'].iloc[0])
for i in codigosZA:
    df = df_final.loc[df_final['ZA Núm person'] == i]
    df['Socio N.'] = df['Socio N.'].replace({'No': 'AAAAA'})
    df = df.sort_values(['Socio N.', 'promedio'],ascending=False)
    df['Socio N.'] = df['Socio N.'].replace({'AAAAA':'No'})
    if df.shape[0] > topclientes:
        df = df.head(topclientes)
    df_directa = pd.concat([df_directa, df])
for i in codigosZ1:
    df = df_final.loc[df_final['Z1 Núm person'] == i]
    df['Socio N.'] = df['Socio N.'].replace({'No': 'AAAAA'})
    df = df.sort_values(['Socio N.', 'promedio'],ascending=False)
    df['Socio N.'] = df['Socio N.'].replace({'AAAAA':'No'})
    if df.shape[0] > topclientes:
        df = df.head(topclientes)
    df_directa = pd.concat([df_directa, df])

pbar.update(10) #Actualizar barra de progreso

    
#---Exportar maestra de la directa---
df_directa.rename(columns = {'promedio':'Ventas'}, inplace = True)
df_directa.to_excel("Maestra marketeam directa.xlsx",index=False)

pbar.update(10) #Actualizar barra de progreso

