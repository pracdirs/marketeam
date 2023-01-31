# marketeam
Material necesario para generar la maestra de clientes potenciales a visitar de la DIRECTA.

Ejecutable final:

-marketeam.exe: ejecutable .exe con la capacidad de funcionar si necesidad de tener instalado python o paqueterias en el computador, solo windows.

Scripts:

-ventas Directa.xlsm: Se utiliza VBA, el cual toma tres parámetros de entrada proporcionados por el usuario, credenciales y fechas evaluadas, genera información de utilidad para el ordenamiento de los resultados finales, esta información se obtiene con una consulta usando SAP Analysis for Microsoft Office

-credenciales.py: Escrito en Python, utiliza la librería tkinter para crear una interfaz gráfica de usuario (GUI) en la que se presentan advertencias para el usuario. Si este decide continuar, se le pide que ingrese sus credenciales de usuario y contraseña. Estas credenciales son almacenadas como variables globales, luego se utiliza esta información para autenticarse en la plataforma de SAP. También se controla el comportamiento de refresco de datos y se establecen filtros para la plataforma de SAP. 

-marketeam.py: Escrito en Python, este es el código principal que importa varias librerías, entre ellas "pandas" que es una librería de análisis de datos, "easygui" que es una librería para crear interfaces gráficas de usuario fácilmente, "numpy" que es una librería para cálculo científico y manipulación de datos, "xlwings" que es una librería para interactuar con Microsoft Excel en Python, "ctypes" que es una librería para llamar funciones de sistema en C desde Python y "tqdm" es una librería para mostrar una barra de progreso en un bucle de Python. En resumen, el script completo en un inicio importa datos de archivos de Excel, el archivo info y la maestra de clientes, ejecuta los otros dos scrips, controla totalmente el archivo con la macro e importa la información, realiza operaciones de limpieza y filtrado de datos, y crea nuevos dataframes y listas a partir de los datos importados. Utiliza estos nuevos datos para actualizar los valores de columnas específicas en otro dataframe, y excluye ciertas filas del dataframe en base a los valores de columnas específicas, por último, los datos de interés se exportan en un archivo de Excel listos para su uso en el siguiente proceso. Además, brinda retroalimentación al usuario sobre el progreso del script utilizando una barra de progreso.


Fuentes de información de entrada:

-P_DQ_Maestro_Maestra_Piloto: Maestra de cientes, debe contar con las siguientes columnas, en el siguiente ordeny nombres
Grp Cta Deudor,	Cliente,	Coordenadas Lugar,	Nombre1 Cliente,	Observacion,	Num Ident Fiscal1,	Mn T Clientes.Codigo Postal,	Poblacion,	Num Calle,	Telefono1,	Distrito,	Canal,	Subcanal,	Tipologia,	Grp4 Cliente,	Of Ventas,	Oficina Ventas,	Num Z1,	Nomb Z1,	Cedula Z1,	Num ZA,	Nomb ZA,	Cedula ZA,	Num Z6,	Nomb Z6,	Cedula Z6,	Num Y9,	Nomb Y9,	Cedula Y9,	Num Y3,	Nomb Y3,	Cedula Y3,	Num Y3 - 2,	Nomb Mercaderista 2,	Cedula Y3 - 2,	Num Y3 - 3,	Nomb Mercaderista 3,	Cedula Y3 - 3,	Num Y3 - 4,	Nomb Mercaderista 4,	Cedula Y3 - 4,	Num Y8,	Nomb Y8,	Cedula Y8,	Coordenada X,	Coordenada Y,	Bloqueo Clientes Pedido,	Grupo clientes 5,	Código Dane,	Nivel socioeconómico,	Grp Precios



-info.xlsx: driver con información para el filtrado de registros y con registro zapatocas, en este archivo se encuentran las siguientes variables.
![image](https://user-images.githubusercontent.com/86368935/215811323-334146d6-e9c3-46f3-a6f7-1c9025bd6fc0.png) esto es una imagen de las variables y sus grupos, se manejan bloques de colores para cada grupo, las columna no deben cambiar de nombre o de orden.

Poblaciones a excluir: Se ingresan los códigos DANE de las poblaciones que no se medirán
Nit a excluir: Clientes que se excluyen a nivel de NIT con el fin de no medirle ningún punto de venta
Tipologias seleccionadas: Tipologias actuales que manej comercial nutresa
Num cliente: (Clientes oxxo)  
Clientes Oxxo: categoriza los clientes oxxo por Base, Hogar, Receso (esta información la pasa la analista de cadenas)
Vendedores a excluir: Son vendedores genericos, de oficina y aquellos que no se deben medir en marketeam
Clientes a excluir: son reportados por la fuerza de venta para no medición, se utiliza el código de cliente
Ofic venta: todas las oficinas de venta que sirven como insumo para variable portafolio clave
Población: todas las poblaciones que sirven como insumo para variable portafolio clave
Concatenado: unión de Ofic venta + Población
Portafolio Clave: portafolio clave asignado según el resultado del concatenado
ZA-Z1 Núm person: código de los vendedores ZA y Z1
Can_Estruct: Depende de la variable ZA-Z1 Núm person, representa el canal real del vendedor, sale de la estructura de ventas
Subc_Estruc:Depende de la variable ZA-Z1 Núm person, representa el sub-canal real del vendedor, sale de la estructura de ventas
Socios No Plus: son los códigos de clientes socios excluyendo los plus, y proviene de la maestra socios generada cada mes
Segmento socios: es la categoria que le dan a dicho socio (variable Socios No Plus), y
Grp4 Cliente: este es el segmento de necesidades 
Segmento Vital: es la descripción de la variable Grp4 Cliente
Tipologia: Respresenta las tipologias reales de los cientes
Segmento:	es la descripción de la variable Tipologia
meses ventas: se ponen los últimos seis meses de ventas, cada mes va en una fila independiente y debe tener el siguiente formato ejemplo: 2022.07, esto es insumo el archivo ventas Directa.xlsm
of ventas a excluir: Son las oficinas de ventas que no se tienen en cuenta para las mediciones ejemplo: 00 
Top de clientes: en la primera fila se escribe en números el top máximo de clientes por zona
Clientes a excluir por bloqueo: Causales de clientes bloqueados para no tenerlos en cuenta en la medición, ejemplo: 06, 16, 17
Clientes desarrollados: Códigos de clientes, es acumulativo no se deben borrar los que ya se establecieron anteriormente, información compartida por equipo apoyo a procesos logisticos y comerciales (Eduardo Ramirez Espeleta)  



Programas necesarios
• Analysis for Microsoft Office
• Microsoft Excel

