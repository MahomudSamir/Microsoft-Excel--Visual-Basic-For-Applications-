# Microsoft Excel (Visual Basic For Applications)
 
*************************************
             Bienvenidos 
*************************************

1) Buscador_Texto.bas

 Hace referencia a la funcion BUSCADORTEXTO() que permite extraer la primera coincidencia de
una cadena de caracteres definido dentro de la formula en una celda de caracteres.

 Su declaracion es a traves de =BUSCADORTEXTO(Celda, Inicio, Texto()), donde:

- Celda: Celda de la hoja donde se encuentra la cadena de caracteres a buscar.
- Inicio: Valor numerico de la posicion inicial de la busqueda dentro de la cadena de caracteres.
- Texto(): Cadena de caracter a buscar dentro de la celda almacenados dentro de un vector de tamaño indefinido.


2) Importacion_SQL_Server.bas

 Hace referencia a la funcion IMPORTACION_SQLSERVER() que permite facilitar la importacion de datos dentro de
una tabla y/o rango a Microsoft SQL Server a traves de un conjunto de sentencias INSERT con los valores de
la entidad y los registros

 Su declaracion es a traves de =IMPORTACION_SQLSERVER(Tabla, Columnas, Valores()), donde:

- Tabla: Cadena de caracteres que define el nombre que tendra dicha entidad y/o tabla.
- Columnas: Valor numerico que define la cantidad de atributos y/o campos que tendra la tabla.
- Valores(): Valores define 1) las celdas donde estan contenidas los atributos de la tabla 2) las celdas donde
estan contenido los registros que cada atributo tiene y 3) la definicion del tipo de valor contenido en los registros
de cada entidad, siendo 1 para valores tipo String y 0 para valores tipo Int.

Ejemplo:

-A----B----------C

1-COD-YEAR-PAIS 


2-002-01/01/1997-BRA


3-003-06/06/1998-ITA


En el rango que va de A1:C3 para la extraccion de datos se define la formula =IMPORTACION_SQL() como:

=IMPORTACION_SQLSERVER("Country_Table";3;$A$1;$B$1;$C$1;$A2;$B2;$C2;0;0;1)

Siendo el resultado:

-A----B----------C---D

1-COD-YEAR-PAIS

2-001-01/01/1997-BRA-INSERT INTO Country_Table (COD,YEAR,PAIS) VALUES (001,19970101,'BRA');

3-003-06/06/1998-ITA-INSERT INTO Country_Table (COD,YEAR,PAIS) VALUES (002,19980606,'ITA');
