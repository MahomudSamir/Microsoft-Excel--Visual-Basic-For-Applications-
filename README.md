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
- Valores(): Valores define a) Las celdas donde estan contenidas los atributos de la tabla b) Las celdas donde
estan contenido los registros que cada atributo tiene y c) La definicion del tipo de valor contenido en los registros
de cada entidad, siendo 1 para valores tipo String y 0 para valores tipo Numeric.
