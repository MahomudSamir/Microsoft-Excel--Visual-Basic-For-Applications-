Attribute VB_Name = "Importacion_SQL_Server"

Function IMPORTACION_SQLSERVER(Tabla As String, Columnas As Integer, ParamArray Valores() As Variant) As String

Dim Posicion, Posicion2, Campos, Conteo, CantidadC, ConteoB As Integer

Dim Cadena, Cadena2, Atributos, Registros As String

Dim Txt As Variant

Cadena = Join(Valores(), ",")

'Extraccion de los atributos:

CantidadC = 0

Conteo = 0

Posicion = 0

For Each Txt In Valores()
    
    Conteo = Conteo + 1
    
    CantidadC = CantidadC + Len(Txt)
    
        If Conteo = Columnas Then
        
            Posicion = CantidadC + Conteo
            
            GoTo Siguiente
        
        End If

Next Txt

Siguiente:

Atributos = Mid(Cadena, 1, Posicion - 1)

'Extraccion de los registros:

Conteo = 0

CantidadC = 0

For Each Txt In Valores()

    Conteo = Conteo + 1
    
    CantidadC = CantidadC + Len(Txt)

    If Conteo = (Columnas * 2) Then
    
    Posicion2 = CantidadC + Conteo
    
    GoTo Proceder
    
    End If

Next Txt

Proceder:

Posicion2 = Posicion2 - Posicion

Conteo = 0

ConteoB = 0

Cadena = Cadena + ","

Cadena2 = ""

For Each Txt In Valores()

    Conteo = Conteo + 1

    If Conteo > Columnas And Conteo < ((Columnas * 2) + 1) Then
    
        If (Mid(Cadena, ((Posicion + Posicion2 + 1) + ConteoB), 2) = "0,") Then
        
            Cadena2 = Cadena2 + CStr(Txt) + ","
        
        ElseIf (Mid(Cadena, ((Posicion + Posicion2 + 1) + ConteoB), 2) = "1,") Then
        
            If (Mid(CStr(Txt), 3, 1) = "/") And (Mid(CStr(Txt), 6, 1) = "/") Then 'Formato Strings Fecha
            
                Cadena2 = Cadena2 + "'" + Mid(CStr(Txt), 7, 4) + Mid(CStr(Txt), 4, 2) + Mid(CStr(Txt), 1, 2) + "',"
        
            Else
        
                Cadena2 = Cadena2 + "'" + CStr(Txt) + "',"
            
            End If
        
        End If

    ConteoB = ConteoB + 2

    ElseIf Conteo = ((Columnas * 2) + 1) Then
    
        GoTo Terminar
    
    End If

Next Txt

Terminar:

Registros = Mid(Cadena2, 1, Len(Cadena2) - 1)

IMPORTACION_SQLSERVER = "INSERT INTO " + Tabla + " (" + Atributos + ") VALUES (" + Registros + ");"

End Function

