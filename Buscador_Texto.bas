Attribute VB_Name = "Buscador_Texto"
Function BUSCADORTEXTO(Celda As String, Inicio As Integer, ParamArray Texto() As Variant) As String

'Declaraciones

Dim Txt As Variant

'Busqueda

    For Each Txt In Texto()
    
            If (InStr(Inicio, Celda, CStr(Txt), vbTextCompare) > 0) Then
                
                BUSCADORTEXTO = Mid(Celda, InStr(Inicio, Celda, CStr(Txt), vbTextCompare), Len(CStr(Txt)))
                
                GoTo Terminar
    
        End If
    
    Next Txt
    
    If BUSCADORTEXTO = "" Then
    
    BUSCADORTEXTO = "Sin Registro"
    
    End If
    
Terminar: End Function

