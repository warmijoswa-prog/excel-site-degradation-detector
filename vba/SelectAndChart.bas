Attribute VB_Name = "Módulo3"
'ENVIAR VARIABLES Y GRAFICAR
Sub coordenada()

On Error Resume Next

Dim red As String
Dim eebb As String
Dim indicador As String
Dim fila As Integer
Dim nombre_columna As String

fila = ActiveCell.Row
columna = ActiveCell.Column
If fila = 1 Then MsgBox "Seleccione registro válido" Else GoTo VER_COLUMNA
GoTo FIN
VER_COLUMNA:
If columna < 3 Or columna > 10 Then MsgBox "Seleccione columna de indicador" Else GoTo CONTINUAR
GoTo FIN

CONTINUAR:
nombre_columna = Mid(Split(Columns(columna).Address, ":")(1), 2)
red = Range("A" & fila).Value
eebb = Range("B" & fila).Value
indicador = Range(nombre_columna & "1").Value

dtmNext = DateAdd("s", 5, Now)

Sheets("GRAFICO").Select

ActiveSheet.PivotTables("TablaDinámica2").PivotCache.Refresh
Range("B1").Value = indicador
Range("B2").Value = red
Range("B3").Value = eebb

Call tabla

FIN:
End Sub
