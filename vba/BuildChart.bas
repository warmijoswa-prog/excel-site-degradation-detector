Attribute VB_Name = "Módulo4"
'GRAFICAR
Sub tabla()

Dim cn As New ADODB.Connection
Dim fila As Integer

Set cn = New ADODB.Connection

Sheets("DATA_").Select
Range("A2:B1048576").Select
Selection.ClearContents

ConnectionString = "Provider=OraOLEDB.Oracle;Data Source=DB4;User Id=MYUSER;Password=PASS"
cn.Open ConnectionString

Dim data As New ADODB.Recordset
Set data = New ADODB.Recordset
consulta = Sheets("CONSULTA").Range("A8").Value
data.Open consulta, cn
Sheets("DATA_").Range("A2").CopyFromRecordset data
Columns("A:A").Select
Selection.NumberFormat = "dd/mm/yyyy hh:mm"
data.Close
cn.Close

If Sheets("GRAFICO").Range("B1").Value = "RTWP" Then GoTo NEGATIVO Else GoTo POSITIVO

NEGATIVO:
Range("A1048576").Select
Selection.End(xlUp).Select
fila = ActiveCell.Row
Range("B2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Range("C2").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "=RC[-1]*-1"
Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C" & fila)
Range("C2:C" & fila).Select
Selection.Copy
Range("B2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Columns("C:C").Select
Application.CutCopyMode = False
Selection.ClearContents
Range("E1").Select
ActiveSheet.PivotTables("TablaDinámica1").PivotCache.Refresh
Sheets("GRAFICO").Select
ActiveSheet.ChartObjects("Gráfico 1").Activate
ActiveChart.Axes(xlValue).Select
ActiveChart.Axes(xlValue).MinimumScale = -110
ActiveChart.Axes(xlValue).MaximumScale = -65
Selection.TickLabels.NumberFormat = "General"
GoTo FIN

POSITIVO:
Range("E1").Select
ActiveSheet.PivotTables("TablaDinámica1").PivotCache.Refresh
Sheets("GRAFICO").Select

If (Range("B1").Value = "GB" Or Range("B1").Value = "ERL") Then GoTo UNIDAD Else GoTo PORCENTAJE

PORCENTAJE:
ActiveSheet.ChartObjects("Gráfico 1").Activate
ActiveChart.Axes(xlValue).Select
ActiveChart.Axes(xlValue).MinimumScale = 0
ActiveChart.Axes(xlValue).MaximumScale = 1.05
Selection.TickLabels.NumberFormat = "0%"
GoTo FIN

UNIDAD:
ActiveSheet.ChartObjects("Gráfico 1").Activate
ActiveChart.Axes(xlValue).Select
ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
Selection.TickLabels.NumberFormat = "General"
GoTo FIN

FIN:
Range("A10").Select
ActiveCell.FormulaR1C1 = "=R[-7]C[1]&"" | ""&R[-1]C&"" | ""&R[-9]C[1]&"" ""&R[-8]C[1]"
Selection.Copy
Range("A11").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Application.CutCopyMode = False
Range("B1").Select
 
End Sub

