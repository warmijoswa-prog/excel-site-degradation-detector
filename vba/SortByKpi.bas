Attribute VB_Name = "Módulo1"
'ORDENAR POR INDICADOR Y PINTARLO
Sub EFI_VOZ()
Attribute EFI_VOZ.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Add Key:=Range("Tabla_Consulta_desde_ODIN[EFI_VOZ]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("C1:J1").Interior.Color = xlColorIndexNone
    Range("C1").Interior.Color = RGB(0, 0, 0)
End Sub
Sub EFI_DAT()
Attribute EFI_DAT.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Add Key:=Range("Tabla_Consulta_desde_ODIN[EFI_DAT]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("C1:J1").Interior.Color = xlColorIndexNone
    Range("D1").Interior.Color = RGB(0, 0, 0)
End Sub
Sub RTWP()
Attribute RTWP.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Add Key:=Range("Tabla_Consulta_desde_ODIN[RTWP]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("C1:J1").Interior.Color = xlColorIndexNone
    Range("E1").Interior.Color = RGB(0, 0, 0)
End Sub
Sub RSSR()
Attribute RSSR.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Add Key:=Range("Tabla_Consulta_desde_ODIN[RSSR]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("C1:J1").Interior.Color = xlColorIndexNone
    Range("F1").Interior.Color = RGB(0, 0, 0)
End Sub
Sub CCR()
Attribute CCR.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Add Key:=Range("Tabla_Consulta_desde_ODIN[CCR]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("C1:J1").Interior.Color = xlColorIndexNone
    Range("G1").Interior.Color = RGB(0, 0, 0)
End Sub
Sub CSFR()
Attribute CSFR.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Add Key:=Range("Tabla_Consulta_desde_ODIN[CSFR]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("C1:J1").Interior.Color = xlColorIndexNone
    Range("H1").Interior.Color = RGB(0, 0, 0)
End Sub
Sub ERL()
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Add Key:=Range("Tabla_Consulta_desde_ODIN[ERL]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("C1:J1").Interior.Color = xlColorIndexNone
    Range("I1").Interior.Color = RGB(0, 0, 0)
End Sub
Sub GB()
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN").Sort _
        .SortFields.Add Key:=Range("Tabla_Consulta_desde_ODIN[GB]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA").ListObjects("Tabla_Consulta_desde_ODIN") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("C1:J1").Interior.Color = xlColorIndexNone
    Range("J1").Interior.Color = RGB(0, 0, 0)
End Sub
