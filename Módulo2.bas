Attribute VB_Name = "Módulo2"
Sub procura_vertical()
Attribute procura_vertical.VB_ProcData.VB_Invoke_Func = " \n14"
'
' procura_vertical Macro
'

'
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R[-2]C,'Bd-operações'!C[-6]:C[-3],4,0)"
    Range("G11").Select
End Sub
Sub monta_tabela()
Attribute monta_tabela.VB_ProcData.VB_Invoke_Func = " \n14"
'
' monta_tabela Macro
'

'
    Sheets("Bd-operações").Select
    Columns("A:D").Select
    Selection.Copy
    Sheets("tmp-print").Select
    Columns("A:A").Select
    ActiveSheet.Paste
    Range("A1:D1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G1").Select
    ActiveSheet.Paste
    Columns("H:I").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "='Consulta dados'!R[7]C"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "='Consulta dados'!R[9]C[-1]"
    Range("G2:H2").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
End Sub
Sub gera_filtro()
Attribute gera_filtro.VB_ProcData.VB_Invoke_Func = " \n14"
'
' gera_filtro Macro
'

'
    Columns("A:D").Select
    Columns("A:D").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "G1:H2"), CopyToRange:=Columns("J:M"), Unique:=False
    Columns("M:M").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
End Sub
Sub Gera_tabela()
Attribute Gera_tabela.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Gera_tabela Macro
'

'
    Columns("J:M").Select
    Selection.Copy
    Sheets("tabela").Select
    Columns("G:G").Select
    ActiveSheet.Paste
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G5:J5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -9.99786370433668E-02
        .PatternTintAndShade = 0
    End With
    Range("G2:J3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("G2:J3").Select
    ActiveCell.FormulaR1C1 = "Tabela de uso de maquina"
    Range("G2:J3").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    Selection.Font.Size = 14
    Selection.Font.Size = 16
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayGridlines = False
End Sub
