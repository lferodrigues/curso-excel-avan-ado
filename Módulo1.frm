Attribute VB_Name = "Módulo1"
Sub busca_data()
Attribute busca_data.VB_ProcData.VB_Invoke_Func = " \n14"
'
' busca_data Macro
'

'
    Range("G15").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("G15").Select
End Sub
Sub Inseri_dados()
Attribute Inseri_dados.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Inseri_dados Macro
'

'
    Sheets("Bd-operações").Select
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Cadastro").Select
    Range("G9,G11,G13,G15").Select
    Range("G15").Activate
    Selection.Copy
    Sheets("Bd-operações").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Sheets("Cadastro").Select
    Range("G9").Select
End Sub
Sub limpeza()
Attribute limpeza.VB_ProcData.VB_Invoke_Func = " \n14"
'
' limpeza Macro
'

'
    Range("G9").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("G11").Select
    Selection.ClearContents
    Range("G13").Select
    Selection.ClearContents
    Range("G15").Select
    Selection.ClearContents
    Range("G9").Select
End Sub
