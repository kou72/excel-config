Attribute VB_Name = "Module1"

Sub Macro1()
'
' Macro1 Macro
'

'
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "test"
    Range("B3").Select
End Sub
