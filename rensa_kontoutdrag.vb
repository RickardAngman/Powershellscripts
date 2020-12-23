Sub rensa_kontoutdrag()
' rensa_kontoutdrag Makro
    Windows("kontoutdrag.xlsx").Activate
    Range("A8").Select
    ActiveCell.FormulaR1C1 = "Datum"
    Rows("1:7").Select
    Selection.Delete Shift:=xlUp
    Columns("A:F").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$F"), , xlYes).Name = "Tabell1"
    Columns("B:C").Select
    Selection.ListObject.ListColumns(2).Delete
    Selection.ListObject.ListColumns(2).Delete
    Columns("D:D").Select
    Selection.ListObject.ListColumns(4).Delete
    ActiveWorkbook.SaveAs Filename:="C:\Users\ricka\Downloads\kontoutdrag.csv", FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
End Sub
