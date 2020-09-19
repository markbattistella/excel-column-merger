' ExcelColumnMerger Macro
' Copyright (c) 2020
' Created by Matthew Baldwin
' Do not alter anything or else the world will break


Sub copyMatterIds()

    Dim ws As Worksheet
    Dim wsDest As Worksheet

	' create a output worksheet
    If Not WorksheetExists("MatterIds") Then
        Sheets.Add.Name = "MatterIds"
    End If

	' set the destination
    Set wsDest = Sheets("MatterIds")

	' loop all sheets in Workbook
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name <> wsDest.Name Then

			' from A1 to last cell entry
            ws.Range("A1", ws.Range("A1").End(xlDown).Offset(0, 1)).Copy

			' paste into destination sheet, offset 1 cell down
            wsDest.Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial xlPasteValues
        End If
    Next ws

End Sub

' see if the sheet name exists
Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function
