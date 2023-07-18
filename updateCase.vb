Sub Update()

' === cross check if all data provided === '
Dim ws As Worksheet
Set ws = ThisWorkbook.ActiveSheet

If ws.Range("J6") = "" Then
    MsgBox "Please provide case number in cell J6."
    Exit Sub
End If
' === end checking === '

Dim CaseNo As Integer
CaseNo = ws.Range("J6")
 
    Call ByCase

MsgBox ("Macro has finished working.")

End Sub

Sub ByCase()

    Dim CaseNum As String
    Dim trackerSheet As Worksheet
    Dim formSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set trackerSheet = ThisWorkbook.Sheets("Tracker")
    Set formSheet = ThisWorkbook.Sheets("Form")
    
    CaseNum = formSheet.Range("J6").Value
    
    ' Search for Case details '
    lastRow = trackerSheet.Cells(trackerSheet.Rows.Count, "A").End(xlUp).Row
        For i = 1 To lastRow
        If trackerSheet.Range("A" & i).Value = CaseNum Then
            ' If found, assign to variable
            
            Dim CoCo As String
            CoCo = trackerSheet.Range("B" & i).Value
            
            If CoCo = 1 Or CoCo = 2 Then
                Call Update.north
            ElseIf CoCo = 3 Or CoCo = 4 Then
                Call Update.sbc
            ElseIf CoCo = 5 Or CoCo = 6 Then
                Call Update.mbc
            Else
                Debug.Print "Company Code not recognized."
            End If

            Exit Sub ' End macro
        End If
    Next i
    
    ' If not found
    MsgBox "CaseNo not found in Tracker."

End Sub
