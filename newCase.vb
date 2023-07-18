Sub Reminder()

'This is the superior macro in order to determine path for Reminder Validator'
'Based on Company Code provided, 3 paths to be followed: North, SBC or MBC check'

' === cross check if all data provided === '
Dim ws As Worksheet
Dim lr As Long
Set ws = ThisWorkbook.ActiveSheet
Range("B9").Select
lr = Cells(Rows.Count, ActiveCell.Column).End(xlUp).Row - ActiveCell.Row

If lr = 0 Then
    MsgBox "Please provide invoice details in the table"
    Exit Sub
ElseIf ws.Range("C6") = "" Then '
    MsgBox "Please provide company code in cell C6"
    Exit Sub
ElseIf ws.Range("C7") = "" Then
    MsgBox "Please provide the reminder date in cell C7"
    Exit Sub
ElseIf ws.Range("C8") = "" Then
    MsgBox "Please provide the dunning level in cell C8"
    Exit Sub
End If
' === end checking === '

' === path recognition === '
Dim cc As Integer
cc = ws.Range("C6")

If cc = 1 Or cc = 2 Then
    Call north_check
ElseIf cc = 3 Or cc = 4 Then
    Call sbc_check
ElseIf cc = 5 Or cc = 6 Then
    Call mbc_check
Else
    Debug.Print "Company Code not recognized."
End If

MsgBox ("Macro has finished working.")
'Application.Wait (Now + TimeValue("00:00:15"))
'Kill "C:\Reminders\temp\temp.xlsx"

End Sub

Sub north_check()



Dim cc As Integer
Dim layout As String
Dim path As String
Dim temp As String

Dim ReminderWorkbook As Workbook
Dim ReminderTrackerSheet As Worksheet
Dim ReminderFormSheet As Worksheet
    
' Set ReminderValidator workbook and sheet references
Set ReminderWorkbook = Workbooks("Reminders Validator.xlsm")
Set ReminderFormSheet = ReminderWorkbook.Sheets("Form")
Set ReminderTrackerSheet = ReminderWorkbook.Sheets("Tracker")
    
cc = ReminderFormSheet.Range("C6")
layout = "/REMIND_RPA"
path = "C:\Reminders\temp\"
temp = "temp.xlsx"

Dim data As Date
Dim reference As String

Dim lr As Long
lr = Cells(Rows.Count, "B").End(xlUp).Row
Dim cell As Range


' === Connection to SAP === '
On Error Resume Next
    If Not IsObject(SAP) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SAP = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
        Set Connection = SAP.Children(0)
    End If
    If Not IsObject(session) Then
        Set session = Connection.Children(0)
    End If
        If IsObject(WScript) Then
         WScript.ConnectObject session, "on"
         WScript.ConnectObject SAP, "on"
    End If
If Err.Number <> 0 Then
    MsgBox "Please connect to SAP!"
    Exit Sub
End If

' === master loop, to verify status of invoices given by the User === '
For Each cell In Range("B10:B" & lr)
    
    Dim lastRow As Long
    lastRow = ReminderTrackerSheet.Cells(ReminderTrackerSheet.Rows.Count, "B").End(xlUp).Row
    Dim unusedRow As Long
    unusedRow = lastRow + 1
        
    Dim dfb, CaseNo, dfp As Variant
    CaseNo = ReminderTrackerSheet.Cells(unusedRow, 1).Value
    
' === FBL1N transaction (Vendor Line Items) === '
    session.findById("wnd[0]").resizeWorkingPane 132, 25, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n FBL1N" '#transaction'
    session.findById("wnd[0]").sendVKey 0 '#execute'
    session.findById("wnd[0]/tbar[1]/btn[16]").press '#dynamic selection'
    session.findById("wnd[0]/usr/radX_AISEL").Select
    session.findById("wnd[0]/usr/chkX_SHBV").Selected = True
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN013-LOW").Text = cell.Value 'date of document'
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN015-LOW").Text = cell.Offset(0, 1).Value 'reference'
    session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKD_LIFNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").Text = cc 'Company Code provided in cell C6'
    session.findById("wnd[0]/usr/ctxtPA_VARI").Text = layout 'Layout of report - /Remind_RPA'
    session.findById("wnd[0]/usr/ctxtPA_VARI").SetFocus
    session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 11
    session.findById("wnd[0]/tbar[1]/btn[8]").press ' execute '
    
    errMessage = session.findById("wnd[0]/sbar").Text
    If errMessage = "No items selected (see long text)" Then
    
    ' case: invoice not in SAP '
            ReminderTrackerSheet.Cells(unusedRow, "B").Value = cc ' Company Code '
            ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
            ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
            ReminderTrackerSheet.Cells(unusedRow, "E").Value = cell.Value ' invoice date '
            ReminderTrackerSheet.Cells(unusedRow, "F").Value = cell.Offset(0, 1).Value ' reference / invoice number '
            ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: invoice missing from SAP."  ' Comment '
            ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
            ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice not final posted in SAP. Case registered under number: " & CaseNo, , "Result" ' message '
        
    Else ' invoice in SAP, extract to temp file '
    
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = path 'Path of Macro folder'
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = temp 'Name of temp file'
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6 '#Replace'
        session.findById("wnd[1]/tbar[0]/btn[11]").press
    
' === open temp file, copy data to tracker === '
        Workbooks.Open path & temp
    
        ' Set Temp workbook and sheet references
           Dim TempSheet As Worksheet
           Dim TempWorkbook As Workbook
           Set TempWorkbook = Workbooks("temp.xlsx")
           Set TempSheet = TempWorkbook.Sheets(1)
    
    ' = verification of status: Posted or Paid? = '
        Dim cellValue As String
        Dim cellToCheck As Range
        Set cellToCheck = TempSheet.Range("R2")
                
        If Len(Trim(cellToCheck.Value)) = 0 Then
        ' case: invoice not paid
                        
                ' Copy data from Temp sheet to Tracker sheet
                ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("B2").Value ' Company Code '
                ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
                ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
                ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("J2").Value ' invoice date '
                ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("P2").Value ' reference / invoice number '
                ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("C2").Value ' Vendor number '
                ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("D2").Value ' Vendor name '
                ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("L2").Value ' due date '
                ReminderTrackerSheet.Cells(unusedRow, "M").Value = TempSheet.Range("Q2").Value ' booking date '
                ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
                ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
                ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
                
                If ReminderTrackerSheet.Cells(unusedRow, "N").Value < 8 Then
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: invoice posted, should be paid out with next payment run." ' Comment '
                Else
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: invoice posted, not paid. Verify the VMD/booking data." ' Comment '
                End If
                dfb = ReminderTrackerSheet.Cells(unusedRow, 14).Value
                MsgBox "Invoice posted finally in SAP " & dfb & " days ago. Case registered under number: " & CaseNo, , "Result" ' message '
                
        Workbooks(temp).Close SaveChanges:=False
        
        Else
        ' case: invoice paid '
        
                ' Copy data from Temp sheet to Tracker sheet
                ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("B2").Value ' Company Code '
                ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
                ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
                ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("J2").Value ' invoice date '
                ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("P2").Value ' reference / invoice number '
                ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("C2").Value ' Vendor number '
                ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("D2").Value ' Vendor name '
                ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("L2").Value ' due date '
                ReminderTrackerSheet.Cells(unusedRow, "M").Value = TempSheet.Range("Q2").Value ' booking date '
                ReminderTrackerSheet.Cells(unusedRow, "O").Value = TempSheet.Range("S2").Value ' payment/clearing date '
                ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
                ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
                ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
                
                If TempSheet.Range("R2").Value <= 3200000000# Or TempSheet.Range("R2").Value >= 3299999999# Then
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice cleared manually. Verify the case and respond to supplier." ' Comment '
                            dfp = ReminderTrackerSheet.Cells(unusedRow, 16).Value
                    MsgBox "Invoice cleared manually in SAP " & dfp & " days ago. Check in SAP. Case registered under number: " & CaseNo, , "Result" ' message '
                Else
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice paid." ' Comment '
                            dfp = ReminderTrackerSheet.Cells(unusedRow, 16).Value
                    MsgBox "Invoice paid " & dfp & " days ago. Case registered under number: " & CaseNo, , "Result" ' message '
                End If
                    
        Workbooks(temp).Close SaveChanges:=False
                
        End If
        
    End If

Next cell

End Sub

Sub sbc_check()


Dim cc As Integer
Dim layout As String
Dim path As String
Dim temp As String

Dim ReminderWorkbook As Workbook
Dim ReminderTrackerSheet As Worksheet
Dim ReminderFormSheet As Worksheet
    
' Set ReminderValidator workbook and sheet references
Set ReminderWorkbook = Workbooks("Reminders Validator.xlsm")
Set ReminderFormSheet = ReminderWorkbook.Sheets("Form")
Set ReminderTrackerSheet = ReminderWorkbook.Sheets("Tracker")
    
cc = ReminderFormSheet.Range("C6")
layout = "/REMINDERS"
path = "C:\Reminders\temp\"
temp = "temp.xlsx"

Dim data As Date
Dim reference As String

Dim lr As Long
lr = Cells(ReminderFormSheet.Rows.Count, "B").End(xlUp).Row
Dim cell As Range


' === Connection to SAP === '
On Error Resume Next
    If Not IsObject(SAP) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SAP = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
        Set Connection = SAP.Children(0)
    End If
    If Not IsObject(session) Then
        Set session = Connection.Children(0)
    End If
        If IsObject(WScript) Then
         WScript.ConnectObject session, "on"
         WScript.ConnectObject SAP, "on"
    End If
If Err.Number <> 0 Then
    MsgBox "Please connect to SAP!"
    Exit Sub
End If
    
' === master loop, to verify status of invoices given by the User === '
For Each cell In ReminderFormSheet.Range("B10:B" & lr)

    ' Disable screen updating and alerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim lastRow As Long
    lastRow = ReminderTrackerSheet.Cells(ReminderTrackerSheet.Rows.Count, "B").End(xlUp).Row
    Dim unusedRow As Long
    unusedRow = lastRow + 1
    
    Dim dfb As Variant
    Dim CaseNo As Variant
    Dim dfp As Variant
    CaseNo = ReminderTrackerSheet.Cells(unusedRow, 1).Value
    
    ' === VIM ANALYTICS 2 transaction (VIM_VA2) === '
    session.findById("wnd[0]").resizeWorkingPane 132, 25, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n/opt/vim_va2" ' Transaction '
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").Text = cell.Value ' Date of invoice '
    session.findById("wnd[0]/usr/txtS_XBLNR-LOW").Text = cell.Offset(0, 1).Value ' Reference / Invoice number '
    session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").Text = cc ' Company Code '
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = layout ' Layout of raport '
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").SetFocus
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Scenario 7: Invoice not in VIM '
    errMessage = session.findById("wnd[0]/sbar").Text
    If errMessage = "No data found for specified select-option/parameter" Then
        ReminderTrackerSheet.Cells(unusedRow, "B").Value = cc ' Company Code '
        ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
        ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
        ReminderTrackerSheet.Cells(unusedRow, "E").Value = cell.Value ' invoice date '
        ReminderTrackerSheet.Cells(unusedRow, "F").Value = cell.Offset(0, 1).Value ' reference / invoice number '
        ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: invoice missing from SAP."  ' Comment '
        ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
        ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
        ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
            ' ==== sending e-mails part ==== '
        Dim mail1 As Object ' Add a variable to hold the mail object
        mail.AddOutlookReference
        mail.AddWordReference
            Set mail1 = CreateObject("mail.MissInv") ' Create and set the mail object
        mail.MissInv unusedRow, CaseNo
            Set mail1 = Nothing ' Release the mail object
        mail.RemoveOutlookReference
        mail.RemoveWordReference
        MsgBox "Invoice missing from VIM. Contact supplier. Case registered under number: " & CaseNo, , "Result" ' message '
        
    Else ' invoice in SAP, extract to temp file '
    
    session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&XXL"
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = path ' path to the folder '
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = temp ' file name '
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[11]").press

    ' === open temp file, copy data to tracker === '
        Workbooks.Open path & temp
    
        ' Set Temp workbook and sheet references
           Dim TempSheet As Worksheet
           Dim TempWorkbook As Workbook
           Set TempWorkbook = Workbooks("temp.xlsx")
           Set TempSheet = TempWorkbook.Sheets(1)
    
        ' = verification of VIM status = '
        Dim cellValue As String
        Dim cellToCheck, AppFN, AppLN, clearingDoc As Range
        Set cellToCheck = TempSheet.Range("M2") ' VIM Status '
        Set AppFN = TempSheet.Range("P2") ' First Name of Approver '
        Set AppLN = TempSheet.Range("Q2") ' Last Name of Approver '
        Set clearingDoc = TempSheet.Range("U2") ' Clearing Doc '
        
        If cellToCheck = "Indexed" Or cellToCheck = "Created" Or cellToCheck = "Parked" Then
        ' - scenario 1 - '
            ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("A2").Value ' Company Code '
            ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
            ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
            ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("I2").Value ' Invoice Date '
            ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("J2").Value ' Invoice no '
            ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("G2").Value ' Vendor no '
            ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("H2").Value ' Vendor name '
            ReminderTrackerSheet.Cells(unusedRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(unusedRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice registered in the system, proceed with processing it in VIM." ' Comment '
            ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
            ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice received in system, proceed with processing in VIM. Case registered under number: " & CaseNo, , "Result" ' message '
            
        ElseIf cellToCheck = "Rejected by Approver" Or cellToCheck = "Blocked" Or cellToCheck = "Approval recalled" Then
        ' - scenario 2 - '
            ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("A2").Value ' Company Code '
            ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
            ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
            ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("I2").Value ' Invoice Date '
            ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("J2").Value ' Invoice no '
            ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("G2").Value ' Vendor no '
            ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("H2").Value ' Vendor name '
            ReminderTrackerSheet.Cells(unusedRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(unusedRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice not fully processed in system, validate the case manually." ' Comment '
            ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
            ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice not fully processed in system. Current status: " & cellToCheck & ". Case registered under number: " & CaseNo, , "Result" ' message '
            
        ElseIf cellToCheck = "Awaiting Approval - Parked" Then
        ' - scenario 3 - '
            ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("A2").Value ' Company Code '
            ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
            ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
            ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("I2").Value ' Invoice Date '
            ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("J2").Value ' Invoice no '
            ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("G2").Value ' Vendor no '
            ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("H2").Value ' Vendor name '
            ReminderTrackerSheet.Cells(unusedRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(unusedRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Approval is undergoing. Current approver: " & AppFN & " " & AppLN ' Comment '
            ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
            ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
            ' ==== sending e-mails part ==== '
                ApprovalMail.AddOutlookReference
                ApprovalMail.AddWordReference
                    Set mail1 = CreateObject("ApprovalMail.MissAppr") ' Create and set the mail object
                ApprovalMail.MissAppr unusedRow, CaseNo, AppFN, AppLN
                    Set mail1 = Nothing ' Release the mail object
                ApprovalMail.RemoveOutlookReference
                ApprovalMail.RemoveWordReference
            MsgBox "Invoice is pending approval in VIM. Case registered under number: " & CaseNo, , "Result" ' message '
            
        ElseIf cellToCheck = "Approval Complete" Then
        ' - scenario 4 - '
            ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("A2").Value ' Company Code '
            ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
            ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
            ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("I2").Value ' Invoice Date '
            ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("J2").Value ' Invoice no '
            ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("G2").Value ' Vendor no '
            ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("H2").Value ' Vendor name '
            ReminderTrackerSheet.Cells(unusedRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(unusedRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Approval is completed - post invoice via VIM." ' Comment '
            ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
            ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice is approved and can be posted in SAP. Case registered under number: " & CaseNo, , "Result" ' message '
   
        ElseIf cellToCheck = "Posted" Then
        ' - scenario 5 - '
            ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("A2").Value ' Company Code '
            ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
            ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
            ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("I2").Value ' Invoice Date '
            ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("J2").Value ' Invoice no '
            ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("G2").Value ' Vendor no '
            ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("H2").Value ' Vendor name '
            ReminderTrackerSheet.Cells(unusedRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(unusedRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(unusedRow, "M").Value = TempSheet.Range("S2").Value ' End date '
            ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
            ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status

            If clearingDoc = "" Then
            ' invoice posted finally, but not yet paid  '
                If ReminderTrackerSheet.Cells(unusedRow, "N").Value < 8 Then
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: invoice posted, should be paid out with next payment run." ' Comment '
                Else
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: invoice posted, not paid. Verify the VMD/booking data." ' Comment '
                End If
                        dfb = ReminderTrackerSheet.Cells(unusedRow, 14).Value
                MsgBox "Invoice posted finally in SAP " & dfb & " days ago. Case registered under number: " & CaseNo, , "Result" ' message '
            ElseIf clearingDoc <= 3200000000# Or clearingDoc >= 3299999999# Then
            ' invoice cleared manually '
                    ReminderTrackerSheet.Cells(unusedRow, "O").Value = TempSheet.Range("T2").Value ' payment/clearing date '
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice cleared manually. Verify the case and respond to supplier." ' Comment '
                            dfb = ReminderTrackerSheet.Cells(unusedRow, 14).Value
                            dfp = ReminderTrackerSheet.Cells(unusedRow, 16).Value
                    MsgBox "Invoice cleared manually in SAP " & dfp & " days ago. Check in SAP. Case registered under number: " & CaseNo, , "Result" ' message '
            Else
            ' invoice paid '
                    ReminderTrackerSheet.Cells(unusedRow, "O").Value = TempSheet.Range("T2").Value ' payment/clearing date '
                            dfb = ReminderTrackerSheet.Cells(unusedRow, 14).Value
                            dfp = ReminderTrackerSheet.Cells(unusedRow, 16).Value
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice paid." ' Comment '
                    MsgBox "Invoice paid " & dfp & " days ago. Case registered under number: " & CaseNo, , "Result" ' message '
            End If
            
        Else ' Obsolete invoice '
        ' - scenario 6 - '
        
            Workbooks(temp).Close SaveChanges:=False
            Kill "C:\Reminders\temp\temp.xlsx"
            
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            session.findById("wnd[0]/usr/txtS_XBLNR-LOW").Text = cell.Offset(0, 1).Value & "*"
            session.findById("wnd[0]/usr/txtS_XBLNR-LOW").SetFocus
            session.findById("wnd[0]/usr/txtS_XBLNR-LOW").caretPosition = 11
            session.findById("wnd[0]/tbar[1]/btn[8]").press
            session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").setCurrentCell -1, "OVERALL_STATUS_TEXT"
            session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").firstVisibleColumn = "DOC_DATE"
            session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").selectColumn "OVERALL_STATUS_TEXT"
            session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").selectedRows = ""
            session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").pressToolbarButton "&MB_FILTER"
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").Select
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "Obsolete"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").Text = "Deleted"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").Text = "Cancelled"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").Text = "Suspected Duplicate"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").Text = "Confirmed Duplicate"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").SetFocus
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").caretPosition = 19
            session.findById("wnd[2]/tbar[0]/btn[8]").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            If errMessage = "No data found for specified select-option/parameter" Then
                MsgBox "ERROR! Invoice deleted from the system, no new invoice scanned. Verify manually.", , "Result" ' message '
            Else
                session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").pressToolbarContextButton "&MB_EXPORT"
                session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&XXL"
                session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "\\oe-divisions\OBSE\Shared Resources\AP\ROBOT\Reminders\temp\" ' path to the folder '
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "temp.xlsx" ' file name '
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
                session.findById("wnd[1]/tbar[0]/btn[11]").press
                
                Workbooks.Open "C:\Reminders\temp\temp.xlsx"
                Dim temp2 As Workbook
                Dim temp2s As Worksheet
                Set temp2 = Workbooks("temp.xlsx")
                Set temp2s = temp2.Sheets(1)
                If temp2.Range("A2").Value = "" Then
                    ReminderTrackerSheet.Cells(unusedRow, "B").Value = cc ' Company Code '
                    ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
                    ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
                    ReminderTrackerSheet.Cells(unusedRow, "E").Value = cell.Value ' invoice date '
                    ReminderTrackerSheet.Cells(unusedRow, "F").Value = cell.Offset(0, 1).Value ' reference / invoice number '
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice deleted from the system, no new invoice scanned."  ' Comment '
                    ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
                    ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
                    ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
                    MsgBox "ERROR! Invoice deleted from the system, no new invoice scanned. Verify manually. Case registered under number: " & CaseNo, , "Result" ' message '                Else
                ReminderTrackerSheet.Cells(unusedRow, "B").Value = temp2s.Range("A2").Value ' Company Code '
                ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
                ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
                ReminderTrackerSheet.Cells(unusedRow, "E").Value = temp2s.Range("I2").Value ' Invoice Date '
                ReminderTrackerSheet.Cells(unusedRow, "F").Value = temp2s.Range("J2").Value ' Invoice no '
                ReminderTrackerSheet.Cells(unusedRow, "G").Value = temp2s.Range("G2").Value ' Vendor no '
                ReminderTrackerSheet.Cells(unusedRow, "H").Value = temp2s.Range("H2").Value ' Vendor name '
                ReminderTrackerSheet.Cells(unusedRow, "I").Value = temp2s.Range("R2").Value ' Invoice scan date '
                ReminderTrackerSheet.Cells(unusedRow, "J").Value = temp2s.Range("N2").Value ' Due date '
                ReminderTrackerSheet.Cells(unusedRow, "Q").Value = temp2s.Range("M2").Value ' VIM Status '
                ReminderTrackerSheet.Cells(unusedRow, "M").Value = temp2.Range("S2").Value ' End date '
                ReminderTrackerSheet.Cells(unusedRow, "O").Value = temp2.Range("T2").Value ' payment/clearing date '
                ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Original document deleted from the system. New docID found." ' Comment '
                ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
                ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
                ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
                MsgBox "Original document deleted. New docID found. Current status: " & temp2s.Range("M2").Value & ". Validate if more steps needed. Case registered under number: " & CaseNo, , "Result" ' message '
                End If
            End If
        End If
    
    End If
    
    ' === Close and clean up === '
    TempWorkbook.Close SaveChanges:=False
    
    ' Enable screen updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    


Next cell
   
    ' Disconnect from SAP
    session.Disconnect
    
    ' Clean up objects
    Set session = Nothing
    Set Connection = Nothing
    Set SAP = Nothing
        
End Sub

Sub mbc_check()

Dim cc As Integer
Dim layout As String
Dim path As String
Dim temp As String

Dim ReminderWorkbook As Workbook
Dim ReminderTrackerSheet As Worksheet
Dim ReminderFormSheet As Worksheet
    
' Set ReminderValidator workbook and sheet references
Set ReminderWorkbook = Workbooks("Reminders Validator.xlsm")
Set ReminderFormSheet = ReminderWorkbook.Sheets("Form")
Set ReminderTrackerSheet = ReminderWorkbook.Sheets("Tracker")
    
cc = ReminderFormSheet.Range("C6")
layout = "/ROBOTREMIND"
path = "C:\Reminders\temp\"
temp = "temp.xlsx"

Dim data As Date
Dim reference As String

Dim lr As Long
lr = Cells(ReminderFormSheet.Rows.Count, "B").End(xlUp).Row
Dim cell As Range


' === Connection to SAP === '
On Error Resume Next
    If Not IsObject(SAP) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SAP = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
        Set Connection = SAP.Children(0)
    End If
    If Not IsObject(session) Then
        Set session = Connection.Children(0)
    End If
        If IsObject(WScript) Then
         WScript.ConnectObject session, "on"
         WScript.ConnectObject SAP, "on"
    End If
If Err.Number <> 0 Then
    MsgBox "Please connect to SAP!"
    Exit Sub
End If

' === master loop, to verify status of invoices given by the User === '
For Each cell In ReminderFormSheet.Range("B10:B" & lr)
    
    Dim lastRow As Long
    lastRow = ReminderTrackerSheet.Cells(ReminderTrackerSheet.Rows.Count, "B").End(xlUp).Row
    Dim unusedRow As Long
    unusedRow = lastRow + 1
        
    Dim dfb, CaseNo, dfp As Variant
    CaseNo = ReminderTrackerSheet.Cells(unusedRow, 1).Value
    
    
    ' === VIM ANALYTICS 2 transaction (VIM_VA2) === '
    session.findById("wnd[0]").resizeWorkingPane 132, 25, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n/opt/vim_va2" ' Transaction '
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").Text = cell.Value ' Date of invoice '
    session.findById("wnd[0]/usr/txtS_XBLNR-LOW").Text = cell.Offset(0, 1).Value ' Reference / Invoice number '
    session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtS_LIFNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").Text = cc ' Company Code '
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = layout ' Layout of raport '
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").SetFocus
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Scenario 7: Invoice not in VIM '
    errMessage = session.findById("wnd[0]/sbar").Text
    If errMessage = "No data found for specified select-option/parameter" Then
        ReminderTrackerSheet.Cells(unusedRow, "B").Value = cc ' Company Code '
        ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
        ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
        ReminderTrackerSheet.Cells(unusedRow, "E").Value = cell.Value ' invoice date '
        ReminderTrackerSheet.Cells(unusedRow, "F").Value = cell.Offset(0, 1).Value ' reference / invoice number '
        ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: invoice missing from SAP."  ' Comment '
        ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
        ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
        ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
            ' ==== sending e-mails part ==== '
            Dim mail1 As Object ' Add a variable to hold the mail object
            mail.AddOutlookReference
            mail.AddWordReference
                Set mail1 = CreateObject("mail.MissInv") ' Create and set the mail object
            mail.MissInv unusedRow, CaseNo
                Set mail1 = Nothing ' Release the mail objects
            mail.RemoveOutlookReference
            mail.RemoveWordReference
        MsgBox "Invoice missing from VIM. Contact supplier. Case registered under number: " & CaseNo, , "Result" ' message '
        
    Else ' invoice in SAP, extract to temp file '
    
    session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&XXL"
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = path ' path to the folder '
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = temp ' file name '
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[11]").press

    ' === open temp file, copy data to tracker === '
        Workbooks.Open path & temp
    
        ' Set Temp workbook and sheet references
           Dim TempSheet As Worksheet
           Dim TempWorkbook As Workbook
           Set TempWorkbook = Workbooks("temp.xlsx")
           Set TempSheet = TempWorkbook.Sheets(1)
    
        ' = verification of VIM status = '
        Dim cellValue As String
        Dim cellToCheck, role, AppFN, AppLN, clearingDoc As Range
        Set role = TempSheet.Range("F2") ' current role '
        Set cellToCheck = TempSheet.Range("M2") ' VIM Status '
        Set AppFN = TempSheet.Range("P2") ' First Name of Approver '
        Set AppLN = TempSheet.Range("Q2") ' Last Name of Approver '
        Set clearingDoc = TempSheet.Range("U2") ' Clearing Doc '
        
        If (cellToCheck = "Indexed" Or cellToCheck = "Created" Or cellToCheck = "Parked") And (role = "" Or role = "PO_AP_PROC" Or role = "NPO_AP_PROC" Or role = "AP_PROCESSOR") Then
        ' - scenario 1 - '
            ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("A2").Value ' Company Code '
            ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
            ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
            ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("I2").Value ' Invoice Date '
            ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("J2").Value ' Invoice no '
            ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("G2").Value ' Vendor no '
            ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("H2").Value ' Vendor name '
            ReminderTrackerSheet.Cells(unusedRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(unusedRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice registered in the system, proceed with processing it in VIM." ' Comment '
            ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
            ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice received in system, proceed with processing in VIM. Case registered under number: " & CaseNo, , "Result" ' message '
            
        ElseIf cellToCheck = "Rejected by Approver" Or cellToCheck = "Blocked" Or cellToCheck = "Approval Recalled" Then
        ' - scenario 2 - '
            ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("A2").Value ' Company Code '
            ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
            ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
            ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("I2").Value ' Invoice Date '
            ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("J2").Value ' Invoice no '
            ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("G2").Value ' Vendor no '
            ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("H2").Value ' Vendor name '
            ReminderTrackerSheet.Cells(unusedRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(unusedRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice not fully processed in system, validate the case manually. Current approver: " & AppFN & " " & AppLN ' Comment '
            ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
            ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice not fully processed in system. Current status: " & cellToCheck & ". Case registered under number: " & CaseNo, , "Result" ' message '
            
        ElseIf (cellToCheck = "Awaiting Approval - Parked" Or cellToCheck = "Sent for Doc Creation") Or (role = "RECEIVER" Or role = "REQUISITIONER" Or role = "PO_BUYER" Or role = "BUYER" Or role = "Z_MANAGER" Or role = "INFO_PROVIDER") Then
        ' - scenario 3 - '
            ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("A2").Value ' Company Code '
            ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
            ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
            ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("I2").Value ' Invoice Date '
            ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("J2").Value ' Invoice no '
            ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("G2").Value ' Vendor no '
            ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("H2").Value ' Vendor name '
            ReminderTrackerSheet.Cells(unusedRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(unusedRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Approval is undergoing. Current approver: " & AppFN & " " & AppLN ' Comment '
            ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
            ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
            ' ==== sending e-mails part ==== '
                ApprovalMail.AddOutlookReference
                ApprovalMail.AddWordReference
                    Set mail1 = CreateObject("mail.MissInv") ' Create and set the mail object
                ApprovalMail.MissAppr unusedRow, CaseNo, AppFN, AppLN
                    Set mail1 = Nothing ' Release the mail object
                ApprovalMail.RemoveOutlookReference
                ApprovalMail.RemoveWordReference
            MsgBox "Invoice is pending approval in VIM. Case registered under number: " & CaseNo, , "Result" ' message '
            
        ElseIf cellToCheck = "Approval Complete" Then
        ' - scenario 4 - '
            ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("A2").Value ' Company Code '
            ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
            ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
            ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("I2").Value ' Invoice Date '
            ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("J2").Value ' Invoice no '
            ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("G2").Value ' Vendor no '
            ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("H2").Value ' Vendor name '
            ReminderTrackerSheet.Cells(unusedRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(unusedRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Approval is completed - post invoice via VIM." ' Comment '
            ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
            ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice is approved and can be posted in SAP. Case registered under number: " & CaseNo, , "Result" ' message '
   
        ElseIf cellToCheck = "Posted" Then
        ' - scenario 5 - '
            ReminderTrackerSheet.Cells(unusedRow, "B").Value = TempSheet.Range("A2").Value ' Company Code '
            ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
            ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
            ReminderTrackerSheet.Cells(unusedRow, "E").Value = TempSheet.Range("I2").Value ' Invoice Date '
            ReminderTrackerSheet.Cells(unusedRow, "F").Value = TempSheet.Range("J2").Value ' Invoice no '
            ReminderTrackerSheet.Cells(unusedRow, "G").Value = TempSheet.Range("G2").Value ' Vendor no '
            ReminderTrackerSheet.Cells(unusedRow, "H").Value = TempSheet.Range("H2").Value ' Vendor name '
            ReminderTrackerSheet.Cells(unusedRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(unusedRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(unusedRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(unusedRow, "M").Value = TempSheet.Range("S2").Value ' End date '
            ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
            ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status
            
            If clearingDoc = "" Then
            ' invoice posted finally, but not yet paid  '
                If ReminderTrackerSheet.Cells(unusedRow, "N").Value < 8 Then
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: invoice posted, should be paid out with next payment run." ' Comment '
                Else
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: invoice posted, not paid. Verify the VMD/booking data." ' Comment '
                End If
                dfb = ReminderTrackerSheet.Cells(unusedRow, 14).Value
                MsgBox "Invoice posted finally in SAP " & dfb & " days ago. Case registered under number: " & CaseNo, , "Result" ' message '
            ElseIf (clearingDoc <= 2000000000# Or clearingDoc >= 2099999999#) And cc = "5" Then
            ' invoice cleared manually 5'
                    ReminderTrackerSheet.Cells(unusedRow, "O").Value = TempSheet.Range("T2").Value ' payment/clearing date '
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice cleared manually. Verify the case and respond to supplier." ' Comment '
                            dfb = ReminderTrackerSheet.Cells(unusedRow, 14).Value
                            dfp = ReminderTrackerSheet.Cells(unusedRow, 16).Value
                    MsgBox "Invoice cleared manually in SAP " & dfp & "days ago. Check in SAP. Case registered under number: " & CaseNo, , "Result" ' message '
            ElseIf (clearingDoc <= 1500000000# Or clearingDoc >= 1599999999#) And cc = "6" Then
            ' invoice cleared manually 6'
                    ReminderTrackerSheet.Cells(unusedRow, "O").Value = TempSheet.Range("T2").Value ' payment/clearing date '
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice cleared manually. Verify the case and respond to supplier." ' Comment '
                            dfb = ReminderTrackerSheet.Cells(unusedRow, 14).Value
                            dfp = ReminderTrackerSheet.Cells(unusedRow, 16).Value
                    MsgBox "Invoice cleared manually in SAP " & dfp & "days ago. Check in SAP. Case registered under number: " & CaseNo, , "Result" ' message '
            Else
            ' invoice paid '
                    ReminderTrackerSheet.Cells(unusedRow, "O").Value = TempSheet.Range("T2").Value ' payment/clearing date '
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice paid." ' Comment '
                            dfb = ReminderTrackerSheet.Cells(unusedRow, 14).Value
                            dfp = ReminderTrackerSheet.Cells(unusedRow, 16).Value
                    MsgBox "Invoice paid " & dfp & " days ago. Case registered under number: " & CaseNo, , "Result" ' message '
            End If
            
        Else ' Obsolete invoice '
        ' - scenario 6 - '
        
            Workbooks(temp).Close SaveChanges:=False
            Kill "\\oe-divisions\OBSE\Shared Resources\AP\ROBOT\Reminders\temp\temp.xlsx"
            
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            session.findById("wnd[0]/usr/txtS_XBLNR-LOW").Text = cell.Offset(0, 1).Value & "*"
            session.findById("wnd[0]/usr/txtS_XBLNR-LOW").SetFocus
            session.findById("wnd[0]/usr/txtS_XBLNR-LOW").caretPosition = 11
            session.findById("wnd[0]/tbar[1]/btn[8]").press
            session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").setCurrentCell -1, "OVERALL_STATUS_TEXT"
            session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").firstVisibleColumn = "DOC_DATE"
            session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").selectColumn "OVERALL_STATUS_TEXT"
            session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").selectedRows = ""
            session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").pressToolbarButton "&MB_FILTER"
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").Select
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "Obsolete"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").Text = "Deleted"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").Text = "Confirmed Duplicate"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").Text = "Suspected Duplicate"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").SetFocus
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").caretPosition = 19
            session.findById("wnd[2]/tbar[0]/btn[8]").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            If errMessage = "No data found for specified select-option/parameter" Then
                MsgBox "ERROR! Invoice deleted from the system, no new invoice scanned. Verify manually.", , "Result" ' message '
            Else
                session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").pressToolbarContextButton "&MB_EXPORT"
                session.findById("wnd[0]/usr/cntlCCTRL_MAIN/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&XXL"
                session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Reminders\temp\" ' path to the folder '
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "temp.xlsx" ' file name '
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
                session.findById("wnd[1]/tbar[0]/btn[11]").press
                
                Workbooks.Open "C:\Reminders\temp\temp.xlsx"
                Dim temp2 As Workbook
                Dim temp2s As Worksheet
                Set temp2 = Workbooks("temp.xlsx")
                Set temp2s = temp2.Sheets(1)
                If temp2.Range("A2").Value = "" Then
                    ReminderTrackerSheet.Cells(unusedRow, "B").Value = cc ' Company Code '
                    ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
                    ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
                    ReminderTrackerSheet.Cells(unusedRow, "E").Value = cell.Value ' invoice date '
                    ReminderTrackerSheet.Cells(unusedRow, "F").Value = cell.Offset(0, 1).Value ' reference / invoice number '
                    ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Invoice deleted from the system, no new invoice scanned."  ' Comment '
                    ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
                    ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
                    ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
                    MsgBox "ERROR! Invoice deleted from the system, no new invoice scanned. Verify manually. Case registered under number: " & CaseNo, , "Result" ' message '
                Else
                ReminderTrackerSheet.Cells(unusedRow, "B").Value = temp2s.Range("A2").Value ' Company Code '
                ReminderTrackerSheet.Cells(unusedRow, "C").Value = ReminderFormSheet.Range("C7").Value 'Reminder Date'
                ReminderTrackerSheet.Cells(unusedRow, "D").Value = ReminderFormSheet.Range("C8").Value 'Dunning level'
                ReminderTrackerSheet.Cells(unusedRow, "E").Value = temp2s.Range("I2").Value ' Invoice Date '
                ReminderTrackerSheet.Cells(unusedRow, "F").Value = temp2s.Range("J2").Value ' Invoice no '
                ReminderTrackerSheet.Cells(unusedRow, "G").Value = temp2s.Range("G2").Value ' Vendor no '
                ReminderTrackerSheet.Cells(unusedRow, "H").Value = temp2s.Range("H2").Value ' Vendor name '
                ReminderTrackerSheet.Cells(unusedRow, "I").Value = temp2s.Range("R2").Value ' Invoice scan date '
                ReminderTrackerSheet.Cells(unusedRow, "J").Value = temp2s.Range("N2").Value ' Due date '
                ReminderTrackerSheet.Cells(unusedRow, "Q").Value = temp2s.Range("M2").Value ' VIM Status '
                ReminderTrackerSheet.Cells(unusedRow, "M").Value = temp2.Range("S2").Value ' End date '
                ReminderTrackerSheet.Cells(unusedRow, "O").Value = temp2.Range("T2").Value ' payment/clearing date '
                ReminderTrackerSheet.Cells(unusedRow, "T").Value = Date & " Macro: Original document deleted from the system. New docID found." ' Comment '
                ReminderTrackerSheet.Cells(unusedRow, "U").Value = Date ' Reminder Receive Date '
                ReminderTrackerSheet.Cells(unusedRow, "V").Value = Date ' Case Update Date '
                ReminderTrackerSheet.Cells(unusedRow, "W").Value = "In progress" ' Status '
                MsgBox "Original document deleted. New docID found. Current status: " & temp2s.Range("M2").Value & ". Validate if more steps needed. Case registered under number: " & CaseNo, , "Result" ' message '
                End If
            End If
        End If
    
    End If
    
    ' === Close and clean up === '
    TempWorkbook.Close SaveChanges:=False
    
    ' Enable screen updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
Next cell
   
    ' Disconnect from SAP
    session.Disconnect
    
    ' Clean up objects
    Set session = Nothing
    Set Connection = Nothing
    Set SAP = Nothing
    
End Sub
