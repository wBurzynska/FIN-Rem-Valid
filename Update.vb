Sub north()
'This Macro verifies the invoice status in SAP (If posted/paid), for Company Codes in North region (4012, 4018, 4019, 4027, 4031 or 4036)

Dim CaseNo As Integer
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
    
CaseNo = ReminderFormSheet.Range("J6").Value
layout = "/REMIND_RPA"
path = "C:\Reminders\temp\"
temp = "temp.xlsx"

Dim InvNo, CoCo As String
Dim InvDate As Date

Dim lastRow, caseRow As Long
Dim i As Long
lastRow = ReminderTrackerSheet.Cells(Rows.Count, "A").End(xlUp).Row

For i = 1 To lastRow
    If ReminderTrackerSheet.Range("A" & i).Value = CaseNo Then
        ' IF found
        CoCo = ReminderTrackerSheet.Range("B" & i).Value
        InvDate = ReminderTrackerSheet.Range("E" & i).Value
        InvNo = ReminderTrackerSheet.Range("F" & i).Value
        caseRow = i
        Exit For ' End loop
    End If
Next i

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

        Dim SAPDate As String
        SAPDate = Format(InvDate, "DD.MM.YYYY")
    
' === FBL1N transaction (Vendor Line Items) === '
    session.findById("wnd[0]").resizeWorkingPane 132, 25, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n FBL1N" '#transaction'
    session.findById("wnd[0]").sendVKey 0 '#execute'
    session.findById("wnd[0]/tbar[1]/btn[16]").press '#dynamic selection'
    session.findById("wnd[0]/usr/radX_AISEL").Select
    session.findById("wnd[0]/usr/chkX_SHBV").Selected = True
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN013-LOW").Text = SAPDate 'date of document'
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN015-LOW").Text = InvNo 'reference'
    session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKD_LIFNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").Text = CoCo 'Company Code provided in cell C6'
    session.findById("wnd[0]/usr/ctxtPA_VARI").Text = layout 'Layout of report - /Remind_RPA'
    session.findById("wnd[0]/usr/ctxtPA_VARI").SetFocus
    session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 11
    session.findById("wnd[0]/tbar[1]/btn[8]").press ' execute '
    
    errMessage = session.findById("wnd[0]/sbar").Text
    If errMessage = "No items selected (see long text)" Then
    
    ' case: invoice not in SAP '
            ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: invoice missing from SAP." '
            ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice not final posted in SAP. Check Basware.", , "Result" ' message '
        
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
        
            Dim dfb, dfp As Variant
                
        If Len(Trim(cellToCheck.Value)) = 0 Then
        ' case: invoice not paid
                        
                ' Copy data from Temp sheet to Tracker sheet
                ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
                ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
                
                If ReminderTrackerSheet.Cells(caseRow, "N").Value < 8 Then
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: invoice posted, should be paid out with next payment run." ' Comment '
                Else
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: invoice posted, not paid. Verify the VMD/booking data." ' Comment '
                End If
                dfb = ReminderTrackerSheet.Cells(caseRow, 14).Value
                MsgBox "Invoice posted finally in SAP " & dfb & " days ago.", , "Result" ' message '
                
        Workbooks(temp).Close SaveChanges:=False
        
        Else
        ' case: invoice paid '
        
                ' Copy data from Temp sheet to Tracker sheet
                ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
                ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
                
                If TempSheet.Range("R2").Value <= 3200000000# Or TempSheet.Range("R2").Value >= 3299999999# Then
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice cleared manually. Verify the case and respond to supplier." ' Comment '
                                dfp = ReminderTrackerSheet.Cells(caseRow, 16).Value
                    MsgBox "Invoice cleared manually in SAP " & dfp & " days ago. Check in SAP.", , "Result" ' message '
                Else
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice paid." ' Comment '
                                dfp = ReminderTrackerSheet.Cells(caseRow, 16).Value
                    MsgBox "Invoice paid " & dfp & " days ago.", , "Result" ' message '
                End If
                    
        Workbooks(temp).Close SaveChanges:=False
                
        End If
        
    End If
        
End Sub

Sub sbc()

Dim CaseNo As Integer
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
    
CaseNo = ReminderFormSheet.Range("J6").Value
layout = "/REMINDERS"
path = "C:\Reminders\temp\"
temp = "temp.xlsx"

Dim InvNo, CoCo As String
Dim InvDate As Date

Dim lastRow, caseRow As Long
Dim i As Long
lastRow = ReminderTrackerSheet.Cells(Rows.Count, "A").End(xlUp).Row

For i = 1 To lastRow
    If ReminderTrackerSheet.Range("A" & i).Value = CaseNo Then
        ' IF found
        CoCo = ReminderTrackerSheet.Range("B" & i).Value
        InvDate = ReminderTrackerSheet.Range("E" & i).Value
        InvNo = ReminderTrackerSheet.Range("F" & i).Value
        caseRow = i
        Exit For ' End loop
    End If
Next i


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

    ' Disable screen updating and alerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim dfb, dfp As Variant


    Dim unusedRow As Long
    unusedRow = caseRow
    
        Dim SAPDate As String
        SAPDate = Format(InvDate, "DD.MM.YYYY")
    
        
    ' === VIM ANALYTICS 2 transaction (VIM_VA2) === '
    session.findById("wnd[0]").resizeWorkingPane 132, 25, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n/opt/vim_va2" ' Transaction '
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").Text = SAPDate ' Date of invoice '
    session.findById("wnd[0]/usr/txtS_XBLNR-LOW").Text = InvNo ' Reference / Invoice number '
    session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").Text = CoCo ' Company Code '
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = layout ' Layout of raport '
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").SetFocus
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Scenario 7: Invoice not in VIM '
    errMessage = session.findById("wnd[0]/sbar").Text
    If errMessage = "No data found for specified select-option/parameter" Then
        ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: invoice missing from SAP."  ' Comment '
        ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
        ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
        
        
        ' ==== sending e-mails part ==== '
        Dim mail1 As Object ' Add a variable to hold the mail object
        mail.AddOutlookReference
        mail.AddWordReference
            Set mail1 = CreateObject("mail.MissInv") ' Create and set the mail object
        mail.MissInv unusedRow, CaseNo
            Set mail1 = Nothing ' Release the mail object
        mail.RemoveOutlookReference
        mail.RemoveWordReference
        MsgBox "Invoice missing from VIM. Contact supplier.", , "Result" ' message '
        
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
            ReminderTrackerSheet.Cells(caseRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(caseRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(caseRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice registered in the system, proceed with processing it in VIM." ' Comment '
            ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice received in system, proceed with processing in VIM.", , "Result" ' message '
            
        ElseIf cellToCheck = "Rejected by Approver" Or cellToCheck = "Blocked" Or cellToCheck = "Approval recalled" Then
        ' - scenario 2 - '
            ReminderTrackerSheet.Cells(caseRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(caseRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(caseRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice not fully processed in system, validate the case manually." ' Comment '
            ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice not fully processed in system. Current status: " & cellToCheck, , "Result" ' message '
            
        ElseIf cellToCheck = "Awaiting Approval - Parked" Then
        ' - scenario 3 - '
            ReminderTrackerSheet.Cells(caseRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(caseRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(caseRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Approval is undergoing. Current approver: " & AppFN & " " & AppLN ' Comment '
            ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
                 ' ==== sending e-mails part ==== '
                ApprovalMail.AddOutlookReference
                ApprovalMail.AddWordReference
                    Set mail1 = CreateObject("ApprovalMail.MissAppr") ' Create and set the mail object
                ApprovalMail.MissAppr unusedRow, CaseNo, AppFN, AppLN
                    Set mail1 = Nothing ' Release the mail object
                ApprovalMail.RemoveOutlookReference
                ApprovalMail.RemoveWordReference
            MsgBox "Invoice is pending approval in VIM.", , "Result" ' message '
            
        ElseIf cellToCheck = "Approval Complete" Then
        ' - scenario 4 - '
            ReminderTrackerSheet.Cells(caseRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(caseRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(caseRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Approval is completed - post invoice via VIM." ' Comment '
            ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice is approved and can be posted in SAP.", , "Result" ' message '
   
        ElseIf cellToCheck = "Posted" Then
        ' - scenario 5 - '
            ReminderTrackerSheet.Cells(caseRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(caseRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(caseRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(caseRow, "M").Value = TempSheet.Range("S2").Value ' End date '
            ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status
            
            If clearingDoc = "" Then
            ' invoice posted finally, but not yet paid  '
                If ReminderTrackerSheet.Cells(caseRow, "N").Value < 8 Then
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: invoice posted, should be paid out with next payment run." ' Comment '
                Else
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: invoice posted, not paid. Verify the VMD/booking data." ' Comment '
                End If
                    dfb = ReminderTrackerSheet.Cells(caseRow, 14).Value
                MsgBox "Invoice posted finally in SAP " & dfb & " days ago.", , "Result" ' message '
            ElseIf clearingDoc <= 3200000000# Or clearingDoc >= 3299999999# Then
            ' invoice cleared manually '
                    ReminderTrackerSheet.Cells(caseRow, "O").Value = TempSheet.Range("T2").Value ' payment/clearing date '
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice cleared manually. Verify the case and respond to supplier." ' Comment '
                        dfp = ReminderTrackerSheet.Cells(caseRow, 16).Value
                    MsgBox "Invoice cleared manually in SAP " & dfp & " days ago. Check in SAP.", , "Result" ' message '
            Else
            ' invoice paid '
                    ReminderTrackerSheet.Cells(caseRow, "O").Value = TempSheet.Range("T2").Value ' payment/clearing date '
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice paid." ' Comment '
                        dfp = ReminderTrackerSheet.Cells(caseRow, 16).Value
                    MsgBox "Invoice paid " & dfp & " days ago.", , "Result" ' message '
            End If
            
        Else ' Obsolete invoice '
        ' - scenario 6 - '
        
            Workbooks(temp).Close SaveChanges:=False
            Kill "C:\Reminders\temp\temp.xlsx"
            
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            session.findById("wnd[0]/usr/txtS_XBLNR-LOW").Text = InvNo & "*"
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
                session.findById("wnd[1]/usr/ctxtDY_PATH").Text = path = "C:\Reminders\temp\" ' path to the folder '
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "temp.xlsx" ' file name '
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
                session.findById("wnd[1]/tbar[0]/btn[11]").press
                
                Workbooks.Open path = "C:\Reminders\temp\temp.xlsx"
                Dim temp2 As Workbook
                Dim temp2s As Worksheet
                Set temp2 = Workbooks("temp.xlsx")
                Set temp2s = temp2.Sheets(1)
                If temp2.Range("A2").Value = "" Then
                    ReminderTrackerSheet.Cells(caseRow, "B").Value = cc ' Company Code '
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice deleted from the system, no new invoice scanned."  ' Comment '
                    ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
                    ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
                    MsgBox "ERROR! Invoice deleted from the system, no new invoice scanned. Verify manually. Case registered under number: " & CaseNo, , "Result" ' message '                Else
                ReminderTrackerSheet.Cells(caseRow, "B").Value = temp2s.Range("A2").Value ' Company Code '
                ReminderTrackerSheet.Cells(caseRow, "E").Value = temp2s.Range("I2").Value ' Invoice Date '
                ReminderTrackerSheet.Cells(caseRow, "F").Value = temp2s.Range("J2").Value ' Invoice no '
                ReminderTrackerSheet.Cells(caseRow, "G").Value = temp2s.Range("G2").Value ' Vendor no '
                ReminderTrackerSheet.Cells(caseRow, "H").Value = temp2s.Range("H2").Value ' Vendor name '
                ReminderTrackerSheet.Cells(caseRow, "I").Value = temp2s.Range("R2").Value ' Invoice scan date '
                ReminderTrackerSheet.Cells(caseRow, "J").Value = temp2s.Range("N2").Value ' Due date '
                ReminderTrackerSheet.Cells(caseRow, "Q").Value = temp2s.Range("M2").Value ' VIM Status '
                ReminderTrackerSheet.Cells(caseRow, "M").Value = temp2.Range("S2").Value ' End date '
                ReminderTrackerSheet.Cells(caseRow, "O").Value = temp2.Range("T2").Value ' payment/clearing date '
                ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Original document deleted from the system. New docID found." ' Comment '
                ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
                ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
                MsgBox "Original document deleted. New docID found. Current status: " & temp2s.Range("M2").Value & ". Validate if more steps needed.", , "Result"  ' message '
                End If
            End If
        End If
    
    End If
    
    ' === Close and clean up === '
    TempWorkbook.Close SaveChanges:=False
    
    ' Enable screen updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Disconnect from SAP
    session.Disconnect
    
    ' Clean up objects
    Set session = Nothing
    Set Connection = Nothing
    Set SAP = Nothing

End Sub




Sub mbc()


Dim CaseNo As Integer
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
    
CaseNo = ReminderFormSheet.Range("J6").Value
layout = "/ROBOTREMIND"
path = "C:\Reminders\temp\"
temp = "temp.xlsx"

Dim InvNo, CoCo As String
Dim InvDate As Date

Dim lastRow, caseRow As Long
Dim i As Long
lastRow = ReminderTrackerSheet.Cells(Rows.Count, "A").End(xlUp).Row

For i = 1 To lastRow
    If ReminderTrackerSheet.Range("A" & i).Value = CaseNo Then
        ' IF found
        CoCo = ReminderTrackerSheet.Range("B" & i).Value
        InvDate = ReminderTrackerSheet.Range("E" & i).Value
        InvNo = ReminderTrackerSheet.Range("F" & i).Value
        caseRow = i
        Exit For ' End loop
    End If
Next i



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
    

    ' Disable screen updating and alerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

        
    Dim dfb, dfp As Variant
    
    Dim unusedRow As Long
    unusedRow = caseRow
    
        Dim SAPDate As String
        SAPDate = Format(InvDate, "DD.MM.YYYY")
    
    ' === VIM ANALYTICS 2 transaction (VIM_VA2) === '
    session.findById("wnd[0]").resizeWorkingPane 132, 25, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n/opt/vim_va2" ' Transaction '
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").Text = SAPDate ' Date of invoice '
    session.findById("wnd[0]/usr/txtS_XBLNR-LOW").Text = InvNo ' Reference / Invoice number '
    session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtS_LIFNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").Text = CoCo ' Company Code '
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = layout ' Layout of raport '
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").SetFocus
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Scenario 7: Invoice not in VIM '
    errMessage = session.findById("wnd[0]/sbar").Text
    If errMessage = "No data found for specified select-option/parameter" Then
        ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: invoice missing from SAP."  ' Comment '
        ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
        ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
            ' ==== sending e-mails part ==== '
            Dim mail1 As Object ' Add a variable to hold the mail object
            mail.AddOutlookReference
            mail.AddWordReference
                Set mail1 = CreateObject("mail.MissInv") ' Create and set the mail object
            mail.MissInv unusedRow, CaseNo
                Set mail1 = Nothing ' Release the mail object
            mail.RemoveOutlookReference
            mail.RemoveWordReference
        MsgBox "Invoice missing from VIM. Contact supplier.", , "Result" ' message '
        
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
            ReminderTrackerSheet.Cells(caseRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(caseRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(caseRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice registered in the system, proceed with processing it in VIM." ' Comment '
            ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice received in system, proceed with processing in VIM.", , "Result" ' message '
            
        ElseIf cellToCheck = "Rejected by Approver" Or cellToCheck = "Blocked" Or cellToCheck = "Approval Recalled" Then
        ' - scenario 2 - '
            ReminderTrackerSheet.Cells(caseRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(caseRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(caseRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice not fully processed in system, validate the case manually. Current approver: " & AppFN & " " & AppLN ' Comment '
            ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice not fully processed in system. Current status: " & cellToCheck & ". ", , "Result" ' message '
            
        ElseIf (cellToCheck = "Awaiting Approval - Parked" Or cellToCheck = "Sent for Doc Creation") Or (role = "RECEIVER" Or role = "REQUISITIONER" Or role = "PO_BUYER" Or role = "BUYER" Or role = "Z_MANAGER" Or role = "INFO_PROVIDER") Then
        ' - scenario 3 - '
            ReminderTrackerSheet.Cells(caseRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(caseRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(caseRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Approval is undergoing. Current approver: " & AppFN & " " & AppLN ' Comment '
            ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
            ' ==== sending e-mails part ==== '
                ApprovalMail.AddOutlookReference
                ApprovalMail.AddWordReference
                    Set mail1 = CreateObject("mail.MissInv") ' Create and set the mail object
                ApprovalMail.MissAppr unusedRow, CaseNo, AppFN, AppLN
                    Set mail1 = Nothing ' Release the mail object
                ApprovalMail.RemoveOutlookReference
                ApprovalMail.RemoveWordReference
            MsgBox "Invoice is pending approval in VIM. ", , "Result" ' message '
            
        ElseIf cellToCheck = "Approval Complete" Then
        ' - scenario 4 - '
            ReminderTrackerSheet.Cells(caseRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(caseRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(caseRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Approval is completed - post invoice via VIM." ' Comment '
            ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
            MsgBox "Invoice is approved and can be posted in SAP. ", , "Result" ' message '
   
        ElseIf cellToCheck = "Posted" Then
        ' - scenario 5 - '
            ReminderTrackerSheet.Cells(caseRow, "I").Value = TempSheet.Range("R2").Value ' Invoice scan date '
            ReminderTrackerSheet.Cells(caseRow, "J").Value = TempSheet.Range("N2").Value ' Due date '
            ReminderTrackerSheet.Cells(caseRow, "Q").Value = TempSheet.Range("M2").Value ' VIM Status '
            ReminderTrackerSheet.Cells(caseRow, "M").Value = TempSheet.Range("S2").Value ' End date '
            ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
            ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status
            
            If clearingDoc = "" Then
            ' invoice posted finally, but not yet paid  '
                If ReminderTrackerSheet.Cells(caseRow, "N").Value < 8 Then
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: invoice posted, should be paid out with next payment run." ' Comment '
                Else
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: invoice posted, not paid. Verify the VMD/booking data." ' Comment '
                End If
                    dfb = ReminderTrackerSheet.Cells(caseRow, 14).Value
                MsgBox "Invoice posted finally in SAP " & dfb & " days ago. ", , "Result" ' message '
            ElseIf (clearingDoc <= 2000000000# Or clearingDoc >= 2099999999#) And cc = "5" Then
            ' invoice cleared manually 4007'
                    ReminderTrackerSheet.Cells(caseRow, "O").Value = TempSheet.Range("T2").Value ' payment/clearing date '
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice cleared manually. Verify the case and respond to supplier." ' Comment '
                        dfp = ReminderTrackerSheet.Cells(caseRow, 16).Value
                    MsgBox "Invoice cleared manually in SAP " & dfp & "days ago. Check in SAP. ", , "Result" ' message '
            ElseIf (clearingDoc <= 1500000000# Or clearingDoc >= 1599999999#) And cc = "6" Then
            ' invoice cleared manually 4007'
                    ReminderTrackerSheet.Cells(caseRow, "O").Value = TempSheet.Range("T2").Value ' payment/clearing date '
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice cleared manually. Verify the case and respond to supplier." ' Comment '
                        dfp = ReminderTrackerSheet.Cells(caseRow, 16).Value
                    MsgBox "Invoice cleared manually in SAP " & dfp & "days ago. Check in SAP. ", , "Result" ' message '
            Else
            ' invoice paid '
                    ReminderTrackerSheet.Cells(caseRow, "O").Value = TempSheet.Range("T2").Value ' payment/clearing date '
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice paid." ' Comment '
                        dfp = ReminderTrackerSheet.Cells(caseRow, 16).Value
                    MsgBox "Invoice paid " & dfp & " days ago. ", , "Result" ' message '
            End If
            
        Else ' Obsolete invoice '
        ' - scenario 6 - '
        
            Workbooks(temp).Close SaveChanges:=False
            Kill "C:\Reminders\temp\temp.xlsx"
            
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            session.findById("wnd[0]/usr/txtS_XBLNR-LOW").Text = InvNo & "*"
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
                    ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Invoice deleted from the system, no new invoice scanned."  ' Comment '
                    ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
                    ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
                    MsgBox "ERROR! Invoice deleted from the system, no new invoice scanned. Verify manually. ", , "Result" ' message '
                Else
                ReminderTrackerSheet.Cells(caseRow, "G").Value = temp2s.Range("G2").Value ' Vendor no '
                ReminderTrackerSheet.Cells(caseRow, "H").Value = temp2s.Range("H2").Value ' Vendor name '
                ReminderTrackerSheet.Cells(caseRow, "I").Value = temp2s.Range("R2").Value ' Invoice scan date '
                ReminderTrackerSheet.Cells(caseRow, "J").Value = temp2s.Range("N2").Value ' Due date '
                ReminderTrackerSheet.Cells(caseRow, "Q").Value = temp2s.Range("M2").Value ' VIM Status '
                ReminderTrackerSheet.Cells(caseRow, "M").Value = temp2.Range("S2").Value ' End date '
                ReminderTrackerSheet.Cells(caseRow, "O").Value = temp2.Range("T2").Value ' payment/clearing date '
                ReminderTrackerSheet.Cells(caseRow, "T").Value = ReminderTrackerSheet.Cells(caseRow, "T").Value & " " & Date & " Macro: Original document deleted from the system. New docID found." ' Comment '
                ReminderTrackerSheet.Cells(caseRow, "U").Value = Date ' Reminder Receive Date '
                ReminderTrackerSheet.Cells(caseRow, "V").Value = Date ' Case Update Date '
                ReminderTrackerSheet.Cells(caseRow, "W").Value = "In progress" ' Status '
                MsgBox "Original document deleted. New docID found. Current status: " & temp2s.Range("M2").Value & ". Validate if more steps needed. ", , "Result" ' message '
                End If
            End If
        End If
    
    End If
    
    ' === Close and clean up === '
    TempWorkbook.Close SaveChanges:=False
    
    ' Enable screen updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
   
    ' Disconnect from SAP
    session.Disconnect
    
    ' Clean up objects
    Set session = Nothing
    Set Connection = Nothing
    Set SAP = Nothing

End Sub
