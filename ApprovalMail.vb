Option Explicit

Sub MissAppr(ByVal unusedRow As Long, ByVal CaseNo As Long, ByVal AppFN As String, ByVal AppLN As String)
    
    On Error GoTo ErrorHandler
    
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    
    Dim ReminderWorkbook As Workbook
    Dim ReminderTrackerSheet As Worksheet
    Dim ReminderFormSheet As Worksheet
    
    ' Set ReminderValidator workbook and sheet references
    Set ReminderWorkbook = Workbooks("Reminders Validator.xlsm")
    Set ReminderFormSheet = ReminderWorkbook.Sheets("Form")
    Set ReminderTrackerSheet = ReminderWorkbook.Sheets("Tracker")
        
    Application.EnableEvents = False
        
    ' Create the Outlook objects
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
        
    Dim sTemplatePath, sCompanyName, sToEmail, sSubject, sBody, sInvoiceDate, sTemplateFile, sInvoiceNo, restMail As String
        
    sTemplatePath = "C:\Reminders\Mail template\"
    sCompanyName = ReminderFormSheet.Range("C6").Value
        
    Dim fileContents, filePath As String
    
    ' Find the corresponding email and template in the data sheet
    Dim dataSheet As Worksheet
    Set dataSheet = ReminderWorkbook.Sheets("Data")
    
    Dim companyRange As Range
    Set companyRange = dataSheet.Range("A2:A" & dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row)
    
    Dim matchIndex As Variant
    matchIndex = Application.Match(sCompanyName, companyRange, 0)
    
    If Not IsError(matchIndex) Then
        Dim emailColumn As Range
        Set emailColumn = dataSheet.Range("D2:D" & dataSheet.Cells(dataSheet.Rows.Count, "D").End(xlUp).Row)
        sToEmail = AppFN & "." & AppLN & "@company.com"
    
        Dim templateColumn As Range
        Set templateColumn = dataSheet.Range("F2:F" & dataSheet.Cells(dataSheet.Rows.Count, "F").End(xlUp).Row)
        sTemplateFile = sTemplatePath & templateColumn.Cells(matchIndex).Value
    Else
        ' Handle case when company name is not found
        MsgBox "Error!!!  " & Erl & vbCrLf & Err.Description
        Exit Sub
    End If
    
        
        
    ' Open the file and copy the contents into the variable
    If Dir(sTemplateFile) <> "" Then
        Dim wordApp As Object
        Set wordApp = CreateObject("Word.Application")
        
        Dim wordDoc As Object
        Set wordDoc = wordApp.Documents.Open(sTemplateFile)
        
        ' Copy the entire content of the Word document
        wordDoc.Range.Copy
        
        ' Paste the content into the email body
        sInvoiceDate = ReminderTrackerSheet.Cells(unusedRow, "E").Value
        sInvoiceNo = ReminderTrackerSheet.Cells(unusedRow, "F").Value
        restMail = vbCrLf & "Invoice date = " & sInvoiceDate & vbCrLf & "Invoice no = " & sInvoiceNo
        sBody = wordDoc.Range.Text & restMail
        
        ' Close the Word document
        wordDoc.Close SaveChanges:=False
        wordApp.Quit
        Set wordDoc = Nothing
        Set wordApp = Nothing
    End If
    
    sSubject = sCompanyName & " | Payment reminder (#" & CaseNo & ") | Invoice " & ReminderTrackerSheet.Cells(unusedRow, "F").Value
    
    ' Compose the email
    With OutlookMail
        .To = sToEmail
        .Subject = sSubject
        .GetInspector.WordEditor.Range.Text = sBody
        .Display
    End With
        
    ' Release the Outlook objects
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
            
    ' Enable events again
    Application.EnableEvents = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error occurred at line: " & Erl & vbCrLf & Err.Description
    
    ' Release the Outlook objects (if they were created)
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
    ' Enable events again
    Application.EnableEvents = True
End Sub

Sub AddOutlookReference()
    Dim ref As Object
    
    ' Check if the reference already exists
    On Error Resume Next
    Set ref = ThisWorkbook.VBProject.References("Outlook")
    On Error GoTo 0
    
    ' Add the reference if it doesn't exist
    If ref Is Nothing Then
        ThisWorkbook.VBProject.References.AddFromGuid _
            GUID:="{00062FFF-0000-0000-C000-000000000046}", _
            Major:=9, Minor:=0
    End If
End Sub

Sub RemoveOutlookReference()
    Dim ref As Object
    
    ' Check if the reference exists
    On Error Resume Next
    Set ref = ThisWorkbook.VBProject.References("Outlook")
    On Error GoTo 0
    
    ' Remove the reference if it exists
    If Not ref Is Nothing Then
        ThisWorkbook.VBProject.References.Remove ref
    End If
End Sub

Sub AddWordReference()
    Dim ref As Object
    
    ' Check if the reference already exists
    On Error Resume Next
    Set ref = ThisWorkbook.VBProject.References("Word")
    On Error GoTo 0
    
    ' Add the reference if it doesn't exist
    If ref Is Nothing Then
        ThisWorkbook.VBProject.References.AddFromGuid _
            GUID:="{00020905-0000-0000-C000-000000000046}", _
            Major:=8, Minor:=7
    End If
End Sub

Sub RemoveWordReference()
    Dim ref As Object
    
    ' Check if the reference exists
    On Error Resume Next
    Set ref = ThisWorkbook.VBProject.References("Word")
    On Error GoTo 0
    
    ' Remove the reference if it existsa
    If Not ref Is Nothing Then
        ThisWorkbook.VBProject.References.Remove ref
    End If
End Sub




