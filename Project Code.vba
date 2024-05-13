'*MICROSOFT ACCESS CLASS OBJECTS*

'Form Filter Form
Option Compare Database

Private Sub cmdTableFullForm_Click()
DoCmd.OpenForm "IE_Form"
DoCmd.Close acForm, "Filter_Form"
End Sub

Private Sub Form_Load()
DoCmd.Maximize
End Sub

'-----

'Form IE Form
Option Compare Database
Option Explicit

Private Sub Combo123_AfterUpdate()
DoCmd.SearchForRecord acDataForm, "IE_Form", acFirst, "[Project Name] = " & "'" & [Screen].[ActiveControl] & "'"

End Sub

Private Sub cmdImport_Click()
DoCmd.OpenForm "Import_Form"
DoCmd.Close acForm, "IE_Form"
End Sub

Private Sub cmdMainMenu_Click()
DoCmd.OpenForm "Main_Form"
DoCmd.Close acForm, "IE_Form"
End Sub

'Save to XLSX
Private Sub cmdSaveXLSX_Click()
Dim Filename As String
Dim filepath As String

Filename = "Whole Table"
filepath = "..\Documents\" & Filename & " " & Format(Date, "yyyy-mm-dd") & " User" & ".xlsx"
DoCmd.OutputTo acOutputForm, "Whole_Table", acFormatXLSX, filepath
MsgBox "Table have been saved successfully", vbInformation, "Save confirmed"
End Sub

Private Sub cmdTableFilterForm_Click()
DoCmd.OpenForm "Filter_Form"
DoCmd.Close acForm, "IE_Form"
End Sub

Private Sub Form_Load()
DoCmd.Maximize
End Sub

'-----

'Form IE Login Form
Option Compare Database
Option Explicit

'For login check for username and password
Private Sub cmdLogin_Click()
    On Error GoTo problem
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM tblLogin WHERE username = '" + Me.txtUsername.Value + "'"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If rs.EOF Then
        IncorrectUsernameStyle
        Exit Sub
    End If
    CorrectUsernameStyle
    
    rs.MoveFirst
    If rs("Password") <> Nz(Me.txtPassword, "") Then
        IncorrectPasswordStyle
        Exit Sub
    End If
    
    If DLookup("[Access]", "tblLogin", "[Username] = '" & Me.txtUsername.Value & "'") = "1" Then
    '"1" = Basic User
    '"2" = Advance User
        
    DoCmd.OpenForm "IE_Form"
    DoCmd.Close acForm, "Main_Form"
    DoCmd.Close acForm, "IE_Login_Form"
    
    Else
    If DLookup("[Access]", "tblLogin", "[Username] = '" & Me.txtUsername.Value & "'") = "2" Then
    '"1" = Basic User
    '"2" = Advance User
        
    DoCmd.OpenForm "IE_Form"
    DoCmd.Close acForm, "Main_Form"
    DoCmd.Close acForm, "IE_Login_Form"
    
    
    Else
    MsgBox "You do not have access.", vbInformation
    End If
    End If
    
problem:
    If Err.Number = 94 Then
        MsgBox "No Username & Password input.", vbInformation
        Me.txtUsername.SetFocus
    End If
    
End Sub

'Set Visible and colour for incorrect username
Private Sub IncorrectUsernameStyle()
    Me.lblIncorrectUsername.Visible = True
    Me.txtUsername.BorderColor = RGB(255, 0, 0)
    Me.txtUsername.SetFocus
End Sub

'Set Visible and colour for correct username
Private Sub CorrectUsernameStyle()
    Me.lblIncorrectUsername.Visible = False
    Me.txtUsername.BorderColor = RGB(0, 153, 153)
End Sub

'Set Visible and colour for incorrect password
Private Sub IncorrectPasswordStyle()
    Me.lblIncorrectPassword.Visible = True
    Me.txtPassword.BorderColor = RGB(255, 0, 0)
    Me.txtPassword.SetFocus
End Sub

'Cancel Login Form
Private Sub cmdExit_Click()
DoCmd.Close acForm, "IE_Login_Form"
DoCmd.OpenForm "Main_Form"
End Sub

Private Sub Form_Load()
DoCmd.Maximize

End Sub

'-----

'Form Import Form
Option Compare Database
Option Explicit

'Browse for excel spreadsheet
Private Sub cmdBrowse_Click()
    Dim diag As Office.FileDialog
    Dim item As Variant
    
    Set diag = Application.FileDialog(msoFileDialogFilePicker)
    diag.AllowMultiSelect = False
    diag.Title = "Please select an Excel Spreadsheet"
    diag.Filters.Clear
    diag.Filters.Add "Excel Spreadsheets", "*.xls, *.xlsx"
    
    If diag.Show Then
        For Each item In diag.SelectedItems
            Me.txtFileName = item
        Next
    End If
    
    'Enable Import Spreadsheet
    Me.cmdImportSpreadsheet.Enabled = True
End Sub

Private Sub cmdReset_Click()
Me.Refresh
End Sub




Private Sub cmdDelete_Click()

    'Question to delete imported file
    If MsgBox("Delete Current imported files without Saving to Main Table?", vbExclamation + vbYesNo) = vbYes Then
    
    'Delete all data from curently imported records
    Dim sql As String
    sql = "Delete * From tblImportExcel;"
    DoCmd.RunSQL sql
    'Delete txtFileName Selected
    Me.txtFileName.Value = Null
    
    'Disable Buttons
    Me.cmdSQL.Enabled = False
    Me.cmdDelete.Enabled = False
    
    Me.Refresh
    
    Else
    
    'Undo changes made and abort
    Me.Undo
    MsgBox "Aborted. Data not Saved.", vbInformation
    Me.Refresh
    
    End If
     
End Sub

Private Sub cmdImportSpreadsheet_Click()
    Dim fso As New FileSystemObject
    On Error GoTo problem
    
    'if no File Selected
    If Nz(Me.txtFileName, "") = "" Then
        MsgBox "Please Select a file!", vbInformation
        Exit Sub
    End If
    
    'Import excel file to access
    If fso.FileExists(Nz(Me.txtFileName, "")) Then
        ExcelImport.ImportExcelSpreadsheet Me.txtFileName, "Temp_File"

    'Question to import file
    If MsgBox("Import this " & fso.GetFileName(Me.txtFileName) & "Excel File ?", vbExclamation + vbYesNo) = vbYes Then
    
    'Copy Temp File to Temp_Form
    Dim sql As String
    sql = "INSERT INTO tblImportExcel SELECT * FROM Temp_File"
    DoCmd.RunSQL sql
    'Delete Temp File
    DoCmd.DeleteObject acTable, "Temp_File"
    
    'Enable and disable buttons
    Me.cmdImportSpreadsheet.Enabled = False
    Me.cmdSQL.Enabled = True
    Me.cmdDelete.Enabled = True
    
    Me.Refresh
    
    Else
    
    'Undo Changes made and abort
    Me.Undo
    MsgBox "Aborted. Data not Saved.", vbInformation
    'Delete Temp File
    DoCmd.DeleteObject acTable, "Temp_File"
    Me.Refresh
    
    End If
    
    Else
    
    MsgBox "File not Found!", vbInformation
        
    End If
    
    
problem:
    If Err.Number = 2501 Then
    
    'Undo Changes made and abort
    Me.Undo
        MsgBox "Aborted. Data not Saved.", vbInformation

    'Delete Temp File
    DoCmd.DeleteObject acTable, "Temp_File"
    
    Me.Refresh
    
    End If
End Sub

Private Sub cmdSQL_Click()

    If MsgBox("Save and Delete Currently imported files to Main Table?", vbExclamation + vbYesNo) = vbYes Then
    
    'Copy Temp File to main
    Dim sql As String
    sql = "INSERT INTO MainTable SELECT * FROM tblImportExcel"
    DoCmd.RunSQL sql
    'Delete Imported file
    sql = "Delete * From tblImportExcel;"
    DoCmd.RunSQL sql
    'Delete txtFileName Selected
    Me.txtFileName.Value = Null
    
    'Disable Buttons
    Me.cmdSQL.Enabled = False
    Me.cmdDelete.Enabled = False
    
    Me.Refresh
    
    Else
    
    'Undo Changes made and abort
    Me.Undo
    MsgBox "Aborted. Data not Saved.", vbInformation
    Me.Refresh
    
    End If
    
End Sub

Private Sub cmdTableFullForm_Click()
DoCmd.OpenForm "IE_Form"
DoCmd.Close acForm, "Import_Form"
End Sub

'Remove Warning from queries SQL
Private Sub Form_Load()
DoCmd.Maximize

   Application.SetOption "Confirm Action Queries", 0
   Application.SetOption "Confirm Document Deletions", 0
   Application.SetOption "Confirm Record Changes", 0
   
    'Delete Imported file
    Dim sql As String
    sql = "Delete * From tblImportExcel;"
    DoCmd.RunSQL sql
    Me.Refresh

End Sub

'Remove Warning from queries SQL
Private Sub Form_Unload(Cancel As Integer)
    Application.SetOption "Confirm Action Queries", 1
    Application.SetOption "Confirm Document Deletions", 1
    Application.SetOption "Confirm Record Changes", 1
End Sub

'-----

'Form Main Form
Option Compare Database

Private Sub cmdQuit_Click()
DoCmd.Quit
End Sub

Private Sub cmdRequest_Click()
DoCmd.OpenForm "Request_Form"
DoCmd.Close acForm, "Main_Form"
End Sub

Private Sub cmdIELogin_Click()
DoCmd.OpenForm "IE_Login_Form"
DoCmd.Close acForm, "Main_Form"
End Sub

Private Sub cmdUserLogin_Click()
DoCmd.OpenForm "User_Login_Form"
DoCmd.Close acForm, "Main_Form"
End Sub

Private Sub cmdWorkInstruction_Click()
   Dim LWordDoc As String
   Dim oApp As Object

   'Path to the word document
   LWordDoc = Application.CodeProject.Path & "\Work Instructions.docx"

   If Dir(LWordDoc) = "" Then
      MsgBox "Document not found."

   Else
      'Create an instance of MS Word
      Set oApp = CreateObject(Class:="Word.Application")
      oApp.Visible = True

      'Open the Document
      oApp.Documents.Open Filename:=LWordDoc
   End If

End Sub

Private Sub Form_Load()
DoCmd.Maximize

DoCmd.OpenForm "Main_Form"
End Sub

'-----

'Form Request Form
Option Compare Database
Private Saved As Boolean

Private Sub cmdClearRecord_Click()

'Question to clear record

If MsgBox("Do you want to clear this record?", vbQuestion + vbYesNo + vbDefaultButton1, "Clear Record?") = vbYes Then
Me.Undo
DoCmd.GoToRecord , , acNewRec
Me.cmdClearRecord.Enabled = False
Me.cmdSave.Enabled = False

Else

    MsgBox "Clear Aborted.", vbInformation

End If

End Sub

'Click to goto Main Form
Private Sub cmdMainForm_Click()
DoCmd.Close acForm, "Request_Form"
DoCmd.OpenForm "Main_Form"
End Sub

'Save to XLSX
Private Sub cmdSaveXLSX_Click()
On Error GoTo problem

Dim Filename As String
Dim filepath As String

Filename = "Request"
filepath = "..\Documents\" & Filename & " " & Format(Date, "yyyy-mm-dd") & " User" & ".xlsx"
DoCmd.OutputTo acOutputForm, "Request_Table", acFormatXLSX, filepath
MsgBox "Table have been saved successfully", vbInformation, "Save confirmed"

problem:
If Err.Number = 2302 Then
MsgBox "Unable to save." & vbCrLf & "File of the same name is currently opened in Excel." & vbCrLf & "Please close that Excel file first.", vbInformation
End If

End Sub

Private Sub cmdSendMailwithTable_Click()

On Error GoTo problem

'tools -> refrence -> Microsoft outlook
Dim olApp As Outlook.Application
Dim olMail As MailItem
Dim mailbody As String
Dim rs As DAO.Recordset

Dim Filename As String
Dim Currentpath As String
Dim filepath As String

' <br> used to insert a line ( press enter)
' create a table using html
' check the link below to know more about html tables
' http://www.w3schools.com/html/html_tables.asp
' html color code
'http://www.computerhope.com/htmcolor.htm or http://html-color-codes.info/
'bg color is used for background color
' font color is used for font color
'<b> bold the text  http://www.w3schools.com/html/html_formatting.asp
' &nbsp;  is used to give a single space between text
'<p style="font-size:15px">This is some text!</p> used to reduce for font size

'created header of table
   mailbody = "<TABLE Border=""1"", Cellspacing=""0""><TR>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Log Date &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> CAD Date &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Requestor &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Customer &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Rig/Project Name &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Part Number &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> PN Description &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Order/Job Type &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Sub Type &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Sales Order/Purchase Order &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Work/Service Order &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Serial Number/Equipment Number &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Receiving Memo # &nbsp;</p></Font></TD>" & _
   "<TD Bgcolor=""#FFFF00"", Align=""Center""><Font Color=#000000><b><p style=""font-size:15px""> Scope of Work/Note/Specification &nbsp;</p></Font></TD>" & _
      "</TR>"


'Add the data to the table
Set rs = CurrentDb.OpenRecordset("RequestTable", dbOpenDynaset)
rs.MoveFirst
Do While Not rs.EOF
       
              mailbody = mailbody & "<TR>" & _
               "<TD ><center>" & rs.Fields![Log Date].Value & "</TD>" & _
               "<TD><center>" & rs.Fields![CAD Date].Value & "</TD>" & _
               "<TD><center>" & rs.Fields![Requestor].Value & "</TD>" & _
               "<TD ><center>" & rs.Fields![Customer].Value & "</TD>" & _
               "<TD><center>" & rs.Fields![Rig/Project Name].Value & "</TD>" & _
               "<TD><center>" & rs.Fields![Part Number].Value & "</TD>" & _
               "<TD ><center>" & rs.Fields![PN Description].Value & "</TD>" & _
               "<TD><center>" & rs.Fields![Order/Job Type].Value & "</TD>" & _
               "<TD ><center>" & rs.Fields![Sub Type].Value & "</TD>" & _
               "<TD><center>" & rs.Fields![Sales Order/Purchase Order].Value & "</TD>" & _
               "<TD><center>" & rs.Fields![Work/Service Order].Value & "</TD>" & _
               "<TD><center>" & rs.Fields![Serial Number/Equipment Number].Value & "</TD>" & _
               "<TD><center>" & rs.Fields![Receiving Memo #].Value & "</TD>" & _
               "<TD><center>" & rs.Fields![Scope of Work/Note/Specification].Value & "</TD>" & _
                         "</TR>"
        
rs.MoveNext
Loop
rs.Close

Filename = "Request"
Currentpath = Application.CodeProject.Path
filepath = Currentpath & "\" & Filename & " " & Format(Date, "yyyy-mm-dd") & " User" & ".xlsx"
'create temporary file
DoCmd.OutputTo acOutputForm, "Request_Table", acFormatXLSX, filepath
If olApp Is Nothing Then
    Set olApp = New Outlook.Application
End If

Set olMail = olApp.CreateItem(olMailItem)
With olMail

' <br> used to insert a line ( press enter) and send email
Set olApp = New Outlook.Application
Set olMail = olApp.CreateItem(olMailItem)
With olMail
.To = ""
.CC = ""
.Subject = ""
.Attachments.Add filepath
.HTMLBody = mailbody & "</Table>"
.Display
'.Send
End With

Set olMail = Nothing
Set olApp = Nothing
'delete temporary file
Kill filepath
Exit Sub

End With

problem:
If Err.Number = 3021 Then
MsgBox "There are no records.", vbInformation
End If

End Sub

'When dirty/changes made set cmdSave and cmdClearRecord Enabled
Private Sub Form_Dirty(Cancel As Integer)
Me.cmdSave.Enabled = True

Me.cmdClearRecord.Enabled = True
End Sub

'Save Record
Private Sub cmdSave_Click()

'Question to save changes to record
If MsgBox("Do you want to save the changes on this record?", vbQuestion + vbYesNo + vbDefaultButton1, "Save Changes?") = vbYes Then

'user clicked yes
    Saved = True
    DoCmd.RunCommand (acCmdSaveRecord)
    Me.cmdSave.Enabled = False
    Me.cmdClearRecord.Enabled = False
    Saved = False
    MsgBox "Save Successful.", vbInformation

    DoCmd.GoToRecord , , acNewRec

    Me.Refresh
   
'user clicked no
Else
    Me.Undo
    MsgBox "Save Aborted.", vbInformation
    
End If
End Sub

'Save Form before closing
Private Sub Form_BeforeUpdate(Cancel As Integer)

'Question to save changes to record
Dim Response As Integer
If Saved = False Then
    Response = MsgBox("Do you want to save the changes on this record?", vbQuestion + vbYesNo + vbDefaultButton1, "Save Changes?")
    
    'If no
    If Response = vbNo Then
       Me.Undo
       MsgBox "Save Aborted.", vbInformation
       
    'if yes
    Else
       MsgBox "Save Successful.", vbInformation
       
    End If
    Me.cmdSave.Enabled = False
End If
End Sub

'Open to new record when form open
Private Sub Form_Load()
DoCmd.Maximize

DoCmd.GoToRecord , , acNewRec

    'Delete all data from records on load
    Dim sql As String
    sql = "Delete * From RequestTable;"
    DoCmd.RunSQL sql
    
'Remove Warning from queries SQL
Application.SetOption "Confirm Action Queries", 0
Application.SetOption "Confirm Document Deletions", 0
Application.SetOption "Confirm Record Changes", 0

Me.Refresh
End Sub

Private Sub Order_Job_Type_Change()

If Me.Order_Job_Type = "1" Then
Me.Order_Job_Type.Value = "New Build"

Else
If Me.Order_Job_Type = "2" Then
Me.Order_Job_Type.Value = "RGR"

Else
If Me.Order_Job_Type = "3" Then
Me.Order_Job_Type.Value = "Rental"

Else
If Me.Order_Job_Type = "4" Then
Me.Order_Job_Type.Value = "AO (Add-On)"

Else
If Me.Order_Job_Type = "5" Then
Me.Order_Job_Type.Value = "Conversion"

Else
If Me.Order_Job_Type = "6" Then
Me.Order_Job_Type.Value = "Rework"

Else
If Me.Order_Job_Type = "7" Then
Me.Order_Job_Type.Value = "Asset Mgmt"

End If
End If
End If
End If
End If
End If
End If

End Sub

Private Sub Sub_Type_Change()

If Me.Sub_Type = "1" Then
Me.Sub_Type.Value = "ATP (Assembly-Test-Paint)"

Else
If Me.Sub_Type = "2" Then
Me.Sub_Type.Value = "DCI (Dismantle-Clean-Inspection)"

Else
If Me.Sub_Type = "3" Then
Me.Sub_Type.Value = "Estimate/Pre-Study"

Else
If Me.Sub_Type = "4" Then
Me.Sub_Type.Value = "NCR"

Else
If Me.Sub_Type = "5" Then
Me.Sub_Type.Value = "New Build, Loose Part"

Else
If Me.Sub_Type = "6" Then
Me.Sub_Type.Value = "Repair/ReMfg"

Else
If Me.Sub_Type = "7" Then
Me.Sub_Type.Value = "Rev Change"

End If
End If
End If
End If
End If
End If
End If

End Sub

'-----

'Form Request Table
Option Compare Database

Private Sub Order_Job_Type_Change()

If Me.Order_Job_Type = "1" Then
Me.Order_Job_Type.Value = "New Build"

Else
If Me.Order_Job_Type = "2" Then
Me.Order_Job_Type.Value = "RGR"

Else
If Me.Order_Job_Type = "3" Then
Me.Order_Job_Type.Value = "Rental"

Else
If Me.Order_Job_Type = "4" Then
Me.Order_Job_Type.Value = "AO (Add-On)"

Else
If Me.Order_Job_Type = "5" Then
Me.Order_Job_Type.Value = "Conversion"

Else
If Me.Order_Job_Type = "6" Then
Me.Order_Job_Type.Value = "Rework"

Else
If Me.Order_Job_Type = "7" Then
Me.Order_Job_Type.Value = "Asset Mgmt"

End If
End If
End If
End If
End If
End If
End If

End Sub

Private Sub Sub_Type_Change()

If Me.Sub_Type = "1" Then
Me.Sub_Type.Value = "ATP (Assembly-Test-Paint)"

Else
If Me.Sub_Type = "2" Then
Me.Sub_Type.Value = "DCI (Dismantle-Clean-Inspection)"

Else
If Me.Sub_Type = "3" Then
Me.Sub_Type.Value = "Estimate/Pre-Study"

Else
If Me.Sub_Type = "4" Then
Me.Sub_Type.Value = "NCR"

Else
If Me.Sub_Type = "5" Then
Me.Sub_Type.Value = "New Build, Loose Part"

Else
If Me.Sub_Type = "6" Then
Me.Sub_Type.Value = "Repair/ReMfg"

Else
If Me.Sub_Type = "7" Then
Me.Sub_Type.Value = "Rev Change"

End If
End If
End If
End If
End If
End If
End If

End Sub

'-----

'Form Security
Option Compare Database

Private Sub Form_Load()
DoCmd.Close acForm, "Security"
DoCmd.OpenForm "Main_Form"

End Sub

'-----

'Form Temp Form
Option Compare Database

Private Sub Release_Date_AfterUpdate()
Dim Date1 As Date
Dim Date2 As Date

Dim result As Integer
Dim result2 As Integer

On Error GoTo problem

Date1 = Me.Log_Date.Value
Date2 = Me.Release_Date.Value

result = DateDiff("d", Date1, Date2)
result2 = result - Me.Weekend_Public_Holiday.Value

Me.Lead_Time__Days_.Value = result2

problem:
If Err.Number = 94 Then

Date1 = Me.Log_Date.Value
Date2 = Me.Release_Date.Value

result = DateDiff("d", Date1, Date2)

Me.Lead_Time__Days_.Value = result
End If

End Sub

Private Sub Weekend_Public_Holiday_AfterUpdate()
Dim Date1 As Date
Dim Date2 As Date

Dim result As Integer

If Not IsNull(Me.Release_Date.Value) Then

Date1 = Me.Log_Date.Value
Date2 = Me.Release_Date.Value

result = DateDiff("d", Date1, Date2)

Me.Lead_Time__Days_.Value = result
End If

End Sub

'-----

'Form User Form
Option Compare Database
Private Saved As Boolean

Private Sub cboUsername_AfterUpdate()
'Selected drop down
DoCmd.SearchForRecord , , acFirst, "[Username] = " & "'" & [Screen].[ActiveControl] & "'"
End Sub

Private Sub cboUsername_BeforeUpdate(Cancel As Integer)
Me.Undo
DoCmd.SearchForRecord , , acFirst, "[Username] = " & "'" & [Screen].[ActiveControl] & "'"
End Sub

'Delete Record
Private Sub cmdDelete_Click()
On Error GoTo problem

If MsgBox("Are you sure you want to DELETE this record?", vbExclamation + vbYesNo) = vbYes Then
'User clicked Yes
    Me.Undo
    DoCmd.SetWarnings False
        DoCmd.RunCommand acCmdSelectRecord
        DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
    MsgBox "Deleted Successfully.", vbInformation
'User clicked No
Else
    MsgBox "Deletion was aborted.", vbInformation
End If

DoCmd.RefreshRecord

problem:
    If Err.Number = 2046 Then
        MsgBox "No Selected Record to be Deleted.", vbInformation
    End If

End Sub

Private Sub cmdIEForm_Click()
DoCmd.OpenForm "Main_Form"
DoCmd.Close acForm, "User_Form"
End Sub

Private Sub cmdNewUser_Click()
Me.Undo
DoCmd.GoToRecord , , acNewRec
End Sub

'Save Record
Private Sub cmdSave_Click()
'Check that Username is filled
If IsNull(Username) Or IsNull(Password) Then
    MsgBox "The Username and Password MUST both be Entered.", _
    vbCritical, _
    "Canceling Update"
    Me.Username.SetFocus
    Cancel = True

Else

'Question to save changes to record
If MsgBox("Do you want to save the changes on this record?", vbQuestion + vbYesNo + vbDefaultButton1, "Save Changes?") = vbYes Then

'user clicked yes
   Saved = True
   DoCmd.RunCommand (acCmdSaveRecord)
   Me.cmdSave.Enabled = False
   Saved = False
   Me.cboUsername.Value = Me.Username.Value
   MsgBox "Save Successful.", vbInformation
   Me.Refresh
   
'user clicked no
Else
    Me.Undo
    MsgBox "Save Aborted.", vbInformation


End If
End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

'Question to save changes to record
Dim Response As Integer
If Saved = False Then
    Response = MsgBox("Do you want to save the changes on this record?", vbQuestion + vbYesNo + vbDefaultButton1, "Save Changes?")
    'If no
    
    If Response = vbNo Then
       Me.Undo
       MsgBox "Save Aborted.", vbInformation
       
    'if yes
    Else
    'Check that Project Name is filled
    If IsNull(Username) Then
    MsgBox "The Username Name Must be Entered.", _
    vbCritical, _
    "Canceling Update"
    Me.Username.SetFocus
    Cancel = True
       Else
       MsgBox "Save Successful.", vbInformation

       
    End If
    Me.cmdSave.Enabled = False
End If
End If
End Sub

Private Sub Form_Current()
Me.cboUsername = Me.Username
End Sub

'When dirty/changes made set cmdSave Enabled
Private Sub Form_Dirty(Cancel As Integer)
Me.cmdSave.Enabled = True
End Sub

Private Sub Form_Load()
DoCmd.Maximize

DoCmd.GoToRecord , , acNewRec
End Sub

'-----

'Form User Login Form
Option Compare Database
Option Explicit

'For login check for username and password
Private Sub cmdLogin_Click()
    On Error GoTo problem
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM tblLogin WHERE username = '" + Me.txtUsername.Value + "'"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If rs.EOF Then
        IncorrectUsernameStyle
        Exit Sub
    End If
    CorrectUsernameStyle
    
    rs.MoveFirst
    If rs("Password") <> Nz(Me.txtPassword, "") Then
        IncorrectPasswordStyle
        Exit Sub
    End If
    
    If DLookup("[Access]", "tblLogin", "[Username] = '" & Me.txtUsername.Value & "'") = "2" Then
    '"1" = Basic User
    '"2" = Advance User
    
    DoCmd.OpenForm "User_Form"
    DoCmd.Close acForm, "Main_Form"
    DoCmd.Close acForm, "User_Login_Form"
    
    Else
    MsgBox "You do not have Access!", vbInformation
    End If
    
problem:
    If Err.Number = 94 Then
        MsgBox "No Username & Password input.", vbInformation
        Me.txtUsername.SetFocus
    End If
End Sub

'Set Visible and colour for incorrect username
Private Sub IncorrectUsernameStyle()
    Me.lblIncorrectUsername.Visible = True
    Me.txtUsername.BorderColor = RGB(255, 0, 0)
    Me.txtUsername.SetFocus
End Sub

'Set Visible and colour for correct username
Private Sub CorrectUsernameStyle()
    Me.lblIncorrectUsername.Visible = False
    Me.txtUsername.BorderColor = RGB(0, 0, 0)
End Sub

'Set Visible and colour for incorrect password
Private Sub IncorrectPasswordStyle()
    Me.lblIncorrectPassword.Visible = True
    Me.txtPassword.BorderColor = RGB(255, 0, 0)
    Me.txtPassword.SetFocus
End Sub

'Cancel Login Form
Private Sub cmdExit_Click()
DoCmd.Close acForm, "User_Login_Form"
DoCmd.OpenForm "Main_Form"
End Sub

Private Sub Form_Load()
DoCmd.Maximize
End Sub

'-----

'Form Whole Table
Option Compare Database
Option Explicit

Private Sub Release_Date_AfterUpdate()
Dim Date1 As Date
Dim Date2 As Date

Dim result As Integer
Dim result2 As Integer

On Error GoTo problem

Date1 = Me.Log_Date.Value
Date2 = Me.Release_Date.Value

result = DateDiff("d", Date1, Date2)
result2 = result - Me.Weekend_Public_Holiday.Value

Me.Lead_Time__Days_.Value = result2

problem:
If Err.Number = 94 Then

Date1 = Me.Log_Date.Value
Date2 = Me.Release_Date.Value

result = DateDiff("d", Date1, Date2)

Me.Lead_Time__Days_.Value = result
End If

End Sub

Private Sub Weekend_Public_Holiday_AfterUpdate()
Dim Date1 As Date
Dim Date2 As Date

Dim result As Integer

If Not IsNull(Me.Release_Date.Value) Then

Date1 = Me.Log_Date.Value
Date2 = Me.Release_Date.Value

result = DateDiff("d", Date1, Date2)

Me.Lead_Time__Days_.Value = result
End If

End Sub

'-----
'-----

'*MODULES*

'ExcelImport
Option Compare Database
Option Explicit

Public Sub ImportExcelSpreadsheet(Filename As String, tableName As String)
On Error GoTo BadFormat
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, tableName, Filename, True
    Exit Sub
    
BadFormat:
    MsgBox "The file you try to import was not an excel spreadsheet!", vbInformation
    
End Sub

'-----

'ShiftKey
Option Compare Database

Function ap_DisableShift()
'This function disable the shift at startup. This action causes
'the Autoexec macro and Startup properties to always be executed.

On Error GoTo errDisableShift

Dim db As DAO.Database
Dim prop As DAO.Property
Const conPropNotFound = 3270

Set db = CurrentDb()

'This next line disables the shift key on startup.
db.Properties("AllowByPassKey") = False

'The function is successful.
Exit Function

errDisableShift:
'The first part of this error routine creates the "AllowByPassKey
'property if it does not exist.
If Err = conPropNotFound Then
Set prop = db.CreateProperty("AllowByPassKey", _
dbBoolean, False)
db.Properties.Append prop
Resume Next
Else
MsgBox "Function 'ap_DisableShift' did not complete successfully."
Exit Function
End If

End Function

Function ap_EnableShift()
'This function enables the SHIFT key at startup. This action causes
'the Autoexec macro and the Startup properties to be bypassed
'if the user holds down the SHIFT key when the user opens the database.

On Error GoTo errEnableShift

Dim db As DAO.Database
Dim prop As DAO.Property
Const conPropNotFound = 3270

Set db = CurrentDb()

'This next line of code disables the SHIFT key on startup.
db.Properties("AllowByPassKey") = True

'function successful
Exit Function

errEnableShift:
'The first part of this error routine creates the "AllowByPassKey
'property if it does not exist.
If Err = conPropNotFound Then
Set prop = db.CreateProperty("AllowByPassKey", _
dbBoolean, True)
db.Properties.Append prop
Resume Next
Else
MsgBox "Function 'ap_DisableShift' did not complete successfully."
Exit Function
End If

End Function
