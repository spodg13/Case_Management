VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCloseCase 
   Caption         =   "Close Case"
   ClientHeight    =   9168
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11010
   OleObjectBlob   =   "frmCloseCase.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCloseCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'---------------------------------------------------------------------------------------
' File   : frmCloseCase
' Author : gouldd
' Date   : 9/1/2016
' Purpose: Multi-purpose form to close a case, update a case with tasks, re-open a case or change the due date
'---------------------------------------------------------------------------------------

Option Explicit





Private Sub cmboAttorney_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo cmboAttorney_BeforeUpdate_Error
Dim zMsg As String

Dim rngAtty As Range
Dim strPart As String
Dim intCount As Integer
Dim sAddress As String

   sAddress = ActiveCell.Address
If InStr(1, cmboAttorney.Value, ",") = 0 Then
        MsgBox "Please enter using the format Lastname, firstname"
        Cancel = True
        cmboAttorney.SetFocus
        cmboAttorney.Value = RevAttyName(ActiveCell.Offset(0, 2))
        Exit Sub
End If

Set rngAtty = Attorneys.Range("Attorneys")
strPart = Trim(Me.cmboAttorney.Value)
intCount = Application.WorksheetFunction.CountIf(rngAtty, strPart)
If intCount = 0 Then
    'Add Attorney to table
    addAttorney strPart
    InvestigationLog.Range(sAddress).Select
    
End If
    txtAttorney.Enabled = True
    txtAttorney = AttorneyName(cmboAttorney.Value)
    txtAttorney.Enabled = False

    On Error GoTo 0
Exit Sub

cmboAttorney_BeforeUpdate_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: cmboAttorney_BeforeUpdate Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdCancelClose_Click()
CaseLogs.OLEObjects("cmdFilterCase").Object.Caption = "Show All"
    With InvestigationLog
        .EnableSelection = xlNoRestrictions
    End With
Unload frmCloseCase
End Sub

Private Sub cmdClose_Click()
  

    On Error GoTo cmdClose_Click_Error
    Dim zMsg As String

Dim TheLastCaseRow As Long
Dim wdApp As New Word.Application
Dim wdDoc As Word.Document
Dim CCtrl As Word.ContentControl
Dim strActionPath As String
Dim strFileName, strClose, strTemplateFileName, strActionEntry As String
Dim intLen, intCommaPos, intVersion  As Integer
Dim strActionName As String
Dim strMethod, strOtherMemo, strInvMemo As String
Dim Wksht As Worksheet
Dim lngRow, intCol As Long
Dim strClient, strInvInitials, strClientLast, strClientFirst, strPath As String
Dim result As Integer
Dim strNewInv As String
Dim w As Workbook


Application.ScreenUpdating = False
InvestigationLog.Activate
    If ReAssignable = True Then
            result = MsgBox("Do you want to close and re-assign?", vbYesNoCancel, "Re-Assign?")
            If result = vbYes Then
                strClose = "close and re-assign "
                strNewInv = InputBox("Type in the new Investigator", "Re-assign to:")
                UserChange ("Procedure: cmdClose_Click Within: frmCloseCase" & ActiveWorkbook.Path)
                
            Else
                strClose = "close the "
            End If
    Else
            strClose = "close the "
     End If
     
 Application.EnableEvents = False
    Set Wksht = InvestigationLog
    lngRow = ActiveCell.row
    intCol = 3 ' Place on client name
    Cells(lngRow, intCol).Activate
    If Cells(lngRow, 12).Value = "Closed" Then
        result = MsgBox("This case is already closed.", vbOKOnly, "Cannot close a closed case")
        Unload frmCloseCase
        Exit Sub
    End If
    
    result = MsgBox("Do you want to " & strClose & ActiveCell.Value & " case?", vbYesNoCancel, "Verify Case")
    If result = vbNo Then
         result = MsgBox("Click on the case you need the closure for, then press the command button!", vbOKOnly, "Verify Case")
         Application.EnableEvents = True
         Exit Sub
    End If
    If result = vbCancel Then
        Unload frmCloseCase
        Exit Sub
    End If
'Get Action Log name
        
        strActionName = ActiveCell.Value
        strActionName = Replace(strActionName, ", ", "_")
        strActionName = strActionName & "_" & ActiveCell.Offset(0, -2).Value
        strActionPath = Files.Cells(6, 2).Value
        strActionName = strActionPath & strActionName & ".docx"
        If TemplateExists(strActionName) = False Then
            result = MsgBox("Your Action Log " & strActionName & " is missing." & vbLf & vbLf & "Did you change the client name and not change the file?  Check your Action log folder and correct the spelling.", vbCritical, "Missing Action Log")
            Unload frmCloseCase
            Exit Sub
        End If
    
'Update CaseLog

CaseLogs.Activate 'Go to last line
'Remove filters - done during form initilization, not needed here


'Calculate last row or import from form
 TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
 TheLastCaseRow = TheLastCaseRow + 1
 strActionEntry = "Closed Case - Total Time"
 
 
 UpdateCaseLog strActionEntry, Format(Now(), "h:mm AM/PM"), Format(Now(), "mm/d/yy"), TheLastCaseRow, lngRow, 1, Val(txtMins)
 CaseLogs.Cells(TheLastCaseRow, 7).Value = "End"
     

'Start Action Log copy **** Print only if automatic is selected
'**************************************************************************
       
    
   PrintActionLog strActionName, txtNewDueDate, ActionRange(2, 5), Files.Cells(18, 2).Value
    
'Print witness list
    PrintWitnessList txtCaseNo, txtClient
    InvestigationLog.Activate
    
    ActiveCell.Offset(0, 4).Value = Format(Now(), "mm/d/yy")
    strFileName = ActiveCell.Value
    strFileName = Replace(strFileName, ", ", "_")
    strFileName = strFileName & "_" & ActiveCell.Offset(0, -2).Value
    strFileName = strFileName & "_" & "Closure"
        
    strClient = Cells(lngRow, 3).Value
    InvestigationLog.Cells(lngRow, 12).Value = "Closed"
    strInvInitials = Files.Cells(16, 2).Value
    
    intCommaPos = InStr(strClient, ",")
    intLen = Len(strClient)
    strClientLast = Left$(strClient, intCommaPos - 1)
    strClientFirst = Right$(strClient, intLen - intCommaPos - 1)
    
    strPath = Files.Cells(1, 2).Value
    strTemplateFileName = Files.Cells(5, 2).Value
    If strTemplateFileName = "" Then
        result = MsgBox("Please select the Closure Report Template Location!", vbCritical, "Need the file name")
        strTemplateFileName = FilePicked("Closure Document")
        Files.Cells(5, 2) = strTemplateFileName
    End If
    If strPath = "" Then
        result = MsgBox("Please select the path where the reports will be stored!", vbCritical, "Need the path")
        strPath = PathPicked("Reports") & "\"
        Files.Cells(1, 2) = strPath
    End If
    intVersion = GetVersion(strPath, strFileName)
        If intVersion > 1 Then
            strFileName = strFileName & "_" & CStr(intVersion)
        End If
        
    If txtOtherMemo = "" Then
        strOtherMemo = " "
    Else
        strOtherMemo = txtOtherMemo
    End If
    strInvMemo = txtInvMemo
 
Set wdDoc = wdApp.Documents.Open(FileName:=strTemplateFileName, AddToRecentFiles:=False, Visible:=False)
        With wdDoc
            
            For Each CCtrl In .ContentControls
                Select Case CCtrl.Title
                    Case "CaseNum"
                        CCtrl.Range.Text = Cells(lngRow, 1).Value
                    Case "ClientName"
                        CCtrl.Range.Text = strClientLast & ", " & strClientFirst
                    Case "xref"
                        CCtrl.Range.Text = Cells(lngRow, 4).Value
                    Case "Atty"
                        
                        If strClose = "close and re-assign " Then
                            CCtrl.Range.Text = Cells(lngRow, 5).Value & "   Re-Assign to: " & strNewInv
                        Else
                            CCtrl.Range.Text = Cells(lngRow, 5).Value
                        End If
                    Case "FI"
                        CCtrl.Range.Text = txtFI
                    Case "PI"
                        CCtrl.Range.Text = txtPI
                    Case "PS"
                        CCtrl.Range.Text = txtPS
                    Case "MS"
                        CCtrl.Range.Text = txtMS
                    Case "Other"
                        CCtrl.Range.Text = CStr(Val(txtOther) + Val(txtDueDil))
                    Case "TotRpt"
                        CCtrl.Range.Text = txtRptTotal
                    Case "TotMin"
                        CCtrl.Range.Text = txtMins
                    Case "Testify"
                        CCtrl.Range.Text = txtTestify
                    Case "Photo"
                        CCtrl.Range.Text = txtPhoto
                    'Case "OtherMemo"
                    '   CCtrl.Range.Text = strOtherMemo
                    'Case "InvMemo"
                    '    If txtInvMemo = "" Then txtInvMemo = " "
                    '    CCtrl.Range.Text = txtInvMemo
                    Case "NewOther"
                        If txtNewOther = "" Then txtNewOther = " -"
                        CCtrl.Range.Text = txtNewOther
                    Case "NewPhoto"
                        If txtNewPhoto = "" Then txtNewPhoto = " -"
                        CCtrl.Range.Text = txtNewPhoto
                    Case "NewSub"
                        If txtNewSub = "" Then txtNewSub = " -"
                        CCtrl.Range.Text = txtNewSub
                    Case "NewInt"
                        If txtNewInt = "" Then txtNewInt = " -"
                        CCtrl.Range.Text = txtNewInt
                    Case "NewTestify"
                     If ckbxNewTestify = True Then
                         CCtrl.Range.Text = "Y"
                        Else
                            CCtrl.Range.Text = "N"
                        End If
                    Case "NewDue"
                        If txtNewDueDate = "" Then txtNewDueDate = " -"
                        If ReAssignable = False Then txtNewDueDate = " -"
                        CCtrl.Range.Text = txtNewDueDate
                    Case "Init"
                        CCtrl.Range.Text = strInvInitials
                    Case "Update"
                        If optUpdate = True Then
                            CCtrl.Checked = True
                        Else
                            CCtrl.Checked = False
                        End If
                    Case "CloseCase"
                        If optClose = True Then
                            CCtrl.Checked = True
                        Else
                            CCtrl.Checked = False
                        End If
                    Case "DOC"
                        CCtrl.Range.Text = Format(Now(), "mm/d/yy")
                    
                    Case "CompReq"
                        If optCompleted = True Then
                            CCtrl.Checked = True
                        Else
                            CCtrl.Checked = False
                        End If
                    Case "CB7"
                        If strMethod = "in person" Then
                            CCtrl.Checked = True
                        Else
                            CCtrl.Checked = False
                        End If
                    Case "CB8"
                        If strMethod = "in person" Then
                            CCtrl.Checked = True
                        Else
                            CCtrl.Checked = False
                        End If
                End Select
                
            Next
            .Bookmarks("OtherMemo").Range.Text = strOtherMemo
            .Bookmarks("InvMemo").Range.Text = txtInvMemo
            .Bookmarks("OtherMemoCopy").Range.Text = strOtherMemo
            .Bookmarks("InvMemoCopy").Range.Text = txtInvMemo
                        
        End With
     
      
    'Save with new name
    wdDoc.SaveAs FileName:=strPath & strFileName & ".docx", FileFormat:=wdFormatDocumentDefault
    If Files.Cells(30, 2).Value = True Then
        wdDoc.Activate
        wdDoc.PrintOut
        wdDoc.Close
        wdApp.Quit
        Set wdApp = Nothing
        Set wdDoc = Nothing
    Else
        wdApp.Visible = True
        wdApp.Activate
        Set wdApp = Nothing
        Set wdDoc = Nothing
    End If
    SortByDueDate
    ClearCaseLogFilter
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    With InvestigationLog
        .EnableSelection = xlNoRestrictions
    End With
ActiveWorkbook.Save
Unload frmCloseCase



    On Error GoTo 0
    Exit Sub

cmdClose_Click_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: cmdClose_Click Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
End Sub


Private Sub cmdUpdate_Click()


    On Error GoTo cmdUpdate_Click_Error
    Dim zMsg As String

 

    Dim wdApp As New Word.Application
    Dim wdDoc As Word.Document
    Dim CCtrl As Word.ContentControl
    Dim lngRow, TheLastCaseRow As Long
    Dim strClient, strInvInitials, strClientFirst, strClientLast As String
    Dim result As Integer
    Dim strTask, strComma, strFileName, strPhotoFolder As String
    Dim strPath, strTemplateFileName As String
    Dim CellAltered As Boolean
    Dim intRepCounter, intCommaPos, intLen, intVersion As Integer
    Dim Wksht As Worksheet
    
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    'Set to InvestigationLog, and activate
    Set Wksht = InvestigationLog
    lngRow = ActiveCell.row

    Cells(lngRow, 3).Activate ' Activate Client name
    strFileName = Replace(ActiveCell.Value, ", ", "_")
    strFileName = strFileName & "_" & InvestigationLog.Cells(lngRow, 1).Value
        
    ActiveCell.Offset(0, 3).Value = GetADate(txtNewDueDate)
    ActiveCell.Offset(0, 2).Value = txtAttorney
    ActiveCell.Offset(0, 7).Value = GetADate(txtCtDate)
    ActiveCell.Offset(0, 8).Value = txtDept
    
    
    If optReOpen = True Then
        ActiveCell.Offset(0, 9) = "Re-Open"
        strTask = ""
    End If
    If optAddSupplement = True Then strTask = ActiveCell.Offset(0, 5).Value & ", "
    If optDueDateOnly = True Then
        strTask = ""
        GoTo NoTask:
    End If
    
       
    ActiveCell.Offset(0, 4).ClearContents
      
    
    strComma = ""
    'Take old tasks and add to new tasks.  Then call task Function
    'Reset columns VWXYZ back to zeroes if Re-open
    If optReOpen = True Then
    
        For intRepCounter = 22 To 35
            Cells(lngRow, intRepCounter) = 0
        Next intRepCounter
    End If
    
    If ckbxNewTestify = True Then
       Cells(lngRow, 35).Value = Cells(lngRow, 35).Value + 1
    
    End If
    
    If Val(txtNewInt) > 0 Then
        Cells(lngRow, 31).Value = Cells(lngRow, 31).Value + Val(txtNewInt)
    End If
    
    If Val(txtNewSub) > 0 Then
        Cells(lngRow, 33).Value = Cells(lngRow, 33).Value + Val(txtNewSub)
    End If
    
    If Val(txtNewPhoto) > 0 Then
        Cells(lngRow, 32).Value = Cells(lngRow, 32).Value + Val(txtNewPhoto)
        strPhotoFolder = Files.Cells(29, 2).Value & strFileName
        MakePhotoDir (strPhotoFolder)
    End If
    
    If Val(txtNewOther) > 0 Then
        Cells(lngRow, 34).Value = Cells(lngRow, 34).Value + Val(txtNewOther)
    End If
    strTask = TaskList(Cells(lngRow, 31).Value, Cells(lngRow, 32).Value, Cells(lngRow, 33).Value, Cells(lngRow, 34).Value, ckbxNewTestify)
    Cells(lngRow, 8).Value = strTask
        
NoTask:
    ActiveCell.Offset(0, 3).Value = ActiveCell.Offset(0, 3).Value + TimeSerial(8, 0, 0)
    TheLastCaseRow = CaseLogs.Cells(1048576, "A").End(xlUp).row
    TheLastCaseRow = TheLastCaseRow + 1
           
       
   
    

   
    'End of reset routine
   
    'Update Investigation Log
    
    'Check for short due date
    
    'Need to activate Name of client cell
    CellAltered = False
   
    If WeekdayName(Weekday(ActiveCell.Offset(0, 3)), True) = "Sat" Then
        
        ActiveCell.Offset(0, 3) = ActiveCell.Offset(0, 3).Value - 1
         CellAltered = True
    End If
        
    If WeekdayName(Weekday(ActiveCell.Offset(0, 3)), True) = "Sun" Then
        
        ActiveCell.Offset(0, 3) = ActiveCell.Offset(0, 3).Value - 2
        CellAltered = True
    End If
    If WeekdayName(Weekday(ActiveCell.Offset(0, 3)), True) = "Mon" Then
        
        ActiveCell.Offset(0, 3) = ActiveCell.Offset(0, 3).Value - 3
        CellAltered = True
    End If
    
    If DateDiff("d", ActiveCell.Offset(0, 3).Value, Files.Cells(17, 2).Value) Mod 14 = 0 Then
       
       ActiveCell.Offset(0, 3).Value = ActiveCell.Offset(0, 3).Value - 1
        CellAltered = True
        
    End If
    
    If CellAltered = True Then
        ActiveCell.Offset(0, 3).Value = ActiveCell.Offset(0, 3).Value + TimeSerial(9, 0, 0)
                
    End If
    
 'Update CaseLogs sheet
    If optDueDateOnly = False Then ' skip if Just changing due date
        UpdateCaseLog "Received and reviewed supplemental investigative request", Format(txtNewRecTime, "h:mm AMPM"), GetADate(txtNewRecDate), TheLastCaseRow, lngRow, 1, 0
        
        CaseLogs.Cells(TheLastCaseRow, 4).font.Bold = True
        CaseLogs.Cells(TheLastCaseRow, 7).Value = "Start"
    End If
    '************************************************************************************
    If optDueDateOnly = True Or chkPrintIR = True Then
    
    'Print out update sheet
    strFileName = ActiveCell.Value
    strFileName = Replace(strFileName, ", ", "_")
    strFileName = strFileName & "_" & ActiveCell.Offset(0, -2).Value
    strFileName = strFileName & "_" & "Update"
        
    strClient = Cells(lngRow, 3).Value
    strInvInitials = Files.Cells(16, 2).Value
    
    intCommaPos = InStr(strClient, ",")
    intLen = Len(strClient)
    strClientLast = Left$(strClient, intCommaPos - 1)
    strClientFirst = Right$(strClient, intLen - intCommaPos - 1)
    
    strPath = Files.Cells(1, 2).Value
    strTemplateFileName = Files.Cells(5, 2).Value
    If strTemplateFileName = "" Then
        result = MsgBox("Please select the Closure Report Template Location!", vbCritical, "Need the file name")
        strTemplateFileName = FilePicked("Closure Document")
        Files.Cells(5, 2) = strTemplateFileName
    End If
    If strPath = "" Then
        result = MsgBox("Please select the path where the reports will be stored!", vbCritical, "Need the path")
        strPath = PathPicked("Reports") & "\"
        Files.Cells(1, 2) = strPath
    End If
    
    intVersion = GetVersion(strPath, strFileName)
        If intVersion > 1 Then
            strFileName = strFileName & "_" & CStr(intVersion)
        End If
Set wdDoc = wdApp.Documents.Open(FileName:=strTemplateFileName, AddToRecentFiles:=False, Visible:=False)
        With wdDoc
            
            For Each CCtrl In .ContentControls
                Select Case CCtrl.Title
                    Case "CaseNum"
                        CCtrl.Range.Text = Cells(lngRow, 1).Value
                    Case "ClientName"
                        CCtrl.Range.Text = strClientLast & ", " & strClientFirst
                    Case "xref"
                        CCtrl.Range.Text = Cells(lngRow, 4).Value
                    
                    Case "FI"
                        CCtrl.Range.Text = " - "
                    Case "PI"
                        CCtrl.Range.Text = " - "
                    Case "PS"
                        CCtrl.Range.Text = " - "
                    Case "MS"
                        CCtrl.Range.Text = " - "
                    Case "Other"
                        CCtrl.Range.Text = " - "
                    Case "TotRpt"
                        CCtrl.Range.Text = " - "
                    Case "TotMin"
                        CCtrl.Range.Text = " - "
                    Case "Testify"
                        CCtrl.Range.Text = " - "
                    Case "Photo"
                        CCtrl.Range.Text = " - "
                   Case "Atty"
                        CCtrl.Range.Text = Cells(lngRow, 5).Value
                   ' Case "InvMemo"
                   '     If txtDueDateMemo = "" Then txtDueDateMemo = " "
                   '     CCtrl.Range.Text = txtDueDateMemo
                    Case "NewOther"
                        If txtNewOther = "" Then txtNewOther = " -"
                        CCtrl.Range.Text = txtNewOther
                    Case "NewPhoto"
                        If txtNewPhoto = "" Then txtNewPhoto = " -"
                        CCtrl.Range.Text = txtNewPhoto
                    Case "NewSub"
                        If txtNewSub = "" Then txtNewSub = " -"
                        CCtrl.Range.Text = txtNewSub
                    Case "NewInt"
                        If txtNewInt = "" Then txtNewInt = " -"
                        CCtrl.Range.Text = txtNewInt
                    Case "NewTestify"
                        If ckbxNewTestify = True Then
                            CCtrl.Range.Text = "Y"
                        Else
                            CCtrl.Range.Text = "N"
                        End If
                    Case "NewDue"
                        CCtrl.Range.Text = txtNewDueDate
                    Case "Init"
                        CCtrl.Range.Text = strInvInitials
                    Case "Update"
                        If optUpdate = True Then
                            CCtrl.Checked = True
                        Else
                            CCtrl.Checked = False
                        End If
                    Case "CloseCase"
                        If optClose = True Then
                            CCtrl.Checked = True
                        Else
                            CCtrl.Checked = False
                        End If
                    Case "DOC"
                        CCtrl.Range.Text = Format(Now(), "mm/d/yy")
                    
                    Case "CompReq"
                        If optCompleted = True Then
                            CCtrl.Checked = True
                        Else
                            CCtrl.Checked = False
                        End If
                    
                End Select
                
            Next
            '.Bookmarks("OtherMemo").Range.Text = strOtherMemo
            .Bookmarks("InvMemo").Range.Text = txtDueDateMemo
            '.Bookmarks("OtherMemoCopy").Range.Text = strOtherMemo
            .Bookmarks("InvMemoCopy").Range.Text = txtDueDateMemo
        End With
     
   'Unload frmCloseCase
    
    'Save with new name
    wdDoc.SaveAs FileName:=strPath & strFileName & ".docx", FileFormat:=wdFormatDocumentDefault
    If Files.Cells(30, 2).Value = True Then
        wdDoc.Activate
        wdDoc.PrintOut
        wdDoc.Close
        wdApp.Quit
        Set wdApp = Nothing
        Set wdDoc = Nothing
    Else
        wdApp.Visible = True
        wdApp.Activate
        Set wdApp = Nothing
        Set wdDoc = Nothing
    End If
    'Due date change reflected on action log?
    UpdateCaseLog "Submitted update sheet.  New Due Date " & txtNewDueDate, Format(txtNewRecTime, "h:mm AMPM"), GetADate(txtNewRecDate), TheLastCaseRow, lngRow, 1, 0
    End If
EndRoutine:

    
    If optDueDateOnly = False Then
        ActiveCell.Offset(0, -1).Value = GetADate(txtNewRecDate)
    End If
SortByDueDate
ClearCaseLogFilter

ActiveCell.Offset(1, 0).Select

    'ActiveSheet.Protect "darryl"
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    With InvestigationLog
        .EnableSelection = xlNoRestrictions
    End With
    ActiveWorkbook.Save
    Unload frmCloseCase
    


    On Error GoTo 0
    Exit Sub

cmdUpdate_Click_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: cmdUpdate_Click Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
   
End Sub



Private Sub optAddSupplement_Click()


    On Error GoTo optAddSupplement_Click_Error
    Dim zMsg As String

Dim row As Long
Dim result As Integer
row = ActiveCell.row
If InvestigationLog.Cells(row, 12).Value = "Closed" Then
    optReOpen = xlOn
    result = MsgBox("Can't add a request to a closed case.  Re-open is the appropriate button.", vbOKOnly)
Else
    optAddSupplement = xlOn
End If
txtDueDateMemo.Visible = False
Label27.Visible = False


    On Error GoTo 0
    Exit Sub

optAddSupplement_Click_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: optAddSupplement_Click Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub optClose_Click()


    On Error GoTo optClose_Click_Error

Application.ScreenUpdating = False
If ReAssignable = True Then
    frameUpdate.Top = 138
    frameUpdate.Visible = True
    frameUpdate.Caption = "Remaining tasks"
    optAddSupplement.Caption = "Re-assigning"
    optAddSupplement.Enabled = False
    optDueDateOnly.Visible = False
    optReOpen.Visible = False
    txtNewRecDate.Visible = False
    txtNewRecTime.Visible = False
    Label17.Visible = False
    Label26.Visible = False
    
    
    frameClosure.Enabled = False
Else
    frameUpdate.Top = 400
    frameUpdate.Visible = False
End If
Application.ScreenUpdating = True
  
frameClosure.Visible = True
cmdClose.Visible = True
cmdUpdate.Visible = False

If txtPhoto = "" Then txtPhoto = 0
If txtOther = "" Then txtOther = 0
If txtPI = "" Then txtPI = 0
If txtFI = "" Then txtFI = 0
If txtMS = "" Then txtMS = 0
If txtPS = "" Then txtPS = 0
If txtTestify = "" Then txtTestify = 0
If txtDueDil = "" Then txtDueDil = 0
If txtDueDil > 0 Then
    If txtOther = 0 Then
        txtOtherMemo = txtDueDil & " Due Diligence report(s)"
        Else
        txtOtherMemo = txtOther & " report(s), and " & txtDueDil & " Due Diligence report(s)"
    End If
End If
    
txtRptTotal = Int(txtFI) + Int(txtPI) + Int(txtDueDil) + Int(txtPhoto) + Int(txtOther)

'Blank Update fields
txtNewInt = " "
txtNewOther = " "
txtNewPhoto = " "
txtNewSub = " "
ckbxNewTestify = False

    
txtMins = SumVisible(ActionRange(5, 5), "Action") '5 is startting column and ending column

'If cell value = closed case, check for number



    On Error GoTo 0
    Exit Sub

optClose_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optClose_Click of Sub frmCloseCase"

End Sub



Private Sub optCompleted_Click()


    On Error GoTo optCompleted_Click_Error
    Dim zMsg As String

Application.ScreenUpdating = False
    frameUpdate.Top = 400
    frameUpdate.Visible = False
Application.ScreenUpdating = True

frameClosure.Visible = True
cmdClose.Visible = True
cmdUpdate.Visible = False
optUpdate = False
optClose = False


If txtPhoto = "" Then txtPhoto = 0
If txtOther = "" Then txtOther = 0
If txtPI = "" Then txtPI = 0
If txtFI = "" Then txtFI = 0
If txtMS = "" Then txtMS = 0
If txtPS = "" Then txtPS = 0
If txtTestify = "" Then txtTestify = 0
If txtDueDil = "" Then txtDueDil = 0
If txtDueDil > 0 Then
    If txtOther = 0 Then
        txtOtherMemo = txtDueDil & " Due Diligence report(s)"
        Else
        txtOtherMemo = txtOther & " report(s), and " & txtDueDil & " Due Diligence report(s)"
    End If
End If
    
txtRptTotal = Int(txtFI) + Int(txtPI) + Int(txtDueDil) + Int(txtPhoto) + Int(txtOther)

'Blank Update fields
txtNewInt = " "
txtNewOther = " "
txtNewPhoto = " "
txtNewSub = " "
ckbxNewTestify = False

    
txtMins = SumVisible(ActionRange(5, 5), "Action") '5 is startting column and ending column

'If cell value = closed case, check for number

    On Error GoTo 0
    Exit Sub

optCompleted_Click_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: optCompleted_Click Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

             End Sub



Private Sub optDueDateOnly_Click()

    On Error GoTo optDueDateOnly_Click_Error
    Dim zMsg As String

Dim row As Long
Dim result As Integer
row = ActiveCell.row
If InvestigationLog.Cells(row, 12).Value = "Closed" Then
    optReOpen = xlOn
    result = MsgBox("Can't add a request to a closed case.  Re-open is the appropriate button.", vbOKOnly)
Else
    optDueDateOnly = xlOn
End If

txtDueDateMemo.Visible = True
Label27.Visible = True


    On Error GoTo 0
    Exit Sub

optDueDateOnly_Click_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: optDueDateOnly_Click Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

             End Sub

Private Sub optReOpen_Click()


    On Error GoTo optReOpen_Click_Error
    Dim zMsg As String



Dim lngRow As Long
lngRow = ActiveCell.row
If InvestigationLog.Cells(lngRow, 12).Value = "Open" Then
    MsgBox ("Close the case before your re-open it, or add a request")
    optAddSupplement.Value = True
End If
If InvestigationLog.Cells(lngRow, 12).Value = "Re-Open" Then
    MsgBox ("Close the case before your re-open it, or add a request")
    optAddSupplement.Value = True
End If
txtDueDateMemo.Visible = False
Label27.Visible = False





    On Error GoTo 0
    Exit Sub

optReOpen_Click_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: optReOpen_Click Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
    
End Sub

Private Sub optUpdate_Click()
    frameUpdate.Visible = True
    frameUpdate.Top = 138
    optAddSupplement.Caption = "Additional Request"
    optAddSupplement.Enabled = True
    optDueDateOnly.Visible = True
    optReOpen.Visible = True
    txtNewRecDate.Visible = True
    txtNewRecTime.Visible = True
    Label17.Visible = True
    Label26.Visible = True
    frameClosure.Visible = False
    cmdClose.Visible = False
    cmdUpdate.Visible = True
    optCompleted = False
        
End Sub

Private Sub txtCtDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo txtCtDate_DblClick_Error
Dim zMsg As String
DatePickerForm.Caption = "Court Date"
DatePickerForm.Show vbModal
Select Case [DatePickerForm]![CallingForm].Caption
Case "Form"
    If IsaDate(txtCtDate) = False Then
        txtCtDate = Format(DateValue(Now()), "MMMM d, yyyy")
    End If
 Case Else
    txtCtDate = [DatePickerForm]![CallingForm].Caption
 End Select

    Cancel = True

    On Error GoTo 0
Exit Sub

txtCtDate_DblClick_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtCtDate_DblClick Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtCtDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

On Error GoTo txtCtDate_KeyPress_Error
Dim zMsg As String

If KeyAscii = 43 Then
    txtCtDate = DateAdd("d", 1, txtCtDate)
    KeyAscii = 0
    txtCtDate = Format(DateValue(txtCtDate), "m/d/yy")
End If
If KeyAscii = 45 Then
    txtCtDate = DateAdd("d", -1, txtCtDate)
    KeyAscii = 0
    txtCtDate = Format(DateValue(txtCtDate), "m/d/yy")
End If

    On Error GoTo 0
Exit Sub

txtCtDate_KeyPress_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtCtDate_KeyPress Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDueDil_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo txtDueDil_BeforeUpdate_Error
Dim zMsg As String

    If txtDueDil = "" Then txtDueDil = 0
    txtRptTotal = Int(txtFI) + Int(txtPI) + Int(txtDueDil) + Int(txtPhoto) + Int(txtOther)

    On Error GoTo 0
Exit Sub

txtDueDil_BeforeUpdate_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtDueDil_BeforeUpdate Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub txtDueDil_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtFI_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo txtFI_BeforeUpdate_Error
Dim zMsg As String

    If txtFI = "" Then txtFI = 0
    txtRptTotal = Int(txtFI) + Int(txtPI) + Int(txtDueDil) + Int(txtPhoto) + Int(txtOther)

    On Error GoTo 0
Exit Sub

txtFI_BeforeUpdate_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtFI_BeforeUpdate Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub txtFI_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Only allow integers
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub



Private Sub txtMS_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If txtMS = "" Then txtMS = 0

End Sub

Private Sub txtMS_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub



Private Sub txtNewDueDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    

On Error GoTo txtNewDueDate_DblClick_Error
Dim zMsg As String
DatePickerForm.Caption = "Due Date"
DatePickerForm.Show vbModal

 Select Case [DatePickerForm]![CallingForm].Caption
      Case "Form"
          If IsaDate(txtNewDueDate) = False Then
              txtNewDueDate = Format(DateValue(Now()), "MMMM d, yyyy")
          End If
       Case Else
          txtNewDueDate = [DatePickerForm]![CallingForm].Caption
 End Select

    Cancel = True

    On Error GoTo 0
Exit Sub
txtNewDueDate_DblClick_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtNewDueDate_DblClick Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtNewDueDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

On Error GoTo txtNewDueDate_KeyPress_Error
Dim zMsg As String

If KeyAscii = 43 Then
    txtNewDueDate = DateAdd("d", 1, txtNewDueDate)
    KeyAscii = 0
    txtNewDueDate = Format(DateValue(txtNewDueDate), "m/d/yy")
End If
If KeyAscii = 45 Then
    txtNewDueDate = DateAdd("d", -1, txtNewDueDate)
    KeyAscii = 0
    txtNewDueDate = Format(DateValue(txtNewDueDate), "m/d/yy")
End If


    On Error GoTo 0
Exit Sub

txtNewDueDate_KeyPress_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtNewDueDate_KeyPress Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub





Private Sub txtNewInt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub



Private Sub txtNewOther_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtNewPhoto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtNewRecDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo txtNewRecDate_DblClick_Error
Dim zMsg As String
DatePickerForm.Caption = "Recieved Date"
DatePickerForm.Show vbModal


    Select Case [DatePickerForm]![CallingForm].Caption
      Case "Form"
          If IsaDate(txtNewRecDate) = False Then
              txtNewRecDate = Format(DateValue(Now()), "MMMM d, yyyy")
          End If
       Case Else
          txtNewRecDate = [DatePickerForm]![CallingForm].Caption
 End Select

    
    Cancel = True

    On Error GoTo 0
Exit Sub

txtNewRecDate_DblClick_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtNewRecDate_DblClick Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtNewRecDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

On Error GoTo txtNewRecDate_KeyPress_Error
Dim zMsg As String

If KeyAscii = 43 Then
    txtNewRecDate = DateAdd("d", 1, txtNewRecDate)
    KeyAscii = 0
    txtNewRecDate = Format(DateValue(txtNewRecDate), "m/d/yy")
End If
If KeyAscii = 45 Then
    txtNewRecDate = DateAdd("d", -1, txtNewRecDate)
    KeyAscii = 0
    txtNewRecDate = Format(DateValue(txtNewRecDate), "m/d/yy")
End If



    On Error GoTo 0
Exit Sub

txtNewRecDate_KeyPress_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtNewRecDate_KeyPress Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub txtNewSub_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtOther_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)


On Error GoTo txtOther_BeforeUpdate_Error
Dim zMsg As String

If txtOther = "" Then txtOther = 0
    txtRptTotal = Int(txtFI) + Int(txtPI) + Int(txtDueDil) + Int(txtPhoto) + Int(txtOther)

    On Error GoTo 0
Exit Sub

txtOther_BeforeUpdate_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtOther_BeforeUpdate Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub txtOther_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtPhoto_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo txtPhoto_BeforeUpdate_Error
Dim zMsg As String

If txtPhoto = "" Then txtPhoto = 0
    txtRptTotal = Int(txtFI) + Int(txtPI) + Int(txtDueDil) + Int(txtPhoto) + Int(txtOther)

    On Error GoTo 0
Exit Sub

txtPhoto_BeforeUpdate_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtPhoto_BeforeUpdate Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub txtPhoto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtPI_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo txtPI_BeforeUpdate_Error
Dim zMsg As String

If txtPI = "" Then txtPI = 0
    txtRptTotal = Int(txtFI) + Int(txtPI) + Int(txtDueDil) + Int(txtPhoto) + Int(txtOther)

    On Error GoTo 0
Exit Sub

txtPI_BeforeUpdate_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtPI_BeforeUpdate Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub txtNewRecTime_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)


On Error GoTo txtNewRecTime_KeyPress_Error
Dim zMsg As String

Dim LMin, LMinNew, IntervalAdd As Integer

If KeyAscii = 43 Then
    LMin = Minute(txtNewRecTime)
    LMinNew = Round(LMin / 5, 0) * 5
    IntervalAdd = (LMin - LMinNew) * -1
    txtNewRecTime = DateAdd("N", IntervalAdd, txtNewRecTime)
    
    txtNewRecTime = DateAdd("N", 5, txtNewRecTime)
    KeyAscii = 0
    txtNewRecTime = Format(txtNewRecTime, "h:mm AM/PM")
End If
If KeyAscii = 42 Then
   
    txtNewRecTime = DateAdd("N", 1, txtNewRecTime)
    KeyAscii = 0
    txtNewRecTime = Format(txtNewRecTime, "h:mm AM/PM")
End If
If KeyAscii = 45 Then
    LMin = Minute(txtNewRecTime)
    LMinNew = Round(LMin / 5, 0) * 5
    IntervalAdd = (LMin - LMinNew) * -1
    txtNewRecTime = DateAdd("N", IntervalAdd, txtNewRecTime)
    
    txtNewRecTime = DateAdd("N", -5, txtNewRecTime)
    KeyAscii = 0
    txtNewRecTime = Format(txtNewRecTime, "h:mm AM/PM")
End If
If KeyAscii = 47 Then
      
    txtNewRecTime = DateAdd("N", -1, txtNewRecTime)
    KeyAscii = 0
    txtNewRecTime = Format(txtNewRecTime, "h:mm AM/PM")
End If

    On Error GoTo 0
Exit Sub

txtNewRecTime_KeyPress_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtNewRecTime_KeyPress Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtPI_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub



Private Sub txtPS_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If txtPS = "" Then txtPS = 0

End Sub

Private Sub txtPS_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub



Private Sub txtTestify_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    
    Case 48 To 57
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub UserForm_Initialize()

On Error GoTo UserForm_Initialize_Error
Dim zMsg As String

Dim ws As Worksheet
Dim row As Long
Dim TheLastRow As Long
Dim TheColumn, result As Integer
Dim rng As Range
Dim strDocket As String

If IsInternetConnected = False Then
  MsgBox "No Network Connection Detected! Shut down and re-start", vbExclamation, "No Connection"
  Exit Sub
End If
CenterForm Me

Set ws = InvestigationLog
row = ActiveCell.row
Cells(row, 1).Activate
strDocket = Cells(row, 1).Value
If strDocket = "" Then
    result = MsgBox("You need to pick a case first!", vbCritical, "No case selected")
    cmdClose.Visible = False
    cmdUpdate.Visible = False
    cmdCancelClose.SetFocus
    Exit Sub
End If


Application.EnableEvents = False
Application.ScreenUpdating = False    'allow case logs to be seen - true
With ws
  .EnableSelection = xlNoSelection
End With

optUpdate = xlOn
If InvestigationLog.Cells(row, 12).Value = "Closed" Then
    optReOpen = xlOn
    chkPrintIR.Visible = True
    chkPrintIR.Value = False
Else
    optAddSupplement = xlOn
    chkPrintIR.Visible = False
    chkPrintIR.Value = False
End If
frameUpdate.Visible = True
frameUpdate.Top = 138
txtDueDateMemo.Visible = False
Label27.Visible = False
frameClosure.Visible = False
cmdClose.Visible = False


txtCaseNo = Cells(row, 1)
txtXref = Cells(row, 4)
txtClient = Cells(row, 3)

'Populate closure fields
txtFI = Cells(row, 24)
txtDueDil = Cells(row, 25)
txtPI = Cells(row, 23)
txtPhoto = Cells(row, 26)
txtPS = Cells(row, 27)
txtMS = Cells(row, 28)
txtOther = Cells(row, 22)  ' 25 adds due diligence to other category
txtTestify = Cells(row, 30)

    txtAttorney = ActiveCell.Offset(0, 4).Value
    cmboAttorney.Value = RevAttyName(txtAttorney)
    txtAttorney.Enabled = False
    txtNewRecDate = Format(Now(), "m/d/yy")
    txtNewRecTime = Format(Now(), "h:mm AMPM")
    txtNewDueDate = Format(ActiveCell.Offset(0, 5).Value, "m/d/yy")
    txtCtDate = ActiveCell.Offset(0, 9).Value
    txtDept = ActiveCell.Offset(0, 10).Value
    optDueDateOnly = False
    
Cells(row, 3).Activate

'Total Mins
'Filter Case logs 1022 = TheLastRow


Set ws = CaseLogs
    With ws
         .ListObjects("Table2").AutoFilter.ShowAllData
         .AutoFilterMode = False
         '.ListObjects("Table2").Range.AutoFilter Field:=4
         '.ListObjects("Table2").Range.AutoFilter Field:=6
    End With

TheColumn = 1
TheLastRow = CaseLogs.UsedRange.SpecialCells(xlCellTypeLastCell).row
Set rng = CaseLogs.Range("a4", "f" & TheLastRow)


            With rng '
                .AutoFilter TheColumn, strDocket
                
            End With
  
  'To be used with command button to close -- Remove after testing
  

Application.EnableEvents = True


    On Error GoTo 0
Exit Sub

UserForm_Initialize_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: UserForm_Initialize Within: frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"



End Sub

Private Sub UserForm_Terminate()


    On Error GoTo UserForm_Terminate_Error
    Dim zMsg As String


With InvestigationLog
    .EnableSelection = xlNoRestrictions
End With
ActiveCell.Offset(0, 0).Select

Application.ScreenUpdating = True


    On Error GoTo 0
    Exit Sub

UserForm_Terminate_Error:

    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserForm_Terminate Within: Sub frmCloseCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

            
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
    
    End Sub

