VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOpenCase 
   Caption         =   "Open Case"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10635
   OleObjectBlob   =   "frmOpenCase.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOpenCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit



Private Sub cmboAttorney_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)


On Error GoTo cmboAttorney_BeforeUpdate_Error
Dim zMsg As String

Dim rngAtty As Range
Dim strPart As String
Dim lCount As Integer
Dim sAddress As String

  sAddress = ActiveCell.Address
If InStr(1, cmboAttorney.Value, ",") = 0 Then
    MsgBox "Please enter using the format Lastname, firstname"
    Cancel = True
    cmboAttorney.SetFocus
    Exit Sub
End If

Set rngAtty = Attorneys.Range("Attorneys")
strPart = Trim(Me.cmboAttorney.Value)
lCount = Application.WorksheetFunction.CountIf(rngAtty, strPart)
If lCount = 0 Then
    'Add Attorney to table
    addAttorney strPart
    InvestigationLog.Range(sAddress).Select
    
End If

    On Error GoTo 0
Exit Sub

cmboAttorney_BeforeUpdate_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: cmboAttorney_BeforeUpdate Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdCancel_Click()

    Unload frmOpenCase

End Sub

Private Sub cmdCreateCase_Click()


On Error GoTo cmdCreateCase_Click_Error
Dim zMsg As String

 Dim wdApp As New Word.Application
    Dim wdDoc As Word.Document
    Dim CCtrl As Word.ContentControl
    Dim strFileName, strPath, strTemplateFileName, strVersion As String
    Dim Wksht As Worksheet, lngRow As Long, intCol As Integer, TheLastRow, TheLastCaseRow As Long
    Dim result, CounterColumn As Integer
    Dim AnyColumn As Integer
    Dim CellAltered As Boolean
    Dim strPhotoFolder As String
   
     
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Verify case does not exist
    If CaseExists(txtCaseNo) Then
        result = MsgBox("Case exists, use the re-open button!", vbOKOnly, "Case Exists")
        Unload frmOpenCase
        frmCloseCase.Show
        Exit Sub
    End If
    
    'Populate InvestigationLog
    ' TheLastRow = InvestigationLog.Cells(Rows.Count, 1).End(xlUp).row
     'TheLastRow = TheLastRow + 1
     'Cells(TheLastRow, 1).Activate
     InvestigationLog.Unprotect
     AddRowToTable "Investigation_Log", txtCaseNo
     InvestigationLog.Protect
     'Exit Sub
     'Cells(TheLastRow + 1, 1).Activate
     ActiveCell.Offset(0, 0).Activate
     TheLastRow = InvestigationLog.Cells(Rows.Count, 1).End(xlUp).row
     'TheLastRow = ActiveCell.row
     
      If IsActiveCellInTable = False Then
          MsgBox "Error putting case in table.  Please try again.  If this is the second try, shut down and restart."
          Unload frmOpenCase
          Application.ScreenUpdating = True
          Application.EnableEvents = True
          Exit Sub
      End If
          
     ActiveCell.Value = UCase(txtCaseNo)
     ActiveCell.NumberFormat = "@"
     Cells(TheLastRow, 2).Value = DateValue(txtNewRecDate)
     Cells(TheLastRow, 3).Value = txtClientLast & ", " & txtClientFirst
     Cells(TheLastRow, 4).Value = Val(txtXref)
     'Change order of attorney
     Cells(TheLastRow, 5).Value = AttorneyName(cmboAttorney)
     Cells(TheLastRow, 6).Value = DateValue(txtNewDueDate)
     'Phrase task requests -
          Cells(TheLastRow, 8).Value = TaskList(txtNewInt, txtNewPhoto, txtNewSub, txtNewOther, False)
     
     
     Cells(TheLastRow, 9).Value = txtCharges
     Cells(TheLastRow, 10).Value = DateValue(txtCtDate)
     Cells(TheLastRow, 11).Value = txtDept
     'Request Numbers
     Cells(TheLastRow, 31).Value = Val(txtNewInt)
     Cells(TheLastRow, 32).Value = Val(txtNewPhoto)
     If Val(txtNewPhoto) > 0 Then
        strPhotoFolder = Files.Cells(29, 2).Value & txtClientLast & "_" & txtClientFirst & "_" & txtCaseNo
        MakePhotoDir (strPhotoFolder)
     End If
     Cells(TheLastRow, 33).Value = Val(txtNewSub)
     Cells(TheLastRow, 34).Value = Val(txtNewOther)
     If ckbxNewTestify = True Then
        Cells(TheLastRow, 35).Value = 1
     Else
        Cells(TheLastRow, 35).Value = 0
     End If
     If ckBoxJuv = True Then
        Cells(TheLastRow, 15).Value = "Juv"
        Else
        Cells(TheLastRow, 15).Value = "Adult"
     End If
     InvestigationLog.Cells(TheLastRow, 12).Value = "Open"
     
     
    TheLastCaseRow = CaseLogs.Cells(65536, "A").End(xlUp).row
    TheLastCaseRow = TheLastCaseRow + 1
    Set Wksht = ActiveSheet
    TheLastRow = ActiveCell.row
    intCol = 4 ' Place on  xref client name
    strVersion = "1"
    Cells(TheLastRow, intCol).Activate
    Cells(TheLastRow, 3).Activate 'Place on client
    
    
    'Does an Action Log exist

    strFileName = ActiveCell.Value
    strFileName = Replace(strFileName, ", ", "_")
    strFileName = strFileName & "_" & ActiveCell.Offset(0, -2).Value
    
    If strVersion > 1 Then
        strFileName = strFileName & "_" & strVersion
    End If
    strPath = Files.Cells(6, 2).Value
    strTemplateFileName = Files.Cells(7, 2).Value
    If strTemplateFileName = "" Then
        result = MsgBox("Please select the Action Log Template Location!", vbCritical, "Need the file name")
        strTemplateFileName = FilePicked("Action Log Template")
        Files.Cells(7, 2).Value = strTemplateFileName
    End If
    If strPath = "" Then
        result = MsgBox("Please select the path where the Action Log will be stored!", vbCritical, "Need the path")
        strPath = PathPicked("Action Log") & "\"
        Files.Cells(6, 2).Value = strPath
    End If
    'Check due date and adjust for Mondays and day off
    
    AnyColumn = ActiveCell.Column + 3
    CellAltered = False
    ActiveCell.Offset(0, 3).Value = ActiveCell.Offset(0, 3).Value + TimeSerial(8, 0, 0)
  
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
    
    'Set report totals
    For CounterColumn = 22 To 26 Step 1
        Cells(TheLastRow, CounterColumn).Value = 0
    Next
    
    'Update CaseLogs
    
        
    CaseLogs.Cells(TheLastCaseRow, 1).Value = InvestigationLog.Cells(TheLastRow, 1).Text
    CaseLogs.Cells(TheLastCaseRow, 2).Value = InvestigationLog.Cells(TheLastRow, 2).Value
    CaseLogs.Cells(TheLastCaseRow, 3).Value = Format(txtNewRecTime, "h:mm AMPM")
    CaseLogs.Cells(TheLastCaseRow, 4).Value = "Received and reviewed investigative request"
    CaseLogs.Cells(TheLastCaseRow, 4).font.Bold = True
    CaseLogs.Cells(TheLastCaseRow, 5).Value = Val(txtDuration)
    CaseLogs.Cells(TheLastCaseRow, 7).Value = "Start"
    PopulateCombo
    
     
    
        'W:\Investigations\DG\Action Logs\ActionLog.dotx
    
        Set wdDoc = wdApp.Documents.Open(FileName:=strTemplateFileName, AddToRecentFiles:=True, Visible:=False)
        With wdDoc
            
            For Each CCtrl In .ContentControls
                Select Case CCtrl.Title
                    Case "CaseNum"
                        CCtrl.Range.Text = Cells(TheLastRow, 1).Value
                    Case "Client"
                        CCtrl.Range.Text = Cells(TheLastRow, 3).Value
                    Case "xref"
                        CCtrl.Range.Text = Cells(TheLastRow, 4).Value
                    Case "Atty"
                        CCtrl.Range.Text = Cells(TheLastRow, 5).Value
                    Case "DueDate"
                        CCtrl.Range.Text = Cells(TheLastRow, 6).Value
                    Case "Charges"
                        CCtrl.Range.Text = Cells(TheLastRow, 9).Value
                    Case "InvName"
                        CCtrl.Range.Text = Files.Cells(20, 2).Value
                    Case "InvPhone"
                        CCtrl.Range.Text = Files.Cells(23, 2).Value
                    Case "InvCell"
                        CCtrl.Range.Text = Files.Cells(24, 2).Value
                    
                End Select
             Next
        End With
        'Save with new name
        wdDoc.SaveAs FileName:=strPath & strFileName & ".docx", FileFormat:=wdFormatDocumentDefault, AddToRecentFiles:=True
        ActiveCell.Offset(0, 10) = "Yes"
        ActiveCell.Offset(0, 11) = 1
               
    wdDoc.Close
    wdApp.Quit
    Set wdApp = Nothing
    Set wdDoc = Nothing
    
    
    ActiveCell.Offset(0, -2).Activate
    SortByDueDate
    ActiveCell.Offset(1, 0).Select
    ActiveWorkbook.Save
    Unload frmOpenCase
    AppActivate "Microsoft Excel"
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    

    On Error GoTo 0
Exit Sub

cmdCreateCase_Click_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: cmdCreateCase_Click Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub




Private Sub txtCaseNo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'Permits only numbers, letters, space and underscore characters

Select Case KeyAscii
    Case 32  'space
        KeyAscii = 32
    Case 45  'hyphen
        KeyAscii = 45
    Case 48 To 57  ' Numbers
        KeyAscii = KeyAscii
    Case 95  'Underscore
        KeyAscii = 95
    Case 65 To 90 ' Cap letters
        KeyAscii = KeyAscii
    Case 97 To 122  'Lowercase letters : convert to upper case
        KeyAscii = KeyAscii - 32
    Case Else
        KeyAscii = 0
End Select

End Sub

Private Sub txtClientFirst_Change()


On Error GoTo txtClientFirst_Change_Error
Dim zMsg As String

    txtClientFirst = Replace(txtClientFirst, ",", " ")
    txtClientFirst = Replace(txtClientFirst, "/", "-")
    'txtClientFirst = Replace(txtClientFirst, ".", "")
    

    On Error GoTo 0
Exit Sub

txtClientFirst_Change_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtClientFirst_Change Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtClientFirst_Exit(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo txtClientFirst_Exit_Error
Dim zMsg As String

If ValidName(txtClientFirst) = False Then
        Cancel = True
        txtClientFirst.SetFocus
        MsgBox "Check for illegal characters!"
        
    End If
    txtClientFirst = Trim(txtClientFirst)
    txtClientFirst = StrConv(txtClientFirst, vbProperCase)

    On Error GoTo 0
Exit Sub

txtClientFirst_Exit_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtClientFirst_Exit Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtClientLast_Change()

On Error GoTo txtClientLast_Change_Error
Dim zMsg As String

    txtClientLast = Replace(txtClientLast, ",", " ")
    txtClientLast = Replace(txtClientLast, "/", "-")
    'txtClientLast = Replace(txtClientLast, ".", "")

    On Error GoTo 0
Exit Sub

txtClientLast_Change_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtClientLast_Change Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtClientLast_Exit(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo txtClientLast_Exit_Error
Dim zMsg As String

If ValidName(txtClientLast) = False Then
        Cancel = True
        txtClientLast.SetFocus
        MsgBox "Check for illegal characters!"
     End If
     txtClientLast = Trim(txtClientLast)
     txtClientLast = ProperCase(txtClientLast) 'StrConv(txtClientLast, vbProperCase)

    On Error GoTo 0
Exit Sub

txtClientLast_Exit_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtClientLast_Exit Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub





Private Sub txtClientLast_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = vbKeyT And Shift = 2 Then
        txtNewInt = Cells(ActiveCell.row, 23).Value + Cells(ActiveCell.row, 24).Value & " of " & Cells(ActiveCell.row, 31).Value & " Int. Completed."
        txtNewPhoto = Cells(ActiveCell.row, 26).Value & " of " & Cells(ActiveCell.row, 32).Value & " Photos Completed."
        txtNewSub = Cells(ActiveCell.row, 27).Value + Cells(ActiveCell.row, 28).Value & " of " & Cells(ActiveCell.row, 33).Value & " Subs Served."
        txtNewOther = Cells(ActiveCell.row, 22).Value & " of " & Cells(ActiveCell.row, 34).Value & " Other Completed."
        txtXref = Cells(ActiveCell.row, 4).Value
        txtCaseNo = Cells(ActiveCell.row, 1).Value
        txtClientLast = Cells(ActiveCell.row, 3).Value
        cmdCreateCase.Visible = False
        txtNewRecTime.Visible = False
        txtDuration.Visible = False
        Label31.Visible = False
        Label32.Visible = False

        cmboAttorney = Cells(ActiveCell.row, 5).Value
        txtNewDueDate = Cells(ActiveCell.row, 6).Value
        txtNewRecDate = Cells(ActiveCell.row, 2).Value
        txtNewRecTime = ""
        txtCtDate = Cells(ActiveCell.row, 10).Value
        txtDept = Cells(ActiveCell.row, 11).Value
        txtCharges = Cells(ActiveCell.row, 9).Value
        If Cells(ActiveCell.row, 25).Value > 0 Then
          txtCharges = txtCharges & "  NOTE: " & Cells(ActiveCell.row, 25).Value & " due diligence(s) completed."
        End If
 End If
End Sub

Private Sub txtCtDate_AfterUpdate()
Dim varDate As Date
If IsDate(txtCtDate) Then
    varDate = DateValue(txtCtDate)
    txtCtDate = Format(varDate, "MMMM d, yyyy")
Else
    MsgBox "Invalid date"
    txtCtDate.SetFocus
    Exit Sub
    
End If
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
              "Procedure: txtCtDate_DblClick Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

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
              "Procedure: txtCtDate_KeyPress Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDept_AfterUpdate()
If txtDept.Value > 69 Then
    ckBoxJuv = True
End If

End Sub


Private Sub txtNewDueDate_AfterUpdate()

On Error GoTo txtNewDueDate_AfterUpdate_Error
Dim zMsg As String

Dim varDate As Date
If IsDate(txtNewDueDate) Then
    varDate = DateValue(txtNewDueDate)
    txtNewDueDate = Format(varDate, "MMMM d, yyyy")
Else
    MsgBox "Invalid date"
    txtNewDueDate.SetFocus
    Exit Sub
    
End If

    On Error GoTo 0
Exit Sub

txtNewDueDate_AfterUpdate_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtNewDueDate_AfterUpdate Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

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
              "Procedure: txtNewDueDate_DblClick Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

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
              "Procedure: txtNewDueDate_KeyPress Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub





Private Sub txtNewRecDate_AfterUpdate()

On Error GoTo txtNewRecDate_AfterUpdate_Error
Dim zMsg As String

Dim varDate As Date
If IsDate(txtNewRecDate) Then
    varDate = DateValue(txtNewRecDate)
    txtNewRecDate = Format(varDate, "MMMM d, yyyy")
Else
    MsgBox "Invalid date"
    txtNewRecDate.SetFocus
    Exit Sub
    
End If

    On Error GoTo 0
Exit Sub

txtNewRecDate_AfterUpdate_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtNewRecDate_AfterUpdate Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtNewRecDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo txtNewRecDate_DblClick_Error
Dim zMsg As String
DatePickerForm.Caption = "Received Date"
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
              "Procedure: txtNewRecDate_DblClick Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

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
              "Procedure: txtNewRecDate_KeyPress Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub txtNewRecTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo txtNewRecTime_Exit_Error
Dim zMsg As String

Dim result As Integer
If IsTime(txtNewRecTime) = False Then
    Cancel = True
    txtNewRecTime.SetFocus
    result = MsgBox("Invalid Time, check your entry! " & txtNewRecTime, vbOKOnly, "Check time entry")
End If

    On Error GoTo 0
Exit Sub

txtNewRecTime_Exit_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: txtNewRecTime_Exit Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

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
              "Procedure: txtNewRecTime_KeyPress Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub UserForm_Initialize()


On Error GoTo UserForm_Initialize_Error
Dim zMsg As String

If IsInternetConnected = False Then
  MsgBox "No Network Connection Detected! Shut down and re-start", vbExclamation, "No Connection"
  Exit Sub
End If
CenterForm Me

     txtNewRecDate = Format(Now(), "m/d/yy")
     txtNewRecTime = Format(Now(), "h:mm AM/PM")
     txtNewDueDate = Format(Now(), "m/d/yy")
     txtCtDate = Format(Now(), "m/d/yy")
      If Files.Cells(28, 2).Value = True Then
         ckBoxJuv = True
     Else
        ckBoxJuv = False
     End If
     

     

    On Error GoTo 0
Exit Sub

UserForm_Initialize_Error:

   
    Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

    zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
              Format(Erl, "###") & vbCrLf & _
              "Procedure: UserForm_Initialize Within: frmOpenCase" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

    Print #1, zMsg

    Close #1

      
    MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

