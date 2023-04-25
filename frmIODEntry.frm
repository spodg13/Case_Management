VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIODEntry 
   Caption         =   "IOD Action"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7770
   OleObjectBlob   =   "frmIODEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIODEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdIODEntry_Click()

10    On Error GoTo cmdIODEntry_Click_Error
      Dim zMsg As String

      Dim TheLastCaseRow, TheLastIODRow As Long
20    Application.ScreenUpdating = False

30        TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
40        TheLastCaseRow = TheLastCaseRow + 1
          
50        TheLastIODRow = IOD.Cells(Rows.Count, 1).End(xlUp).row
60        TheLastIODRow = TheLastIODRow + 1
          
          
70        CaseLogs.Cells(TheLastCaseRow, 1).Value = "IOD"
80        CaseLogs.Cells(TheLastCaseRow, 2).Value = GetADate(txtDOA) 'Format(txtDOA, "m/d/yy")
90        CaseLogs.Cells(TheLastCaseRow, 3).Value = Format(txtTOA, "h:mm AMPM")
100       CaseLogs.Cells(TheLastCaseRow, 4).Value = txtActions
110       CaseLogs.Cells(TheLastCaseRow, 5).Value = Val(txtDuration)
          
120       IOD.Cells(TheLastIODRow, 1).Value = txtClient & ", " & txtDocket & ", " & txtAttorney
130       IOD.Cells(TheLastIODRow, 2).Value = txtActions
140       IOD.Cells(TheLastIODRow, 4).Value = GetADate(txtDOA) 'Format(txtDOA, "m/d/yy")
150       If chkCompleted = True Then
160           IOD.Cells(TheLastIODRow, 3).Value = "Yes"
170       Else
180           IOD.Cells(TheLastIODRow, 3).Value = "No"
190       End If
          
200       Application.ScreenUpdating = True
210       ActiveWorkbook.Save
220       Unload frmIODEntry



230       On Error GoTo 0
240   Exit Sub

cmdIODEntry_Click_Error:

         
250       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

260       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdIODEntry_Click Within: frmIODEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

270       Print #1, zMsg

280       Close #1

            
290       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdCancel_Click()
300       Unload frmIODEntry
End Sub



Private Sub txtClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)

310   On Error GoTo txtClient_Exit_Error
      Dim zMsg As String

320       txtClient = ProperCase(txtClient)

330       On Error GoTo 0
340       Exit Sub

txtClient_Exit_Error:
350       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1
           
360       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtClient_Exit within: Sub - frmIODEntry " & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

370       Print #1, zMsg

380       Close #1

            
390       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
          
End Sub

Private Sub txtDOA_AfterUpdate()

400   On Error GoTo txtDOA_AfterUpdate_Error
      Dim zMsg As String

      Dim varDate As Date
410   If IsDate(txtDOA) Then
420       varDate = DateValue(txtDOA)
430       txtDOA = Format(varDate, "MMMM d, yyyy")
440   Else
450       MsgBox "Invalid date"
460       txtDOA.SetFocus
470       Exit Sub
          
480   End If

490       On Error GoTo 0
500   Exit Sub

txtDOA_AfterUpdate_Error:

         
510       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

520       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDOA_AfterUpdate Within: frmIODEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

530       Print #1, zMsg

540       Close #1

            
550       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Private Sub txtDOA_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

560   On Error GoTo txtDOA_DblClick_Error
      Dim zMsg As String
570   DatePickerForm.Caption = "IOD Date"
580   DatePickerForm.Show vbModal
590       Select Case [DatePickerForm]![CallingForm].Caption
            Case "Form"
600             If IsaDate(txtDOA) = False Then
610                 txtDOA = Format(DateValue(Now()), "MMMM d, yyyy")
620             End If
630          Case Else
640             txtDOA = [DatePickerForm]![CallingForm].Caption
650    End Select

660       Cancel = True

670       On Error GoTo 0
680   Exit Sub

txtDOA_DblClick_Error:

         
690       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

700       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDOA_DblClick Within: frmEnterAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

710       Print #1, zMsg

720       Close #1

            
730       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Private Sub txtDOA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

740   On Error GoTo txtDOA_KeyPress_Error
      Dim zMsg As String

750   If KeyAscii = 43 Then
760       txtDOA = DateAdd("d", 1, txtDOA)
770       KeyAscii = 0
780       txtDOA = Format(DateValue(txtDOA), "MMMM d, yyyy")
790   End If
800   If KeyAscii = 45 Then
810       txtDOA = DateAdd("d", -1, txtDOA)
820       KeyAscii = 0
830       txtDOA = Format(DateValue(txtDOA), "MMMM d, yyyy")
840   End If

850       On Error GoTo 0
860   Exit Sub

txtDOA_KeyPress_Error:

         
870       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

880       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDOA_KeyPress Within: frmIODEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

890       Print #1, zMsg

900       Close #1

            
910       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub






Private Sub txtDocket_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Permits only numbers, letters, space and underscore characters

920   Select Case KeyAscii
          Case 32  'space
930           KeyAscii = 32
940       Case 45  'hyphen
950           KeyAscii = 45
960       Case 48 To 57  ' Numbers
970           KeyAscii = KeyAscii
980       Case 95  'Underscore
990           KeyAscii = 95
1000      Case 65 To 90 ' Cap letters
1010          KeyAscii = KeyAscii
1020      Case 97 To 122  'Lowercase letters : convert to upper case
1030          KeyAscii = KeyAscii - 32
1040      Case Else
1050          KeyAscii = 0
1060  End Select

End Sub

Private Sub txtTOA_Exit(ByVal Cancel As MSForms.ReturnBoolean)

1070  On Error GoTo txtTOA_Exit_Error
      Dim zMsg As String

      Dim result As Integer
1080  If IsTime(txtTOA) = False Then
1090      Cancel = True
1100      txtTOA.SetFocus
1110      result = MsgBox("Invalid Time, check your entry!", vbOKOnly, "Check time entry")
1120  End If

1130      On Error GoTo 0
1140  Exit Sub

txtTOA_Exit_Error:

         
1150      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1160      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtTOA_Exit Within: frmIODEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1170      Print #1, zMsg

1180      Close #1

            
1190      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtTOA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

1200  On Error GoTo txtTOA_KeyPress_Error
      Dim zMsg As String

      Dim LMin, LMinNew, IntervalAdd As Integer

1210  If KeyAscii = 43 Then
1220      LMin = Minute(txtTOA)
1230      LMinNew = Round(LMin / 5, 0) * 5
1240      IntervalAdd = (LMin - LMinNew) * -1
1250      txtTOA = DateAdd("N", IntervalAdd, txtTOA)
          
1260      txtTOA = DateAdd("N", 5, txtTOA)
1270      KeyAscii = 0
1280      txtTOA = Format(txtTOA, "h:mm AM/PM")
1290  End If
1300  If KeyAscii = 42 Then
               
1310         txtTOA = DateAdd("N", 1, txtTOA)
1320         KeyAscii = 0
1330         txtTOA = Format(txtTOA, "h:mm AM/PM")
1340     End If



1350  If KeyAscii = 45 Then
1360      LMin = Minute(txtTOA)
1370      LMinNew = Round(LMin / 5, 0) * 5
1380      IntervalAdd = (LMin - LMinNew) * -1
1390      txtTOA = DateAdd("N", IntervalAdd, txtTOA)
          
1400      txtTOA = DateAdd("N", -5, txtTOA)
1410      KeyAscii = 0
1420      txtTOA = Format(txtTOA, "h:mm AM/PM")
1430  End If
1440  If KeyAscii = 47 Then
                  
1450         txtTOA = DateAdd("N", -1, txtTOA)
1460         KeyAscii = 0
1470         txtTOA = Format(txtTOA, "h:mm AM/PM")
1480     End If

1490      On Error GoTo 0
1500  Exit Sub

txtTOA_KeyPress_Error:

         
1510      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1520      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtTOA_KeyPress Within: frmIODEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1530      Print #1, zMsg

1540      Close #1

            
1550      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub UserForm_Initialize()

1560  On Error GoTo UserForm_Initialize_Error
      Dim zMsg As String
1570  CenterForm Me

1580      txtDOA = Format(Now(), "MMMM d, yyyy")
1590      txtTOA = Format(Now(), "h:mm AM/PM")
1600      chkCompleted = True
1610      txtClient.SetFocus
          
          

1620      On Error GoTo 0
1630  Exit Sub

UserForm_Initialize_Error:

         
1640      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1650      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserForm_Initialize Within: frmIODEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1660      Print #1, zMsg

1670      Close #1

            
1680      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

