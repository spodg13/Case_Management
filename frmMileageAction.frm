VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMileageAction 
   Caption         =   "Mileage Action"
   ClientHeight    =   3996
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7050
   OleObjectBlob   =   "frmMileageAction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMileageAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub cmdAddMileage_Click()

10    On Error GoTo cmdAddMileage_Click_Error
      Dim zMsg As String

      Dim TheLastMileRow As Double
      '************************ERROR CHECK*********************
20    If ValidMileage(txtStartM, txtEndM, txtMileageAddress, Me) = False Then
30        Exit Sub
40    End If
      '***********************FINISHED ERROR CHECK*****************


50        TheLastMileRow = MileageLog.Cells(Rows.Count, 1).End(xlUp).row
60        TheLastMileRow = TheLastMileRow + 1
70        MileageLog.Unprotect
80        MileageLog.Cells(TheLastMileRow, 1).Value = Format(txtDate, "m/d/yy")
90        MileageLog.Cells(TheLastMileRow, 2).Value = txtMileageAddress
100       MileageLog.Cells(TheLastMileRow, 3).Value = txtDocket
110       MileageLog.Cells(TheLastMileRow, 4).Value = Val(txtStartM)
120       MileageLog.Cells(TheLastMileRow, 5).Value = Val(txtEndM)
130       MileageLog.Protect
          
140       ActiveCell.Offset(0, 1).Value = "Mileage Entry"
150       ActiveCell.Offset(0, 2).Value = Val(txtStartM)
160       ActiveCell.Offset(0, 3).Value = Val(txtEndM)
          
170   Application.ScreenUpdating = True
180   ActiveWorkbook.Save
190   Unload frmMileageAction
          

200       On Error GoTo 0
210   Exit Sub

cmdAddMileage_Click_Error:

         
220       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

230       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdAddMileage_Click Within: frmMileageAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

240       Print #1, zMsg

250       Close #1

            
260       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdCancel_Click()
270   Application.ScreenUpdating = True
280   Unload frmMileageAction

End Sub



Private Sub cmdCopy_Click()
290       txtMileageAddress = txtAction
End Sub



Private Sub txtDate_AfterUpdate()

300   On Error GoTo txtDate_AfterUpdate_Error
      Dim zMsg As String

      Dim varDate As Date
310   If IsDate(txtDate) Then
320       varDate = DateValue(txtDate)
330       txtDate = Format(varDate, "MMMM d, yyyy")
340   Else
350       MsgBox "Invalid date"
360       txtDate.SetFocus
370       Exit Sub
          
380   End If

390       On Error GoTo 0
400   Exit Sub

txtDate_AfterUpdate_Error:

         
410       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

420       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_AfterUpdate Within: frmMileageAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

430       Print #1, zMsg

440       Close #1

            
450       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

460   On Error GoTo txtDate_DblClick_Error
      Dim zMsg As String
470   DatePickerForm.Caption = "Mileage Date"
480   DatePickerForm.Show vbModal
490      Select Case [DatePickerForm]![CallingForm].Caption
            Case "Form"
500             If IsaDate(txtDate) = False Then
510                 txtDate = Format(DateValue(Now()), "MMMM d, yyyy")
520             End If
530          Case Else
540             txtDate = [DatePickerForm]![CallingForm].Caption
550    End Select

560       Cancel = True

570       On Error GoTo 0
580   Exit Sub

txtDate_DblClick_Error:

         
590       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

600       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_DblClick Within: frmMileageAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

610       Print #1, zMsg

620       Close #1

            
630       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

640   On Error GoTo txtDate_KeyPress_Error
      Dim zMsg As String

650   If KeyAscii = 43 Then
660       txtDate = DateAdd("d", 1, txtDate)
670       KeyAscii = 0
680       txtDate = Format(DateValue(txtDate), "MMMM d, yyyy")
690   End If
700   If KeyAscii = 45 Then
710       txtDate = DateAdd("d", -1, txtDate)
720       KeyAscii = 0
730       txtDate = Format(DateValue(txtDate), "MMMM d, yyyy")
740   End If

750       On Error GoTo 0
760   Exit Sub

txtDate_KeyPress_Error:

         
770       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

780       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_KeyPress Within: frmMileageAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

790       Print #1, zMsg

800       Close #1

            
810       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub




Private Sub txtDocket_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Permits only numbers, letters, space and underscore characters

820   Select Case KeyAscii
          Case 32  'space
830           KeyAscii = 32
840       Case 45  'hyphen
850           KeyAscii = 45
860       Case 48 To 57  ' Numbers
870           KeyAscii = KeyAscii
880       Case 95  'Underscore
890           KeyAscii = 95
900       Case 65 To 90 ' Cap letters
910           KeyAscii = KeyAscii
920       Case 97 To 122  'Lowercase letters : convert to upper case
930           KeyAscii = KeyAscii - 32
940       Case Else
950           KeyAscii = 0
960   End Select

End Sub

Private Sub txtEndM_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Only allow numbers or a decimal
970   Select Case KeyAscii
          Case Is = 46
980           KeyAscii = KeyAscii
990       Case 48 To 57
1000          KeyAscii = KeyAscii
1010      Case Else
1020          KeyAscii = 0
1030      End Select
End Sub




Private Sub txtMileageAddress_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1040  On Error GoTo txtMileageAddress_KeyDown_Error
      Dim zMsg As String

1050  Select Case KeyCode
              Case vbKeyS And Shift = 2 'Ctrl S
1060              txtMileageAddress = txtMileageAddress & "Sacramento, CA "
                              
1070          Case vbKeyC And Shift = 2 'Ctrl C
1080              txtMileageAddress = txtMileageAddress & "Citrus Heights, CA "
                  
1090          Case vbKeyC And Shift = 3 'Ctrl Shift C
1100              txtMileageAddress = txtMileageAddress & "Carmichael, CA "
                  
1110          Case vbKeyE And Shift = 2 'Ctrl E
1120              txtMileageAddress = txtMileageAddress & "Elk Grove, CA "

1130          Case vbKeyR And Shift = 2 'Ctrl R
1140              txtMileageAddress = txtMileageAddress & "Rancho Cordova, CA "

1150          Case vbKeyF And Shift = 2 'Ctrl F
1160              txtMileageAddress = txtMileageAddress & "Folsom, CA "
       
1170          Case vbKeyA And Shift = 2 'Ctrl A
1180              txtMileageAddress = txtMileageAddress & "Antelope, CA "

1190          Case vbKeyN And Shift = 2 'Ctrl N
1200              txtMileageAddress = txtMileageAddress & "North Highlands, CA "
                  
1210          Case vbKeyR And Shift = 3 'Ctrl Shift R
1220              txtMileageAddress = txtMileageAddress & "Roseville, CA "

1230          Case vbKeyO And Shift = 2 'Ctrl O
1240              txtMileageAddress = txtMileageAddress & "Orangevale, CA "
                  
1250          Case vbKeyG And Shift = 2 'Ctrl G
1260              txtMileageAddress = txtMileageAddress & "Galt, CA "
                  
1270          Case vbKeyF And Shift = 3 'Ctrl Shift F
1280              txtMileageAddress = txtMileageAddress & "Fair Oaks, CA "
                  
                  
1290          Case Else
                'do nothing
1300      End Select

1310      On Error GoTo 0
1320      Exit Sub

txtMileageAddress_KeyDown_Error:
1330      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1
           
1340      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtMileageAddress_KeyDown within: Sub - frmWitnessEntry " & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1350      Print #1, zMsg

1360      Close #1

            
1370      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"


End Sub

Private Sub txtStartM_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Only allow numbers or a decimal
1380  Select Case KeyAscii
          Case Is = 46
1390          KeyAscii = KeyAscii
1400      Case 48 To 57
1410          KeyAscii = KeyAscii
1420      Case Else
1430          KeyAscii = 0
1440      End Select
End Sub

Private Sub UserForm_Initialize()


1450  On Error GoTo UserForm_Initialize_Error
      Dim zMsg As String
1460  CenterForm Me

1470      txtDate = Format(ActiveCell.Offset(0, -4).Value, "m/d/yy")
1480      txtDocket = ActiveCell.Offset(0, -5).Value
1490      txtAction = ActiveCell.Offset(0, -2).Value

          'txtAction.Locked = True  unlocked allows copy and pasting of address
1500      txtDate.Locked = True
1510      txtDocket.Locked = True
          


1520      On Error GoTo 0
1530  Exit Sub

UserForm_Initialize_Error:

         
1540      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1550      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserForm_Initialize Within: frmMileageAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1560      Print #1, zMsg

1570      Close #1

            
1580      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

