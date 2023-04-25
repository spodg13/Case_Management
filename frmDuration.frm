VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDuration 
   Caption         =   "Add Duration"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9090
   OleObjectBlob   =   "frmDuration.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'---------------------------------------------------------------------------------------
' File   : frmDuration
' Author : gouldd
' Date   : 9/1/2016
' Purpose: Only called when closing a case and the duration of a task is missing
'---------------------------------------------------------------------------------------

Private Sub cmdCancelClose_Click()
          
10        Unload frmDuration
20        End
End Sub



Private Sub cmdNext_Click()


30    On Error GoTo cmdNext_Click_Error
      Dim zMsg As String

40        ActiveCell.Value = Val(txtDuration)
50        ActiveCell.Offset(0, -1).Value = txtActions
60        ActiveWorkbook.Save
70        Unload frmDuration
          

80        On Error GoTo 0
90    Exit Sub

cmdNext_Click_Error:

         
100       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

110       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdNext_Click Within: frmDuration" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

120       Print #1, zMsg

130       Close #1

            
140       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub txtDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

150   On Error GoTo txtDate_DblClick_Error
      Dim zMsg As String
160   DatePickerForm.Caption = "Action Date"
170   DatePickerForm.Show vbModal
180       Select Case [DatePickerForm]![CallingForm].Caption
            Case "Form"
190             If IsaDate(txtDate) = False Then
200                 txtDate = Format(DateValue(Now()), "MMMM d, yyyy")
210             End If
220          Case Else
230             txtDate = [DatePickerForm]![CallingForm].Caption
240    End Select

250       Cancel = True

260       On Error GoTo 0
270   Exit Sub

txtDate_DblClick_Error:

         
280       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

290       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_DblClick Within: frmDuration" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

300       Print #1, zMsg

310       Close #1

            
320       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)


330   On Error GoTo txtDate_KeyPress_Error
      Dim zMsg As String

340   If KeyAscii = 43 Then
350       txtDate.Locked = False
360       txtDate = DateAdd("d", 1, txtDate)
370       KeyAscii = 0
380       txtDate = Format(DateValue(txtDate), "m/d/yy")
390       txtDate.Locked = True
400   End If
410   If KeyAscii = 45 Then
420       txtDate.Locked = False
430       txtDate = DateAdd("d", -1, txtDate)
440       KeyAscii = 0
450       txtDate = Format(DateValue(txtDate), "m/d/yy")
460       txtDate.Locked = True
470   End If

480       On Error GoTo 0
490   Exit Sub

txtDate_KeyPress_Error:

         
500       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

510       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_KeyPress Within: frmDuration" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

520       Print #1, zMsg

530       Close #1

            
540       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtTime_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)


550   On Error GoTo txtTime_KeyPress_Error
      Dim zMsg As String

      Dim LMin, LMinNew, IntervalAdd As Integer

560   If KeyAscii = 43 Then
570       txtTime.Locked = False
580       LMin = Minute(txtTime)
590       LMinNew = Round(LMin / 5, 0) * 5
600       IntervalAdd = (LMin - LMinNew) * -1
610       txtTime = DateAdd("N", IntervalAdd, txtTime)
          
620       txtTime = DateAdd("N", 5, txtTime)
630       KeyAscii = 0
640       txtTime = Format(txtTime, "h:mm AM/PM")
650       txtTime.Locked = True
660   End If

670   If KeyAscii = 42 Then
         
680       txtTime = DateAdd("N", 1, txtTime)
690       KeyAscii = 0
700       txtTime = Format(txtTime, "h:mm AM/PM")
710       txtTime.Locked = True
720   End If

730   If KeyAscii = 47 Then
            
740       txtTime = DateAdd("N", -1, txtTime)
750       KeyAscii = 0
760       txtTime = Format(txtTime, "h:mm AM/PM")
770       txtTime.Locked = True
780   End If
790   If KeyAscii = 45 Then
800       txtTime.Locked = False
810       LMin = Minute(txtTime)
820       LMinNew = Round(LMin / 5, 0) * 5
830       IntervalAdd = (LMin - LMinNew) * -1
840       txtTime = DateAdd("N", IntervalAdd, txtTime)
          
850       txtTime = DateAdd("N", -5, txtTime)
860       KeyAscii = 0
870       txtTime = Format(txtTime, "h:mm AM/PM")
880       txtTime.Locked = True
          
890   End If

900       On Error GoTo 0
910   Exit Sub

txtTime_KeyPress_Error:

         
920       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

930       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtTime_KeyPress Within: frmDuration" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

940       Print #1, zMsg

950       Close #1

            
960       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub UserForm_Initialize()


970   On Error GoTo UserForm_Initialize_Error
      Dim zMsg As String
980   CenterForm Me

990       txtActions = ActiveCell.Offset(0, -1).Value
1000      txtDate = ActiveCell.Offset(0, -3).Value
1010      txtTime = Format(ActiveCell.Offset(0, -2).Value, "h:mm AM/PM")
1020      txtTime.Locked = True
1030      txtDate.Locked = True
1040      txtActions.Locked = False
1050      txtDuration.SetFocus
          
          

1060      On Error GoTo 0
1070  Exit Sub

UserForm_Initialize_Error:

         
1080      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1090      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserForm_Initialize Within: frmDuration" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1100      Print #1, zMsg

1110      Close #1

            
1120      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

