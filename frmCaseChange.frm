VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCaseChange 
   Caption         =   "Case Edit"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9615
   OleObjectBlob   =   "frmCaseChange.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCaseChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub cmdCancel_Click()
                  
10        Unload frmCaseChange
20        InvestigationLog.togProtection.Value = True
          
End Sub



Private Sub cmdMakeChanges_Click()


30    On Error GoTo cmdMakeChanges_Click_Error
      Dim zMsg As String

      Dim result As Integer
      Dim intRow As Long
      Dim strNewName, strOldName As String

40    intRow = ActiveCell.row
50    Application.EnableEvents = False
60    If txtOldCaseNo = "IOD" Or txtOldCaseNo = "Admin" Then
70        MsgBox ("Cannot delete IOD case or ADMIN case")
80        GoTo exitroutine
90    End If
100   strOldName = Replace(txtOldClientName, ", ", "_") & "_"

110   If optChangeCaseNo = True Then

              'verify txtNewCaseNo is not blank
120           If txtNewCaseNo = "" Then
130               txtNewCaseNo.SetFocus
140               Exit Sub
150           End If
              
160           result = MsgBox("Are you sure you want to change case " & txtOldCaseNo & " to " & txtNewCaseNo & " ?", vbOKCancel, "Confirm change in case numbers for the " & txtOldClientName & " case")
              
170       If result = vbOK Then
180           InvestigationLog.Cells(intRow, 1).Value = txtNewCaseNo
190           ChangeCaseNumber txtOldCaseNo, txtNewCaseNo 'Update Case Logs
200           RenameFilesNewCase txtOldCaseNo, txtNewCaseNo, strOldName  'Update Window Files
210           GoTo exitroutine
220       Else
230           cmdCancel.SetFocus
240       End If

250   End If

260   If optChangeClientName = True Then
          
270       strNewName = txtNewLastName & "_" & txtNewFirstName & "_"
          
          
280       result = MsgBox("Are you sure you want to change case " & "  " & txtOldClientName & " to " & txtNewLastName & ", " & txtNewFirstName & "?", vbOKCancel, "Confirm change in name for case  " & txtOldCaseNo)
          
290       If result = vbOK Then
300           InvestigationLog.Cells(intRow, 3).Value = txtNewLastName & ", " & txtNewFirstName
310           RenameFilesNewClient txtOldCaseNo, strNewName, strOldName
320           GoTo exitroutine
330       Else
340           cmdCancel.SetFocus
350       End If



360   End If

370   If optDeleteCase = True Then
          
380       result = MsgBox("Are you sure you want to delete case " & txtOldCaseNo & " " & txtOldClientName & "?", vbOKCancel, "Confirm delete case")
390       If result = vbOK Then
400           InvestigationLog.Cells(intRow, 1).EntireRow.Delete
410       Else
420           cmdCancel.SetFocus
430       End If
440   End If

exitroutine:
450   Application.EnableEvents = True
460   Unload frmCaseChange

470       On Error GoTo 0
480   Exit Sub

cmdMakeChanges_Click_Error:

         
490       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

500       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdMakeChanges_Click Within: frmCaseChange" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

510       Print #1, zMsg

520       Close #1

            
530       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub optChangeCaseNo_Click()
540       txtNewCaseNo.Visible = True
550       txtNewFirstName.Visible = False
560       txtNewLastName.Visible = False
570       Label2.Visible = True
580       Label4.Visible = False
590       Label5.Visible = False
          
600       cmdMakeChanges.Caption = "Change Case Number"
610       txtNewCaseNo.SetFocus
End Sub

Private Sub optChangeClientName_Click()
620       txtNewCaseNo.Visible = False
630       Label2.Visible = False
640       txtNewFirstName.Visible = True
650       txtNewLastName.Visible = True
660       Label4.Visible = True
670       Label5.Visible = True
          
680       cmdMakeChanges.Caption = "Change Client Name"
690       txtNewFirstName.SetFocus
          
End Sub

Private Sub optDeleteCase_Click()
700       txtNewCaseNo.Visible = False
710       txtNewFirstName.Visible = False
720       txtNewLastName.Visible = False
730       Label2.Visible = False
740       Label4.Visible = False
750       Label5.Visible = False
          
760       cmdMakeChanges.Caption = "Delete Case"
770       cmdMakeChanges.SetFocus
End Sub

Private Sub txtNewCaseNo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

      'Prevents illegal file characters

780   Select Case KeyAscii
          Case 32  'space
790           KeyAscii = 32
800       Case 45  'hyphen
810           KeyAscii = 45
820       Case 48 To 57   'Numbers
830           KeyAscii = KeyAscii
840       Case 95   'Underscore
850           KeyAscii = 95
860       Case 65 To 90   ' Capital letters
870           KeyAscii = KeyAscii
880       Case 97 To 122  'Lowercase
890           KeyAscii = KeyAscii - 32
900       Case Else
910           KeyAscii = 0
920   End Select

End Sub
Private Sub txtNewFirstName_Change()

930   On Error GoTo txtNewFirstName_Change_Error
      Dim zMsg As String

940       txtNewFirstName = Replace(txtNewFirstName, ",", " ")
950       txtNewFirstName = Replace(txtNewFirstName, "/", "-")


960       On Error GoTo 0
970   Exit Sub

txtNewFirstName_Change_Error:

         
980       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

990       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtNewFirstName_Change Within: frmCaseChange" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1000      Print #1, zMsg

1010      Close #1

            
1020      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtNewFirstName_Exit(ByVal Cancel As MSForms.ReturnBoolean)

1030  On Error GoTo txtNewFirstName_Exit_Error
      Dim zMsg As String

1040  If ValidName(txtNewFirstName) = False Then
1050          Cancel = True
1060          txtNewFirstName.SetFocus
1070          MsgBox "Check for illegal characters!"
              
1080      End If
1090      txtNewFirstName = Trim(txtNewFirstName)
1100      txtNewFirstName = ProperCase(txtNewFirstName) 'StrConv(txtNewFirstName, vbProperCase)

1110      On Error GoTo 0
1120  Exit Sub

txtNewFirstName_Exit_Error:

         
1130      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1140      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtNewFirstName_Exit Within: frmCaseChange" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1150      Print #1, zMsg

1160      Close #1

            
1170      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtNewLastName_Change()
1180      txtNewLastName = Replace(txtNewLastName, ",", " ")
1190      txtNewLastName = Replace(txtNewLastName, "/", "-")
End Sub

Private Sub txtNewLastName_Exit(ByVal Cancel As MSForms.ReturnBoolean)

1200  On Error GoTo txtNewLastName_Exit_Error
      Dim zMsg As String

1210  If ValidName(txtNewLastName) = False Then
1220          Cancel = True
1230          txtNewLastName.SetFocus
1240          MsgBox "Check for illegal characters!"
1250       End If
1260       txtNewLastName = Trim(txtNewLastName)
1270       txtNewLastName = ProperCase(txtNewLastName) 'strConv(txtNewLastName, vbPropercase)

1280      On Error GoTo 0
1290  Exit Sub

txtNewLastName_Exit_Error:

         
1300      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1310      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtNewLastName_Exit Within: frmCaseChange" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1320      Print #1, zMsg

1330      Close #1

            
1340      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub UserForm_Initialize()
      Dim intRow As Long
1350  CenterForm Me

1360  intRow = ActiveCell.row
1370  txtOldCaseNo = InvestigationLog.Cells(intRow, 1).Value
1380  txtOldClientName = InvestigationLog.Cells(intRow, 3).Value

1390  txtOldCaseNo.Locked = True
1400  txtOldClientName.Locked = True

1410  txtNewCaseNo.Visible = False
1420  txtNewFirstName.Visible = False
1430  txtNewLastName.Visible = False
1440  Label2.Visible = False
1450  Label4.Visible = False
1460  Label5.Visible = False

1470  cmdMakeChanges.Caption = "Edit Case Info"




End Sub
