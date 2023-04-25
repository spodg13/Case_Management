VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMileageError 
   Caption         =   "Check Mileage Entries"
   ClientHeight    =   2964
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8130
   OleObjectBlob   =   "frmMileageError.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMileageError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' File   : frmMileageError
' Author : gouldd
' Date   : 9/1/2016
' Purpose: Catches mileage equal to zero when creating mileage log.  Likely not going to be used as starting and ending mileage is mandatory.
'---------------------------------------------------------------------------------------




Private Sub cmdCancelMileage_Click()
10        Unload frmMileageError
20        End
End Sub

Private Sub cmdNextMileage_Click()


30    On Error GoTo cmdNextMileage_Click_Error
      Dim zMsg As String

40        ActiveCell.Value = Val(txtStartingM)
50        ActiveCell.Offset(0, 1).Value = Val(txtEndingM)
60        ActiveWorkbook.Save
70        Unload frmMileageError


80        On Error GoTo 0
90    Exit Sub

cmdNextMileage_Click_Error:

         
100       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

110       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdNextMileage_Click Within: frmMileageError" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

120       Print #1, zMsg

130       Close #1

            
140       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub






Private Sub txtEndingM_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

150   On Error GoTo txtEndingM_BeforeUpdate_Error
      Dim zMsg As String

160   If Val(txtEndingM) < Val(txtStartingM) Then
170       MsgBox "Enter a value greater than your starting mileage!"
180       txtEndingM.SetFocus
190       Cancel = True
200   End If

210       On Error GoTo 0
220   Exit Sub

txtEndingM_BeforeUpdate_Error:

         
230       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

240       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtEndingM_BeforeUpdate Within: frmMileageError" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

250       Print #1, zMsg

260       Close #1

            
270       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub UserForm_Initialize()

280   On Error GoTo UserForm_Initialize_Error
      Dim zMsg As String
290   CenterForm Me

300       txtStartingM = ActiveCell.Value
310       txtEndingM = ActiveCell.Offset(0, 1).Value
320       txtAddress = ActiveCell.Offset(0, -2).Value
330       txtDate = ActiveCell.Offset(0, -3).Value
340       txtCase = ActiveCell.Offset(0, -1).Value
350       txtCase.Locked = True
360       txtDate.Locked = True
370       txtAddress.Locked = True
          

380       On Error GoTo 0
390   Exit Sub

UserForm_Initialize_Error:

         
400       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

410       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserForm_Initialize Within: frmMileageError" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

420       Print #1, zMsg

430       Close #1

            
440       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

