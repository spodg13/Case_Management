VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTemplateFiles 
   Caption         =   "Template Files and Location"
   ClientHeight    =   6636
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7830
   OleObjectBlob   =   "frmTemplateFiles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTemplateFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdCancel_Click()
10        Unload Me
End Sub

Private Sub cmdMakeChanges_Click()

20    On Error GoTo cmdMakeChanges_Click_Error
      Dim zMsg As String

      Dim contr As Control
      Dim strTemplateFileName As String
      Dim strTemplatePath As String
      Dim result As Integer

30    result = MsgBox("Folder and name changes might be against policy.  Do you want to continue?", vbYesNo, "Policy warning")
40    If result = vbNo Then
50        Unload Me
60        Exit Sub
70    End If

80    For Each contr In frmTemplateFiles.Controls
90        If TypeName(contr) = "CheckBox" And contr.Value = True Then
100           If InStr(1, Files.Cells(contr.Tag, 1).Value, "Folder") Then
110               result = MsgBox("Please select the " & contr.Caption & "!", vbOKCancel, "Need the file or path")
120                   If result = vbCancel Then GoTo Nexti
130                   strTemplatePath = PathPicked(contr.Caption)
140                   Files.Cells(contr.Tag, 2).Value = strTemplatePath
150           End If
                 
                 
160              If InStr(1, Files.Cells(contr.Tag, 1).Value, "Template") Then
170                   result = MsgBox("Please select the " & contr.Caption & "!", vbOKCancel, "Need the file or path")
180                   If result = vbCancel Then GoTo Nexti
190                   strTemplateFileName = FilePicked(contr.Caption)
200                   Files.Cells(contr.Tag, 2).Value = strTemplateFileName
210              End If
220     End If
Nexti:

230   Next

240   Unload Me


250       On Error GoTo 0
260   Exit Sub

cmdMakeChanges_Click_Error:

         
270       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

280       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdMakeChanges_Click Within: frmTemplateFiles" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

290       Print #1, zMsg

300       Close #1

            
310       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub UserForm_Initialize()
320   CenterForm Me

End Sub
