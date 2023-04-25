VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MileageLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdMileageLog_Click()

10        frmCreateMileageLog.Show

End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
20    On Error GoTo Worksheet_BeforeDoubleClick_Error
      Dim zMsg As String
      Dim TheLastMileageRow As Long
      Dim result As Integer

      Dim lRow As Long

30    TheLastMileageRow = MileageLog.Cells(Rows.Count, 1).End(xlUp).row
40    TheLastMileageRow = TheLastMileageRow + 1


50    If Target.Column < 7 And Target.row < TheLastMileageRow Then
60        lRow = Target.row
70        If lRow = 1 Then lRow = 2 ' To not allow editing of header
80        Cancel = True
90        result = MsgBox("Would you like to adjust the mileage entry to " & Cells(lRow, 2).Value & " on " & Cells(lRow, 1).Value & " for case " & Cells(lRow, 3).Value & " ?", vbYesNo, Cells(lRow, 3).Value & " " & Cells(lRow, 1).Value)
100       Application.DisplayAlerts = False
110       If result = vbYes Then
120          Cells(lRow, 4).Activate
130          frmMileageError.Show
140       End If
          
150   End If
          


160       Application.DisplayAlerts = True



170       On Error GoTo 0
180       Exit Sub

Worksheet_BeforeDoubleClick_Error:
190       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1
           
200       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Worksheet_BeforeDoubleClick within: Sub - InvestigationLog " & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

210       Print #1, zMsg

220       Close #1

            
230       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
End Sub




Private Sub Worksheet_Activate()
240   MileageLog.Protect UserInterfaceOnly:=True
End Sub



