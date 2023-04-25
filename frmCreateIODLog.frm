VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateIODLog 
   Caption         =   "Verify the date before creating the log"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmCreateIODLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateIODLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' File   : frmCreateIODLog
' Author : gouldd
' Date   : 9/1/2016
' Purpose: Creates the IOD Log
'---------------------------------------------------------------------------------------




Private Sub cmdCancel_Click()
                 
10        Unload frmCreateIODLog
End Sub

Private Sub cmdCreateIODLog_Click()

20    On Error GoTo cmdCreateIODLog_Click_Error
      Dim zMsg As String

      Dim strFileName, strPath, strComment As String
      Dim TheColumn, TheLastIODRow As Long
      Dim rng, rangeToCopy As Range
      Dim varDate As Variant

      Dim CCtrl As Word.ContentControl

30    strPath = Files.Cells(8, 2).Value
40    strFileName = Files.Cells(9, 2).Value

      'Sort Table
50     IOD.ListObjects("IODTable").Sort.SortFields.Clear
60        IOD.ListObjects("IODTable").Sort.SortFields.Add _
              Key:=Range("IODTable[Date]"), SortOn:=xlSortOnValues, Order:=xlAscending _
              , DataOption:=xlSortNormal
70        With IOD.ListObjects("IODTable").Sort
80            .Header = xlYes
90            .MatchCase = False
100           .Orientation = xlTopToBottom
110           .SortMethod = xlPinYin
120           .Apply
130       End With


140   TheColumn = 4


150   TheLastIODRow = IOD.UsedRange.SpecialCells(xlCellTypeLastCell).row
160   varDate = CDate(txtDate)
170   IOD.Activate

180   Set rng = IOD.Range("a2", "d" & TheLastIODRow)


190               With rng '
200                   .AutoFilter TheColumn, varDate
                      
210               End With
220      On Error Resume Next
230       Set rangeToCopy = IOD.Range("a2", "c" & TheLastIODRow).SpecialCells(xlCellTypeVisible)
          
240       If rangeToCopy Is Nothing Then
250           strComment = "No IOD Actions Today"
260       Else
270           strComment = "Append"
280           rangeToCopy.Copy
290      End If
         '  Set rangeToCopy = IOD.Range("a2", "c" & TheLastIODRow)
300      TheColumn = TheColumn / 0
         
         
          Dim wdApp As Word.Application
310       Set wdApp = New Word.Application
320       wdApp.Visible = True
          Dim wdDoc As Word.Document
          
          
330       Set wdDoc = wdApp.Documents.Open(FileName:=strFileName, AddToRecentFiles:=True, Visible:=False)
340       wdDoc.Activate
350       With wdDoc
                  
360               For Each CCtrl In .ContentControls
370                   Select Case CCtrl.Title
                          Case "Date"
380                           CCtrl.Range.Text = Format(varDate, "MMMM d, yyyy")
390                       Case "InvName"
400                           CCtrl.Range.Text = Files.Cells(20, 2).Value
410                       Case "InvPhone"
420                           CCtrl.Range.Text = Files.Cells(23, 2).Value
430                       Case "InvCell"
440                           CCtrl.Range.Text = Files.Cells(24, 2).Value
450                   End Select
460               Next
470       End With
480       strFileName = "IODLog_" & Format(varDate, "MM_dd_yy") 'set new filename
          
          
          'wdDoc.Application.Selection.Find.Execute "InsertHere"
          'wdApp.Selection.MoveRight Unit:=wdCharacter, Count:=1
          'wdApp.Selection.TypeParagraph 'duplicate?
          

                
490               wdDoc.Tables(1).Select
500               wdApp.Selection.EndOf Unit:=wdCell
510               If strComment = "Append" Then
520                   wdApp.Selection.PasteAppendTable
530               Else
540                   wdApp.Selection.MoveLeft Unit:=wdCell, Count:=2, Extend:=wdMove
550                   wdApp.Selection.Text = "No IOD Actions"
560                End If
570               wdApp.Selection.EndOf Unit:=wdCell
580               wdApp.Selection.InsertRowsBelow (1)
590               wdDoc.SaveAs FileName:=strPath & strFileName, FileFormat:=wdFormatDocumentDefault
600               wdDoc.PrintOut
610               wdDoc.Close
620               wdApp.Quit
630               Set wdApp = Nothing
640               Set wdDoc = Nothing
EndRoutine:
      'Optimize Code
      'Clear The Clipboard
650   Application.CutCopyMode = False

660   ActiveSheet.ListObjects("IODTable").Range.AutoFilter
          
670       Application.ScreenUpdating = True
680       Application.EnableEvents = True
690       ActiveWorkbook.Save
       
700    Unload frmCreateIODLog
          


710       On Error GoTo 0
720       Exit Sub

cmdCreateIODLog_Click_Error:

         
730       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

740       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdCreateIODLog_Click Within: frmCreateIODLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

750       Print #1, zMsg

760       Close #1

            
770       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub txtDate_AfterUpdate()

780   On Error GoTo txtDate_AfterUpdate_Error
      Dim zMsg As String

      Dim varDate As Date
790   If IsDate(txtDate) Then
800       varDate = DateValue(txtDate)
810       txtDate = Format(varDate, "MMMM d, yyyy")
820   Else
830       MsgBox "Invalid date"
840       txtDate.SetFocus
850       Exit Sub
          
860   End If

870       On Error GoTo 0
880   Exit Sub

txtDate_AfterUpdate_Error:

         
890       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

900       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_AfterUpdate Within: frmCreateIODLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

910       Print #1, zMsg

920       Close #1

            
930       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

940   On Error GoTo txtDate_DblClick_Error
      Dim zMsg As String
950   DatePickerForm.Caption = "IOD Date"
960   DatePickerForm.Show vbModal
970       [DatePickerForm]![CallingTextBox].Caption = "txtDate"
          '[DatePickerForm]![CallingForm].Caption = "frmCloseCase"
980       txtDate = [DatePickerForm]![CallingForm].Caption
990       If txtDate = "Form" Then txtDate = Format(DateValue(Now()), "MMMM d, yyyy")
1000      Cancel = True

1010      On Error GoTo 0
1020  Exit Sub

txtDate_DblClick_Error:

         
1030      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1040      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_DblClick Within: frmCreateIODLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1050      Print #1, zMsg

1060      Close #1

            
1070      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

1080  On Error GoTo txtDate_KeyPress_Error
      Dim zMsg As String

1090  If KeyAscii = 43 Then
1100      txtDate = DateAdd("d", 1, txtDate)
1110      KeyAscii = 0
          
1120  End If
1130  If KeyAscii = 45 Then
1140      txtDate = DateAdd("d", -1, txtDate)
1150      KeyAscii = 0
          
1160  End If

1170      On Error GoTo 0
1180  Exit Sub

txtDate_KeyPress_Error:

         
1190      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1200      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_KeyPress Within: frmCreateIODLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1210      Print #1, zMsg

1220      Close #1

            
1230      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub UserForm_Initialize()

1240  On Error GoTo UserForm_Initialize_Error
      Dim zMsg As String
1250  CenterForm Me

1260      txtDate = Format(Now(), "MMMM d, yyyy")

1270      On Error GoTo 0
1280  Exit Sub

UserForm_Initialize_Error:

         
1290      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1300      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserForm_Initialize Within: frmCreateIODLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1310      Print #1, zMsg

1320      Close #1

            
1330      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

