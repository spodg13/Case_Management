VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateMileageLog 
   Caption         =   "Create Mileage Log"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6645
   OleObjectBlob   =   "frmCreateMileageLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateMileageLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' File   : frmCreateMileageLog
' Author : gouldd
' Date   : 9/1/2016
' Purpose: Creates the Mileage Log
'---------------------------------------------------------------------------------------
Option Explicit


Private Sub cmdCancel_Click()
10        Unload frmCreateMileageLog
End Sub

Private Sub cmdCreateMileageLog_Click()

20    On Error GoTo cmdCreateMileageLog_Click_Error
      Dim zMsg As String

      Dim strFileName, strPath As String
      Dim TheLastMileageRow As Long
      Dim result As Integer
      Dim LastRange, rangeToCopy, MileageRange As Range
      Dim CCtrl As Word.ContentControl
      Dim firstDateOfMonth As Long, lastDateOfMonth As Long
      Dim boolTemp As Boolean

      Dim intCount, intTotal As Integer
      Dim wdApp As Word.Application
          
      Dim wdDoc As Word.Document
      Dim OutlookApp As Object
      Dim OutlookMessage As Object
      Dim strMailTo As String

30    Application.ScreenUpdating = True

      'Sort Table
40     MileageLog.ListObjects("MileageLog").Sort. _
              SortFields.Clear
50        MileageLog.ListObjects("MileageLog").Sort. _
              SortFields.Add Key:=Range("A2"), SortOn:=xlSortOnValues, Order:= _
              xlAscending, DataOption:=xlSortNormal
60        With MileageLog.ListObjects("MileageLog").Sort
70            .Header = xlYes
80            .MatchCase = False
90            .Orientation = xlTopToBottom
100           .SortMethod = xlPinYin
110           .Apply
120       End With

130   strPath = Files.Cells(14, 2).Value
140   strFileName = Files.Cells(15, 2).Value
150   strFileName = strFileName
160   strMailTo = Files.Cells(19, 2)


170   TheLastMileageRow = MileageLog.UsedRange.SpecialCells(xlCellTypeLastCell).row
180   MileageLog.Activate

          
          'First date of specified month in current year
          
190       firstDateOfMonth = DateSerial(Year(txtDate), Month(txtDate), 1)
          
          'Last date of specified month in current year
          
200       lastDateOfMonth = DateSerial(Year(txtDate), Month(txtDate) + 1, 0)
              
210       With MileageLog
220           .AutoFilterMode = False
230           .Range("A1").Select             'Select a cell within the data to be autofiltered
240       End With
250       Selection.AutoFilter Field:=1, Criteria1:=">=" & firstDateOfMonth, Operator:=xlAnd, Criteria2:="<=" & lastDateOfMonth
          
260       On Error Resume Next
270           Set LastRange = MileageLog.Range("d2", "f" & TheLastMileageRow).SpecialCells(xlCellTypeVisible)
          
280       If LastRange Is Nothing Then
290              intCount = 0
300              intTotal = 0
310              result = MsgBox("You are submitting a mileage log to " & strMailTo & " with no entries. Is this correct?", vbYesNo, "Confirm no mileage this month.")
320              If result = vbNo Then
330                   Selection.AutoFilter
340                   Exit Sub
350              End If
360       Else
              
370           boolTemp = MileageFinished(LastRange)
              
              
380           Set MileageRange = ActiveSheet.Range(Cells(2, 6), Cells(TheLastMileageRow, 6))
390           intTotal = SumVisible(MileageRange, "Mileage")
400           intCount = MileageLog.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
410           result = MsgBox("Sending mileage log to " & strMailTo & " with " & intCount & " trip(s), total miles: " & intTotal, vbYesNo, "Send mileage confirmation")
420           If result = vbNo Then
430                   Selection.AutoFilter
440                   Exit Sub
450              End If
                  
460       End If
          
470       Set rangeToCopy = MileageLog.Range("a2", "f" & TheLastMileageRow)
480       rangeToCopy.Copy
          
          
490       Set wdApp = New Word.Application
500       wdApp.Visible = True
510       Set wdDoc = wdApp.Documents.Open(FileName:=strFileName, AddToRecentFiles:=True, Visible:=True)
520       wdDoc.Activate
530       With wdDoc
                  
540               For Each CCtrl In .ContentControls
550                   Select Case CCtrl.Title
                          Case "txtDate"
560                           CCtrl.Range.Text = Format(txtDate, "MMMM, yyyy")
570                       Case "Odometer"
580                           CCtrl.Range.Text = txtOdometer
590                       Case "Total"
600                           CCtrl.Range.Text = CStr(intTotal)
610                       Case "Count"
620                           CCtrl.Range.Text = CStr(intCount)
630                       Case "InvName"
640                           CCtrl.Range.Text = Files.Cells(20, 2).Value
650                       Case "InvPhone"
660                           CCtrl.Range.Text = Files.Cells(23, 2).Value
670                       Case "InvCell"
680                           CCtrl.Range.Text = Files.Cells(24, 2).Value
690                       Case "InvLP"
700                           CCtrl.Range.Text = Files.Cells(21, 2).Value
710                       Case "InvVehID"
720                           CCtrl.Range.Text = Files.Cells(22, 2).Value
                              
                              
730                   End Select
740               Next
750       End With
760       strFileName = "MileageLog_" & Format(txtDate, "MM_yyyy") 'set new filename
          
               
                
770               wdDoc.Tables(2).Select
780               wdApp.Selection.StartOf Unit:=wdCell
790               wdApp.Selection.MoveRight Unit:=wdCell, Count:=10 'Move to line below header
800               If intCount > 0 Then
810                   wdApp.Selection.PasteAppendTable
820               End If
830               wdApp.Selection.EndOf Unit:=wdCell
840               wdApp.Selection.InsertRowsBelow (1)
850               wdDoc.SaveAs FileName:=strPath & strFileName, FileFormat:=wdFormatDocumentDefault
                  'wdDoc.PrintOut
860               wdDoc.Close
870               wdApp.Quit
880               Set wdApp = Nothing
890               Set wdDoc = Nothing
                  
                  
900     strFileName = strPath & strFileName & ".docx"
       
        
        
910     On Error Resume Next
          'Set OutlookApp = GetObject(class:="Outlook.Application") 'Handles if Outlook is already open
        'Err.Clear
          'If OutlookApp Is Nothing Then
920       Set OutlookApp = CreateObject(Class:="Outlook.Application") 'If not, open Outlook
          
930       If Err.Number = 429 Then
940         MsgBox "Outlook could not be found, aborting.", 16, "Outlook Not Found"
950         GoTo ExitSub
960       End If
970     On Error GoTo 0

      'Create a new email message
980     Set OutlookMessage = OutlookApp.CreateItem(0)

      'Create Outlook email with attachment
990     On Error Resume Next
1000      With OutlookMessage
1010       .To = strMailTo
1020       .CC = ""
1030       .BCC = ""
1040       .Subject = "Mileage Log"
1050       .Body = "Please find my attached mileage log." & vbNewLine & vbNewLine & Files.Cells(20, 2).Value
1060       .Attachments.Add strFileName
1070       .send  'Display to just open up email
1080      End With
1090    On Error GoTo 0


      'Clear Memory
1100    Set OutlookMessage = Nothing
1110    Set OutlookApp = Nothing
           
          
ExitSub:
1120      Application.CutCopyMode = False
1130      If ActiveSheet.FilterMode Then
1140      ActiveSheet.ShowAllData
1150      End If
1160      ActiveWorkbook.Save
1170      Unload frmCreateMileageLog
1180  Application.ScreenUpdating = True

1190      On Error GoTo 0
1200  Exit Sub

cmdCreateMileageLog_Click_Error:

         
1210      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1220      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdCreateMileageLog_Click Within: frmCreateMileageLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1230      Print #1, zMsg

1240      Close #1

            
1250      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDate_Change()

1260  On Error GoTo txtDate_Change_Error
      Dim zMsg As String

1270      txtDate = Format(txtDate, "MMMM, yyyy")

1280      On Error GoTo 0
1290  Exit Sub

txtDate_Change_Error:

         
1300      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1310      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_Change Within: frmCreateMileageLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1320      Print #1, zMsg

1330      Close #1

            
1340      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

1350  On Error GoTo txtDate_DblClick_Error
      Dim zMsg As String
1360  DatePickerForm.Caption = "Mileage Month"
1370  DatePickerForm.Show vbModal
1380      [DatePickerForm]![CallingTextBox].Caption = "txtDate"
          '[DatePickerForm]![CallingForm].Caption = "frmCloseCase"
1390      txtDate = [DatePickerForm]![CallingForm].Caption
1400      If txtDate = "Form" Then txtDate = Format(DateValue(Now()), "MMMM d, yyyy")
1410      Cancel = True

1420      On Error GoTo 0
1430  Exit Sub

txtDate_DblClick_Error:

         
1440      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1450      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_DblClick Within: frmCreateMileageLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1460      Print #1, zMsg

1470      Close #1

            
1480      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

1490  On Error GoTo txtDate_KeyPress_Error
      Dim zMsg As String

1500  If KeyAscii = 43 Then
1510      txtDate = DateAdd("m", 1, txtDate)
1520      KeyAscii = 0
          
1530  End If
1540  If KeyAscii = 45 Then
1550      txtDate = DateAdd("m", -1, txtDate)
1560      KeyAscii = 0
          
1570  End If

1580      On Error GoTo 0
1590  Exit Sub

txtDate_KeyPress_Error:

         
1600      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1610      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_KeyPress Within: frmCreateMileageLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1620      Print #1, zMsg

1630      Close #1

            
1640      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub UserForm_Initialize()

1650  On Error GoTo UserForm_Initialize_Error
      Dim zMsg As String
1660  CenterForm Me

1670  If Day(Now()) > 15 Then
1680             txtDate = Format(Now(), "MMMM, yyyy")
1690         Else
1700          txtDate = DateAdd("m", -1, Now())
1710          txtDate = Format(txtDate, "MMMM, yyyy")
1720      End If

1730      On Error GoTo 0
1740  Exit Sub

UserForm_Initialize_Error:

         
1750      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1760      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserForm_Initialize Within: frmCreateMileageLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1770      Print #1, zMsg

1780      Close #1

            
1790      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

