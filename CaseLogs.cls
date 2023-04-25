VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdIODEntry_Click()

10    On Error GoTo cmdIODEntry_Click_Error
      Dim zMsg As String

20        frmIODEntry.Show

30        On Error GoTo 0
40    Exit Sub

cmdIODEntry_Click_Error:

         
50        Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

60        zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdIODEntry_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

70        Print #1, zMsg

80        Close #1

            
90        MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Private Sub cmdEnterActionCaseSheet_Click()

100   On Error GoTo cmdEnterActionCaseSheet_Click_Error
      Dim zMsg As String
110   ResetButton cmdEnterActionCaseSheet
      Dim lngRow As Long
120   lngRow = ActiveCell.row
130   If CaseLogs.Cells(lngRow, 1).Value = "" Then
140           MsgBox ("Pick a case before you add an action")
150           Exit Sub
160       End If
170       If CaseLogs.Cells(lngRow, 1).Value = "IOD" Then
180           frmIODEntry.Show
190       Else
200           frmEnterAction.Show
210       End If

220       On Error GoTo 0
230   Exit Sub

cmdEnterActionCaseSheet_Click_Error:

         
240       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

250       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdEnterActionCaseSheet_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

260       Print #1, zMsg

270       Close #1

            
280       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Private Sub cmdMileageInc_Click()

290   On Error GoTo cmdMileageInc_Click_Error
      Dim zMsg As String
300   ResetButton cmdMileageInc
      Dim lngRow As Long
310   Application.ScreenUpdating = False
          
320       lngRow = ActiveCell.row
330       Cells(lngRow, 6).Activate
          
340      frmMileageAction.Caption = "Create Mileage Action - Client: " & ActiveCell.Value
350      frmMileageAction.Show
         
         

360       On Error GoTo 0
370   Exit Sub

cmdMileageInc_Click_Error:

         
380       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

390       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdMileageInc_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

400       Print #1, zMsg

410       Close #1

            
420       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub cmdMinus_Click()

430   On Error GoTo cmdMinus_Click_Error
      Dim zMsg As String
440   ResetButton cmdMinus
      Dim LMin, LMinNew, IntervalAdd As Integer
450   If ActiveCell.Column = 2 Then
460       ActiveCell.Value = ActiveCell.Value - 1
470   End If
480   If ActiveCell.Column = 3 Then
490       LMin = Minute(ActiveCell.Value)
500       LMinNew = Round(LMin / 5, 0) * 5
510       IntervalAdd = (LMin - LMinNew) * -1
520       ActiveCell.Value = DateAdd("N", IntervalAdd, ActiveCell.Value)
          
530       ActiveCell.Value = DateAdd("N", -5, ActiveCell.Value)
          
540       ActiveCell.Value = Format(ActiveCell.Value, "h:mm AM/PM")
550   End If

560       On Error GoTo 0
570   Exit Sub

cmdMinus_Click_Error:

         
580       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

590       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdMinus_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

600       Print #1, zMsg

610       Close #1

            
620       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdMinus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

630   On Error GoTo cmdMinus_DblClick_Error
      Dim zMsg As String
640   ResetButton cmdMinus
      Dim LMin, LMinNew, IntervalAdd As Integer
650   If ActiveCell.Column = 2 Then
660       ActiveCell.Value = ActiveCell.Value - 2
670   End If
680   If ActiveCell.Column = 3 Then
690       LMin = Minute(ActiveCell.Value)
700       LMinNew = Round(LMin / 5, 0) * 5
710       IntervalAdd = (LMin - LMinNew) * -1
720       ActiveCell.Value = DateAdd("N", IntervalAdd, ActiveCell.Value)
          
730       ActiveCell.Value = DateAdd("N", -10, ActiveCell.Value)
          
740       ActiveCell.Value = Format(ActiveCell.Value, "h:mm AM/PM")
750   End If

760       On Error GoTo 0
770   Exit Sub

cmdMinus_DblClick_Error:

         
780       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

790       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdMinus_DblClick Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

800       Print #1, zMsg

810       Close #1

            
820       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub cmdPlus_Click()

830   On Error GoTo cmdPlus_Click_Error
      Dim zMsg As String
840   ResetButton cmdPlus
      Dim LMin, LMinNew, IntervalAdd As Integer

850   If ActiveCell.Column = 2 Then
860       ActiveCell.Value = ActiveCell.Value + 1
870   End If
880   If ActiveCell.Column = 3 Then
890       LMin = Minute(ActiveCell.Value)
900       LMinNew = Round(LMin / 5, 0) * 5
910       IntervalAdd = (LMin - LMinNew) * -1
920       ActiveCell.Value = DateAdd("N", IntervalAdd, ActiveCell.Value)
          
930       ActiveCell.Value = DateAdd("N", 5, ActiveCell.Value)
          
940       ActiveCell.Value = Format(ActiveCell.Value, "h:mm AM/PM")
950   End If


960       On Error GoTo 0
970   Exit Sub

cmdPlus_Click_Error:

         
980       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

990       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdPlus_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1000      Print #1, zMsg

1010      Close #1

            
1020      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdPlus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

1030  On Error GoTo cmdPlus_DblClick_Error
      Dim zMsg As String
1040  ResetButton cmdPlus
      Dim LMin, LMinNew, IntervalAdd As Integer

1050  If ActiveCell.Column = 2 Then
1060      ActiveCell.Value = ActiveCell.Value + 2
1070  End If
1080  If ActiveCell.Column = 3 Then
1090      LMin = Minute(ActiveCell.Value)
1100      LMinNew = Round(LMin / 5, 0) * 5
1110      IntervalAdd = (LMin - LMinNew) * -1
1120      ActiveCell.Value = DateAdd("N", IntervalAdd, ActiveCell.Value)
          
1130      ActiveCell.Value = DateAdd("N", 10, ActiveCell.Value)
          
1140      ActiveCell.Value = Format(ActiveCell.Value, "h:mm AM/PM")
1150  End If

1160      On Error GoTo 0
1170  Exit Sub

cmdPlus_DblClick_Error:

         
1180      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1190      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdPlus_DblClick Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1200      Print #1, zMsg

1210      Close #1

            
1220      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub




Private Sub cmdPrintActionDraft_Click()


1230  On Error GoTo cmdPrintActionDraft_Click_Error
      Dim zMsg As String
1240  ResetButton cmdPrintActionDraft
      Dim TheColumn As Integer
      Dim TheLastRow As Long
      Dim rng As Range
      Dim strDocket As String

1250  strDocket = Cells(ActiveCell.row, 1)
1260  Application.ScreenUpdating = False
1270  Application.EnableEvents = False

1280  TheColumn = 1
1290  TheLastRow = CaseLogs.UsedRange.SpecialCells(xlCellTypeLastCell).row

1300  Set rng = CaseLogs.Range("a4", "f" & TheLastRow)


1310              With rng
1320                  .AutoFilter TheColumn, strDocket
                      
1330              End With

1340  PrintActionDraft ActionRange(2, 4), strDocket, Cells(ActiveCell.row, 6)
1350  ClearCaseLogFilter
1360  Application.ScreenUpdating = True
1370  Application.EnableEvents = True


1380      On Error GoTo 0
1390  Exit Sub

cmdPrintActionDraft_Click_Error:

         
1400      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1410      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdPrintActionDraft_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1420      Print #1, zMsg

1430      Close #1

            
1440      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub





Private Sub CommandButton1_Click()

End Sub

Private Sub Worksheet_Activate()


1450  On Error GoTo Worksheet_Activate_Error
      Dim zMsg As String

1460     RemoveColumnDelete


1470      On Error GoTo 0
1480  Exit Sub

Worksheet_Activate_Error:

         
1490      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1500      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Worksheet_Activate Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1510      Print #1, zMsg

1520      Close #1

            
1530      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

1540  On Error GoTo Worksheet_Change_Error
      Dim zMsg As String

      'ActiveSheet.Unprotect "darryl"
1550  If Target.Column = 4 Then Target.Locked = True
1560    If Target.Address = Target.EntireColumn.Address Then
         
         
1570    With Application
1580        .EnableEvents = False
1590        .Undo
1600        MsgBox "No deleting rows or columns", 16
1610        .EnableEvents = True
1620    End With
         
1630      Else
         
1640    Exit Sub
         
1650      End If
1660  ActiveSheet.Range("A:C").Locked = True
      'ActiveSheet.Range("E:J").Locked = False
      'ActiveSheet.Protect "darryl"

1670      On Error GoTo 0
1680  Exit Sub

Worksheet_Change_Error:

         
1690      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1700      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Worksheet_Change Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1710      Print #1, zMsg

1720      Close #1

            
1730      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Private Sub cboClientSearch_Click()


1740  On Error GoTo cboClientSearch_Click_Error
      Dim zMsg As String
1750  ResetButton cboClientSearch

      Dim rng As Range
      Dim LastRow, firstrow As Long
      Dim strFoundDocket As String

1760  ActiveSheet.AutoFilterMode = False
1770  Application.ScreenUpdating = True

      '****** Clear all filters
1780  ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=6
1790  ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=1
1800  If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
1810    ActiveSheet.ShowAllData
1820  End If
      '*******

1830  LastRow = ActiveSheet.UsedRange.Rows.Count + 1 'Account for headers, add 3

1840  Set rng = Application.Intersect(ActiveSheet.UsedRange, Range("f4", "f" & LastRow))

1850  With rng 'CaseLogs.UsedRange

1860      .AutoFilter 6, cboClientSearch.Value
          
1870  End With
1880  On Error Resume Next

1890  If Err.Number = 1004 Then
               ' No cells returned by filter
1900    txtDueDate.Visible = False
1910    Exit Sub
1920  End If

1930  Set rng = rng.SpecialCells(xlCellTypeVisible)
1940  firstrow = rng.Cells(1, 1).row

      'firstrow = rsltRng.Cells(1, 1).Row
1950  strFoundDocket = Cells(firstrow, 1)


1960  Range("a65336").End(xlUp).Select
1970  ActiveWindow.ScrollRow = 1

1980  ActiveCell.Offset(1, 0).Select


1990  CaseLogs.AutoFilterMode = False
      'Selection.AutoFilter Field:=6, Criteria1:=cboClientSearch.Value, Operator:=xlAnd
2000  cmdFilterCase.Caption = "Show All"
2010  txtDueDate.Visible = True
2020  txtDueDate = "Due: " & GetDueDate(strFoundDocket)
2030  Application.EnableEvents = True
2040  Exit Sub
errHandler:
2050  If Err.Number = 1004 Then
               ' No cells returned by filter
2060    txtDueDate.Visible = False
2070    Application.EnableEvents = True
2080    Application.ScreenUpdating = True
2090    Exit Sub
2100  End If
2110    Application.EnableEvents = True


2120      On Error GoTo 0
2130  Exit Sub

cboClientSearch_Click_Error:

         
2140      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2150      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cboClientSearch_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2160      Print #1, zMsg

2170      Close #1

            
2180      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub cboClientSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

2190  On Error GoTo cboClientSearch_KeyDown_Error
      Dim zMsg As String

      Dim rng As Range
2200  ResetButton cboClientSearch
2210  Application.ScreenUpdating = False

      '****** Clear all filters
2220  ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=6
2230  ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=1
2240  If ActiveSheet.FilterMode Then
2250      ActiveSheet.ShowAllData
2260    End If
      '*******

      Dim LastRow As Long
2270   Select Case KeyCode
              Case 9 'Tab
2280              LastRow = ActiveSheet.UsedRange.Rows.Count + 10 'Account for headers, add 3

2290              Set rng = Application.Intersect(ActiveSheet.UsedRange, Range("f4", "f" & LastRow))
                 
2300              With rng 'CaseLogs.UsedRange
2310                  .AutoFilter 6, cboClientSearch.Value
          
2320              End With

2330              Range("a1").End(xlDown).Select
2340              Application.GoTo ActiveCell
2350              ActiveCell.Offset(1, 0).Select

2360              CaseLogs.AutoFilterMode = False
                  'Selection.AutoFilter Field:=6, Criteria1:=cboClientSearch.Value, Operator:=xlAnd
2370              cmdFilterCase.Caption = "Show All"
2380          Case 13 'Enter
2390             LastRow = ActiveSheet.UsedRange.Rows.Count + 10 'Account for headers, add 3

2400              Set rng = Application.Intersect(ActiveSheet.UsedRange, Range("f4", "f" & LastRow))
                 
2410              With rng 'CaseLogs.UsedRange
2420                  .AutoFilter 6, cboClientSearch.Value
2430              End With
                  
2440             Range("a65536").End(xlUp).Select
2450             ActiveWindow.ScrollRow = 1
2460              ActiveCell.Offset(1, 0).Select

2470              CaseLogs.AutoFilterMode = False
                  'Selection.AutoFilter Field:=6, Criteria1:=cboClientSearch.Value, Operator:=xlAnd
2480              cmdFilterCase.Caption = "Show All"
                  
2490          Case Else
                  'do nothing
2500      End Select
2510  Application.ScreenUpdating = True

2520      On Error GoTo 0
2530  Exit Sub

cboClientSearch_KeyDown_Error:

         
2540      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2550      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cboClientSearch_KeyDown Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2560      Print #1, zMsg

2570      Close #1

            
2580      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub cmdFilterCase_Click()

2590  On Error GoTo cmdFilterCase_Click_Error
      Dim zMsg As String
2600  ResetButton cmdFilterCase
      Dim col As Integer, iCount As Long
      Dim lngRow, TheLastRow, TempRow, rc As Long
      Dim TheColumn As Integer
      Dim strFoundDocket As String
      Dim intR As Long
      Dim rng As Range
2610  ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=6
2620   Application.Calculation = xlCalculationManual
2630   TheColumn = ActiveCell.Column
2640   TempRow = ActiveCell.row
2650   strFoundDocket = Cells(TempRow, 1)
2660   TheLastRow = CaseLogs.UsedRange.SpecialCells(xlCellTypeLastCell).row
         
2670     For rc = TheLastRow To 1 Step -1    ' Step down to 1 to insure all deletions * For speed, down to last 10
2680        If UCase(Cells(rc, 6).Value) = "-" Then Rows(rc).Delete 'Column 6 ="-"
            'If UCase(Cells(rc, 1).Value) = "" Then Rows(rc).Delete  - need to insure rc greater than first two rows
2690     Next rc
2700  Application.Calculation = xlCalculationAutomatic
2710  Application.ScreenUpdating = False


       
2720  If cmdFilterCase.Caption = "Show All" Then GoTo unhide

      '****** Desired code for speed - need to correct  - adjust Range below to apply to any choice
         
2730  If TheColumn > 1 And TheColumn < 4 Then GoTo DateCriteria
          
2740  Set rng = Application.Intersect(ActiveSheet.UsedRange, Range("a4", "f" & TheLastRow))
                 
2750              With rng 'CaseLogs.UsedRange
2760                  .AutoFilter TheColumn, ActiveCell.Value
2770              End With
2780              Range("a1").End(xlDown).Select
2790              Application.GoTo ActiveCell
2800              ActiveCell.Offset(1, 0).Select

2810              CaseLogs.AutoFilterMode = False
2820              cmdFilterCase.Caption = "Show All"
2830              txtDueDate.Visible = True
2840              txtDueDate = "Due: " & GetDueDate(strFoundDocket)
2850  Exit Sub

      '******OLD slow method, but allows more robust selection - Old method needed for date and time
DateCriteria:

2860  With ActiveCell.CurrentRegion
2870      lngRow = .Rows.Count
2880      col = ActiveCell.Column
2890      For intR = 2 To lngRow
2900          If .Cells(intR, col) <> ActiveCell Then
2910          .Cells(intR, col).EntireRow.Hidden = True
2920          Else
2930          iCount = iCount + 1
2940          End If
2950      Next intR
2960  End With
2970  cmdFilterCase.Caption = "Show All"
2980  txtDueDate.Visible = True
2990  txtDueDate = "Due: " & GetDueDate(strFoundDocket)
3000  CaseLogs.AutoFilterMode = False
3010  Exit Sub

unhide:

3020  Reset_Used_Range

3030  ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=4


3040  TheLastRow = CaseLogs.UsedRange.SpecialCells(xlCellTypeLastCell).row
3050  Cells(TheLastRow, "A").Select
3060  CaseLogs.AutoFilterMode = False
      '*****OLD method of removing hidden rows
      'With ActiveCell.CurrentRegion
      '.EntireRow.Hidden = False
      'End With
3070  If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
3080    ActiveSheet.ShowAllData
3090  End If
3100  Range("a65536").End(xlUp).Select
3110  ActiveCell.Offset(1, 0).Select

3120  cmdFilterCase.Caption = "Filter Case"
3130  txtDueDate.Visible = False
3140  Application.ScreenUpdating = True


3150      On Error GoTo 0
3160  Exit Sub

cmdFilterCase_Click_Error:

         
3170      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3180      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdFilterCase_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3190      Print #1, zMsg

3200      Close #1

            
3210      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdNow_Click()


3220  On Error GoTo cmdNow_Click_Error
      Dim zMsg As String
3230  ResetButton cmdNow
3240  If ActiveCell.Column = 2 Then
3250      ActiveCell.Offset(0, 1).Activate
3260  End If
3270      ActiveCell.Value = Format(Now(), "h:mm AMPM")
3280      ActiveCell.Offset(0, -1).Value = Format(Now(), "mm/dd/yy")
3290      ActiveCell.Offset(0, 1).Activate
3300      cmdNow.Visible = False
3310      cmdToday.Visible = False


3320      On Error GoTo 0
3330  Exit Sub

cmdNow_Click_Error:

         
3340      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3350      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdNow_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3360      Print #1, zMsg

3370      Close #1

            
3380      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdSortCaseLogs_Click()
          

3390  On Error GoTo cmdSortCaseLogs_Click_Error
      Dim zMsg As String
3400  ResetButton cmdSortCaseLogs
3410      ClearCaseLogFilter
3420      SortCaseLogs
3430      cmdFilterCase.Caption = "Filter Case"
3440      txtDueDate.Visible = False


3450      On Error GoTo 0
3460  Exit Sub

cmdSortCaseLogs_Click_Error:

         
3470      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3480      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdSortCaseLogs_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3490      Print #1, zMsg

3500      Close #1

            
3510      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub




Private Sub cmdToday_Click()


3520  On Error GoTo cmdToday_Click_Error
      Dim zMsg As String
3530  ResetButton cmdToday
3540  If ActiveCell.Column = 3 Then
3550      ActiveCell.Offset(0, -1).Activate
3560  End If

3570  ActiveCell.Value = Format(Now(), "mm/dd/yy")
3580  cmdToday.Visible = False
3590  cmdNow.Visible = False


3600      On Error GoTo 0
3610  Exit Sub

cmdToday_Click_Error:

         
3620      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3630      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdToday_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3640      Print #1, zMsg

3650      Close #1

            
3660      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdWitness_Click()

3670  On Error GoTo cmdWitness_Click_Error
      Dim zMsg As String
3680  ResetButton cmdWitness
      Dim strWitness As String
      Dim rng As Range
      Dim LastRow As Long

3690  strWitness = InputBox("Please enter the first or last name of the witness.", "Witness name")
3700  strWitness = "*" & strWitness & "*"
3710  LastRow = ActiveSheet.UsedRange.Rows.Count + 10 'Account for headers, add 3

3720  Set rng = Application.Intersect(ActiveSheet.UsedRange, Range("d4", "d" & LastRow))
3730  With rng 'CaseLogs.UsedRange
3740      .AutoFilter 4, strWitness
          
3750  End With
3760  AutoFilterMode = False
3770  If cmdFilterCase.Caption = "Filter Case" Then
3780      cmdFilterCase.Caption = "Show All"
3790  End If

       
3800   Range("a65536").End(xlUp).Select
3810   ActiveWindow.ScrollRow = 1
3820   ActiveCell.Offset(1, 0).Select


3830      On Error GoTo 0
3840  Exit Sub

cmdWitness_Click_Error:

         
3850      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3860      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdWitness_Click Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3870      Print #1, zMsg

3880      Close #1

            
3890      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, _
  Cancel As Boolean)


3900  On Error GoTo Worksheet_BeforeDoubleClick_Error
      Dim zMsg As String

      Dim result As Integer
      Dim cmdTemp1, cmdTemp2 As OLEObject
      Dim ws As Worksheet
3910  Set ws = ActiveSheet


3920  Set cmdTemp1 = ws.OLEObjects("cmdNow")
3930  Set cmdTemp2 = ws.OLEObjects("cmdToday")

3940    On Error Resume Next
3950    With cmdTemp1
3960      .Visible = False
3970    End With
3980    With cmdTemp2
3990      .Visible = False
4000    End With
       

cmdCheck:
4010  On Error GoTo errHandler

4020    If Target.Column = 2 Or Target.Column = 3 Then
          
4030          With cmdTemp2
4040              .Visible = True
4050              .Left = Target.Left + Target.Width
4060              .Top = Target.Top + 15
4070              .Height = Target.Height + 5
4080              .Width = Target.Width
4090          End With
4100          With cmdTemp1
4110           .Visible = True
4120              .Left = Target.Left
4130              .Top = Target.Top + 15
4140              .Height = Target.Height + 5
4150              .Width = Target.Width
4160          End With
4170          cmdTemp2.Activate
4180          cmdTemp1.Activate
          
            
4190    End If
        
4200    If Target.Column = 4 Then
4210      result = MsgBox("Are you sure you want to edit this entry?", vbOKCancel, "Confirm Edit Action")
4220      If result = vbOK Then
4230          ActiveCell.Activate
4240      Else
4250          Target.Locked = True
4260          Exit Sub
4270      End If
4280    End If

errHandler:
4290    Application.EnableEvents = True
4300    Exit Sub


4310      On Error GoTo 0
4320  Exit Sub

Worksheet_BeforeDoubleClick_Error:

         
4330      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

4340      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Worksheet_BeforeDoubleClick Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

4350      Print #1, zMsg

4360      Close #1

            
4370      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub Worksheet_Deactivate()

4380  On Error GoTo Worksheet_Deactivate_Error
      Dim zMsg As String

4390      ResetMenu

4400      On Error GoTo 0
4410  Exit Sub

Worksheet_Deactivate_Error:

         
4420      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

4430      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Worksheet_Deactivate Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

4440      Print #1, zMsg

4450      Close #1

            
4460      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

'=========================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

4470  On Error GoTo Worksheet_SelectionChange_Error
      Dim zMsg As String

      Dim cmdTemp1, cmdTemp2 As OLEObject
      Dim ws As Worksheet
4480  Set ws = ActiveSheet
4490  Application.EnableEvents = False
4500  Application.ScreenUpdating = True

4510  If Application.CutCopyMode Then
       'allow copying and pasting on the worksheet
4520    GoTo errHandler
4530  End If

4540  Set cmdTemp1 = ws.OLEObjects("cmdNow")
4550  Set cmdTemp2 = ws.OLEObjects("cmdToday")
4560    On Error Resume Next

4570    With cmdTemp1
4580          .Visible = False
4590    End With
4600    With cmdTemp2
4610          .Visible = False
4620    End With
       

errHandler:
4630    Application.EnableEvents = True
4640    Exit Sub


4650      On Error GoTo 0
4660  Exit Sub

Worksheet_SelectionChange_Error:

         
4670      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

4680      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Worksheet_SelectionChange Within: CaseLogs" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

4690      Print #1, zMsg

4700      Close #1

            
4710      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


