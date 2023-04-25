Attribute VB_Name = "Module1"
Option Explicit


Function FilePicked(strAnytype As String) As String


10    On Error GoTo FilePicked_Error
      Dim zMsg As String


      Dim fd As Office.FileDialog

20        Set fd = Application.FileDialog(msoFileDialogFilePicker)

30       With fd

40          .AllowMultiSelect = False
50          .InitialFileName = Files.Cells(36, 2).Value & "\Templates"

            ' Set the title of the dialog box.
60          .Title = "Please select the " & strAnytype & "."

            ' Clear out the current filters, and add our own.
70          .Filters.Clear
80          .Filters.Add "Word Templates", "*.dotx"
90          .Filters.Add "Word", "*.doc*"
100         .Filters.Add "Excel", "*.xls*"
110         .Filters.Add "All Files", "*.*"

            ' Show the dialog box. If the .Show method returns True, the
            ' user picked at least one file. If the .Show method returns
            ' False, the user clicked Cancel.
120         If .Show = True Then
130           FilePicked = .SelectedItems(1) 'replace txtFileName with your textbox

140         End If
150      End With


160       On Error GoTo 0
170   Exit Function

FilePicked_Error:

         
180       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

190       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: FilePicked Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

200       Print #1, zMsg

210       Close #1

            
220       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function
Function PathPicked(strAnyName As String) As String


230   On Error GoTo PathPicked_Error
      Dim zMsg As String


      Dim fd As Office.FileDialog

240       Set fd = Application.FileDialog(msoFileDialogFolderPicker)

250      With fd

260         .AllowMultiSelect = False
270         .InitialFileName = Files.Cells(36, 2).Value

            ' Set the title of the dialog box.
280         .Title = "Please select the folder where you will store your " & strAnyName & "."

            ' Clear out the current filters, and add our own.
290         .Filters.Clear
            

            ' Show the dialog box. If the .Show method returns True, the
            ' user picked at least one file. If the .Show method returns
            ' False, the user clicked Cancel.
300         If .Show = True Then
310           PathPicked = .SelectedItems(1) & "\" 'replace txtFileName with your textbox

320         End If
330      End With


340       On Error GoTo 0
350   Exit Function

PathPicked_Error:

         
360       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

370       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: PathPicked Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

380       Print #1, zMsg

390       Close #1

            
400       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function
'---------------------------------------------------------------------------------------
' Method : GetADate
' Author : gouldd
' Date   : 9/1/2016
' Purpose: Verifies a valid date and formats it correctly
'---------------------------------------------------------------------------------------
Function GetADate(AnyDate As String) As Date


410   On Error GoTo GetADate_Error
      Dim zMsg As String

420   If IsDate(AnyDate) Then
430       GetADate = Format(DateValue(AnyDate), "m/d/yy")
          
440   Else
450       MsgBox "Invalid date"
460   End If

470       On Error GoTo 0
480   Exit Function

GetADate_Error:

         
490       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

500       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: GetADate Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

510       Print #1, zMsg

520       Close #1

            
530       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function
Function IsaDate(anyString As String) As Boolean


540   On Error GoTo IsaDate_Error
      Dim zMsg As String

550       IsaDate = IsDate(anyString)


560       On Error GoTo 0
570   Exit Function

IsaDate_Error:

         
580       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

590       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: IsaDate Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

600       Print #1, zMsg

610       Close #1

            
620       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function
Function IsTime(anyTime As String) As Boolean
          
630       On Error Resume Next
640       IsTime = IsDate(TimeValue(anyTime))
650       On Error GoTo 0
End Function


Sub Copy_Case_Above()
Attribute Copy_Case_Above.VB_ProcData.VB_Invoke_Func = "c\n14"


660   On Error GoTo Copy_Case_Above_Error
      Dim zMsg As String

      Dim TheLastRow, rc, lngRow As Long
      Dim TheColumn, TempRow As Variant
      Dim rng As Range

       'Application.ScreenUpdating = True
670    Application.EnableEvents = False
680    Application.Calculation = xlCalculationManual
690    TheColumn = ActiveCell.Column
700    TempRow = ActiveCell.row
710    TheLastRow = CaseLogs.UsedRange.SpecialCells(xlCellTypeLastCell).row
         
720      For rc = TheLastRow To (TheLastRow - 5) Step -1    ' Step down to 1 to insure all deletions * For speed, down to last 10
730         If UCase(Cells(rc, 6).Value) = "-" Then Rows(rc).Delete  'Column 6 ="-"
740      Next rc
750   Application.Calculation = xlCalculationAutomatic


760   Cells(TempRow, TheColumn).Select
      'Do While IsEmpty(ActiveCell.Value)
      '  ActiveCell.Offset(-1, 0).Select
        
      'Loop
770   If IsEmpty(ActiveCell.Value) = True Then

780   Set rng = ActiveCell
790   For lngRow = 1 To TheLastRow    'Rows.Count
800       Set rng = rng.Offset(-1, 0)
810       If rng.EntireRow.Hidden = False Then
820           rng.Select
830           GoTo sel
840       End If
850   Next
860   Else
870       Exit Sub   'Not allow copying of a cell in the middle of the case logs
880   End If
      ' Copy_Case_Above Macro
      ' Copy the cell above
      '
      ' Keyboard Shortcut: Ctrl+Shift+C
      '

sel:
890   Selection.Copy
          
900   With ActiveCell
910           With .Offset(1, 0).Resize(Rows.Count - .row, 1)
920               .SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
930    End With
940   End With
          
950       ActiveSheet.Paste
960       Application.EnableEvents = True
              

970       On Error GoTo 0
980       Exit Sub

Copy_Case_Above_Error:

         
990       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1000      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Copy_Case_Above Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1010      Print #1, zMsg

1020      Close #1

            
1030      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
End Sub

Sub Reset_Used_Range()


1040  On Error GoTo Reset_Used_Range_Error
      Dim zMsg As String

      Dim intCount As Long
1050  intCount = ActiveSheet.UsedRange.Rows.Count

1060      On Error GoTo 0
1070  Exit Sub

Reset_Used_Range_Error:

         
1080      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1090      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Reset_Used_Range Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1100      Print #1, zMsg

1110      Close #1

            
1120      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Sub FindLastRow()

1130  On Error GoTo FindLastRow_Error
      Dim zMsg As String

      Dim TheLastRow As Long
1140  TheLastRow = CaseLogs.UsedRange.SpecialCells(xlCellTypeLastCell).row
1150  MsgBox ("Last  " & TheLastRow)


1160      On Error GoTo 0
1170  Exit Sub

FindLastRow_Error:

         
1180      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1190      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: FindLastRow Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1200      Print #1, zMsg

1210      Close #1

            
1220      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Sub PopulateCombo()


1230  On Error GoTo PopulateCombo_Error
      Dim zMsg As String

      Dim LastRow As Long
      Dim ListArray As Variant
      Dim Temp As String
      Dim lngRow As Long, intCol As Integer

1240  If Files.Cells(20, 2).Value = "New User" Then Exit Sub


1250  LastRow = InvestigationLog.Range("A" & Rows.Count).End(xlUp).row  'UsedRange.Rows.Count + 2
1260  If LastRow = 3 Then Exit Sub


1270  CaseLogs.cboClientSearch.List = InvestigationLog.Range("c3", "c" & LastRow).Value

1280  With CaseLogs.cboClientSearch
1290  ListArray = Application.Transpose(.List)
1300      For lngRow = 1 To UBound(ListArray) - 1
1310          For intCol = lngRow To UBound(ListArray)
1320              If ListArray(intCol) < ListArray(lngRow) Then
1330                  Temp = ListArray(lngRow)
1340                  ListArray(lngRow) = ListArray(intCol)
1350                  ListArray(intCol) = Temp
1360              End If
1370          Next intCol
1380      Next lngRow
1390  .List = ListArray
1400  End With

1410  InvestigationLog.cboClientLocate.List = InvestigationLog.Range("c3", "c" & LastRow).Value

1420  With InvestigationLog.cboClientLocate
1430  ListArray = Application.Transpose(.List)
1440      For lngRow = 1 To UBound(ListArray) - 1
1450          For intCol = lngRow To UBound(ListArray)
1460              If ListArray(intCol) < ListArray(lngRow) Then
1470                  Temp = ListArray(lngRow)
1480                  ListArray(lngRow) = ListArray(intCol)
1490                  ListArray(intCol) = Temp
1500              End If
1510          Next intCol
1520      Next lngRow
1530  .List = ListArray
1540  End With
1550       If ActiveWorkbook.ReadOnly = False Then
1560        ActiveWorkbook.Save
1570       End If


1580      On Error GoTo 0
1590      Exit Sub

PopulateCombo_Error:

             
1600      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1610      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: PopulateCombo Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1620      Print #1, zMsg

1630      Close #1

1640      If Err.Number = 1004 Then
1650          Exit Sub
1660      Else
1670          MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
1680      End If
        
  End Sub

'---------------------------------------------------------------------------------------
' Method : ClearCaseLogFilter
' Author : gouldd
' Date   : 9/1/2016
' Purpose: Clears filter from CaseLogs sheet
'---------------------------------------------------------------------------------------
Public Sub ClearCaseLogFilter()


1690  On Error GoTo ClearCaseLogFilter_Error
      Dim zMsg As String

      Dim ws As Worksheet
1700  Set ws = CaseLogs

1710  With ws
1720      .AutoFilterMode = False
1730      .ListObjects("Table2").Range.AutoFilter Field:=6
1740      .ListObjects("Table2").Range.AutoFilter Field:=4
1750      .ListObjects("Table2").Range.AutoFilter Field:=1
1760  End With
1770  ws.OLEObjects("cmdFilterCase").Object.Caption = "Filter Case"


1780      On Error GoTo 0
1790      Exit Sub

ClearCaseLogFilter_Error:

         
1800      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1810      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: ClearCaseLogFilter Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1820      Print #1, zMsg

1830      Close #1

            
1840      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


'---------------------------------------------------------------------------------------
' Method : SortCaseLogs
' Author : gouldd
' Date   : 9/1/2016
' Purpose: Sorts the CaseLogs Sheet by date and time
'---------------------------------------------------------------------------------------
Public Sub SortCaseLogs()


1850  On Error GoTo SortCaseLogs_Error
      Dim zMsg As String

      Dim LastRow, rc As Long
      Dim rngSortRange As Range
1860  Application.EnableEvents = False
1870  Application.ScreenUpdating = False
      Dim ws As Worksheet

1880  Set ws = CaseLogs

1890  LastRow = ws.UsedRange.Rows.Count
1900  Set rngSortRange = ws.Range("a3", "a" & LastRow)
1910      ws.ListObjects("Table2").Sort.SortFields.Clear
1920      ws.ListObjects("Table2").Sort.SortFields.Add _
              Key:=Range("Table2[Date]"), SortOn:=xlSortOnValues, Order:= _
              xlAscending, DataOption:=xlSortNormal
1930      ws.ListObjects("Table2").Sort.SortFields.Add _
              Key:=Range("Table2[Time]"), SortOn:=xlSortOnValues, Order:= _
              xlAscending, DataOption:=xlSortNormal
1940      With ws.ListObjects("Table2").Sort
1950          .Header = xlYes
1960          .MatchCase = False
1970          .Orientation = xlTopToBottom
1980          .SortMethod = xlPinYin
1990          .Apply
2000      End With
2010  ws.Range("D2:D" & LastRow).Rows.AutoFit
2020  ws.Activate

      'Remove blank rows
2030  For rc = LastRow To 3 Step -1    ' Step down to 1 to insure all deletions * For speed, down to last 10
2040        If UCase(ws.Cells(rc, 6).Value) = "-" Then ws.Rows(rc).Delete 'Column 6 ="-"
2050        If UCase(ws.Cells(rc, 1).Value) = "" Then ws.Rows(rc).Delete  '- need to insure rc greater than first two rows
2060  Next rc


2070  Application.ScreenUpdating = True
2080  ws.Range("A62500").End(xlUp).Select
2090  Application.EnableEvents = True
2100  ActiveWorkbook.Save


2110      On Error GoTo 0
2120      Exit Sub

SortCaseLogs_Error:

         
2130      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2140      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: SortCaseLogs Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2150      Print #1, zMsg

2160      Close #1

            
2170      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
'---------------------------------------------------------------------------------------
' Method : ValidName
' Author : gouldd
' Date   : 9/1/2016
' Purpose: Verifies name is longer than one character and no illegal characters
'---------------------------------------------------------------------------------------
Function ValidName(strAnyName As String) As Boolean

2180  On Error GoTo ValidName_Error
      Dim zMsg As String

      Dim DQ As String
2190  DQ = Chr$(34)
2200  ValidName = True ' assume valid
      Const InvalidChars As String = "*[|\/?*():;]*"
2210  If InStr(strAnyName, DQ) > 0 Then ValidName = False
2220  If ValidName = True Then
2230      ValidName = Not strAnyName Like InvalidChars
2240  End If


2250      On Error GoTo 0
2260  Exit Function

ValidName_Error:

         
2270      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2280      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: ValidName Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2290      Print #1, zMsg

2300      Close #1

            
2310      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function

Function ValidMileage(txtStartM, txtEndM, txtMileageAddress As String, anyfrm As UserForm)

2320  On Error GoTo ValidMileage_Error
      Dim zMsg As String

      Dim result As Integer

2330  If Val(txtStartM) < 2 Then
2340          ValidMileage = False
2350          result = MsgBox("Check your starting mileage entry!", vbExclamation, "Mileage issue")
2360          anyfrm![txtStartM].SetFocus
2370          Exit Function
2380  End If
          
2390  If Val(txtEndM) < 2 Then
2400          ValidMileage = False
2410          result = MsgBox("Check your ending mileage entry! ", vbExclamation, "Mileage issue")
2420          anyfrm![txtEndM].SetFocus
2430          Exit Function
2440  End If

2450  If Trim(txtMileageAddress) = "" Or Len(Trim(txtMileageAddress)) < 3 Then
2460      ValidMileage = False
2470      result = MsgBox("Can't have a blank address!", vbExclamation, "Mileage issue")
2480      anyfrm![txtMileageAddress].SetFocus
2490      Exit Function
2500  End If

2510  If (Val(txtEndM) - Val(txtStartM)) < 1 Then
2520          ValidMileage = False
2530          result = MsgBox("Can't have a mileage less than one!", vbExclamation, "Mileage issue")
2540          anyfrm![txtStartM].SetFocus
2550          Exit Function
2560      Else
2570          ValidMileage = True
2580      End If

2590      On Error GoTo 0
2600  Exit Function

ValidMileage_Error:

         
2610      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2620      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: ValidMileage Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2630      Print #1, zMsg

2640      Close #1

            
2650      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function

Function ParseOutNames(FullName As String, intAnyPos As Integer) As Variant


2660  On Error GoTo ParseOutNames_Error
      Dim zMsg As String

      Dim FirstName As String
      Dim LastName As String
      Dim MidInitial As String
      Dim Suffix As String
      Dim pos As Integer
      Dim Pos2 As Integer
      Dim Pos3 As Integer

2670  pos = InStr(1, FullName, " ", vbTextCompare)
2680  If pos = 0 Then
2690      pos = Len(FullName) + 1
2700  End If
2710  FirstName = Trim(Left(FullName, pos - 1))

2720  Pos2 = InStr(pos + 1, FullName, " ", vbTextCompare)
2730  If Pos2 Then
2740      Pos3 = InStr(Pos2 + 1, FullName, " ", vbTextCompare)
2750      If Pos3 Then
2760          Suffix = Right(FullName, Len(FullName) - Pos3)
2770          LastName = Mid(FullName, Pos2 + 1, Pos3 - Pos2)
2780      Else
2790          LastName = Right(FullName, Len(FullName) - Pos2)
2800          FirstName = Left(FirstName, Pos2 - 1)
2810      End If
2820  End If

      'Pos2 = InStr(Pos + 2, FullName, " ", vbTextCompare)
2830  If Pos2 = 0 Then
2840      Pos2 = Len(FullName)
2850      LastName = Right(FullName, Len(FullName) - pos)
2860  End If

2870  If Pos2 > pos Then
2880      MidInitial = Mid(FullName, pos + 1, Pos2 - pos)
2890      If LastName = MidInitial Then
2900              MidInitial = ""
2910          End If
2920  End If

2930  pos = InStr(1, FirstName, "-", vbTextCompare)
2940  If pos Then
2950      LastName = Trim(StrConv(Left(LastName, pos), vbProperCase)) & _
          Trim(StrConv(Right(LastName, Len(LastName) - pos), vbProperCase))
2960  Else
2970      FirstName = Trim(StrConv(FirstName, vbProperCase))
2980  End If

2990  FirstName = Trim(StrConv(FirstName, vbProperCase))
3000  MidInitial = Trim(StrConv(MidInitial, vbProperCase))
3010  LastName = Trim(StrConv(LastName, vbProperCase))
3020  Suffix = Trim(StrConv(Suffix, vbProperCase))
      '
      ' suffix handling
      '

3030  Select Case intAnyPos
          Case 1
3040          ParseOutNames = FirstName
3050      Case 2
3060          ParseOutNames = MidInitial
3070      Case 3
3080          ParseOutNames = LastName
3090      Case 4
3100          ParseOutNames = Suffix
3110  End Select


3120      On Error GoTo 0
3130      Exit Function

ParseOutNames_Error:

         
3140      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3150      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: ParseOutNames Within: Module1" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3160      Print #1, zMsg

3170      Close #1

            
3180      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function


