Attribute VB_Name = "Module2"
Option Explicit


Function GetDueDate(AnyDocket) As Variant


10    On Error GoTo GetDueDate_Error
      Dim zMsg As String

      Dim FindString As String
          Dim rng As Range
20        FindString = AnyDocket
30        If Trim(FindString) <> "" Then
40            With InvestigationLog.Range("A:A")
50                Set rng = .Find(What:=FindString, _
                                  After:=.Cells(.Cells.Count), _
                                  LookIn:=xlValues, _
                                  LookAt:=xlWhole, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlNext, _
                                  MatchCase:=False)
60                If Not rng Is Nothing Then
70                    GetDueDate = InvestigationLog.Cells(rng.row, 6).Value
80                    Else
90                    GetDueDate = "Unknown"
100               End If
110           End With
120       End If


130       On Error GoTo 0
140   Exit Function

GetDueDate_Error:

         
150       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

160       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: GetDueDate Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

170       Print #1, zMsg

180       Close #1

            
190       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function
Sub UpdateCaseLog(strActionEntry, strCurrentTime, strAnyDate As String, TheLastCaseRow, anyRow, anyCol, anyDuration As Long)


200   On Error GoTo UpdateCaseLog_Error
      Dim zMsg As String

210       CaseLogs.Cells(TheLastCaseRow, 1).Value = InvestigationLog.Cells(anyRow, anyCol).Value
220       CaseLogs.Cells(TheLastCaseRow, 2).Value = GetADate(strAnyDate) 'Format(txtDOInt, "m/d/yy")
230       CaseLogs.Cells(TheLastCaseRow, 3).Value = Format(strCurrentTime, "h:mm AMPM")
240       CaseLogs.Cells(TheLastCaseRow, 4).Value = strActionEntry
250       If anyDuration > 0 Then
260           CaseLogs.Cells(TheLastCaseRow, 5).Value = anyDuration
270       End If
          

280       On Error GoTo 0
290   Exit Sub

UpdateCaseLog_Error:

         
300       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

310       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UpdateCaseLog Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

320       Print #1, zMsg

330       Close #1

            
340       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Public Function NDow(Y As Integer, m As Integer, _
     N As Integer, DOW As Integer) As Date


350   On Error GoTo NDow_Error
      Dim zMsg As String

360   NDow = DateSerial(Y, m, (8 - Weekday(DateSerial(Y, m, 1), _
        (DOW + 1) Mod 8)) + ((N - 1) * 7))


370       On Error GoTo 0
380   Exit Function

NDow_Error:

         
390       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

400       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: NDow Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

410       Print #1, zMsg

420       Close #1

            
430       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function
Public Function HolidayCheck(AnyDate As Date) As Boolean
          
440   Select Case AnyDate
          Case Month(AnyDate) = 1 And Day(AnyDate) = 1 'New Years day
450           HolidayCheck = True
460       Case NDow(Year(AnyDate), 1, 3, 1) '3rd Monday in January MLK
470           HolidayCheck = True
480       Case Month(AnyDate) = 2 And Day(AnyDate) = 12 'Abraham Lincoln
490           HolidayCheck = True
500       Case NDow(Year(AnyDate), 2, 3, 1) '3rd Monday in February George Washington
510           HolidayCheck = True
520       Case Month(AnyDate) = 3 And Day(AnyDate) = 31 'Cesar Chavez
530           HolidayCheck = True
540       Case Month(AnyDate) = 7 And Day(AnyDate) = 4 'July 4
              'Check if Sunday or Saturday
550           HolidayCheck = True
560       Case NDow(Year(AnyDate), 9, 1, 1) ' 1st Monday in September, Labor Day
570           HolidayCheck = True
580       Case Month(AnyDate) = 11 And Day(AnyDate) = 11 'Veterans Day
590           HolidayCheck = True
600       Case NDow(Year(AnyDate), 11, 4, 4) ' 4th Thursday in November
610           HolidayCheck = True
620       Case NDow(Year(AnyDate), 11, 4, 5) ' 4th Friday in November
630           HolidayCheck = True
640       Case Month(AnyDate) = 12 And Day(AnyDate) = 25 'Christmas
650           HolidayCheck = True
660       Case Else
670           HolidayCheck = False
680   End Select



End Function
Function SumVisible(CRange As Range, strAction As String)


690   On Error GoTo SumVisible_Error
      Dim zMsg As String

      Dim TotalSum As Integer
      Dim cell As Range
700       Application.Volatile
          
StartLoop:
710       TotalSum = 0
720       For Each cell In CRange
730          If cell.Columns.Hidden = False Then
740             If cell.Rows.Hidden = False Then
750                If cell.Value = "" Then
760                   If strAction = "Action" Then
770                           cell.Activate
780                           frmDuration.Show
790                   End If
800                End If
810                TotalSum = TotalSum + cell.Value
820             End If
830          End If
840       Next
850       SumVisible = TotalSum

860       On Error GoTo 0
870       Exit Function

SumVisible_Error:

         
880       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

890       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: SumVisible Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

900       Print #1, zMsg

910       Close #1

            
920       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function
Function GetVersion(ByVal strAnyPath As String, ByVal strAnyFile As String) As Integer


930   On Error GoTo GetVersion_Error
      Dim zMsg As String

      Dim strOldFile, strFileRet As String
      Dim FileExists As Boolean
      Dim result As Integer
      Dim BtnCode As Integer
      Dim shl As Object
940   Set shl = CreateObject("wscript.shell")





950   strFileRet = strAnyPath & strAnyFile & ".docx"

960   If Len(Dir(strFileRet)) > 0 Then
970           strOldFile = strAnyFile
980           FileExists = True
990           If IsNumeric(Right(strAnyFile, 1)) Then
                  'strip last two
1000              If Mid(strAnyFile, Len(strAnyFile) - 1, 1) = "_" Then
1010                  strAnyFile = Left(strAnyFile, Len(strAnyFile) - 2)
1020              End If
1030          End If
1040          GetVersion = 1
1050      Do
1060          GetVersion = GetVersion + 1
1070          strFileRet = strAnyPath & strAnyFile & "_" & CStr(GetVersion) & ".docx"
1080      Loop Until Len(Dir(strFileRet)) = 0
              'result = MsgBox("A previous version of " & strOldFile & ".docx " & " exists." & Chr$(13) & "Creating version " & CStr(GetVersion), vbOKOnly, "File exists")
1090          BtnCode = shl.Popup("A previous version of " & strOldFile & ".docx " & " exists." & Chr$(13) & "Creating version " & CStr(GetVersion), 5, "File exists", 0 + 48)
1100  Else
1110          FileExists = False
1120          GetVersion = 1
              
1130  End If


1140      On Error GoTo 0
1150      Exit Function

GetVersion_Error:

         
1160      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1170      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: GetVersion Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1180      Print #1, zMsg

1190      Close #1

            
1200      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function

Sub PrintActionLog(strAnyFile, strAnyDate As String, rngAnyRange As Range, isAuto As Boolean)


1210  On Error GoTo PrintActionLog_Error
      Dim zMsg As String

      Dim wdApp As New Word.Application
      Dim wdDoc As Word.Document
      Dim CCtrl As Word.ContentControl

1220  CaseLogs.Activate
1230  rngAnyRange.Copy

1240  Set wdDoc = wdApp.Documents.Open(FileName:=strAnyFile, AddToRecentFiles:=True, Visible:=False)
1250              wdDoc.Activate
1260              With wdDoc
1270              For Each CCtrl In .ContentControls
1280                  If CCtrl.Title = "DueDate" Then
1290                      CCtrl.Range.Text = strAnyDate
1300                  End If
1310              Next
1320              End With
1330              wdDoc.Tables(1).Select
1340              wdApp.Selection.EndOf Unit:=wdCell
1350              wdApp.Selection.InsertRowsBelow (1)
1360              wdApp.Selection.PasteAppendTable
1370              wdApp.Selection.EndOf Unit:=wdCell
1380              wdApp.Selection.InsertRowsBelow (1)
1390              wdDoc.Save
1400              If isAuto = True Then
1410                  wdDoc.PrintOut
1420                  wdDoc.Close
1430                  wdApp.Quit
1440                  Set wdApp = Nothing
1450              Else
1460                  wdApp.Visible = True
1470                  wdApp.Documents(strAnyFile).Activate
                      'wdApp.Activate - old single line code
1480                  Set wdApp = Nothing
1490                  Set wdDoc = Nothing
1500              End If
                  

1510      On Error GoTo 0
1520      Exit Sub

PrintActionLog_Error:

         
1530      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1540      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: PrintActionLog Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1550      Print #1, zMsg

1560      Close #1

            
1570      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Function ActionRange(intSColumn, intEndColumn) As Range


1580  On Error GoTo ActionRange_Error
      Dim zMsg As String

      Dim LastStart As Long
      Dim EndCell, StartCell As Range
      Dim TheLastRow As Long

1590  CaseLogs.Activate
1600  TheLastRow = CaseLogs.UsedRange.SpecialCells(xlCellTypeLastCell).row

      'Look for a previous end.
1610  Set EndCell = Columns(7).Find(What:="End", LookAt:=xlWhole, SearchDirection:=xlPrevious, MatchCase:=False)
      'If no end, then look for look forward for first Start.
1620  If EndCell Is Nothing Then
          'LastStart = Columns(7).Find(What:="Start", LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False).row
1630      Set StartCell = Columns(7).Find(What:="Start", LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)
1640      If StartCell Is Nothing Then
1650          LastStart = Rows("4:64000").SpecialCells(xlCellTypeVisible).row
1660      Else
1670          LastStart = StartCell.row
1680      End If
          
1690  Else
          'If end exists, activate cell and look forward for end cell for the start
          
1700      If EndCell.row = TheLastRow Then
1710          Set EndCell = Columns(7).FindPrevious(After:=EndCell)
1720      End If
1730      EndCell.Activate
1740      LastStart = Columns(7).Find(What:="Start", After:=ActiveCell, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False).row
1750  End If
1760  Set ActionRange = ActiveSheet.Range(Cells(LastStart, intSColumn), Cells(TheLastRow, intEndColumn))


1770      On Error GoTo 0
1780      Exit Function

ActionRange_Error:

         
1790      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1800      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: ActionRange Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1810      Print #1, zMsg

1820      Close #1

            
1830      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function


Public Function GetFileNameBase(ByVal strClient As String, ByVal strDocket As String) As String

1840  On Error GoTo GetFileNameBase_Error
      Dim zMsg As String

      Dim strFileName As String
1850      strFileName = Replace(strClient, ", ", "_")
1860      GetFileNameBase = Trim(strFileName & "_" & strDocket)

1870      On Error GoTo 0
1880  Exit Function

GetFileNameBase_Error:

         
1890      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1900      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: GetFileNameBase Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1910      Print #1, zMsg

1920      Close #1

            
1930      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function

Sub PrintWitnessList(strAnyDocket, strAnyClient As String)


1940  On Error GoTo PrintWitnessList_Error
      Dim zMsg As String

      Dim rng As Range
      Dim TheLastRow As Long
      Dim TheColumn As Integer
      Dim SelString As String
1950  TheColumn = 1

1960  WitnessLog.Activate
1970  ActiveCell.SpecialCells(xlLastCell).Select
1980  TheLastRow = Selection.row
1990  Set rng = WitnessLog.Range("a1", "h" & TheLastRow)

2000              With rng '
2010                  .AutoFilter TheColumn, strAnyDocket
                      
2020              End With

2030  ActiveCell.SpecialCells(xlLastCell).Select
2040  TheLastRow = Selection.row
2050  SelString = "b1:H" & TheLastRow

2060  ActiveSheet.PageSetup.PrintArea = SelString

2070  With ActiveSheet.PageSetup
2080      .PrintTitleRows = "$1:$1"
          '.PrintTitleColumns = "$B:$B"
2090      .Orientation = xlLandscape
2100      .CenterHeader = "&16 &b" & "Witness List for " & strAnyClient 'remove 16 to be back to normal
2110      .RightHeader = "Printed: " & Date
2120      .LeftHeader = "&12 &b" & strAnyDocket
2130      .Zoom = False
2140      .FitToPagesWide = 1
          '.FitToPagesTall = 1
2150  End With
2160  ActiveSheet.PrintOut
2170  With rng
2180      .AutoFilter
2190  End With
                              

2200      On Error GoTo 0
2210      Exit Sub

PrintWitnessList_Error:

         
2220      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2230      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: PrintWitnessList Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2240      Print #1, zMsg

2250      Close #1

            
2260      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Public Sub addAttorney(ByVal strData As String)


2270  On Error GoTo addAttorney_Error
      Dim zMsg As String

          Dim lLastRow As Long
2280      Application.ScreenUpdating = False

          'Dim Tbl As ListObject  '**
          'Dim NewRow As ListRow   '**

         '** Set NewRow = Attorneys.ListObjects("Attorneys").ListRows.Add(AlwaysInsert:=True)
         '**NewRow.Range.Cells(1, 1).Value = strData
           
2290      lLastRow = Attorneys.Cells(Rows.Count, 1).End(xlUp).row
2300      lLastRow = lLastRow + 1
2310      Attorneys.Cells(lLastRow, 1).Value = strData
2320      SortAttorneys
2330      ActiveWorkbook.Save
2340      Application.ScreenUpdating = True

2350      On Error GoTo 0
2360      Exit Sub

addAttorney_Error:

         
2370      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2380      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: addAttorney Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2390      Print #1, zMsg

2400      Close #1

            
2410      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Public Function AttorneyName(anyString As String) As String


2420  On Error GoTo AttorneyName_Error
      Dim zMsg As String

      Dim icomma, iremove As Integer
      Dim strFirst, strLast As String

2430  If anyString = "" Then anyString = "-,-"
         
2440  icomma = InStr(anyString, ",")
2450  strLast = Left(anyString, icomma - 1)
2460  iremove = Len(anyString) - Len(strLast) - 1
2470  strFirst = Trim(Right(anyString, iremove))

2480  AttorneyName = strFirst & " " & strLast


2490      On Error GoTo 0
2500      Exit Function

AttorneyName_Error:

         
2510      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2520      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: AttorneyName Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2530      Print #1, zMsg

2540      Close #1

            
2550      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function
Public Function RevAttyName(ByVal anyString As String) As String

2560  On Error GoTo RevAttyName_Error
      Dim zMsg As String

      Dim ispace As Integer
      Dim strFirst, strLast As String

2570  If anyString = "" Then anyString = "- -"
             
2580  ispace = InStr(anyString, " ")
2590  strFirst = Left(anyString, ispace - 1)
      'iremove = Len(anyString) - Len(strLast) - 1
2600  strLast = Right(anyString, Len(anyString) - Len(strFirst))

2610  RevAttyName = Trim(strLast & ", " & strFirst)


2620      On Error GoTo 0
2630  Exit Function

RevAttyName_Error:

         
2640      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2650      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: RevAttyName Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2660      Print #1, zMsg

2670      Close #1

            
2680      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function


Public Function TaskList(anyInt As String, anyPhoto As String, anySubs As String, anyOther As String, anyTest As Boolean) As String


2690  On Error GoTo TaskList_Error
      Dim zMsg As String

      Dim strInt, strPhoto, strSubs, strOther, strTest, strTemp As String
      Dim strPlural As String
2700  strPlural = "s"
2710  strInt = ""
2720  strPhoto = ""
2730  strSubs = ""
2740  strOther = ""

2750  If Len(anyInt) > 0 And Val(anyInt) > 0 Then
2760      strInt = anyInt & " Interview"
2770      If Val(anyInt) > 1 Then
2780          strInt = strInt & strPlural
2790      End If
2800      strInt = strInt & ", "
2810  End If

2820  If Len(anyPhoto) > 0 And Val(anyPhoto) > 0 Then
2830      strPhoto = anyPhoto & " Photo"
2840      If Val(anyPhoto) > 1 Then
2850          strPhoto = strPhoto & strPlural
2860      End If
2870      strPhoto = strPhoto & ", "
2880  End If

2890  If Len(anySubs) > 0 And Val(anySubs) > 0 Then
2900      strSubs = anySubs & " Subpoena"
2910      If Val(anySubs) > 1 Then
2920          strSubs = strSubs & strPlural
2930      End If
2940      strSubs = strSubs & ", "
2950  End If

2960  If Len(anyOther) > 0 And Val(anyOther) > 0 Then
2970      strOther = anyOther & " Other task"
2980      If Val(anyOther) > 1 Then
2990          strOther = strOther & strPlural
3000      End If
3010      strOther = strOther & ", "
3020  End If

3030  If anyTest = True Then
3040      strTest = "Testimony"
3050      Else
3060      strTest = ""
3070  End If

3080  strTemp = strInt & strPhoto & strSubs & strOther & strTest

3090  If Right(strTemp, 2) = ", " Then
3100      TaskList = Left(strTemp, Len(strTemp) - 2)
3110  Else
3120      TaskList = strTemp
          
3130  End If


3140      On Error GoTo 0
3150      Exit Function

TaskList_Error:

         
3160      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3170      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: TaskList Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3180      Print #1, zMsg

3190      Close #1

            
3200      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function



Public Sub SortByDueDate()

3210  On Error GoTo SortByDueDate_Error
      Dim zMsg As String

      Dim LastRow As Long
      Dim rngSortRange As Range
3220  Application.ScreenUpdating = False

3230  InvestigationLog.Activate

3240  LastRow = ActiveSheet.UsedRange.Rows.Count
3250  Set rngSortRange = Range("a3", "a" & LastRow)


3260      InvestigationLog.ListObjects("Investigation_Log").Sort.SortFields.Clear
3270      InvestigationLog.ListObjects("Investigation_Log").Sort.SortFields.Add _
              Key:=Range("Investigation_Log[DATE COMPLETE]"), SortOn:=xlSortOnValues, Order:= _
              xlAscending, DataOption:=xlSortNormal
3280      InvestigationLog.ListObjects("Investigation_Log").Sort.SortFields.Add _
              Key:=Range("Investigation_Log[DUE DATE]"), SortOn:=xlSortOnValues, Order:= _
              xlAscending, DataOption:=xlSortNormal
3290      With InvestigationLog.ListObjects("Investigation_Log").Sort
3300          .Header = xlYes
3310          .MatchCase = False
3320          .Orientation = xlTopToBottom
3330          .SortMethod = xlPinYin
3340          .Apply
3350      End With

3360  ActiveWorkbook.Save
3370  Range("a65536").End(xlUp).Select



3380      On Error GoTo 0
3390      Exit Sub

SortByDueDate_Error:

         
3400      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3410      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: SortByDueDate Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3420      Print #1, zMsg

3430      Close #1

            
3440      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Sub SortAttorneys()


3450  On Error GoTo SortAttorneys_Error
      Dim zMsg As String

      '
      'Application.ScreenUpdating = False
      ' SortTable Macro
          
3460      Attorneys.ListObjects("Attorneys").Sort.SortFields _
              .Clear
3470      Attorneys.ListObjects("Attorneys").Sort.SortFields _
              .Add Key:=Range("Attorneys[[#All],[Attorney Table]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
              DataOption:=xlSortNormal
3480      With Attorneys.ListObjects("Attorneys").Sort
3490          .Header = xlYes
3500          .MatchCase = False
3510          .Orientation = xlTopToBottom
3520          .SortMethod = xlPinYin
3530          .Apply
3540      End With
3550      Range("a65536").End(xlUp).Select
              
      'Application.ScreenUpdating = True


3560      On Error GoTo 0
3570      Exit Sub

SortAttorneys_Error:

         
3580      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3590      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: SortAttorneys Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3600      Print #1, zMsg

3610      Close #1

            
3620      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Public Sub AddMileage(AnyDate, AnyAddress, AnyDocket, anyStartM, anyEndM As String)

3630  On Error GoTo AddMileage_Error
      Dim zMsg As String

      Dim TheLastMileRow As Double

3640      TheLastMileRow = MileageLog.Cells(Rows.Count, 1).End(xlUp).row
3650      TheLastMileRow = TheLastMileRow + 1
3660      MileageLog.Cells(TheLastMileRow, 1).Value = Format(AnyDate, "m/d/yy")
3670      MileageLog.Cells(TheLastMileRow, 2).Value = AnyAddress
3680      MileageLog.Cells(TheLastMileRow, 3).Value = AnyDocket
3690      MileageLog.Cells(TheLastMileRow, 4).Value = Val(anyStartM)
3700      MileageLog.Cells(TheLastMileRow, 5).Value = Val(anyEndM)
          
         ' ActiveCell.Offset(0, 1).Value = "Mileage Entry"
         ' ActiveCell.Offset(0, 2).Value = Val(AnyStartM)
         ' ActiveCell.Offset(0, 3).Value = Val(AnyEndM)

3710      On Error GoTo 0
3720  Exit Sub

AddMileage_Error:

         
3730      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3740      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: AddMileage Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3750      Print #1, zMsg

3760      Close #1

            
3770      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Public Function DefTitle(AnyStr As String)

3780  On Error GoTo DefTitle_Error
      Dim zMsg As String

3790  If AnyStr = "Juv" Then
3800      DefTitle = "Minor"
3810      Else
3820      DefTitle = "Defendant"
3830  End If
          

3840      On Error GoTo 0
3850  Exit Function

DefTitle_Error:

         
3860      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3870      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: DefTitle Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3880      Print #1, zMsg

3890      Close #1

            
3900      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function

Public Function CaseDesc(AnyStr As String)

3910  On Error GoTo CaseDesc_Error
      Dim zMsg As String

3920  If AnyStr = "Juv" Then
3930      CaseDesc = "Petition"
3940      Else
3950      CaseDesc = "Case"
3960  End If

3970      On Error GoTo 0
3980  Exit Function

CaseDesc_Error:

         
3990      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

4000      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: CaseDesc Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

4010      Print #1, zMsg

4020      Close #1

            
4030      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function

Public Function DivisionTitle(AnyStr As String)

4040  On Error GoTo DivisionTitle_Error
      Dim zMsg As String

4050  If AnyStr = "Juv" Then
4060      DivisionTitle = "Juvenile Division"
4070      Else
4080      DivisionTitle = " "
4090  End If

4100      On Error GoTo 0
4110  Exit Function

DivisionTitle_Error:

         
4120      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

4130      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: DivisionTitle Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

4140      Print #1, zMsg

4150      Close #1

            
4160      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function
Public Sub MakePhotoDir(strAnyFolder As String)


4170  On Error GoTo MakePhotoDir_Error
      Dim zMsg As String

4180  If Dir(strAnyFolder, vbDirectory) = "" Then
4190      MkDir Path:=strAnyFolder
4200      MsgBox "Photo Directory Created"
4210  Else
4220      MsgBox "Photo folder previously created"
4230  End If


4240      On Error GoTo 0
4250  Exit Sub

MakePhotoDir_Error:

         
4260      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

4270      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: MakePhotoDir Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

4280      Print #1, zMsg

4290      Close #1

            
4300      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Public Sub ClearWitnessFilter()


4310  On Error GoTo ClearWitnessFilter_Error
      Dim zMsg As String

4320      ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=4

4330      On Error GoTo 0
4340  Exit Sub

ClearWitnessFilter_Error:

         
4350      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

4360      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: ClearWitnessFilter Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

4370      Print #1, zMsg

4380      Close #1

            
4390      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Public Sub ChangeCaseNumber(strOldCase, strNewCase As String)
          

4400  On Error GoTo ChangeCaseNumber_Error
      Dim zMsg As String

4410      CaseLogs.Columns("A").Replace _
          What:=strOldCase, Replacement:=strNewCase, _
          SearchOrder:=xlByColumns, MatchCase:=True

4420      WitnessLog.Columns("A").Replace _
          What:=strOldCase, Replacement:=strNewCase, _
          SearchOrder:=xlByColumns, MatchCase:=True


4430      On Error GoTo 0
4440  Exit Sub

ChangeCaseNumber_Error:

         
4450      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

4460      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: ChangeCaseNumber Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

4470      Print #1, zMsg

4480      Close #1

            
4490      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Public Sub RenameFilesNewCase(strOldCase, strNewCase, strName As String)

4500  On Error GoTo RenameFilesNewCase_Error
      Dim zMsg As String

      Dim result As Integer
      'Rename any word files  - need a search routine first
          'GivenLocation = "c:\temp\" 'note the trailing backslash
          'OldFileName = "SomeFileName.xls"
          'NewFileName = "DifferentFileName.xls"
          'Name GivenLocation & OldFileName As GivenLocation & NewFileName
          
          
          Dim RetVal As Variant
          Dim strPathOld, strPathNew As String
                     
4510      ChDir Files.Cells(1, 2).Value        'Change folder to Reports folder
4520      strPathOld = Files.Cells(1, 2).Value
4530      RetVal = Dir(strPathOld & strName & strOldCase & "*")      'Get first file in folder
4540      Do While Len((RetVal)) > 0 'Rename until no more files
4550          result = MsgBox("Rename  " & RetVal & "?", vbYesNo, "Rename File")
4560          If result = vbYes Then
4570              CloseFile strPathOld & RetVal
4580              Name (strPathOld & RetVal) As Replace(strPathOld & RetVal, strOldCase, strNewCase) 'Rename
4590          End If
4600          RetVal = Dir()            'Get next file to rename
4610      Loop
4620      ChDir Files.Cells(6, 2).Value 'Check Action Log
4630      strPathOld = Files.Cells(6, 2).Value
4640      RetVal = Dir(strPathOld & strName & strOldCase & "*")       'Get first file in folder
4650      Do While Len((RetVal)) > 0 'Rename until no more files
4660          result = MsgBox("Renaming  " & RetVal & "!", vbOKOnly, "Rename File")
4670          CloseFile strPathOld & RetVal
4680          Name (strPathOld & RetVal) As Replace(strPathOld & RetVal, strOldCase, strNewCase)   'Renamefile regardless
              
4690          RetVal = Dir()            'Get next file to rename
4700      Loop
4710      ChDir Files.Cells(29, 2).Value 'Check photo folders
4720      strPathOld = Files.Cells(29, 2).Value & strName & strOldCase
4730      strPathNew = Files.Cells(29, 2).Value & strName & strNewCase
4740      RetVal = Dir(strPathOld, vbDirectory)      'Get first file in folder
4750      Do While Len((RetVal)) > 0 'Rename until no more files
4760          If MsgBox("Rename Photo Folder  " & RetVal & "?", vbYesNo, "Rename Folder " & strName & strOldCase) = vbYes Then
4770              Name strPathOld As strPathNew    'Rename
4780          End If
4790          RetVal = Dir()            'Get next file to rename
4800      Loop
           
           

4810      On Error GoTo 0
4820  Exit Sub

RenameFilesNewCase_Error:

         
4830      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

4840      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: RenameFilesNewCase Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

4850      Print #1, zMsg

4860      Close #1

            
4870      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Public Sub RenameFilesNewClient(strOldCase, strNewName, strOldName As String)


4880  On Error GoTo RenameFilesNewClient_Error
      Dim zMsg As String

          Dim result As Integer
          Dim RetVal As Variant
          Dim strOldFileName As String
          Dim strNewFileName As String
          Dim strPathOld, strPathNew As String
         ' Dim doc As Object, app As Object
            
             
4890      strOldFileName = strOldName & strOldCase
4900      strNewFileName = strNewName & strOldCase
             
4910      ChDir Files.Cells(1, 2).Value        'Change folder to Reports folder
4920      strPathOld = Files.Cells(1, 2).Value
          
4930      RetVal = Dir(strPathOld & strOldFileName & "*")       'Get first file in folder
4940      Do While Len((RetVal)) > 0 'Rename until no more files
4950          If MsgBox("Rename " & RetVal & "?", vbYesNo, "Rename File") = vbYes Then  '***
4960              CloseFile strPathOld & RetVal
4970              Name (strPathOld & RetVal) As Replace(strPathOld & RetVal, strOldFileName, strNewFileName)    'Rename
4980          End If
4990          RetVal = Dir()            'Get next file to rename
              
5000      Loop
5010      ChDir Files.Cells(6, 2).Value 'Check Action Log
5020      strPathOld = Files.Cells(6, 2).Value
5030      RetVal = Dir(strPathOld & strOldFileName & "*")       'Get first file in folder
5040      Do While Len((RetVal)) > 0 'Rename until no more files
5050          result = MsgBox("Renaming Action Log " & RetVal & "!", vbOKOnly, "Renaming File")
5060          If result = vbOK Then
5070              CloseFile strPathOld & RetVal
5080              Name (strPathOld & RetVal) As Replace(strPathOld & RetVal, strOldFileName, strNewFileName)    'Rename
5090          End If
5100          RetVal = Dir()            'Get next file to rename
5110      Loop
5120      ChDir Files.Cells(29, 2).Value 'Check photo folders
5130      strPathOld = Files.Cells(29, 2).Value & strOldFileName
5140      strPathNew = Files.Cells(29, 2).Value & strNewFileName
5150      RetVal = Dir(strPathOld, vbDirectory)       'Get first file in folder
5160      Do While Len((RetVal)) > 0 'Rename until no more files
5170          If MsgBox("Rename Photo Folder " & RetVal & "?", vbYesNo, "Rename Folder ") = vbYes Then Name strPathOld As strPathNew    'Rename
5180          RetVal = Dir()            'Get next file to rename
5190      Loop
          
          
          

5200      On Error GoTo 0
5210  Exit Sub

RenameFilesNewClient_Error:

         
5220      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

5230      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: RenameFilesNewClient Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

5240      Print #1, zMsg

5250      Close #1

            
5260      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Public Function CaseExists(strAnyCase As String) As Boolean


5270  On Error GoTo CaseExists_Error
      Dim zMsg As String

      Dim rng As Range
5280  If Trim(strAnyCase) <> "" Then
5290      With InvestigationLog.Range("A:A") 'searches all of column A
5300          Set rng = .Find(What:=strAnyCase, _
                              After:=.Cells(.Cells.Count), _
                              LookIn:=xlValues, _
                              LookAt:=xlWhole, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlNext, _
                              MatchCase:=False)
5310          If Not rng Is Nothing Then
5320              Application.GoTo rng, True
5330              CaseExists = True
5340          Else
5350              CaseExists = False  'value not found
5360          End If
5370      End With
5380  End If

5390      On Error GoTo 0
5400  Exit Function

CaseExists_Error:

         
5410      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

5420      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: CaseExists Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

5430      Print #1, zMsg

5440      Close #1

            
5450      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function



Public Function MileageFinished(ByVal rng As Range) As Boolean

5460  On Error GoTo MileageFinished_Error
      Dim zMsg As String

      Dim cell As Range

5470  Application.Volatile

5480  For Each cell In rng
5490         If cell.Columns.Hidden = False Then
5500            If cell.Rows.Hidden = False Then
5510               If cell.Value = "" Or cell.Value <= 0 Then
                              
5520                         MileageFinished = False
5530                         Cells(cell.row, 4).Activate  ' put on starting mileage
5540                         frmMileageError.Show
                          
5550               End If
5560            End If
5570         End If
5580      Next
5590    MileageFinished = True
          

5600      On Error GoTo 0
5610  Exit Function

MileageFinished_Error:

         
5620      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

5630      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: MileageFinished Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

5640      Print #1, zMsg

5650      Close #1

            
5660      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function
Public Function FolderExists(strFolderPath As String) As Boolean
5670      On Error Resume Next
5680      FolderExists = (GetAttr(strFolderPath) And vbDirectory) = vbDirectory
5690      On Error GoTo 0
End Function
Public Function TemplateExists(strTemplateName As String) As Boolean
5700  On Error Resume Next
5710      If Dir(strTemplateName) = "" Then
5720          TemplateExists = False
5730      Else
5740          TemplateExists = True
5750      End If
5760      On Error GoTo 0
End Function



Sub CopyInitFiles(strDestinationFolder, strSourceFolder As String)

5770  On Error GoTo CopyInitFiles_Error
      Dim zMsg As String

      'Declare Variables
      Dim FSO As Object
5780  Set FSO = CreateObject("Scripting.FileSystemObject")


5790   If FSO.FolderExists(strSourceFolder) = False Then
       
5800          MsgBox "Source folder doesn't exist.  Missing " & strSourceFolder
5810          Exit Sub
5820   End If
          
          
5830      MsgBox "Copying files from " & strSourceFolder & " to " & strDestinationFolder
5840      FSO.CopyFolder source:=strSourceFolder, Destination:=strDestinationFolder
          



5850      On Error GoTo 0
5860  Exit Sub

CopyInitFiles_Error:

         
5870      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

5880      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: CopyInitFiles Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

5890      Print #1, zMsg

5900      Close #1

            
5910      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Public Function JuvCase(anyCase As String) As String

5920  On Error GoTo JuvCase_Error
      Dim zMsg As String

      Dim pos As Integer
      Dim intPos As Integer

5930  pos = InStr(1, anyCase, "-", vbTextCompare)

5940  If pos = 0 Then
5950      JuvCase = anyCase
5960      Else
5970      JuvCase = Left(anyCase, pos - 1)
5980  End If
5990  For intPos = 5 To Len(anyCase)
6000          If Mid(anyCase, intPos, 1) Like "[A-Z]" Then
6010              JuvCase = Left(anyCase, intPos - 1)
6020              Exit Function
6030          End If
6040      Next intPos



6050      On Error GoTo 0
6060  Exit Function

JuvCase_Error:

         
6070      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

6080      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: JuvCase Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

6090      Print #1, zMsg

6100      Close #1

            
6110      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function

Public Sub PrintActionDraft(rngAnyRange As Range, strAnyDocket, strAnyClient As String)

6120  On Error GoTo PrintActionDraft_Error
      Dim zMsg As String

6130  rngAnyRange.Select

6140  CaseLogs.PageSetup.PrintArea = Selection.Address
6150  With CaseLogs.PageSetup
6160      .PrintTitleRows = "$3:$3"
          '.PrintTitleColumns = "$B:$B"
6170      .Orientation = xlPortrait
6180      .CenterHeader = "&16 &b" & "Action Log Draft for " & strAnyClient 'remove 16 to be back to normal
6190      .RightHeader = "Printed: " & Date
6200      .LeftHeader = "&12 &b" & strAnyDocket
6210      .Zoom = False
6220      .FitToPagesWide = 1
6230      .FitToPagesTall = False
6240  End With

6250  CaseLogs.PrintPreview
6260  ActiveWorkbook.Save
6270  ClearCaseLogFilter
6280  CaseLogs.PageSetup.PrintArea = ""


6290      On Error GoTo 0
6300  Exit Sub

PrintActionDraft_Error:

         
6310      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

6320      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: PrintActionDraft Within: Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

6330      Print #1, zMsg

6340      Close #1

            
6350      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Sub CloseFile(anyfile As String)

6360      On Error GoTo CloseFile_Error
          Dim zMsg As String

      Dim doc As Object
      Dim app As Object
6370  Set doc = GetObject(anyfile)
6380          If Not doc Is Nothing Then
6390              Set app = GetObject(, "Word.Application")
6400               doc.Close
6410              If app.Documents.Count = 0 Then app.Quit
6420          End If

6430      On Error GoTo 0
6440      Exit Sub

CloseFile_Error:

6450      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

6460      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                          Format(Erl, "###") & vbCrLf & _
                          "Procedure: CloseFile Within: Sub Module2" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

6470      Print #1, zMsg

6480      Close #1

                  
6490      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
End Sub
Function Insert(source As String, strI As String) As String
      Dim pos As Integer

6500      With frmEnterAction.ActiveControl

6510          pos = .SelStart
6520      End With
6530      Insert = Mid(source, 1, pos) & strI & Mid(source, pos + 1, Len(source) - pos)
End Function
Function formatTel(anyTel) As String

6540  On Error GoTo formatTel_Error
      Dim zMsg As String

      Dim result As Integer
          
6550      anyTel = Replace(anyTel, "-", "")
6560      anyTel = Replace(anyTel, "(", "")
6570      anyTel = Replace(anyTel, ")", "")
6580      anyTel = Replace(anyTel, " ", "")
          
6590      If Len(anyTel) <> 10 And Len(anyTel) <> 0 Then
6600          result = MsgBox("Standard telephone number uses 10 digits, are you sure you want to continue with " & Format(anyTel.Text, "(000) 000-0000") & " ?", vbYesNo, "Confirm number")
6610          If result = vbNo Then
6620              formatTel = Format(anyTel, "(000) 00000000000000")
6630          End If
6640      End If
          
6650      formatTel = Format(anyTel, "(000) 000-0000")

6660      On Error GoTo 0
6670  Exit Function

formatTel_Error:

         
6680      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

6690      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: anyTel_Exit Within: formatTel" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

6700      Print #1, zMsg

6710      Close #1

            
6720      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function


