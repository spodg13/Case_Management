Attribute VB_Name = "Module3"
Option Explicit
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef _
lpdwFlags As Long, ByVal ipszConnectionName As String, ByVal _
dwNameLen As Integer, ByVal dwReserved As Long) As Long
Sub DeleteFilesFolders()
10    Call RecursiveFolder("C:\Users\YourUSERNAME\AppData\Local\Temp")
End Sub
Sub UserChange(anyProc As String)

20    On Error GoTo UserChange_Error
      Dim zMsg As String


30        Open "W:\Investigations\ICMS\ErrorLogs\ICMSOpenLog.txt" For Append As #1

40          zMsg = Now & " " & Files.Cells(20, 2).Value & " : " & _
                          Environ("UserName") & vbCrLf & _
                          anyProc & vbCrLf

50           Print #1, zMsg

60           Close #1

70        On Error GoTo 0
80    Exit Sub

UserChange_Error:

         
90        Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

100       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserChange Within: Module3" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

110       Print #1, zMsg

120       Close #1

            
130       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Sub RecursiveFolder(MyPath As String)

       Dim FileSys As FileSystemObject
       Dim objfolder As Scripting.Folder
       Dim objSubFolder As Scripting.Folder
       Dim objFile As File
       Dim result As Integer
140    Set FileSys = CreateObject("Scripting.FileSystemObject")
150    Set objfolder = FileSys.GetFolder(MyPath)

160   On Error Resume Next
170    For Each objFile In objfolder.Files
180    If Left(objFile.Name, 1) = "~" And objFile.Name Like "*" & ThisWorkbook.Name Then
       'MsgBox "Instance of tilde found. " & objFile.Name & vbCrLf & objfolder.Path
       
       'result = MsgBox("Delete " & objFile.Name & "?", vbYesNo, "Potential to lose data with temp file")
       'If result = vbYes Then
190       objFile.Delete
       'End If
200    End If
210    Next objFile
       
220    Set FileSys = Nothing
230    Set objfolder = Nothing
240    Set objSubFolder = Nothing
250    Set objFile = Nothing
260   On Error GoTo 0

End Sub

Sub RemoveColumnDelete()
270       Application.CommandBars("List Range Popup").Controls("&Delete").Controls("Table Columns").Enabled = False
280       Application.CommandBars("List Range Popup").Controls("&Insert").Enabled = False
290       Application.CommandBars("List Range Popup").Controls("Se&lect").Enabled = False
300       Application.CommandBars("List Range Popup").Controls("&Paste").Enabled = False
310   With Application
            Dim cButDMV As CommandBarButton
            Dim cButCLU As CommandBarButton
            Dim cButLexis As CommandBarButton
320         Set cButDMV = .CommandBars("List Range Popup").Controls.Add(Temporary:=True)
330             cButDMV.Caption = "DMV"
340             cButDMV.Style = msoButtonCaption
350             cButDMV.OnAction = "RunFile"
360             cButDMV.Parameter = "DMV"
370         Set cButCLU = .CommandBars("List Range Popup").Controls.Add(Temporary:=True)
380             cButCLU.Caption = "CLU"
390             cButCLU.Style = msoButtonCaption
400             cButCLU.OnAction = "RunFile"
410             cButCLU.Parameter = "CLU"
420         Set cButLexis = .CommandBars("List Range Popup").Controls.Add(Temporary:=True)
430             cButLexis.Caption = "Lexis"
440             cButLexis.Style = msoButtonCaption
450             cButLexis.OnAction = "RunFile"
460             cButLexis.Parameter = "Lexis"
470         End With
            
          
End Sub
Sub ResetMenu()
480       Application.CommandBars("List Range Popup").Reset
End Sub
Function ReAssignable() As Boolean
          
490       If Not StrComp(Files.Cells(39, 2).Value, Environ("UserName"), vbTextCompare) = 0 Then
500           ReAssignable = True
510       End If
          
End Function
Function IsWorkBookOpen(FileName As String)
          Dim ff As Long, ErrNo As Long

520       On Error Resume Next
530       ff = FreeFile()
540       Open FileName For Input Lock Read As #ff
550       Close ff
560       ErrNo = Err
570       On Error GoTo 0

580       Select Case ErrNo
          Case 0:    IsWorkBookOpen = False
590       Case 70:   IsWorkBookOpen = True
600       Case Else: Error ErrNo
610       End Select
End Function
Sub AwardCriteria(anyTotal As Long, anyType As String)
      Dim strMsg As String
      Dim result As Integer


620    If anyType = "Sub" Then
       
630                Select Case anyTotal
                          
                          Case 100 To 110
640                           If Files.Cells(41, 2).Value = True Then
650                               strMsg = "Great job " & Files.Cells(20, 2).Value & vbNewLine & "You have served " & anyTotal & " subpoenas." & vbNewLine & "Outstanding Effort on your part."
660                               result = MsgBox(strMsg, vbOKOnly, "Outstanding Effort - Subpoenas")
670                               Files.Cells(41, 2).Value = False
680                           End If
690                       Case 250 To 260
700                           If Files.Cells(40, 2).Value = True Then
710                               strMsg = "Great job " & Files.Cells(20, 2).Value & vbNewLine & "You have served " & anyTotal & " subpoenas." & vbNewLine & "Outstanding Effort - You are an expert in your trade"
720                               result = MsgBox(strMsg, vbOKOnly, "Outstanding Effort - Subpoena Expert Award")
730                               Files.Cells(41, 2).Value = False
740                           End If
750                       Case 500 To 510
760                           If Files.Cells(40, 2).Value = True Then
770                               strMsg = "Great job " & Files.Cells(20, 2).Value & vbNewLine & "You have served " & anyTotal & " subpoenas." & "Outstanding Effort - You are an master in your trade"
780                               result = MsgBox(strMsg, vbOKOnly, "Outstanding Effort - Subpoena Master Award")
790                               Files.Cells(41, 2).Value = False
800                           End If
810                       Case 1000 To 1010
820                           If Files.Cells(40, 2).Value = True Then
830                               strMsg = "Great job " & Files.Cells(20, 2).Value & vbNewLine & "You have served " & anyTotal & " subpoenas." & "Outstanding Effort - You have reached the Lifetime Achievement award"
840                               result = MsgBox(strMsg, vbOKOnly, "Lifetime Achievement - Subpoena Service")
850                               Files.Cells(41, 2).Value = False
860                           End If
870                       Case Else
880                           Files.Cells(41, 2).Value = True
890                   End Select
900    End If
       
910    If anyType = "Int" Then
       
920                   Select Case anyTotal
                          
                          Case 500 To 510
930                           If Files.Cells(40, 2).Value = True Then
940                               strMsg = "Great job " & Files.Cells(20, 2).Value & vbNewLine & "You have interviewed " & anyTotal & " witnesses." & vbNewLine & "Outstanding Effort on your part."
950                               result = MsgBox(strMsg, vbOKOnly, "Outstanding Effort - Interviews")
960                               Files.Cells(40, 2).Value = False
970                           End If
980                       Case 1000 To 1010
990                           If Files.Cells(40, 2).Value = True Then
1000                              strMsg = "Great job " & Files.Cells(20, 2).Value & vbNewLine & "You have interviewed " & anyTotal & " witnesses." & vbNewLine & "Outstanding Effort - You are an expert in your trade"
1010                              result = MsgBox(strMsg, vbOKOnly, "Outstanding Effort - Interview Expert Award")
1020                              Files.Cells(40, 2).Value = False
1030                          End If
1040                      Case 2000 To 2010
1050                          If Files.Cells(40, 2).Value = True Then
1060                              strMsg = "Great job " & Files.Cells(20, 2).Value & vbNewLine & "You have interviewed " & anyTotal & " witnesses." & "Outstanding Effort - You are an master in your trade"
1070                              result = MsgBox(strMsg, vbOKOnly, "Outstanding Effort - Interview Master Award")
1080                              Files.Cells(40, 2).Value = False
1090                          End If
1100                      Case 4000 To 4010
1110                          If Files.Cells(40, 2).Value = True Then
1120                              strMsg = "Great job " & Files.Cells(20, 2).Value & vbNewLine & "You have interviewed " & anyTotal & " witnesses." & "Outstanding Effort - You have reached the Lifetime Achievement award"
1130                              result = MsgBox(strMsg, vbOKOnly, "Lifetime Achievement - Interviews")
1140                              Files.Cells(40, 2).Value = False
1150                           End If
1160                      Case Else
1170                          Files.Cells(40, 2).Value = True
1180                  End Select
                              
1190    End If

End Sub
Function ProperCase(StrTxt As String, Optional Caps As Long, Optional Excl As Long) As String

1200  On Error GoTo ProperCase_Error
      Dim zMsg As String

           'Convert an input string to proper-case.
           'Surnames like O', Mc and hyphenated names are converted to proper case also.
           'If Caps = 0, then upper-case strings like ABC are preserved; otherwise they're converted.
           'If Excl = 0, selected words are retained as lower-case, except when they follow specified punctuation marks.
          Dim i As Long, j As Long, k As Long, l As Long, bChngFlg As Boolean
          Dim StrTmpA As String, StrTmpB As String, StrExcl As String, StrPunct As String, StrChr As String
1210      StrExcl = " a , an , and , as , at , but , by , for , from , if , in , is , of , on , or , the , this , to , with "
1220      StrPunct = "!,:,.,?,"""
1230      If Excl <> 0 Then
1240          StrExcl = ""
1250          StrPunct = ""
1260      End If
1270      If Len(Trim(StrTxt)) = 0 Then
1280          ProperCase = StrTxt
1290          Exit Function
1300      End If
1310      If Caps <> 0 Then StrTxt = LCase(StrTxt)
1320      StrTxt = " " & StrTxt & " "
1330      For i = 1 To UBound(Split(StrTxt, " "))
1340          StrTmpA = " " & Split(StrTxt, " ")(i) & " "
1350          StrTmpB = UCase(Left(StrTmpA, 2)) & Right(StrTmpA, Len(StrTmpA) - 2)
1360          StrTxt = Replace(StrTxt, StrTmpA, StrTmpB)
1370      Next i
1380      StrTxt = Trim(StrTxt)
           'Code for handling O' names
1390      For i = 1 To UBound(Split(StrTxt, "'"))
1400          If InStr(Right(Split(StrTxt, "'")(i - 1), 2), " ") = 1 Then
1410              StrTmpA = Split(StrTxt, "'")(i)
1420              StrTmpB = UCase(Left(StrTmpA, 1)) & Right(StrTmpA, Len(StrTmpA) - 1)
1430              StrTxt = Replace(StrTxt, StrTmpA, StrTmpB)
1440          End If
1450      Next
           'Code for handling hyphenated names
1460      For i = 1 To UBound(Split(StrTxt, "-"))
1470          StrTmpA = Split(StrTxt, "-")(i)
1480          StrTmpB = UCase(Left(StrTmpA, 1)) & Right(StrTmpA, Len(StrTmpA) - 1)
1490          StrTxt = Replace(StrTxt, StrTmpA, StrTmpB)
1500      Next
           'Code for handling names starting with Mc
1510      If Left(StrTxt, 2) = "Mc" Then
1520          Mid(StrTxt, 3, 1) = UCase(Mid(StrTxt, 3, 1))
1530      End If
1540      i = InStr(StrTxt, " Mc")
1550      If i > 0 Then
1560          Mid(StrTxt, i + 3, 1) = UCase(Mid(StrTxt, i + 3, 1))
1570      End If
           'Code for handling names starting with Mac
1580      If Left(StrTxt, 3) = "Mac" Then
1590          If Len(Split(Trim(StrTxt), " ")(0)) > 5 Then
1600              Mid(StrTxt, 4, 1) = UCase(Mid(StrTxt, 4, 1))
1610          End If
1620      End If
1630      i = InStr(StrTxt, " Mac")
1640      If i > 0 Then
1650          If Len(StrTxt) > i + 5 Then
1660              Mid(StrTxt, i + 4, 1) = UCase(Mid(StrTxt, i + 4, 1))
1670          End If
1680      End If
           'Code to restore excluded words to lower case
1690      For i = 0 To UBound(Split(StrExcl, ","))
1700          StrTmpA = Split(StrExcl, ",")(i)
1710          StrTmpB = UCase(Left(StrTmpA, 2)) & Right(StrTmpA, Len(StrTmpA) - 2)
1720          If InStr(StrTxt, StrTmpB) > 0 Then
1730              StrTxt = Replace(StrTxt, StrTmpB, StrTmpA)
                   'Make sure an excluded words following punctution marks are given proper case anyway
1740              For j = 0 To UBound(Split(StrPunct, ","))
1750                  StrChr = Split(StrPunct, ",")(j)
1760                  StrTxt = Replace(StrTxt, StrChr & StrTmpA, StrChr & StrTmpB)
1770              Next
1780          End If
1790      Next
1800      ProperCase = StrTxt

1810      On Error GoTo 0
1820  Exit Function

ProperCase_Error:

         
1830      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1840      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: ProperCase Within: Module3" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1850      Print #1, zMsg

1860      Close #1

            
1870      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Function


Public Sub EmailAttorney(anyCase As String, anyClient As String, anyAttorney As String)
      Dim OutlookApp As Outlook.Application
      Dim OutlookMessage As MailItem


1880  On Error GoTo EmailAttorney_Error
      Dim zMsg As String

1890      Set OutlookApp = CreateObject(Class:="Outlook.Application") 'If not, open Outlook
          
1900      If Err.Number = 429 Then
1910        MsgBox "Outlook could not be found, aborting.", 16, "Outlook Not Found"
1920      Exit Sub
1930      End If
1940    On Error GoTo 0

      'Create a new email message
1950    Set OutlookMessage = OutlookApp.CreateItem(0)

      'Create Outlook email with attachment
1960    On Error Resume Next
1970      With OutlookMessage
1980       .To = anyAttorney
1990       .CC = ""
2000       .BCC = ""
2010       .Subject = anyClient & " - case update, " & anyCase
              
2020       .Display  'Display to just open up email
2030      End With
2040    On Error GoTo 0


      'Clear Memory
2050    Set OutlookMessage = Nothing
2060    Set OutlookApp = Nothing



2070      On Error GoTo 0
2080      Exit Sub

EmailAttorney_Error:
2090      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1
           
2100      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: EmailAttorney within: Sub - Module3 " & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2110      Print #1, zMsg

2120      Close #1

            
2130      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
          
End Sub

Function fn_validate_drive(ByVal sDrive As String) As Boolean

      Dim FSO As FileSystemObject
2140  Set FSO = CreateObject("Scripting.FileSystemObject")

2150  fn_validate_drive = FSO.DriveExists(sDrive)

2160  Set FSO = Nothing

End Function
Function fn_validate_directory(ByVal strPath As String, ByVal bCreate) As Boolean

      'Attempt to find a directory.
      '  Returns TRUE if found, FALSE if not found.
      '  Returns TRUE if user passed bCreate = True and code created the directory.

      Dim FSO As FileSystemObject
2170  Set FSO = CreateObject("Scripting.FileSystemObject")

2180  Select Case FSO.FolderExists(FSO.GetParentFolderName(strPath))
          Case True
2190          fn_validate_directory = True
2200          Exit Function
2210      Case False
2220          If bCreate Then
2230              FSO.CreateFolder (FSO.GetParentFolderName(strPath))
2240              fn_validate_directory = True
2250          Else
2260              fn_validate_directory = False
2270          End If
2280  End Select

2290  Set FSO = Nothing

End Function
Public Function IsInternetConnected() As Boolean
      Dim strConnType As String
      Dim lngReturnStatus As Long
2300      IsInternetConnected = False
2310      lngReturnStatus = InternetGetConnectedStateEx(lngReturnStatus, strConnType, 254, 0)
2320      If lngReturnStatus = 1 Then IsInternetConnected = True
End Function

Public Sub EndofNight()

2330      ActiveWorkbook.Close SaveChanges:=False

End Sub
Function IsActiveCellInTable() As Boolean

        'Function returns true if active cell is in a table and
        'false if it isn't.
          Dim rngActiveCell
2340      Set rngActiveCell = ActiveCell
2350      Debug.Print IsActiveCellInTable
          'Test for table.
          'Statement produces error when active cell is not
          'in a table.
2360      On Error Resume Next
2370      rngActiveCell = (rngActiveCell.ListObject.Name <> "")
2380      On Error GoTo 0
          'Set function's return value.
2390      IsActiveCellInTable = rngActiveCell
End Function
Public Sub AddRowToTable(ByRef tableName As String, ByRef data As Variant)
          Dim tableLO As ListObject
          Dim tableRange As Range
          Dim newRow As Range

2400      Set tableLO = Range(tableName).ListObject
2410      tableLO.AutoFilter.ShowAllData

2420      If (tableLO.ListRows.Count = 0) Then
2430          Set newRow = tableLO.ListRows.Add(AlwaysInsert:=True).Range
2440      Else
2450          Set tableRange = tableLO.Range
2460          tableLO.Resize tableRange.Resize(tableRange.Rows.Count + 1, tableRange.Columns.Count)
2470          Set newRow = tableLO.ListRows(tableLO.ListRows.Count).Range
2480      End If

2490      If TypeName(data) = "Range" Then
2500          newRow.Activate   '= data.Value
2510      Else
2520          newRow.Activate ' =data
2530      End If
End Sub
Public Sub ResetButton(ByRef btn As Object)
      ' Purpose:      Reset button size and font size for form command button on worksheet
      '               Addresses known Excel bug(s) which alters button size and/or apparent font size
      ' Parameters:   Reference to button object
      ' Remarks:      Getting/setting font size fails since font size remains the same; display (apparent) size changes
      '               AutoSize maximizes the font size to fit the current button size in case it has changed
      '               Button size is reset in case it has changed
      '               Finally, font size is reset to adjust for font changes applied by AutoSize
      '               This fix seems to handle shrinking button icon sizes as well
      Dim h As Integer    'command button height
      Dim w As Integer    '               width
      Dim fs As Integer   '               font size
2540      With btn
2550          h = .Height             'capture original values
2560          w = .Width
2570          fs = .font.Size
2580          .AutoSize = True        'apply maximum font size to fit button
2590          .AutoSize = False
2600          .Height = h             'reset original button and font sizes
2610          .Width = w
2620          .font.Size = fs
2630      End With
End Sub
Sub CenterForm(ByRef anyform As Object)

2640  With anyform
2650    .StartUpPosition = 0
2660    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
2670    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
          
2680  End With
End Sub
Sub RunFile()
      Dim p As String
      Dim myFile As String, Cmd As String
      Dim TheLastCaseRow As Long
      Dim strAction As String
      Dim activeRow As Long
      Dim strName As String
2690  Application.ScreenUpdating = False
2700  activeRow = ActiveCell.row

2710  p = Environ("username")
2720  TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
2730  TheLastCaseRow = TheLastCaseRow + 1
                
2740  If Cells(activeRow, 1).Value Like "##?*" Then

2750      strName = InputBox("Enter the name or number you are running", Cells(activeRow, 1).Value)
2760    If StrPtr(strName) = 0 Then Exit Sub    'Cancel pressed
2770      strAction = "Ran a " & CommandBars.ActionControl.Parameter & " check on " & strName
2780         CaseLogs.Cells(TheLastCaseRow, 1).Value = Cells(activeRow, 1).Value
2790         CaseLogs.Cells(TheLastCaseRow, 2).Value = GetADate(Now) 'Format(txtDOInt, "m/d/yy")
2800         CaseLogs.Cells(TheLastCaseRow, 3).Value = Format(Now, "h:mm AMPM")
2810         CaseLogs.Cells(TheLastCaseRow, 4).Value = strAction
2820   End If
       
2830  Application.ScreenUpdating = True
2840  With CommandBars.ActionControl
          

2850  Select Case .Parameter
          Case "DMV"
2860         myFile = "\\pd-fp01\user$\" & p & "\Desktop\DMV.zws"
2870      Case "CLU"
2880          myFile = "\\pd-fp01\user$\" & p & "\Desktop\PD-CLU.appref-ms"
2890      Case "Lexis"
2900          ActiveWorkbook.FollowHyperlink Address:="https://advance.lexis.com/firsttime?crid=6d95376e-9cef-4865-941c-1ede01a3c72a"
2910          GoTo theEnd
              
2920  End Select
2930  End With
          
2940  Cmd = "RunDLL32.EXE shell32.dll,ShellExec_RunDLL "
2950  Shell (Cmd & myFile)
theEnd:
End Sub

