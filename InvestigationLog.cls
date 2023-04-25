VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InvestigationLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdEditCase_Click()
          
          
10    On Error GoTo cmdEditCase_Click_Error
      Dim zMsg As String
20    ResetButton cmdEditCase
30        frmCaseChange.Show

40        On Error GoTo 0
50    Exit Sub

cmdEditCase_Click_Error:

         
60        Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

70        zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdEditCase_Click Within: InvestigationLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

80        Print #1, zMsg

90        Close #1

            
100       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub togProtection_Click()

110   On Error GoTo togProtection_Click_Error
      Dim zMsg As String

      Dim result As Integer
120   ResetButton togProtection
130       If togProtection.Value = True Then
140           ActiveSheet.Unprotect
150           togProtection.Caption = "Edit On"
160           result = MsgBox("Use the 'Change case or Client info' button to edit the case number or defendant name.  Put your cursor on the row you want to change.", vbOKCancel, "Editing turned on")
170           cmdEditCase.Visible = True
180           Application.EnableEvents = True
190           ActiveCell.Select
200       Else
210           ActiveSheet.Protect UserInterfaceOnly:=True
220           togProtection.Caption = "Edit Off"
230           cmdEditCase.Visible = False
240           Application.EnableEvents = True
250       End If
          

260       On Error GoTo 0
270   Exit Sub

togProtection_Click_Error:

         
280       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

290       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: togProtection_Click Within: InvestigationLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

300       Print #1, zMsg

310       Close #1

            
320       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cboClientLocate_Change()


330   On Error GoTo cboClientLocate_Change_Error
      Dim zMsg As String
340   ResetButton cboClientLocate

       Dim FindString As String
          Dim rng As Range
350       FindString = cboClientLocate.Value
360       If Trim(FindString) <> "" Then
370           With InvestigationLog.Range("C:C")
380               Set rng = .Find(What:=FindString, _
                                  After:=.Cells(.Cells.Count), _
                                  LookIn:=xlValues, _
                                  LookAt:=xlWhole, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlNext, _
                                  MatchCase:=False)
390               If Not rng Is Nothing Then
400                   Application.GoTo rng, True
410               Else
420                   MsgBox "Name not found"
430               End If
440           End With
450       End If

460       On Error GoTo 0
470   Exit Sub

cboClientLocate_Change_Error:

         
480       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

490       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cboClientLocate_Change Within: InvestigationLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

500       Print #1, zMsg

510       Close #1

            
520       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub cmdCloseCase_Click()
          
530   On Error GoTo cmdCloseCase_Click_Error
      Dim zMsg As String
540   ResetButton cmdCloseCase
550       frmCloseCase.Show vbModeless

560       On Error GoTo 0
570   Exit Sub

cmdCloseCase_Click_Error:

         
580       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

590       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdCloseCase_Click Within: InvestigationLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

600       Print #1, zMsg

610       Close #1

            
620       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub cmdEnterAction_Click()
          
630   On Error GoTo cmdEnterAction_Click_Error
      Dim zMsg As String
640   ResetButton cmdEnterAction
      Dim lngRow As Long
650   lngRow = ActiveCell.row
660   If InvestigationLog.Cells(lngRow, 1).Value = "" Then
670           MsgBox ("Pick a case before you add an action")
680           Exit Sub
690       End If
700       If InvestigationLog.Cells(lngRow, 1).Value = "IOD" Then
710           frmIODEntry.Show
720       Else
730           frmEnterAction.Show
740       End If

750       On Error GoTo 0
760   Exit Sub

cmdEnterAction_Click_Error:

         
770       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

780       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdEnterAction_Click Within: InvestigationLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

790       Print #1, zMsg

800       Close #1

            
810       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub Worksheet_Activate()

820   On Error GoTo Worksheet_Activate_Error
      Dim zMsg As String

830       InvestigationLog.Protect UserInterfaceOnly:=True
840       togProtection.Value = False
850       togProtection.Caption = "Edit Off"
860       cmdEditCase.Visible = False
870       InvestigationLog.EnableSelection = xlNoRestrictions
880   With Application
            Dim cButDMV As CommandBarButton
            Dim cButCLU As CommandBarButton
            Dim cButLexis As CommandBarButton
890         Set cButDMV = .CommandBars("List Range Popup").Controls.Add(Temporary:=True)
900             cButDMV.Caption = "DMV"
910             cButDMV.Style = msoButtonCaption
920             cButDMV.OnAction = "RunFile"
930             cButDMV.Parameter = "DMV"
940         Set cButCLU = .CommandBars("List Range Popup").Controls.Add(Temporary:=True)
950             cButCLU.Caption = "CLU"
960             cButCLU.Style = msoButtonCaption
970             cButCLU.OnAction = "RunFile"
980             cButCLU.Parameter = "CLU"
990         Set cButLexis = .CommandBars("List Range Popup").Controls.Add(Temporary:=True)
1000            cButLexis.Caption = "Lexis"
1010            cButLexis.Style = msoButtonCaption
1020            cButLexis.OnAction = "RunFile"
1030            cButLexis.Parameter = "Lexis"
1040        End With
             
          
1050      On Error GoTo 0
1060  Exit Sub

Worksheet_Activate_Error:

         
1070      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1080      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Worksheet_Activate Within: InvestigationLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1090      Print #1, zMsg

1100      Close #1

            
1110      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
1120  On Error GoTo Worksheet_BeforeDoubleClick_Error
      Dim zMsg As String
      Dim TheLastCaseRow As Long
      Dim result As Integer

      Dim lRow As Long


1130  If Target.Column = 5 Then
1140      lRow = Target.row
1150      Cancel = True
1160      result = MsgBox("Would you like to email attorney " & Cells(lRow, 5).Value & " for the " & Cells(lRow, 3).Value & " case?", vbYesNo, Cells(lRow, 3).Value & " " & Cells(lRow, 1).Value)
1170      Application.DisplayAlerts = False
1180      If result = vbYes Then
1190          EmailAttorney Cells(lRow, 1).Value, Cells(lRow, 3).Value, RevAttyName(Cells(lRow, 5).Value)
          
          'Update Case Logs
1200          TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
1210          TheLastCaseRow = TheLastCaseRow + 1
                     
1220         CaseLogs.Cells(TheLastCaseRow, 1).Value = CStr(Cells(lRow, 1).Value)
1230         CaseLogs.Cells(TheLastCaseRow, 2).Value = GetADate(Now()) 'Format(txtDOInt, "m/d/yy")
1240         CaseLogs.Cells(TheLastCaseRow, 3).Value = Format(Now(), "h:mm AMPM")
1250         CaseLogs.Cells(TheLastCaseRow, 4).Value = "Emailed attorney."
             
1260      End If
          

1270  Else
1280      If Target.Column = 8 Then
1290          lRow = Target.row
1300          Cancel = True
1310          zMsg = Cells(ActiveCell.row, 23).Value + Cells(ActiveCell.row, 24).Value & " of " & Cells(ActiveCell.row, 31).Value & " Interviews Completed." & vbNewLine
1320          zMsg = zMsg & Cells(ActiveCell.row, 26).Value & " of " & Cells(ActiveCell.row, 32).Value & " Photos Completed." & vbNewLine
1330          zMsg = zMsg & Cells(ActiveCell.row, 27).Value + Cells(ActiveCell.row, 28).Value & " of " & Cells(ActiveCell.row, 33).Value & " Subs Served." & vbNewLine
1340          zMsg = zMsg & Cells(ActiveCell.row, 25).Value & "  Due Diligence(s) written." & vbNewLine
1350          zMsg = zMsg & Cells(ActiveCell.row, 22).Value & " of " & Cells(ActiveCell.row, 34).Value & " Other Completed."
1360          result = MsgBox(zMsg, vbOKOnly, "CaseStatus")
1370      End If
1380      Application.DisplayAlerts = True
1390  End If


1400      On Error GoTo 0
1410      Exit Sub

Worksheet_BeforeDoubleClick_Error:
1420      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1
           
1430      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Worksheet_BeforeDoubleClick within: Sub - InvestigationLog " & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1440      Print #1, zMsg

1450      Close #1

            
1460      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
          
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)

1470  On Error GoTo Worksheet_Change_Error
      Dim zMsg As String

      Dim result As Integer
1480  Application.EnableEvents = False
1490  If Target.Column = 1 Then
1500    result = MsgBox("You need to use the 'Change case or Client info' button.", vbCritical, "Need to change the case number?")
1510    Application.Undo
             
1520  End If
1530  If Target.Column = 3 Then
1540     result = MsgBox("You need to use the 'Change case or Client info' button.", vbCritical, "Need to change the client's name?")
1550     Application.Undo
             
1560  End If
1570  Application.EnableEvents = True

1580      On Error GoTo 0
1590  Exit Sub

Worksheet_Change_Error:

         
1600      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1610      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Worksheet_Change Within: InvestigationLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1620      Print #1, zMsg

1630      Close #1

            
1640      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdAction_Click()
             
          
1650  On Error GoTo cmdAction_Click_Error
      Dim zMsg As String

1660  ResetButton cmdAction
1670  frmOpenCase.Show
         

1680      On Error GoTo 0
1690  Exit Sub

cmdAction_Click_Error:

         
1700      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1710      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdAction_Click Within: InvestigationLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1720      Print #1, zMsg

1730      Close #1

            
1740      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Private Sub cmdHide_Click()

         
1750  On Error GoTo cmdHide_Click_Error
      Dim zMsg As String
1760  ResetButton cmdHide
1770  SortByDueDate

      'Dim LastRow As Long
      'Dim cell As Range
      'Application.ScreenUpdating = False
      'Dim bResult As Boolean

      'LastRow = ActiveSheet.UsedRange.Rows.Count
      'If cmdHide.Caption = "Hide Old" Then
      '    bResult = True
      'Else
      '    bResult = False
      'End If

      'For Each cell In Range("g3:g" & LastRow)

      'If CDate(cell.Value) <= Now() - 90 And CDate(cell.Value) <> 0 Then
      '    ActiveSheet.Unprotect "darryl"
      '    cell.EntireRow.Hidden = bResult
      '    ActiveSheet.Protect "darryl"
      'End If
      'Next
      'If bResult = False Then
      '    cmdHide.Caption = "Hide Old"
      'Else
      '    cmdHide.Caption = "Reveal Old"
      'End If
      'Application.ScreenUpdating = True

1780      On Error GoTo 0
1790  Exit Sub

cmdHide_Click_Error:

         
1800      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1810      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdHide_Click Within: InvestigationLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1820      Print #1, zMsg

1830      Close #1

            
1840      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub cmdReport_Click()
          
1850  On Error GoTo cmdReport_Click_Error
      Dim zMsg As String
1860  ResetButton cmdReport
      Dim lngRow As Long
1870  lngRow = ActiveCell.row
1880  If InvestigationLog.Cells(lngRow, 1).Value = "" Then
1890          MsgBox ("Pick a case before you create a report")
1900          Exit Sub
1910  Else
            
1920     Cells(lngRow, 3).Activate
          
1930     frmWitnessEntry.Caption = "Create Report - Client: " & ActiveCell.Value
1940     frmWitnessEntry.Show
1950  End If


1960      On Error GoTo 0
1970  Exit Sub

cmdReport_Click_Error:

         
1980      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1990      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdReport_Click Within: InvestigationLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2000      Print #1, zMsg

2010      Close #1

            
2020      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
 
    


Private Sub cmdPathChanger_Click()


2030  On Error GoTo cmdPathChanger_Click_Error
      Dim zMsg As String
2040  ResetButton cmdPathChanger
2050  frmNewUser.Show
        
              

2060      On Error GoTo 0
2070  Exit Sub

cmdPathChanger_Click_Error:

         
2080      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2090      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdPathChanger_Click Within: InvestigationLog" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2100      Print #1, zMsg

2110      Close #1

            
2120      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub Worksheet_Deactivate()
          
2130      ResetMenu

End Sub




