VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEnterAction 
   Caption         =   "Enter Action"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8775
   OleObjectBlob   =   "frmEnterAction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEnterAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'---------------------------------------------------------------------------------------
' File   : frmEnterAction
' Author : gouldd
' Date   : 9/1/2016
' Purpose: Enter an investigative action from either the InvestigationLog sheet or CaseLogs sheet
'---------------------------------------------------------------------------------------



Option Explicit

'---------------------------------------------------------------------------------------
' Method : ckbMileage_Click
' Author : gouldd
' Date   : 9/1/2016
' Purpose: Prepare to add mileage to any action
'---------------------------------------------------------------------------------------
Private Sub ckbMileage_Click()
10        txtStartM.Visible = True
20        txtEndM.Visible = True
30        txtMileageAddress.Visible = True
40        Label5.Visible = True
50        Label6.Visible = True
60        Label7.Visible = True
          'ckbMileage.Visible = False
End Sub



Private Sub cmdCancelAction_Click()

70         Unload frmEnterAction

End Sub



Private Sub cmdEnterAction_Click()

80    On Error GoTo cmdEnterAction_Click_Error
      Dim zMsg As String

      Dim TheLastCaseRow As Long
      Dim result As Integer
      Dim strCaseName As String
      Dim lngRow As Long
      Dim blReActivate As Boolean
90    Application.EnableEvents = False
100   Application.ScreenUpdating = False

      '*******************************Validate Drive****************
110       If fn_validate_directory(Files.Cells(33, 2).Value, False) = False Then
120         MsgBox "Aborting action - drive not accessible "
130       End If
          


140   lngRow = ActiveCell.row

150   If ActiveSheet.CodeName = "InvestigationLog" Then
160       InvestigationLog.Cells(lngRow, 3).Activate
170       strCaseName = InvestigationLog.Cells(lngRow, 3).Value
180       blReActivate = True
190       Else
200           strCaseName = CaseLogs.Cells(lngRow, 6).Value
210           blReActivate = False
220       End If
          
      '**********Check Mileage**************
230   If ckbMileage = True Then
240       If ValidMileage(txtStartM, txtEndM, txtMileageAddress, Me) = False Then
              
250           Exit Sub
260       End If
270   End If

      'Check for blank action entry
280   If txtAction = "" Then
290       MsgBox ("Can't have a blank action.  Cancel or add your action entry.")
300       txtAction.SetFocus
310       Exit Sub
320   End If

330   If txtDuration = "" Then
340       MsgBox ("Can't have a blank duration.  Cancel or add the time your action took.")
350       txtDuration.SetFocus
360       Exit Sub
370   End If
          
          
380       result = MsgBox("Do you want an Action Entry for " & strCaseName & "?", vbYesNoCancel, "Verify Case " & txtCaseNo)
390       If result = vbNo Then
400            result = MsgBox("Click on the case you need the Action Entry for, then press the Enter Action button!", vbOKOnly, "Verify Case")
410            frmEnterAction.Hide
420            Application.EnableEvents = True
430            Exit Sub
440       End If
450       If result = vbCancel Then
460           Unload frmEnterAction
470           Exit Sub
480       End If
          
          
          
          
490       TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
500       TheLastCaseRow = TheLastCaseRow + 1
          
          
510       If ActiveSheet.CodeName = "CaseLogs" Then
520          CaseLogs.Cells(TheLastCaseRow, 1).Value = CaseLogs.Cells(lngRow, 1).Value
530       Else
540           CaseLogs.Cells(TheLastCaseRow, 1).Value = CStr(txtCaseNo)
550       End If
560       CaseLogs.Cells(TheLastCaseRow, 2).Value = GetADate(txtDate) 'Format(txtDOInt, "m/d/yy")
570       CaseLogs.Cells(TheLastCaseRow, 3).Value = Format(txtTime, "h:mm AMPM")
580       CaseLogs.Cells(TheLastCaseRow, 4).Value = txtAction
590       CaseLogs.Cells(TheLastCaseRow, 5).Value = Val(txtDuration)
              
600       If ckbMileage = True Then
610           If Val(txtStartM) > 1 Then
620           Call AddMileage(GetADate(txtDate), txtMileageAddress, txtCaseNo, txtStartM, txtEndM)
630           CaseLogs.Cells(TheLastCaseRow, 7).Value = "Mileage Entry"
640           CaseLogs.Cells(TheLastCaseRow, 8).Value = Val(txtStartM)
650           CaseLogs.Cells(TheLastCaseRow, 9).Value = Val(txtEndM)
660           Else
670               MsgBox ("Mileage not added")
680           End If
690        End If
700        SortCaseLogs
710        If blReActivate = True Then
720           InvestigationLog.Activate
730        End If
740   Application.ScreenUpdating = True
750   Application.EnableEvents = True
760   ActiveWorkbook.Save
770   Unload frmEnterAction


780       On Error GoTo 0
790   Exit Sub

cmdEnterAction_Click_Error:

         
800       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

810       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdEnterAction_Click Within: frmEnterAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

820       Print #1, zMsg

830       Close #1

            
840       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtAction_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim tempstr As String
          
850       Application.ScreenUpdating = False
860         If UCase(txtAction) = "LM" Then
870             tempstr = InputBox("Name ?")
880             txtAction = "I attempted to contact " & ProperCase(tempstr)
890             tempstr = InputBox("Phone?", "Phone number")
900             tempstr = formatTel(tempstr)
910             txtAction = txtAction & " at " & tempstr & ". The call went to a voicemail message.  I left my contact information and the reason for my call."
920             Cancel = True
930         End If
940         If UCase(txtAction) = "LBC" Then
950             tempstr = InputBox("Address ?", "Address")
960             txtAction = "I responded to " & tempstr
970             tempstr = InputBox("Name ?", "Name")
980             txtAction = txtAction & " in an attempt to contact " & ProperCase(tempstr) & ". I knocked and .  I left my business card"
990             Cancel = True
1000        End If
          'Spell check
1010      If Files.Cells(31, 2).Value = True Then
1020          Spelling.Range("A1") = txtAction
1030          Spelling.Range("A1").CheckSpelling
1040          txtAction = Spelling.Range("A1")
1050          Spelling.Range("A1").ClearContents
1060      End If
exitroutine:
1070      Application.ScreenUpdating = True

End Sub
Private Sub txtAction_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
      Dim lngRow As Long

1080  If ActiveSheet.CodeName = "CaseLogs" Then
1090  If KeyCode = vbKeyC And Shift = 2 Then
1100      lngRow = ActiveCell.row
1110      txtAction = ActiveSheet.Cells(lngRow, 4).Value
1120  End If
1130  End If
1140      Select Case KeyCode
              Case vbKeyG And Shift = 2 'Ctrl G
1150           txtAction = Insert(txtAction, "generic carrier ")
                              
1160       Case vbKeyP And Shift = 2 'Ctrl P
1170            txtAction = Insert(txtAction, "personalized ")
                
1180        Case vbKeyN And Shift = 2 'Ctrl N
1190            txtAction = Insert(txtAction, "nobody answered")
                 
                  
1200       Case Else
                'do nothing
1210   End Select

End Sub






Private Sub txtDate_AfterUpdate()
      Dim varDate As Date
1220  If IsDate(txtDate) Then
1230      varDate = DateValue(txtDate)
1240      txtDate = Format(varDate, "MMMM d, yyyy")
1250  Else
1260      MsgBox "Invalid date"
1270      txtDate.SetFocus
1280      Exit Sub
          
1290  End If
End Sub



Private Sub txtDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

1300  On Error GoTo txtDate_DblClick_Error
      Dim zMsg As String

1310  DatePickerForm.Show vbModal
1320      [DatePickerForm]![CallingTextBox].Caption = "txtDate"
          '[DatePickerForm]![CallingForm].Caption = "frmCloseCase"
1330      txtDate = [DatePickerForm]![CallingForm].Caption
1340      If txtDate = "Form" Then txtDate = Format(DateValue(Now()), "MMMM d, yyyy")
1350      Cancel = True

1360      On Error GoTo 0
1370  Exit Sub

txtDate_DblClick_Error:

         
1380      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1390      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_DblClick Within: frmEnterAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1400      Print #1, zMsg

1410      Close #1

            
1420      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)


1430  On Error GoTo txtDate_KeyPress_Error
      Dim zMsg As String

1440  If KeyAscii = 43 Then
1450      txtDate = DateAdd("d", 1, txtDate)
1460      KeyAscii = 0
1470      txtDate = Format(DateValue(txtDate), "m/d/yy")
1480  End If
1490  If KeyAscii = 45 Then
1500      txtDate = DateAdd("d", -1, txtDate)
1510      KeyAscii = 0
1520      txtDate = Format(DateValue(txtDate), "m/d/yy")
1530  End If

1540      On Error GoTo 0
1550  Exit Sub

txtDate_KeyPress_Error:

         
1560      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1570      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDate_KeyPress Within: frmEnterAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1580      Print #1, zMsg

1590      Close #1

            
1600      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub txtEndM_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
1610  Select Case KeyAscii
          Case Is = 46
1620          KeyAscii = KeyAscii
1630      Case 48 To 57
1640          KeyAscii = KeyAscii
1650      Case Else
1660          KeyAscii = 0
1670      End Select
End Sub



Private Sub txtMileageAddress_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1680  On Error GoTo txtMileageAddress_KeyDown_Error
      Dim zMsg As String

1690  Select Case KeyCode
              Case vbKeyS And Shift = 2 'Ctrl S
1700              txtMileageAddress = txtMileageAddress & "Sacramento, CA "
                              
1710          Case vbKeyC And Shift = 2 'Ctrl C
1720              txtMileageAddress = txtMileageAddress & "Citrus Heights, CA "
                  
1730          Case vbKeyC And Shift = 3 'Ctrl Shift C
1740              txtMileageAddress = txtMileageAddress & "Carmichael, CA "
                  
1750          Case vbKeyE And Shift = 2 'Ctrl E
1760              txtMileageAddress = txtMileageAddress & "Elk Grove, CA "

1770          Case vbKeyR And Shift = 2 'Ctrl R
1780              txtMileageAddress = txtMileageAddress & "Rancho Cordova, CA "

1790          Case vbKeyF And Shift = 2 'Ctrl F
1800              txtMileageAddress = txtMileageAddress & "Folsom, CA "
       
1810          Case vbKeyA And Shift = 2 'Ctrl A
1820              txtMileageAddress = txtMileageAddress & "Antelope, CA "

1830          Case vbKeyN And Shift = 2 'Ctrl N
1840              txtMileageAddress = txtMileageAddress & "North Highlands, CA "
                  
1850          Case vbKeyR And Shift = 3 'Ctrl Shift R
1860              txtMileageAddress = txtMileageAddress & "Roseville, CA "

1870          Case vbKeyO And Shift = 2 'Ctrl O
1880              txtMileageAddress = txtMileageAddress & "Orangevale, CA "
                  
1890          Case vbKeyG And Shift = 2 'Ctrl G
1900              txtMileageAddress = txtMileageAddress & "Galt, CA "
                  
1910          Case vbKeyF And Shift = 3 'Ctrl Shift F
1920              txtMileageAddress = txtMileageAddress & "Fair Oaks, CA "
                  
                  
1930          Case Else
                'do nothing
1940      End Select

1950      On Error GoTo 0
1960      Exit Sub

txtMileageAddress_KeyDown_Error:
1970      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1
           
1980      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtMileageAddress_KeyDown within: Sub - frmWitnessEntry " & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1990      Print #1, zMsg

2000      Close #1

            
2010      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"


End Sub

Private Sub txtStartM_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
2020  Select Case KeyAscii
          Case Is = 46
2030          KeyAscii = KeyAscii
2040      Case 48 To 57
2050          KeyAscii = KeyAscii
2060      Case Else
2070          KeyAscii = 0
2080      End Select
          
End Sub

Private Sub txtTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)

2090  On Error GoTo txtTime_Exit_Error
      Dim zMsg As String

      Dim result As Integer
2100  If IsTime(txtTime) = False Then
2110      Cancel = True
2120      txtTime.SetFocus
2130      result = MsgBox("Invalid Time, check your entry!", vbOKOnly, "Check time entry")
2140  End If

2150      On Error GoTo 0
2160  Exit Sub

txtTime_Exit_Error:

         
2170      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2180      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtTime_Exit Within: frmEnterAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2190      Print #1, zMsg

2200      Close #1

            
2210      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtTime_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      Dim LMin, LMinNew, IntervalAdd As Integer

2220  If KeyAscii = 43 Then
2230      LMin = Minute(txtTime)
2240      LMinNew = Round(LMin / 5, 0) * 5
2250      IntervalAdd = (LMin - LMinNew) * -1
2260      txtTime = DateAdd("N", IntervalAdd, txtTime)
          
2270      txtTime = DateAdd("N", 5, txtTime)
2280      KeyAscii = 0
2290      txtTime = Format(txtTime, "h:mm AM/PM")
2300  End If
2310  If KeyAscii = 42 Then
         
2320      txtTime = DateAdd("N", 1, txtTime)
2330      KeyAscii = 0
2340      txtTime = Format(txtTime, "h:mm AM/PM")
2350  End If
2360  If KeyAscii = 45 Then
2370      LMin = Minute(txtTime)
2380      LMinNew = Round(LMin / 5, 0) * 5
2390      IntervalAdd = (LMin - LMinNew) * -1
2400      txtTime = DateAdd("N", IntervalAdd, txtTime)
          
2410      txtTime = DateAdd("N", -5, txtTime)
2420      KeyAscii = 0
2430      txtTime = Format(txtTime, "h:mm AM/PM")
2440  End If
2450  If KeyAscii = 47 Then
            
2460  txtTime = DateAdd("N", -1, txtTime)
2470      KeyAscii = 0
2480      txtTime = Format(txtTime, "h:mm AM/PM")
2490  End If
End Sub

Private Sub UserForm_Activate()
2500  If IsInternetConnected = False Then
2510       MsgBox "No Network Connection Detected! Shut down and re-start", vbExclamation, "No Connection"
2520      Exit Sub
2530  End If


2540  On Error GoTo UserForm_Activate_Error
      Dim zMsg As String
2550      CenterForm Me
      Dim lngRow As Long
2560  lngRow = ActiveCell.row
2570  txtCaseNo = ActiveSheet.Cells(lngRow, 1).Value
2580  frmEnterAction.Caption = "Action for " & InvestigationLog.Cells(lngRow, 3).Value & " " & txtCaseNo


2590      On Error GoTo 0
2600  Exit Sub

UserForm_Activate_Error:

         
2610      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2620      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserForm_Activate Within: frmEnterAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2630      Print #1, zMsg

2640      Close #1

            
2650      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub UserForm_Initialize()

2660  On Error GoTo UserForm_Initialize_Error
      Dim zMsg As String

      Dim lngRow As Long

2670  lngRow = ActiveCell.row

          'txtCaseNo = InvestigationLog.Cells(lngRow, 1).Value
2680      txtCaseNo = ActiveSheet.Cells(lngRow, 1).Value
              
2690      txtTime = Format(Now(), "h:mm AM/PM")
2700      txtDate = Format(Now(), "mm/dd/yy")
2710      frmEnterAction.Caption = "Action for " & InvestigationLog.Cells(lngRow, 3).Value & " " & txtCaseNo
          
2720      ckbMileage.Visible = True
2730      ckbMileage = False
2740      txtStartM.Visible = False
2750      txtEndM.Visible = False
2760      txtMileageAddress.Visible = False
2770      Label5.Visible = False
2780      Label6.Visible = False
2790      Label7.Visible = False
          

2800      On Error GoTo 0
2810  Exit Sub

UserForm_Initialize_Error:

         
2820      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2830      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserForm_Initialize Within: frmEnterAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2840      Print #1, zMsg

2850      Close #1

            
2860      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

