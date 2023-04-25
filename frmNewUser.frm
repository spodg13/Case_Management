VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewUser 
   Caption         =   "Welcome - Information Needed"
   ClientHeight    =   8016
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6525
   OleObjectBlob   =   "frmNewUser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' File   : frmNewUser
' Author : gouldd
' Date   : 9/1/2016
' Purpose: First time users will see this screen and enter their data specific to them
'---------------------------------------------------------------------------------------

Option Explicit





'---------------------------------------------------------------------------------------
' Method : cmdNext_Click
' Author : gouldd
' Date   : 9/2/2016
' Purpose: Write initial user information, verify root path exists
'---------------------------------------------------------------------------------------
Private Sub cmdNext_Click()

10    On Error GoTo cmdNext_Click_Error
      Dim zMsg As String

      Dim result As Integer
      Dim lngRow As Long
      Dim intA As Integer
      Dim intInitS As Integer
      Dim strTemplatePath As String
      Dim strActionPath As String
      Dim strReportsPath As String
      Dim strCheckFolder As Variant
      Dim strRootFilePath As String
      Dim strTemplateFileName As String
      Dim strUserPath As String
      Dim strInstallationPath As String

       
      '----------------------------------------------------------------------------------------
      'Directory Structure in each users folder.  Example using initials "DG" in Investigations folder, and last of Gould for Photos
      '   W:\Investigations\DG                Stores the main ICMS Excel program
      '   W:\Investigations\DG\Reports\       Stores all reports
      '   W:\Investigations\DG\Action_Logs\   Stores the action logs
      '   W:\Investigations\DG\Mileage_Logs\  Stores the mileage logs
      '   W:\Investigations\DG\IOD_Logs\      Stores the IOD Logs
      '   W:\Investigations\DG\Templates\     Stores the templates
      '   S:\Investigations\PHOTOS\Photos\Gould
      '   Root path is currently W:\Investigations, Media Drive is S:\
      '   Application.ActiveWorkbook.FullName or Application.ActiveWorkbook.Path
      '   Verify Root directory exists and has not changed.
      '   Checking for W:\Investigations

      'strRootFilePath verified at form initialization
20    strRootFilePath = Files.Cells(33, 2).Value


      'Verify Initials have been entered correctly and the corresponding folder exists.
      'Looking for W:\Investigations\DG  **Initials changed to user

       

30    If FolderExists(strRootFilePath) = False Then
40        result = MsgBox("Your individual folder is missing or your are using the wrong initials(" & Trim(txtInitials) & ")" & vbLf & vbLf & "Press OK to review your initials or Cancel and request IT verify your folder exists", vbOKCancel, "Missing Folder")
50        If result = vbCancel Then
60            MsgBox "Seek assistance from the IT Department.", vbOKOnly
70            Unload Me
80            ActiveWorkbook.Close (False)  'Close workbook without further changes
90            Exit Sub
100       End If
110           txtInitials.SetFocus
120           Exit Sub
130   End If
140   strUserPath = Files.Cells(36, 2).Value
150   strInstallationPath = Files.Cells(38, 2).Value
160   strCheckFolder = Array("\Templates", "\IOD_Logs", "\Mileage_Logs", "\Reports", "\Action_Logs")


      'Full installation section
170   If Files.Cells(37, 2) = False Then    'Not installed yet, copy files and directory

180       For intA = 0 To UBound(strCheckFolder)
          
190       If FolderExists(strUserPath & strCheckFolder(intA)) = False Then
200           CopyInitFiles strUserPath & strCheckFolder(intA), strInstallationPath & strCheckFolder(intA)
210       End If
220           Select Case intA
                  Case 0
230                   strTemplatePath = strUserPath & strCheckFolder(intA) & "\"
240                   Files.Cells(34, 2).Value = strTemplatePath
250               Case 1  'Establish folders in spreadsheet - IOD Logs
260                   Files.Cells(8, 2).Value = strUserPath & strCheckFolder(intA) & "\"
270               Case 2  'Mileage Logs
280                   Files.Cells(14, 2).Value = strUserPath & strCheckFolder(intA) & "\"
290               Case 3  'Reports
300                   Files.Cells(1, 2).Value = strUserPath & strCheckFolder(intA) & "\"
310               Case 4  'Action Logs
320                   Files.Cells(6, 2).Value = strUserPath & strCheckFolder(intA) & "\"
330           End Select
          
340       Next intA
          
350       Files.Cells(2, 2).Value = strTemplatePath & "InvestigativeReport.dotx"
360       Files.Cells(3, 2).Value = strTemplatePath & "PhotoReport.dotx"
370       Files.Cells(4, 2).Value = strTemplatePath & "DueDiligenceReport.dotx"
380       Files.Cells(5, 2).Value = strTemplatePath & "ClosureForm.dotx"
390       Files.Cells(7, 2).Value = strTemplatePath & "ActionLog.dotx"
400       Files.Cells(9, 2).Value = strTemplatePath & "IODLog.dotx"
410       Files.Cells(10, 2).Value = strTemplatePath & "CR_125.dotx"
420       Files.Cells(11, 2).Value = strTemplatePath & "ContactLetter.dotx"
430       Files.Cells(12, 2).Value = strTemplatePath & "InvestigativeReportJuv.dotx"
440       Files.Cells(13, 2).Value = strTemplatePath & "FaxCover.dotx"
450       Files.Cells(15, 2).Value = strTemplatePath & "MileageLog.dotx"
460       Files.Cells(32, 2).Value = strTemplatePath & "ContactLetterSubpoena.dotx"
          
          'Look for Photo Folder
          
470       If FolderExists("S:\Investigations\PHOTOS\Photos\" & ParseOutNames(txtName, 3)) = False Then
480           Files.Cells(29, 2).Value = PathPicked("Photo Folder")
490       Else
500           Files.Cells(29, 2).Value = "S:\Investigations\PHOTOS\Photos\" & ParseOutNames(txtName, 3) & "\"
510       End If
520       Files.Cells(37, 2) = True ' Installed
         
530   End If
      'End Full installation

      'Look for any missing files or folders

540   For intA = 1 To 35
                                 
550              If InStr(1, Files.Cells(intA, 1).Value, "Folder") > 0 And FolderExists(Files.Cells(intA, 2).Value) = False Then
560                   result = MsgBox("Please select the " & Files.Cells(intA, 1).Value & "!", vbOKCancel, "Need the file or path")
570                   If result = vbCancel Then GoTo Nexti
580                   strTemplatePath = PathPicked(Files.Cells(intA, 1).Value)
590                   Files.Cells(intA, 2).Value = strTemplatePath
600              End If
                 
                 
610              If InStr(1, Files.Cells(intA, 1).Value, "Template") > 0 And TemplateExists(Files.Cells(intA, 2).Value) = False Then
620                   result = MsgBox("Please select the " & Files.Cells(intA, 1).Value & "!", vbOKCancel, "Need the file or path")
630                   If result = vbCancel Then GoTo Nexti
640                   strTemplateFileName = FilePicked(Files.Cells(intA, 1).Value)
650                   Files.Cells(intA, 2).Value = strTemplateFileName
660              End If
           
Nexti:
670        Next intA

      'Write values to Files worksheet

680   Files.Cells(20, 2).Value = txtName
690   Files.Cells(16, 2).Value = Trim(txtInitials)
700   Files.Cells(18, 2).Value = ckAutoAction
710   Files.Cells(21, 2).Value = txtLicense
720   Files.Cells(22, 2).Value = txtVehID
730   Files.Cells(23, 2).Value = txtOfficePhone
740   Files.Cells(24, 2).Value = txtCellPhone
750   Files.Cells(19, 2).Value = txtMileageEmail
760   Files.Cells(25, 2).Value = txtEmail
770   Files.Cells(26, 2).Value = txtInvTitle
780   Files.Cells(27, 2).Value = txtInvFax
790   Files.Cells(28, 2).Value = ckJuvHall
800   Files.Cells(30, 2).Value = ckAutoClosure


810   If txtDayOff > Now() Then
820       txtDayOff = DateAdd("d", -14, txtDayOff)
830   End If
840   Files.Cells(17, 2).Value = txtDayOff


850   result = MsgBox("Do you want to update where your files are located or named?", vbYesNo, "Update information")

860   If result = vbYes Then
870       frmTemplateFiles.Show
880   End If
          

890   ActiveWorkbook.Save
900   PopulateCombo
910   SortAttorneys
920   ActiveWorkbook.Save

930   Unload frmNewUser


       

940       On Error GoTo 0
950   Exit Sub

cmdNext_Click_Error:

         
960       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

970       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdNext_Click Within: frmNewUser" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

980       Print #1, zMsg

990       Close #1

            
1000      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdSupervisor_Click()

1010  On Error GoTo cmdSupervisor_Click_Error
      Dim zMsg As String
      Dim strFile As String
      Dim result As Integer

      Dim strAny As String
1020  strFile = FilePicked(strAny)
1030  If IsWorkBookOpen(strFile) = True Then
1040      MsgBox ("Cannot re-assign until case is closed.")
1050      Application.DisplayAlerts = False
1060      Workbooks.Open FileName:=strFile, ReadOnly:=True
1070      UserChange ("Procedure: cmdSupervisor_Click Within: frmNewUser" & ActiveWorkbook.Path & "ReadOnly = True-open workbook")
1080  Else
1090      result = MsgBox("Would you like to Close and re-assign?", vbYesNoCancel, "Read Only")
1100      If result = vbCancel Then Exit Sub
1110      End If
1120      If result = vbYes Then
1130          Workbooks.Open FileName:=strFile, ReadOnly:=False
1140          UserChange ("Procedure: cmdSupervisor_Click Within: frmNewUser" & ActiveWorkbook.Path & "ReadOnly = False")
1150      Else
1160          Workbooks.Open FileName:=strFile, ReadOnly:=True
1170          UserChange ("Procedure: cmdSupervisor_Click Within: frmNewUser" & ActiveWorkbook.Path & "ReadOnly = True")
1180      End If
1190  Application.DisplayAlerts = True
1200  Unload frmNewUser

1210      On Error GoTo 0
1220  Exit Sub

cmdSupervisor_Click_Error:

         
1230      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1240      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdSupervisor_Click Within: frmNewUser" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1250      Print #1, zMsg

1260      Close #1

            
1270      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtCellPhone_Exit(ByVal Cancel As MSForms.ReturnBoolean)

1280  On Error GoTo txtCellPhone_Exit_Error
      Dim zMsg As String

1290      txtCellPhone.Text = Replace(txtCellPhone, "-", "")
1300      txtCellPhone.Text = Replace(txtCellPhone, " ", "")
1310      txtCellPhone.Text = Format(txtCellPhone.Text, "(000) 000-0000")

1320      On Error GoTo 0
1330  Exit Sub

txtCellPhone_Exit_Error:

         
1340      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1350      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtCellPhone_Exit Within: frmNewUser" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1360      Print #1, zMsg

1370      Close #1

            
1380      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDayOff_AfterUpdate()

1390      txtDayOff = GetADate(txtDayOff)
End Sub



Private Sub txtDayOff_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

1400  On Error GoTo txtDayOff_DblClick_Error
      Dim zMsg As String
1410  DatePickerForm.Caption = "Day Off"
1420  DatePickerForm.Show vbModal
1430      [DatePickerForm]![CallingTextBox].Caption = "txtDayOff"
          '[DatePickerForm]![CallingForm].Caption = "frmCloseCase"
1440      txtDayOff = [DatePickerForm]![CallingForm].Caption
1450      If txtDayOff = "Form" Then txtDayOff = Format(DateValue(Now()), "MMMM d, yyyy")
1460      Cancel = True

1470      On Error GoTo 0
1480  Exit Sub

txtDayOff_DblClick_Error:

         
1490      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1500      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDayOff_DblClick Within: frmEnterAction" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1510      Print #1, zMsg

1520      Close #1

            
1530      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"


End Sub

Private Sub txtDayOff_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

1540  On Error GoTo txtDayOff_KeyPress_Error
      Dim zMsg As String

1550  If KeyAscii = 43 Then
1560      txtDayOff = DateAdd("d", 1, txtDayOff)
1570      KeyAscii = 0
1580      txtDayOff = Format(DateValue(txtDayOff), "m/d/yy")
1590  End If
1600  If KeyAscii = 45 Then
1610      txtDayOff = DateAdd("d", -1, txtDayOff)
1620      KeyAscii = 0
1630      txtDayOff = Format(DateValue(txtDayOff), "m/d/yy")
1640  End If

1650      On Error GoTo 0
1660  Exit Sub

txtDayOff_KeyPress_Error:

         
1670      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1680      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDayOff_KeyPress Within: frmNewUser" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1690      Print #1, zMsg

1700      Close #1

            
1710      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtInvFax_Exit(ByVal Cancel As MSForms.ReturnBoolean)

1720  On Error GoTo txtInvFax_Exit_Error
      Dim zMsg As String

1730      txtInvFax.Text = Replace(txtInvFax, "-", "")
1740      txtInvFax.Text = Replace(txtInvFax, " ", "")
1750      txtInvFax.Text = Format(txtInvFax.Text, "(000) 000-0000")

1760      On Error GoTo 0
1770  Exit Sub

txtInvFax_Exit_Error:

         
1780      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1790      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtInvFax_Exit Within: frmNewUser" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1800      Print #1, zMsg

1810      Close #1

            
1820      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub txtInvTitle_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
1830  On Error GoTo txtInvTitle_Change_Error
      Dim zMsg As String

1840  If InStr(1, txtInvTitle, "Super") > 0 Or InStr(1, txtInvTitle, "Chief") > 0 Then
1850          ActiveWorkbook.SaveAs ActiveWorkbook.Path & "\Investigation Log-" & txtInitials
1860          MsgBox "Please use file Investigation Log-" & txtInitials & " from now on."
1870      End If

1880      On Error GoTo 0
1890  Exit Sub

txtInvTitle_Change_Error:

         
1900      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1910      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtInvTitle_Change Within: frmNewUser" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1920      Print #1, zMsg

1930      Close #1

            
1940      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
End Sub

Private Sub txtLicense_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
1950  Select Case KeyAscii
          
          Case 48 To 57  ' Numbers
1960          KeyAscii = KeyAscii
1970      Case 65 To 90 ' Cap letters
1980          KeyAscii = KeyAscii
1990      Case 97 To 122  'Lowercase letters : convert to upper case
2000          KeyAscii = KeyAscii - 32
2010      Case Else
2020          KeyAscii = 0
2030  End Select

End Sub

Private Sub txtName_AfterUpdate()

2040  On Error GoTo txtName_AfterUpdate_Error
      Dim zMsg As String

      Dim intMidl, length As Integer
2050  intMidl = InStrRev(txtName, " ")
2060  length = Len(txtName)

2070  txtEmail = Right(txtName, length - intMidl) & Left(txtName, 1) & "@saccounty.net"
2080  If Files.Cells(39, 2).Value = "" Then
2090        Files.Cells(39, 2).Value = Environ("UserName")
2100   End If

2110      On Error GoTo 0
2120  Exit Sub

txtName_AfterUpdate_Error:

         
2130      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2140      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtName_AfterUpdate Within: frmNewUser" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2150      Print #1, zMsg

2160      Close #1

            
2170      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub txtName_Change()


2180  On Error GoTo txtName_Change_Error
      Dim zMsg As String
2190      Files.Cells(39, 2).Value = Environ("UserName")


2200      On Error GoTo 0
2210  Exit Sub

txtName_Change_Error:

         
2220      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2230      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtName_Change Within: frmNewUser" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2240      Print #1, zMsg

2250      Close #1

            
2260      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

2270  On Error GoTo txtName_KeyDown_Error
      Dim zMsg As String

      '****  Add escape to New User form by hitting Ctrl E when in the name box
2280  If KeyCode = vbKeyE And Shift = 2 Then
2290      Unload frmNewUser
2300      End
2310  End If
2320  If KeyCode = vbKeyU And Shift = 2 Then
2330     If cmdSupervisor.Visible = False Then
2340            cmdSupervisor.Visible = True
2350     Else
2360            cmdSupervisor.Visible = False
2370     End If
2380  End If
2390      If KeyCode = vbKeyT And Shift = 2 Then
2400        MsgBox ("Interviews - " & Application.WorksheetFunction.Sum(Columns("W:W")) + Application.WorksheetFunction.Sum(Columns("X:X")))
2410        MsgBox ("Subs - " & Application.WorksheetFunction.Sum(Columns("AA:AA")) + Application.WorksheetFunction.Sum(Columns("AB:AB")))
2420        KeyCode = 0
2430      End If
2440      On Error GoTo 0
2450  Exit Sub

txtName_KeyDown_Error:

         
2460      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2470      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtName_KeyDown Within: frmNewUser" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2480      Print #1, zMsg

2490      Close #1

            
2500      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub
Private Sub txtOfficePhone_Exit(ByVal Cancel As MSForms.ReturnBoolean)

2510  On Error GoTo txtOfficePhone_Exit_Error
      Dim zMsg As String

2520      txtOfficePhone.Text = Replace(txtOfficePhone, "-", "")
2530      txtOfficePhone.Text = Replace(txtOfficePhone, " ", "")
2540      txtOfficePhone.Text = Format(txtOfficePhone.Text, "(000) 000-0000")

2550      On Error GoTo 0
2560  Exit Sub

txtOfficePhone_Exit_Error:

         
2570      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

2580      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtOfficePhone_Exit Within: frmNewUser" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

2590      Print #1, zMsg

2600      Close #1

            
2610      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub UserForm_Initialize()

2620  On Error GoTo UserForm_Initialize_Error
      Dim zMsg As String

      Dim strRootFilePath As String
      Dim result As Integer
      Dim intInitS As Integer

      '------------------------------------------------------------------------------------
      'Verify Root directory, W:\Investigations is the correct directory.
      'If it is not, then pick the new root and write to spreadsheet
      'If cancel, close workbook without saving
      '------------------------------------------------------------------------------------
2630  CenterForm Me

2640  cmdSupervisor.Visible = False

2650  strRootFilePath = Files.Cells(33, 2).Value

2660  If FolderExists(strRootFilePath) = False Then
2670      result = MsgBox("The Investigations Folder is missing, or has changed." & vbLf & "Please select the Investigations main folder, ex: 'W:\Investigations\'" & vbLf & vbLf & "Your individual folder will be added later.", vbOKCancel, "Missing Folder")
2680      If result = vbCancel Then
2690          MsgBox "Seek assistance from the IT Department.", vbOKOnly
2700          ActiveWorkbook.Close (False)  'Close workbook without saving changes
2710          Exit Sub
2720      End If
2730      strRootFilePath = PathPicked("Root Investigation Folder - ie 'W:\Investigations'")
2740      Files.Cells(33, 2).Value = strRootFilePath
2750      ActiveWorkbook.Save
2760  End If

      '--------------------------------------------------------------------
      'Pre-populate fields
      '--------------------------------------------------------------------

2770      txtName = Files.Cells(20, 2).Value
2780      intInitS = InStrRev(ActiveWorkbook.Path, "\")
          
          'Set user path
2790      If Files.Cells(37, 2) = False Then
2800          Files.Cells(36, 2).Value = ActiveWorkbook.Path  'Sets user path
2810      End If
          
          
          
          'Excel Workbooked is saved in the the Investigations\UserInitials Folder
2820      If txtName = "New User" Then
2830          txtInitials = Right(ActiveWorkbook.Path, Len(ActiveWorkbook.Path) - intInitS)
2840      Else
2850          txtInitials = Files.Cells(16, 2).Value
2860      End If
          
2870      txtInitials.Locked = False
          'txtInitials = Files.Cells(16, 2).Value ' Alternative method
          
2880      txtLicense = Files.Cells(21, 2).Value
2890      txtVehID = Files.Cells(22, 2).Value
2900      If Files.Cells(17, 2).Value = "" Then
2910          txtDayOff = Now() + 8 - Weekday(Now(), vbFriday)
2920          txtDayOff = Format(txtDayOff, "mm/dd/yy")
2930          Else
2940          txtDayOff = Files.Cells(17, 2).Value
2950      End If
          
2960      txtOfficePhone = Files.Cells(23, 2).Value
2970      txtCellPhone = Files.Cells(24, 2).Value
2980      txtMileageEmail = Files.Cells(19, 2).Value
2990      txtEmail = Files.Cells(25, 2).Value
3000      txtInvTitle = Files.Cells(26, 2).Value
3010      txtInvFax = Files.Cells(27, 2).Value
          
3020       If Files.Cells(28, 2).Value = "" Then
3030          ckJuvHall = False
3040      Else
3050          ckJuvHall = Files.Cells(28, 2).Value
3060      End If
          
3070      If Files.Cells(18, 2).Value = "" Then
3080          ckAutoAction = False
3090      Else
3100          ckAutoAction = Files.Cells(18, 2).Value
3110      End If
3120      If Files.Cells(30, 2).Value = "" Then
3130          ckAutoClosure = False
3140      Else
3150          ckAutoClosure = Files.Cells(30, 2).Value
3160      End If
          
3170      If InStr(1, Files.Cells(26, 2).Value, "Super") > 0 Or InStr(1, Files.Cells(26, 2).Value, "Chief") > 0 Then
3180          cmdSupervisor.Visible = True
3190      End If
          

3200      On Error GoTo 0
3210  Exit Sub

UserForm_Initialize_Error:

         
3220      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3230      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserForm_Initialize Within: frmNewUser" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3240      Print #1, zMsg

3250      Close #1

            
3260      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

