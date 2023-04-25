VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWitnessEntry 
   ClientHeight    =   8496
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10065
   OleObjectBlob   =   "frmWitnessEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWitnessEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'---------------------------------------------------------------------------------------
' File   : frmWitnessEntry
' Author : gouldd
' Date   : 9/1/2016
' Purpose: A multi-use form to generate witness reports, contact letters, photo reports, fax covers, subpoena service, due diligence reports and misc. reports
'---------------------------------------------------------------------------------------

Option Explicit


Private Sub ckAddMileage_Click()

10    On Error GoTo ckAddMileage_Click_Error
      Dim zMsg As String

20    If ckAddMileage = True Then
30    If optSubpoena = True Then
40        txtMileageAddress = txtLOS
50    End If
60    If optWitnessReport = True Then
70        txtMileageAddress = txtLocation
80    End If
90    Else
100       txtMileageAddress = ""
110   End If


120       On Error GoTo 0
130   Exit Sub

ckAddMileage_Click_Error:

         
140       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

150       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: ckAddMileage_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

160       Print #1, zMsg

170       Close #1

            
180       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub ckbxSubLetter_AfterUpdate()

190   On Error GoTo ckbxSubLetter_AfterUpdate_Error
      Dim zMsg As String

      Dim lngRow As Long

200   lngRow = ActiveCell.row

210       If ckbxSubLetter = True Then
220           txtDateofApp.Visible = True
230           txtTimeofApp.Visible = True
240           txtDept.Visible = True
250           Label29.Visible = True
260           Label30.Visible = True
270           Label31.Visible = True
280           txtDateofApp = Format(Now(), "MMMM d, yyyy")
290           txtTimeofApp = "8:45 AM"
300           txtDept = InvestigationLog.Cells(lngRow, 11).Value
310       Else
320           txtDateofApp.Visible = False
330           txtTimeofApp.Visible = False
340           txtDept.Visible = False
350           Label29.Visible = False
360           Label30.Visible = False
370           Label31.Visible = False
380       End If

390       On Error GoTo 0
400   Exit Sub

ckbxSubLetter_AfterUpdate_Error:

         
410       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

420       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: ckbxSubLetter_AfterUpdate Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

430       Print #1, zMsg

440       Close #1

            
450       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub cmdCancel_Click()
460       Unload frmWitnessEntry
              
End Sub

Private Sub cmdCancelContact_Click()
470   Unload frmWitnessEntry
End Sub

Private Sub cmdCancelCover_Click()
480   Unload frmWitnessEntry
End Sub

Private Sub cmdCancelSub_Click()
490   Unload frmWitnessEntry

End Sub

Private Sub cmdContactLetter_Click()

500   On Error GoTo cmdContactLetter_Click_Error
      Dim zMsg As String

510       Application.ScreenUpdating = False
          Dim wdApp As New Word.Application
          Dim wdDoc As Word.Document
          Dim CCtrl As Word.ContentControl
          Dim strPath, strTemplateFileName, strActionEntry, strFileName As String
          Dim intLen, intCommaPos, intVersion As Integer
          Dim strCurrentTime, strLetterType As String
          Dim strEnvAddr As String
          
          Dim shl As Object
520       Set shl = CreateObject("wscript.shell")
          
          Dim Wksht As Worksheet, lngRow As Long, intCol As Integer
          Dim result As Integer
          Dim strAddress, strClientFirst, strClientLast, strClient As String
          Dim TheLastCaseRow As Long
         
530       If ckbxSubLetter = True Then
540           strLetterType = "Subpoena Contact Letter"
550           Else
560           strLetterType = "Contact Letter"
570       End If
              
580       strCurrentTime = Format(Now(), "h:mm AMPM")
590       Application.EnableEvents = False
600       Set Wksht = ActiveSheet
610       lngRow = ActiveCell.row
620       intCol = 3 ' Place on client name
630       Cells(lngRow, intCol).Activate

640       strAddress = txtWitAddress & ", " & txtWitCity & ", " & txtWitState & " " & txtWitZip
650       strActionEntry = "Generated and mailed " & strLetterType & " to " & txtWitFirst & " " & txtWitLast & " at " & strAddress
          
          
660       result = MsgBox("Do you want to create a " & strLetterType & " for " & ActiveCell.Value & "?", vbYesNoCancel, "Verify Case")
670       If result = vbNo Then
680            result = MsgBox("Click on the case you need the Investigative Report for, then press the command button!", vbOKOnly, "Verify Case")
690            frmWitnessEntry.Hide
700            Application.EnableEvents = True
710            Exit Sub
720       End If
730       If result = vbCancel Then
740           Unload frmWitnessEntry
750           Exit Sub
760       End If
         
770       TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
780       TheLastCaseRow = TheLastCaseRow + 1
                  
790       strFileName = ActiveCell.Value
800       strFileName = Replace(strFileName, ", ", "_")
810       strFileName = strFileName & "_" & ActiveCell.Offset(0, -2).Value
820       strFileName = strFileName & "_" & txtWitFirst & "_" & txtWitLast & "_" & strLetterType
          'If txtVersion <> "" Then strFileName = strFileName & "_" & txtVersion
          
830       strClient = Cells(lngRow, 3).Value
              
840       intCommaPos = InStr(strClient, ",")
850       intLen = Len(strClient)
860       strClientLast = Left$(strClient, intCommaPos - 1)
870       strClientFirst = Right$(strClient, intLen - intCommaPos - 1)
          
880       strPath = Files.Cells(1, 2).Value
890       If ckbxSubLetter = False Then
900           strTemplateFileName = Files.Cells(11, 2).Value
910           Else
920           strTemplateFileName = Files.Cells(32, 2).Value
930       End If
940       If strTemplateFileName = "" Then
950           result = MsgBox("Please select the Contact Letter Template Location!", vbCritical, "Need the file name")
960           strTemplateFileName = FilePicked("Contact Letter Template")
970           Files.Cells(11, 2).Value = strTemplateFileName
980       End If
990       If strPath = "" Then
1000          result = MsgBox("Please select the path where the Contact Letter will be stored!", vbCritical, "Need the path")
1010          strPath = PathPicked("Contact Letter") & "\"
1020           Files.Cells(1, 2).Value = strPath
1030      End If
          
1040     intVersion = GetVersion(strPath, strFileName)
1050         If intVersion > 1 Then
1060          strFileName = strFileName & "_" & CStr(intVersion)
1070      End If
           
1080      UpdateCaseLog strActionEntry, Format(Now(), "h:mm AMPM"), GetADate(Now()), TheLastCaseRow, lngRow, 1, Val(txtDurationLetter)
          
          
1090      Set wdDoc = wdApp.Documents.Open(FileName:=strTemplateFileName, AddToRecentFiles:=False, Visible:=False)
1100          With wdDoc
                  
1110              For Each CCtrl In .ContentControls
1120                  Select Case CCtrl.Title
                          
                          Case "Client"
1130                          CCtrl.Range.Text = strClientFirst & " " & strClientLast
1140                      Case "WitSalutation"
1150                          CCtrl.Range.Text = txtSalutation
1160                      Case "Publish Date"
1170                          CCtrl.Range.Text = Now()
1180                      Case "WitFirst"
1190                          CCtrl.Range.Text = txtWitFirst
1200                      Case "WitLast"
1210                          CCtrl.Range.Text = txtWitLast
1220                      Case "WitAddress"
1230                          CCtrl.Range.Text = txtWitAddress
1240                      Case "WitCity"
1250                          CCtrl.Range.Text = txtWitCity
1260                      Case "WitState"
1270                          CCtrl.Range.Text = txtWitState
1280                      Case "WitZip"
1290                          CCtrl.Range.Text = txtWitZip
1300                      Case "CourtTime"
1310                          CCtrl.Range.Text = txtTimeofApp
1320                      Case "CourtDate"
1330                          CCtrl.Range.Text = txtDateofApp
1340                      Case "CourtDept"
1350                          CCtrl.Range.Text = txtDept
1360                      Case "InvName"
1370                          CCtrl.Range.Text = Files.Cells(20, 2).Value
1380                      Case "InvPhone"
1390                          CCtrl.Range.Text = Files.Cells(23, 2).Value
1400                      Case "InvCell"
1410                          CCtrl.Range.Text = Files.Cells(24, 2).Value
1420                      Case "InvTitle"
1430                          CCtrl.Range.Text = Files.Cells(26, 2).Value
1440                  End Select
                      
1450              Next
                              
1460          End With
           
          
1470      Application.ScreenUpdating = True
1480      Application.EnableEvents = True
1490      ActiveWorkbook.Save
1500      Unload frmWitnessEntry
          
          'Save with new name
1510      wdDoc.SaveAs FileName:=strPath & strFileName & ".docx", FileFormat:=wdFormatDocumentDefault
1520      wdApp.Visible = True
1530      wdApp.Activate
1540          strEnvAddr = txtWitFirst & " " & txtWitLast & vbCr & txtWitAddress & vbCr & txtWitCity & ", " & txtWitState & " " & txtWitZip
1550      result = shl.Popup("Print an Envelope to:" & vbNewLine & strEnvAddr & " ??", , "Check address", 4)
          'strEnvAddr = txtWitFirst & " " & txtWitLast & vbCr & txtWitAddress & vbCr & txtWitCity & ", " & txtWitState & " " & txtWitZip
1560      If result = 6 Then
1570          wdDoc.Envelope.PrintOut Address:=strEnvAddr, omitReturnAddress:=True, Size:="Size 10"
1580      End If
          'AppActivate("Microsoft Excel" )
          'MsgBox "Envelope Printing Out", vbOKOnly, "Envelope Created"
          'wdApp.Activate
1590      Set shl = Nothing
1600      Set wdApp = Nothing
1610      Set wdDoc = Nothing
          

1620      On Error GoTo 0
1630  Exit Sub

cmdContactLetter_Click_Error:

         
1640      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

1650      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdContactLetter_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

1660      Print #1, zMsg

1670      Close #1

            
1680      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdCreateWitness_Click()

1690  On Error GoTo cmdCreateWitness_Click_Error
      Dim zMsg As String

       'Note: this code requires a reference to the Word object model
1700      Application.ScreenUpdating = False
          Dim wdApp As New Word.Application
          Dim wdDoc As Word.Document
          Dim CCtrl As Word.ContentControl
          Dim myShape As Word.Shape
          Dim strPath, strTemplateFileName, strActionEntry, strIdentification, strFileName As String
          Dim intLen, intCommaPos, intVersion As Integer
          Dim strCurrentTime, strInvInitials As String
          Dim Wksht As Worksheet, lngRow As Long, intCol As Integer
          Dim result As Integer
          Dim strGender, strAddress, strClientFirst, strClientLast, strClient, strMethod As String
          Dim TheLastRow, TheLastIODRow, TheLastCaseRow As Long
          Dim strDept, strXref, strAttorney, strCase As String
          Dim strNextCt As String
           
1710      strCurrentTime = Format(Now(), "h:mm AMPM")
1720      Application.EnableEvents = False
1730      Set Wksht = ActiveSheet
1740      lngRow = ActiveCell.row
1750      intCol = 3 ' Place on client name
1760      Cells(lngRow, intCol).Activate
          'Modified to adapt IOD cases
1770      strCase = Cells(lngRow, 1)
1780      strClient = Cells(lngRow, 3).Value
1790      strAttorney = Cells(lngRow, 5).Value
1800      strXref = Cells(lngRow, 4).Value
1810      strNextCt = Cells(lngRow, 10).Value
1820      strDept = Cells(lngRow, 11).Value
                
1830       If strCase = "ADMIN" Or strCase = "IOD" Then
1840            strCase = UCase(InputBox("Please enter the case number.", "Case Number"))
1850            strClient = ProperCase(InputBox("Please enter the first and last name of client. ex: John Doe", "Client"))
1860            strClient = RevAttyName(strClient)
1870            strAttorney = ProperCase(InputBox("Please enter the first and last name of the attorney. ex: Paulino Duran", "Attorney"))
1880            TheLastIODRow = IOD.Cells(Rows.Count, 1).End(xlUp).row
1890            TheLastIODRow = TheLastIODRow + 1
1900      End If
1910      result = MsgBox("Do you want an Investigative Report for " & ActiveCell.Value & "?", vbYesNoCancel, "Verify Case")
1920      If result = vbNo Then
1930           result = MsgBox("Click on the case you need the Investigative Report for, then press the command button!", vbOKOnly, "Verify Case")
1940           frmWitnessEntry.Hide
1950           Application.EnableEvents = True
1960           Exit Sub
1970      End If
1980      If result = vbCancel Then
1990          Unload frmWitnessEntry
2000          Exit Sub
2010      End If
2020      TheLastRow = WitnessLog.UsedRange.SpecialCells(xlCellTypeLastCell).row
2030      TheLastRow = TheLastRow + 1
2040      TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
2050      TheLastCaseRow = TheLastCaseRow + 1
                  
2060      strFileName = strClient
2070      strFileName = Replace(strFileName, ", ", "_")
2080      strFileName = strFileName & "_" & strCase
2090      strFileName = strFileName & "_" & txtWitFirst & "_" & txtWitLast
          'If txtVersion <> "" Then strFileName = strFileName & "_" & txtVersion
          
          
2100      strInvInitials = Files.Cells(16, 2).Value
          
2110      intCommaPos = InStr(strClient, ",")
2120      intLen = Len(strClient)
2130      strClientLast = Left$(strClient, intCommaPos - 1)
2140      strClientFirst = Right$(strClient, intLen - intCommaPos - 1)
          
2150      strPath = Files.Cells(1, 2).Value
          
          'Check for Juvenile Template
2160      If InvestigationLog.Cells(lngRow, 15) = "Juv" Then
2170              strTemplateFileName = Files.Cells(12, 2).Value
2180          Else
2190              strTemplateFileName = Files.Cells(2, 2).Value
2200      End If
              
2210      If strTemplateFileName = "" Then
2220          result = MsgBox("Please select the Investigative Report Template Location!", vbCritical, "Need the file name")
2230          strTemplateFileName = FilePicked("Investigative Report Template")
2240          Files.Cells(2, 2).Value = strTemplateFileName
2250      End If
2260      If strPath = "" Then
2270          result = MsgBox("Please select the path where the Investigative Report will be stored!", vbCritical, "Need the path")
2280          strPath = PathPicked("Investigative Report") & "\"
2290          Files.Cells(1, 2).Value = strPath
2300      End If
          
2310      intVersion = GetVersion(strPath, strFileName)
2320      If intVersion > 1 Then
2330          strFileName = strFileName & "_" & CStr(intVersion)
2340      End If
           
          
2350      If optMale = True Then
2360          strGender = "him"
2370      Else
2380          strGender = "her"
2390      End If
          
2400      If optInPerson = True Then
2410          strMethod = "an in person interview with"
2420          strAddress = txtLocation & ", telephone: " & txtTel
2430          strIdentification = "I showed " & txtWitLast & " my business identification and gave " & strGender & " my business card."
2440          WitnessLog.Cells(TheLastRow, 4) = "PI"
2450          strActionEntry = "In person interview with " & txtWitFirst & " " & txtWitLast & ", " & txtLocation
              'Report Counting
2460          Cells(lngRow, 24).Value = Cells(lngRow, 24).Value + 1
              'End of counting
2470      Else
2480          strMethod = "a telephone interview with"
2490          strAddress = txtTel & ", address: " & txtLocation
2500          WitnessLog.Cells(TheLastRow, 4) = "TI"
2510          strActionEntry = "Telephone Interview with " & txtWitFirst & " " & txtWitLast & ", " & txtTel
2520          strIdentification = " "
              'Report Counting
2530          Cells(lngRow, 23).Value = Cells(lngRow, 23).Value + 1
              'End of counting
2540      End If
          
          'Start of new detail
          
          
          
2550      If Cells(lngRow, 1).Value = "IOD" Then
                    
2560            IOD.Cells(TheLastIODRow, 1).Value = strClient & ", " & strCase & ", " & RevAttyName(strAttorney)
2570            IOD.Cells(TheLastIODRow, 2).Value = strActionEntry
2580            IOD.Cells(TheLastIODRow, 4).Value = GetADate(txtDOInt) 'Format(txtDOInt, "m/d/yy")
2590            IOD.Cells(TheLastIODRow, 3).Value = "Yes"
2600            strXref = InputBox("Client's Xref? - ", "Xref")
2610            strNextCt = InputBox("Next court date? - ", "Court Date")
2620            strNextCt = GetADate(strNextCt)
2630            strDept = InputBox("Court Department? - ", "Department")
2640       End If
      'Update WitnessLog & CaseLogs
       
2650      WitnessLog.Cells(TheLastRow, 1).Value = InvestigationLog.Cells(lngRow, 1).Value
2660      WitnessLog.Cells(TheLastRow, 5).Value = txtWitLast & ", " & txtWitFirst
2670      WitnessLog.Cells(TheLastRow, 2).Value = DateValue(txtDOInt)
2680      WitnessLog.Cells(TheLastRow, 3).Value = Format(txtTOI, "h:mm AMPM")
          
2690      UpdateCaseLog strActionEntry, Format(txtTOI, "h:mm AMPM"), GetADate(txtDOInt), TheLastCaseRow, lngRow, 1, Val(txtDuration)
              
          'Typed statement
2700      UpdateCaseLog "Typed report for " & txtWitFirst & " " & txtWitLast, Format(Now(), "h:mm AMPM"), Format(Now(), "m/d/yy"), TheLastCaseRow + 1, lngRow, 1, 0
          
          
2710      If IsaDate(txtDOB) = True Then
2720          WitnessLog.Cells(TheLastRow, 6).Value = DateValue(txtDOB)
2730      End If
          
2740      WitnessLog.Cells(TheLastRow, 7).Value = txtLocation
2750      WitnessLog.Cells(TheLastRow, 8).Value = txtTel
2760      If ckAddMileage = True Then
2770          Call AddMileage(GetADate(txtDOInt), txtMileageAddress, InvestigationLog.Cells(lngRow, 1).Value, txtStartM, txtEndM)
2780          CaseLogs.Cells(TheLastCaseRow, 7).Value = "Mileage Entry"
2790          CaseLogs.Cells(TheLastCaseRow, 8).Value = Val(txtStartM)
2800          CaseLogs.Cells(TheLastCaseRow, 9).Value = Val(txtEndM)
2810       End If
      'Generate and update word doc
             
2820         Set wdDoc = wdApp.Documents.Open(FileName:=strTemplateFileName, AddToRecentFiles:=False, Visible:=False)
2830          With wdDoc
2840              wdDoc.Activate
2850              For Each CCtrl In .ContentControls
2860                  Select Case CCtrl.Title
                          Case "CaseNum"
2870                          CCtrl.Range.Text = JuvCase(strCase) 'Cells(lngRow, 1).Value)
2880                      Case "Client"
2890                          CCtrl.Range.Text = strClientFirst & " " & strClientLast
2900                      Case "xref"
2910                          CCtrl.Range.Text = strXref
2920                      Case "Atty"
2930                          CCtrl.Range.Text = strAttorney
2940                      Case "CourtDate"
2950                          CCtrl.Range.Text = strNextCt
2960                      Case "Dept"
2970                          CCtrl.Range.Text = strDept
2980                      Case "Publish Date"
2990                          CCtrl.Range.Text = Now()
3000                      Case "TimeWritten"
3010                          CCtrl.Range.Text = strCurrentTime
3020                      Case "WitnessFirst"
3030                          CCtrl.Range.Text = txtWitFirst
3040                      Case "WitnessLast"
3050                          CCtrl.Range.Text = txtWitLast
3060                      Case "DOI"
3070                          CCtrl.Range.Text = txtDOInt
3080                      Case "TOI"
3090                          CCtrl.Range.Text = txtTOI
3100                      Case "DOB"
3110                          CCtrl.Range.Text = txtDOB
3120                      Case "Location"
3130                          CCtrl.Range.Text = strAddress
3140                      Case "Method"
3150                          CCtrl.Range.Text = strMethod
3160                      Case "HimHer"
3170                          CCtrl.Range.Text = strGender
3180                      Case "IncidentDate"
3190                          CCtrl.Range.Text = txtIncidentDate
3200                      Case "Identification"
3210                          CCtrl.Range.Text = strIdentification
3220                      Case "InvInit"
3230                          CCtrl.Range.Text = strInvInitials
3240                      Case "InvInit2"
3250                          If strMethod = "in person" Then
3260                              CCtrl.Range.Text = strInvInitials
3270                          Else
3280                              CCtrl.Range.Text = "-"
3290                          End If
3300                      Case "InvName"
3310                          CCtrl.Range.Text = Files.Cells(20, 2).Value
3320                      Case "InvPhone"
3330                          CCtrl.Range.Text = Files.Cells(23, 2).Value
3340                      Case "InvCell"
3350                          CCtrl.Range.Text = Files.Cells(24, 2).Value
3360                      Case "TelInt"
3370                          If strMethod = "via telephone" Then
3380                              CCtrl.Checked = True
3390                          Else
3400                              CCtrl.Checked = False
3410                          End If
3420                      Case "CB2"
3430                          If strMethod = "via telephone" Then
3440                              CCtrl.Checked = True
3450                          Else
3460                              CCtrl.Checked = False
3470                          End If
3480                      Case "CB3"
3490                          If strMethod = "via telephone" Then
3500                              CCtrl.Checked = True
3510                          Else
3520                              CCtrl.Checked = False
3530                          End If
3540                      Case "PersInt"
3550                          If strMethod = "in person" Then
3560                              CCtrl.Checked = True
3570                          Else
3580                              CCtrl.Checked = False
3590                          End If
                          
                              
3600                  End Select
                      
3610              Next
                              
3620          End With
           
          
          
          
          'Save with new name
3630      wdDoc.SaveAs FileName:=strPath & strFileName & ".docx", FileFormat:=wdFormatDocumentDefault
3640      Set myShape = wdDoc.Shapes("InitBox")
3650      With myShape
3660                  .TextFrame.TextRange.ContentControls(1).Range.Text = strInvInitials
3670                  .TextFrame.TextRange.ContentControls(2).Range.Text = strInvInitials
3680                  .TextFrame.TextRange.ContentControls(3).Range.Text = strInvInitials
3690                  If optInPerson = True Then
3700                      .TextFrame.TextRange.ContentControls(4).Range.Text = strInvInitials
3710                      .TextFrame.TextRange.ContentControls(5).Range.Text = strInvInitials
3720                  Else
3730                      .TextFrame.TextRange.ContentControls(4).Range.Text = "-"
3740                      .TextFrame.TextRange.ContentControls(5).Range.Text = "-"
3750                  End If
3760           End With
          
3770      wdDoc.Bookmarks("Content").Range.Select
          
3780      wdDoc.Save
3790      SortCaseLogs
3800      InvestigationLog.Activate
3810      Application.ScreenUpdating = True
3820      Application.EnableEvents = True
3830      ActiveWorkbook.Save
3840      Unload frmWitnessEntry
          
3850      wdApp.Visible = True
3860      wdApp.Activate
3870      Set wdApp = Nothing
3880      Set wdDoc = Nothing
          
          'Save with new name
               
          'wdApp.Quit
              
          
          

3890      On Error GoTo 0
3900  Exit Sub

cmdCreateWitness_Click_Error:

         
3910      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

3920      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdCreateWitness_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

3930      Print #1, zMsg

3940      Close #1

            
3950      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub




Private Sub cmdDueDiligence_Click()


3960  On Error GoTo cmdDueDiligence_Click_Error
      Dim zMsg As String

       'Note: this code requires a reference to the Word object model
3970      Application.ScreenUpdating = False
          Dim wdApp As New Word.Application
          Dim wdDoc As Word.Document
          Dim CCtrl As Word.ContentControl
          Dim strPath, strTemplateFileName, strActionEntry, strFileName As String
          Dim intLen, intCommaPos, intVersion As Integer
          Dim strCurrentTime As String
          Dim Wksht As Worksheet, lngRow As Long, intCol As Integer
          Dim result As Integer
          Dim strGender, strAddress, strClientFirst, strClientLast, strClient, strMethod As String
          Dim TheLastCaseRow As Long
          
3980      strCurrentTime = Format(Now(), "h:mm AMPM")
3990      Application.EnableEvents = False ' Disable on change routine
          
4000      Set Wksht = ActiveSheet
4010      lngRow = ActiveCell.row
4020      intCol = 3 ' Place on client name
4030      Cells(lngRow, intCol).Activate
          
4040      result = MsgBox("Do you want a Due Diligence Report for " & ActiveCell.Value & "?", vbYesNo, "Verify Case")
4050      If result = vbNo Then
4060           result = MsgBox("Click on the case you need the Due Diligence Report for, then press the command button!", vbOKOnly, "Verify Case")
4070           Unload frmWitnessEntry
4080           Exit Sub
4090      End If
              
4100      TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
4110      TheLastCaseRow = TheLastCaseRow + 1
              
4120      strFileName = ActiveCell.Value
4130      strFileName = Replace(strFileName, ", ", "_")
4140      strFileName = strFileName & "_" & ActiveCell.Offset(0, -2).Value
4150      strFileName = strFileName & "_" & txtWitFirst & "_" & txtWitLast & "_Due_Diligence"
4160      strPath = Files.Cells(1, 2).Value
          
4170      intVersion = GetVersion(strPath, strFileName)
4180      If intVersion > 1 Then
4190          strFileName = strFileName & "_" & CStr(intVersion)
4200      End If
         
            
4210      strClient = Cells(lngRow, 3).Value
4220      intCommaPos = InStr(strClient, ",")
4230      intLen = Len(strClient)
4240      strClientLast = Left$(strClient, intCommaPos - 1)
4250      strClientFirst = Right$(strClient, intLen - intCommaPos - 1)
4260      strActionEntry = "Prepared Due Diligence for " & txtWitFirst & " " & txtWitLast
          
4270      strPath = Files.Cells(1, 2).Value
4280      strTemplateFileName = Files.Cells(4, 2).Value
4290      If strTemplateFileName = "" Then
4300          result = MsgBox("Please select the Due Diligence Report Template Location!", vbCritical, "Need the file name")
4310          strTemplateFileName = FilePicked("Due Diligence Template")
4320          Files.Cells(4, 2).Value = strTemplateFileName
4330      End If
4340      If strPath = "" Then
4350          result = MsgBox("Please select the path where the Due Diligence Report will be stored!", vbCritical, "Need the path")
4360          strPath = PathPicked("Due Diligence Report") & "\"
4370          Files.Cells(1, 2).Value = strPath
4380      End If
          
4390      If optMale = True Then
4400          strGender = "him"
4410      Else
4420          strGender = "her"
4430      End If
          
4440      If optInPerson = True Then
4450          strMethod = "in person"
4460          strAddress = txtLocation & ", telephone: " & txtTel
4470      Else
4480          strMethod = "via telephone"
4490          strAddress = txtTel & ", address: " & txtLocation
4500      End If
          'Update Case Log with Entry
          
4510      UpdateCaseLog strActionEntry, Format(strCurrentTime, "h:mm AMPM"), GetADate(Now()), TheLastCaseRow, lngRow, 1, Val(txtDuration)
              
          
          'Report Counting
4520       Cells(lngRow, 25).Value = Cells(lngRow, 25).Value + 1
          'End of counting
4530         Set wdDoc = wdApp.Documents.Open(FileName:=strTemplateFileName, AddToRecentFiles:=False, Visible:=False)
4540          With wdDoc
4550              wdDoc.Activate
4560              For Each CCtrl In .ContentControls
4570                  Select Case CCtrl.Title
                          Case "CaseNum"
4580                          CCtrl.Range.Text = JuvCase(Cells(lngRow, 1).Value)
4590                      Case "Client"
4600                          CCtrl.Range.Text = strClientFirst & " " & strClientLast
4610                      Case "ClientTitle"
4620                          CCtrl.Range.Text = DefTitle(InvestigationLog.Cells(lngRow, 15).Value)
4630                      Case "Div"
4640                          CCtrl.Range.Text = DivisionTitle(InvestigationLog.Cells(lngRow, 15).Value)
4650                      Case "CaseDesc"
4660                          CCtrl.Range.Text = CaseDesc(InvestigationLog.Cells(lngRow, 15).Value)
4670                      Case "xref"
4680                          CCtrl.Range.Text = Cells(lngRow, 4).Value
4690                      Case "Atty"
4700                          CCtrl.Range.Text = Cells(lngRow, 5).Value
4710                      Case "CourtDate"
4720                          CCtrl.Range.Text = Cells(lngRow, 10).Value
4730                      Case "Dept"
4740                          CCtrl.Range.Text = Cells(lngRow, 11).Value
4750                      Case "Publish Date"
4760                          CCtrl.Range.Text = Now()
4770                      Case "TimeWritten"
4780                          CCtrl.Range.Text = strCurrentTime
4790                      Case "WitnessFirst"
4800                          If txtWitFirst = "" Then txtWitFirst = " "
4810                          CCtrl.Range.Text = txtWitFirst
4820                      Case "WitnessLast"
4830                          If txtWitLast = "" Then txtWitLast = " "
4840                          CCtrl.Range.Text = txtWitLast
4850                      Case "InvName"
4860                          CCtrl.Range.Text = Files.Cells(20, 2).Value
4870                      Case "InvPhone"
4880                          CCtrl.Range.Text = Files.Cells(23, 2).Value
4890                      Case "InvCell"
4900                          CCtrl.Range.Text = Files.Cells(24, 2).Value
                          
4910                  End Select
                      
4920              Next
                              
4930          End With
4940      wdDoc.Bookmarks("Content").Range.Select
          
4950      Application.ScreenUpdating = True
4960      Application.EnableEvents = True ' Enable on change routine
4970      ActiveWorkbook.Save
4980      Unload frmWitnessEntry
          
          'Save with new name
4990      wdDoc.SaveAs FileName:=strPath & strFileName & ".docx", FileFormat:=wdFormatDocumentDefault
5000      wdApp.Visible = True
5010      wdApp.Activate
5020      Set wdApp = Nothing
5030      Set wdDoc = Nothing
          'wdApp.Quit
          'Set wdDoc = Nothing: Set wdApp = Nothing: Set WkSht = Nothing
          
5040      Exit Sub
          

5050      On Error GoTo 0
5060  Exit Sub

cmdDueDiligence_Click_Error:

         
5070      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

5080      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdDueDiligence_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

5090      Print #1, zMsg

5100      Close #1

            
5110      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub




Private Sub cmdFaxCover_Click()

5120  On Error GoTo cmdFaxCover_Click_Error
      Dim zMsg As String

5130      Application.ScreenUpdating = False
          Dim wdApp As New Word.Application
          Dim wdDoc As Word.Document
          Dim CCtrl As Word.ContentControl
          Dim strPath, strTemplateFileName, strActionEntry, strFileName As String
          Dim intLen, intCommaPos, intVersion As Integer
          Dim strCurrentTime As String
          
          Dim Wksht As Worksheet, lngRow As Long, intCol As Integer
          Dim result As Integer
          Dim strAddress, strClientFirst, strClientLast, strClient As String
          Dim TheLastCaseRow As Long
         
       
5140      strCurrentTime = Format(Now(), "h:mm AMPM")
5150      Application.EnableEvents = False
5160      Set Wksht = ActiveSheet
5170      lngRow = ActiveCell.row
5180      intCol = 3 ' Place on client name
5190      Cells(lngRow, intCol).Activate

5200      strAddress = txtWitAddress & ", " & txtWitCity & ", " & txtWitState & " " & txtWitZip
5210      strActionEntry = "Generated fax cover and sent to " & txtFaxNum
          
          
5220      result = MsgBox("Do you want to create a fax for " & ActiveCell.Value & "?", vbYesNoCancel, "Verify Case")
5230      If result = vbNo Then
5240           result = MsgBox("Click on the case you need the Investigative Report for, then press the command button!", vbOKOnly, "Verify Case")
5250           frmWitnessEntry.Hide
5260           Application.EnableEvents = True
5270           Exit Sub
5280      End If
5290      If result = vbCancel Then
5300          Unload frmWitnessEntry
5310          Exit Sub
5320      End If
         
5330      TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
5340      TheLastCaseRow = TheLastCaseRow + 1
                  
5350      strFileName = ActiveCell.Value
5360      strFileName = Replace(strFileName, ", ", "_")
5370      strFileName = strFileName & "_" & ActiveCell.Offset(0, -2).Value
5380      strFileName = strFileName & "_" & txtWitFirst & "_" & txtWitLast & "_Fax_Cover"
          'If txtVersion <> "" Then strFileName = strFileName & "_" & txtVersion
          
5390      strClient = Cells(lngRow, 3).Value
              
5400      intCommaPos = InStr(strClient, ",")
5410      intLen = Len(strClient)
5420      strClientLast = Left$(strClient, intCommaPos - 1)
5430      strClientFirst = Right$(strClient, intLen - intCommaPos - 1)
          
5440      strPath = Files.Cells(1, 2).Value
5450      strTemplateFileName = Files.Cells(13, 2).Value
5460      If strTemplateFileName = "" Then
5470          result = MsgBox("Please select the Contact Letter Template Location!", vbCritical, "Need the file name")
5480          strTemplateFileName = FilePicked("Contact Letter Template")
5490          Files.Cells(13, 2).Value = strTemplateFileName
5500      End If
5510      If strPath = "" Then
5520          result = MsgBox("Please select the path where the Contact Letter will be stored!", vbCritical, "Need the path")
5530          strPath = PathPicked("Contact Letter") & "\"
5540           Files.Cells(1, 2).Value = strPath
5550      End If
          
5560     intVersion = GetVersion(strPath, strFileName)
5570         If intVersion > 1 Then
5580          strFileName = strFileName & "_" & CStr(intVersion)
5590      End If
          
          'blank content controls in word document if entries are blank
5600      If txtFaxMsg = "" Then txtFaxMsg = " "
5610      If txtPages = "" Then txtPages = " "
5620      If txtFaxDept = "" Then txtFaxDept = " "
           
          ' Update case log
5630      UpdateCaseLog strActionEntry, Format(Now(), "h:mm AMPM"), GetADate(Now()), TheLastCaseRow, lngRow, 1, Val(txtDurationFax)
              
5640      Set wdDoc = wdApp.Documents.Open(FileName:=strTemplateFileName, AddToRecentFiles:=False, Visible:=False)
5650          With wdDoc
5660              wdDoc.Activate
5670              For Each CCtrl In .ContentControls
5680                  Select Case CCtrl.Title
                          
                          Case "FaxTo"
5690                          CCtrl.Range.Text = txtFaxName
5700                      Case "FaxToDept"
5710                          If txtFaxDept = "" Then txtFaxDept = " "
5720                          CCtrl.Range.Text = txtFaxDept
5730                      Case "FaxDate"
5740                          CCtrl.Range.Text = Format(Now(), "mmmm d, yyyy")
5750                      Case "Msg"
5760                          CCtrl.Range.Text = txtFaxMsg
5770                      Case "FaxNumber"
5780                          CCtrl.Range.Text = txtFaxNum
5790                      Case "Pages"
5800                          CCtrl.Range.Text = txtPages
                          
5810                      Case "SenderFax"
5820                          CCtrl.Range.Text = Files.Cells(27, 2).Value
5830                      Case "InvTitle"
5840                          CCtrl.Range.Text = Files.Cells(26, 2).Value
5850                      Case "InvEmail"
5860                          CCtrl.Range.Text = Files.Cells(25, 2).Value
5870                      Case "InvName"
5880                          CCtrl.Range.Text = Files.Cells(20, 2).Value
5890                      Case "InvPhone"
5900                          CCtrl.Range.Text = Files.Cells(23, 2).Value
5910                      Case "InvCell"
5920                          CCtrl.Range.Text = Files.Cells(24, 2).Value
5930                  End Select
                      
5940              Next
                              
5950          End With
           
          
5960      Application.ScreenUpdating = True
5970      Application.EnableEvents = True
5980      ActiveWorkbook.Save
5990      Unload frmWitnessEntry
          
          'Save with new name
6000      wdDoc.SaveAs FileName:=strPath & strFileName & ".docx", FileFormat:=wdFormatDocumentDefault
6010      wdApp.Visible = True
6020      wdApp.Activate
6030      Set wdApp = Nothing
6040      Set wdDoc = Nothing

6050      On Error GoTo 0
6060  Exit Sub

cmdFaxCover_Click_Error:

         
6070      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

6080      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdFaxCover_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

6090      Print #1, zMsg

6100      Close #1

            
6110      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdPhotoReport_Click()

6120  On Error GoTo cmdPhotoReport_Click_Error
      Dim zMsg As String

      'Note: this code requires a reference to the Word object model
6130      Application.ScreenUpdating = False
          Dim wdApp As New Word.Application
          Dim wdDoc As Word.Document
          Dim CCtrl As Word.ContentControl
          Dim strPath, strTemplateFileName, strFileName, strCurrentTime As String
          Dim strCase As String
          Dim strAttorney As String
          Dim intLen, intCommaPos As Integer
          Dim intVersion As Integer
          Dim Wksht As Worksheet, lngRow As Long, intCol As Integer
          Dim result As Integer
          Dim TheLastCaseRow, TheLastIODRow As Long
          Dim strClientFirst, strClientLast, strClient, strActionEntry, strRptType, strRptMsg As String
          Dim strDept, strXref, strNextCt As String
          
6140      strCurrentTime = Format(Now(), "h:mm AMPM")
6150      Application.EnableEvents = False ' Disable on change routine
          
6160      If optOther = False Then
6170          strRptType = "_Photo_Report"
6180          strRptMsg = Replace(strRptType, "_", " ") ' replaces underscore with space
               
6190      Else
6200          strRptMsg = optOther.Caption
6210          strRptType = Replace(optOther.Caption, " ", "_")
6220          strRptType = "_" & strRptType
6230      End If
          
6240      Set Wksht = ActiveSheet
6250      lngRow = ActiveCell.row
6260      intCol = 3 ' Place on client name
6270      Cells(lngRow, intCol).Activate
6280      strCase = Cells(lngRow, 1)
6290      strClient = Cells(lngRow, 3).Value
6300      strAttorney = Cells(lngRow, 5).Value
          
6310      If strCase = "ADMIN" Or strCase = "IOD" Then
6320          strCase = UCase(InputBox("Please enter the case number.", "Case Number"))
6330          strClient = ProperCase(InputBox("Please enter the first and last name of client. ex: John Doe", "Client"))
6340          strClient = RevAttyName(strClient)
6350          strAttorney = ProperCase(InputBox("Please enter the first and last name of the attorney. ex: Paulino Duran", "Attorney"))
6360          TheLastIODRow = IOD.Cells(Rows.Count, 1).End(xlUp).row
6370          TheLastIODRow = TheLastIODRow + 1
6380      End If
          
6390      TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
6400      TheLastCaseRow = TheLastCaseRow + 1
          
6410      result = MsgBox("Do you want a " & strRptMsg & " for " & ActiveCell.Value & "?", vbYesNo, "Verify Case " & strRptMsg)
6420      If result = vbNo Then
6430           result = MsgBox("Click on the case you need the Photo Report for, then press the command button!", vbOKOnly, "Verify Case")
6440           Unload frmWitnessEntry
6450           Exit Sub
6460      End If
         
              
6470      strFileName = GetFileNameBase(strClient, strCase)
6480      strFileName = strFileName & strRptType
          'If txtVersion <> "" Then strFileName = strFileName & "_" & txtVersion
          
6490      strPath = Files.Cells(1, 2).Value
          
6500      intVersion = GetVersion(strPath, strFileName)
          
6510      If intVersion > 1 Then
6520          strFileName = strFileName & "_" & CStr(intVersion)
6530      End If
          
          
          
6540      intCommaPos = InStr(strClient, ",")
6550      intLen = Len(strClient)
6560      strClientLast = Left$(strClient, intCommaPos - 1)
6570      strClientFirst = Right$(strClient, intLen - intCommaPos - 1)
6580      strActionEntry = "Prepared " & strRptMsg
          
          
6590      strPath = Files.Cells(1, 2).Value
6600      strTemplateFileName = Files.Cells(3, 2).Value
6610      If strTemplateFileName = "" Then
6620          result = MsgBox("Please select the Photo Report Template Location!", vbCritical, "Need the file name")
6630          strTemplateFileName = FilePicked("Photo Report Template")
6640          Files.Cells(3, 2).Value = strTemplateFileName
6650      End If
6660      If strPath = "" Then
6670          result = MsgBox("Please select the path where the Photo Report will be stored!", vbCritical, "Need the path")
6680          strPath = PathPicked("Due Diligence Report") & "\"
6690          Files.Cells(1, 2).Value = strPath
6700      End If
          'Report Counting
6710      If strRptType = "_Photo_Report" Then
6720          Cells(lngRow, 26).Value = Cells(lngRow, 26).Value + 1
6730          MakePhotoDir (GetFileNameBase(Files.Cells(29, 2).Value & strClient, strCase))
6740      Else
6750          Cells(lngRow, 22).Value = Cells(lngRow, 22).Value + 1
6760      End If
          'End of counting
          
          'Update Case Log with Entry
6770      UpdateCaseLog strActionEntry, strCurrentTime, GetADate(Now()), TheLastCaseRow, lngRow, 1, Val(txtDuration)
          
          
6780      If Cells(lngRow, 1).Value = "IOD" Then
              
6790          IOD.Cells(TheLastIODRow, 1).Value = strClient & ", " & strCase & ", " & RevAttyName(strAttorney)
6800          IOD.Cells(TheLastIODRow, 2).Value = strActionEntry
6810          IOD.Cells(TheLastIODRow, 4).Value = GetADate(txtDOInt) 'Format(txtDOInt, "m/d/yy")
6820          IOD.Cells(TheLastIODRow, 3).Value = "Yes"
6830          strXref = InputBox("Client's Xref? - ", "Xref")
6840          strNextCt = InputBox("Next court date? - ", "Court Date")
6850          strNextCt = GetADate(strNextCt)
6860          strDept = InputBox("Court Department? - ", "Department")
6870     End If
          
6880      If ckAddMileage = True Then
6890          Call AddMileage(GetADate(txtDOInt), txtMileageAddress, InvestigationLog.Cells(lngRow, 1).Value, txtStartM, txtEndM)
6900          CaseLogs.Cells(TheLastCaseRow, 7).Value = "Mileage Entry"
6910          CaseLogs.Cells(TheLastCaseRow, 8).Value = Val(txtStartM)
6920          CaseLogs.Cells(TheLastCaseRow, 9).Value = Val(txtEndM)
6930       End If
6940         Set wdDoc = wdApp.Documents.Open(FileName:=strTemplateFileName, AddToRecentFiles:=False, Visible:=False)
6950          With wdDoc
6960              wdDoc.Activate
6970              For Each CCtrl In .ContentControls
6980                  Select Case CCtrl.Title
                          Case "CaseNum"
6990                          CCtrl.Range.Text = JuvCase(strCase)
7000                      Case "Client"
7010                          CCtrl.Range.Text = strClientFirst & " " & strClientLast
7020                      Case "xref"
7030                          If Cells(lngRow, 1).Value = "IOD" Or Cells(lngRow, 1).Value = "Admin" Then
7040                              CCtrl.Range.Text = strXref
7050                          Else
7060                              CCtrl.Range.Text = Cells(lngRow, 4).Value
7070                          End If
                              
7080                      Case "Atty"
7090                          CCtrl.Range.Text = strAttorney
7100                      Case "ClientTitle"
7110                          CCtrl.Range.Text = DefTitle(InvestigationLog.Cells(lngRow, 15).Value)
7120                      Case "Div"
7130                          CCtrl.Range.Text = DivisionTitle(InvestigationLog.Cells(lngRow, 15).Value)
7140                      Case "CaseDesc"
7150                          CCtrl.Range.Text = CaseDesc(InvestigationLog.Cells(lngRow, 15).Value)
7160                      Case "CourtDate"
7170                          If Cells(lngRow, 1).Value = "IOD" Or Cells(lngRow, 1).Value = "Admin" Then
7180                              CCtrl.Range.Text = Format(strNextCt, "MM/DD/YYYY")
7190                          Else
7200                              CCtrl.Range.Text = Cells(lngRow, 10).Value
7210                          End If
7220                      Case "Dept"
7230                          If Cells(lngRow, 1).Value = "IOD" Or Cells(lngRow, 1).Value = "Admin" Then
7240                              CCtrl.Range.Text = strDept
7250                          Else
7260                              CCtrl.Range.Text = Cells(lngRow, 11).Value
7270                          End If
                              
7280                      Case "Publish Date"
7290                          CCtrl.Range.Text = Now()
7300                      Case "TimeWritten"
7310                          CCtrl.Range.Text = strCurrentTime
7320                      Case "RptType"
7330                          CCtrl.Range.Text = strRptMsg
7340                      Case "InvName"
7350                          CCtrl.Range.Text = Files.Cells(20, 2).Value
7360                      Case "InvPhone"
7370                          CCtrl.Range.Text = Files.Cells(23, 2).Value
7380                      Case "InvCell"
7390                          CCtrl.Range.Text = Files.Cells(24, 2).Value
7400                  End Select
                      
7410              Next
                              
7420          End With
           
7430      wdDoc.Bookmarks("Content").Range.Select
7440      Application.ScreenUpdating = True
7450      Application.EnableEvents = True ' Enable on change routine
7460      ActiveWorkbook.Save
7470      Unload frmWitnessEntry
          
          'Save with new name
7480      wdDoc.SaveAs FileName:=strPath & strFileName & ".docx", FileFormat:=wdFormatDocumentDefault
7490      wdApp.Visible = True
7500      wdApp.Activate
7510      Set wdApp = Nothing
7520      Set wdDoc = Nothing
          'wdApp.Quit
          'Set wdDoc = Nothing: Set wdApp = Nothing: Set WkSht = Nothing
          
7530      On Error GoTo 0
7540  Exit Sub

cmdPhotoReport_Click_Error:

         
7550      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

7560      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdPhotoReport_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

7570      Print #1, zMsg

7580      Close #1

            
7590      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdSubService_Click()

7600  On Error GoTo cmdSubService_Click_Error
      Dim zMsg As String

      Dim TheLastRow, TheLastCaseRow, TheLastIODRow, lngRow As Long
      Dim intCol As Integer
      Dim strServiceMethod, strName, strLogName As String
      Dim intReportCol As Integer
7610  Application.ScreenUpdating = False
      Dim result As Integer

      'Automate service form
      Dim strCase As String
      Dim wdApp As New Word.Application
      Dim wdDoc As Word.Document
      Dim CCtrl As Word.ContentControl
      Dim strClient, strPath, strFileName, strTemplateFileName As String
      Dim intVersion As Integer
      Dim strAttorney As String


7620  lngRow = ActiveCell.row
7630  intCol = 3 ' Place on client name
7640  Cells(lngRow, intCol).Activate

      'Allow automatic update of service form
7650  strCase = UCase(Cells(lngRow, 1))
7660  strClient = AttorneyName(Cells(lngRow, intCol).Value)
7670  strAttorney = InvestigationLog.Cells(lngRow, 5).Value

7680  If strCase = "ADMIN" Or strCase = "IOD" Then
7690      strCase = InputBox("Please enter the case number.", "Case Number")
7700      strClient = InputBox("Please enter the first and last name of client. ex: John Doe", "Client")
7710      strAttorney = InputBox("Please enter the first and last name of the attorney.  ex: Paulino Duran", "Attorney")
7720      TheLastIODRow = IOD.Cells(Rows.Count, 1).End(xlUp).row
7730      TheLastIODRow = TheLastIODRow + 1
7740  End If
          
7750      TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
7760      TheLastCaseRow = TheLastCaseRow + 1
7770      TheLastRow = WitnessLog.Cells(Rows.Count, 1).End(xlUp).row
7780      TheLastRow = TheLastRow + 1
        
          
7790      result = MsgBox("Do you want to serve a subpoena for " & ActiveCell.Value & "?", vbYesNo, "Verify Case ")
7800      If result = vbNo Then
7810           result = MsgBox("Click on the case you need the service for, then press the command button!", vbOKOnly, "Verify Case")
7820           Unload frmWitnessEntry
7830           Exit Sub
7840      End If
          
7850      If optCLU = True Then
7860          strServiceMethod = "Served subpoena via Court Liaison to "
7870          intReportCol = 27
7880      End If
7890      If optEmailService = True Then
7900          strServiceMethod = "Served subpoena via Email to "
7910          intReportCol = 28
7920      End If
7930      If optFaxService = True Then
7940          strServiceMethod = "Served subpoena via Fax to "
7950          intReportCol = 28
7960      End If
7970      If optInPerService = True Then
7980          strServiceMethod = "Served subpoena in person to "
7990          intReportCol = 27
8000      End If
8010      If optUSMail = True Then
8020          strServiceMethod = "Served subpoena via US Mail to "
8030          intReportCol = 28
8040      End If
8050      If optOtherService = True Then
8060          strServiceMethod = "Served subpoena to "
8070          intReportCol = 27
8080      End If
          
8090      strName = txtWitFirst & " " & txtWitLast
8100      If txtWitLast = "" Then
8110          strLogName = txtWitFirst
8120      Else
8130          strLogName = txtWitLast & ", " & txtWitFirst
8140      End If
         
8150      strServiceMethod = strServiceMethod & strName & " at " & txtLOS
8160      strServiceMethod = strServiceMethod & ", " & txtTelS & ". " & txtNotes
          
      'Update WitnessLog, CaseLogs & Mileage Logs
       
8170      WitnessLog.Cells(TheLastRow, 1).Value = InvestigationLog.Cells(lngRow, 1).Value
8180      WitnessLog.Cells(TheLastRow, 5).Value = strLogName
8190      WitnessLog.Cells(TheLastRow, 2).Value = DateValue(txtDOInt)
8200      WitnessLog.Cells(TheLastRow, 3).Value = Format(txtTOI, "h:mm AMPM")
          
8210      UpdateCaseLog strServiceMethod, Format(txtTOI, "h:mm AMPM"), GetADate(txtDOInt), TheLastCaseRow, lngRow, 1, Val(txtSubDuration)
          
          
8220      WitnessLog.Cells(TheLastRow, 4) = "Sub"
8230      If IsaDate(txtDOB) = True Then
8240          WitnessLog.Cells(TheLastRow, 6).Value = DateValue(txtDOB)
8250      End If
8260      WitnessLog.Cells(TheLastRow, 7).Value = txtLOS
8270      WitnessLog.Cells(TheLastRow, 8).Value = txtTelS
8280      If ckAddMileage = True Then
8290          Call AddMileage(GetADate(txtDOInt), txtMileageAddress, strCase, txtStartM, txtEndM)
8300          CaseLogs.Cells(TheLastCaseRow, 7).Value = "Mileage Entry"
8310          CaseLogs.Cells(TheLastCaseRow, 8).Value = Val(txtStartM)
8320          CaseLogs.Cells(TheLastCaseRow, 9).Value = Val(txtEndM)
8330       End If
          
8340     If Cells(lngRow, 1).Value = "IOD" Then
              
8350          IOD.Cells(TheLastIODRow, 1).Value = strClient & ", " & strCase & ", " & RevAttyName(strAttorney)
8360          IOD.Cells(TheLastIODRow, 2).Value = strServiceMethod
8370          IOD.Cells(TheLastIODRow, 4).Value = GetADate(txtDOInt) 'Format(txtDOInt, "m/d/yy")
8380          IOD.Cells(TheLastIODRow, 3).Value = "Yes"
              
8390     End If
          
8400      Cells(lngRow, intReportCol).Value = Cells(lngRow, intReportCol).Value + 1
          
          'Automate service form
8410      If ckBox125 = True Then
8420          strTemplateFileName = Files.Cells(10, 2).Value
8430          strPath = Files.Cells(1, 2).Value
              'If Cells(lngRow, 1).Value = "ADMIN" Or Cells(lngRow, 1).Value = "IOD" Then
8440              strFileName = GetFileNameBase(RevAttyName(strClient), strCase)
             ' Else
                  'strFileName = GetFileNameBase(strClient, strCase)
              'End If
8450          strFileName = strFileName & "_" & strLogName & "_Subpoena"
8460          strFileName = Replace(strFileName, ".", "")
8470          strFileName = Replace(strFileName, ",", "_")
8480          intVersion = GetVersion(strPath, strFileName)
8490          If intVersion > 1 Then
8500              strFileName = strFileName & "_" & CStr(intVersion)
8510          End If
8520          Set wdDoc = wdApp.Documents.Open(FileName:=strTemplateFileName, AddToRecentFiles:=False, Visible:=False)
8530          With wdDoc
8540              wdDoc.Activate
8550              For Each CCtrl In .ContentControls
8560                  Select Case CCtrl.Title
                          Case "CaseNo"
8570                          CCtrl.Range.Text = JuvCase(strCase)
8580                      Case "Client"
8590                          CCtrl.Range.Text = strClient
8600                      Case "DOB"
8610                          If txtDOB = "" Then txtDOB = " "
8620                          CCtrl.Range.Text = txtDOB
8630                      Case "DOS"
8640                          CCtrl.Range.Text = txtDOInt
8650                      Case "TOS"
8660                          CCtrl.Range.Text = txtTOI
8670                      Case "Addr"
8680                          CCtrl.Range.Text = txtLOS & "  Tel: " & txtTelS
8690                      Case "WitFirst"
8700                          If txtWitFirst = "" Then txtWitFirst = " "
8710                          CCtrl.Range.Text = txtWitFirst
8720                      Case "WitLast"
8730                          If txtWitLast = "" Then txtWitLast = " "
8740                          CCtrl.Range.Text = txtWitLast
8750                      Case "Notes"
8760                          If txtNotes = "" Then txtNotes = " "
8770                          CCtrl.Range.Text = txtNotes
8780                      Case "Date"
8790                          CCtrl.Range.Text = Format(Now(), "MMM d, yyyy")
8800                      Case "InvName"
8810                          CCtrl.Range.Text = Files.Cells(20, 2).Value
8820                      Case "InvPhone"
8830                          CCtrl.Range.Text = Files.Cells(23, 2).Value
8840                      Case "InvCell"
8850                          CCtrl.Range.Text = Files.Cells(24, 2).Value
8860                  End Select
                      
8870              Next
                              
8880          End With
8890      wdDoc.SaveAs FileName:=strPath & strFileName & ".docx", FileFormat:=wdFormatDocumentDefault, AddToRecentFiles:=True
8900      wdDoc.PrintOut Copies:=2
8910      wdApp.Visible = False
8920      wdDoc.Close
8930      wdApp.Quit
8940      Set wdApp = Nothing
8950      Set wdDoc = Nothing
8960      End If
8970      SortCaseLogs
8980      InvestigationLog.Activate
8990      Application.ScreenUpdating = True
9000      ActiveWorkbook.Save
9010      Unload frmWitnessEntry
          

9020      On Error GoTo 0
9030  Exit Sub

cmdSubService_Click_Error:

         
9040      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

9050      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdSubService_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

9060      Print #1, zMsg

9070      Close #1

            
9080      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub cmdTestify_Click()

9090  On Error GoTo cmdTestify_Click_Error
      Dim zMsg As String

      Dim TheLastCaseRow As Long
      Dim strActionEntry As String
      Dim lngRow As Long, intCol As Integer

9100  Application.ScreenUpdating = False

9110  lngRow = ActiveCell.row
9120  intCol = 3 ' Place on client name
9130  Cells(lngRow, intCol).Activate
9140      strActionEntry = "Testified - "
9150      TheLastCaseRow = CaseLogs.Cells(Rows.Count, 1).End(xlUp).row
9160      TheLastCaseRow = TheLastCaseRow + 1
          
9170      UpdateCaseLog strActionEntry, Format(txtTOI, "h:mm AMPM"), GetADate(txtDOInt), TheLastCaseRow, lngRow, 1, Val(txtDuration)
                
9180      Cells(lngRow, 30).Value = Cells(lngRow, 30).Value + 1 ' Add testimony
9190      SortCaseLogs
9200      InvestigationLog.Activate
9210      Application.ScreenUpdating = True
9220      ActiveWorkbook.Save
9230      Unload frmWitnessEntry

9240      On Error GoTo 0
9250  Exit Sub

cmdTestify_Click_Error:

         
9260      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

9270      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: cmdTestify_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

9280      Print #1, zMsg

9290      Close #1

            
9300      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub





Private Sub frameMileage_Exit(ByVal Cancel As MSForms.ReturnBoolean)

9310  On Error GoTo frameMileage_Exit_Error
      Dim zMsg As String

9320  If ckAddMileage = False Then
9330      Exit Sub
9340  Else

9350      If ValidMileage(txtStartM, txtEndM, txtMileageAddress, Me) = False Then
9360          Cancel = True
          
9370      End If
9380  End If

9390      On Error GoTo 0
9400  Exit Sub

frameMileage_Exit_Error:

         
9410      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

9420      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: frameMileage_Exit Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

9430      Print #1, zMsg

9440      Close #1

            
9450      MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub optCLU_Click()
9460      txtLOS.SetFocus
End Sub

Private Sub optContactLetter_Click()

9470  On Error GoTo optContactLetter_Click_Error
      Dim zMsg As String

9480  cmdDueDiligence.Visible = False
9490  cmdPhotoReport.Visible = False
9500  cmdCreateWitness.Visible = False
9510  cmdSubService.Visible = False
9520  cmdTestify.Visible = False
9530  ckbxSubLetter = False
9540  optOther.Caption = "Other Report"
9550  Label1.Caption = "Witness First name"
9560  Label2.Caption = "Last name"
9570  txtDuration.Visible = False
9580  Label13.Visible = False
9590  frameContactLetter.Visible = True
9600  frameContactLetter.Enabled = True
9610  frameContactLetter.Top = 126


      Dim contr As Control

9620  For Each contr In frmWitnessEntry.Controls
9630      If TypeName(contr) = "TextBox" Or TypeName(contr) = "Label" Or TypeName(contr) = "OptionButton" Then
9640          contr.Enabled = True
9650          contr.Visible = True
9660      End If
9670  Next
9680  frameGender.Visible = False
9690  frameGender.Enabled = False
9700  frameServiceInfo.Visible = False
9710  frameFaxCover.Visible = False
9720  frameMileage.Visible = False


9730  txtIncidentDate.Visible = False
9740  txtDuration.Visible = False
9750  optInPerson.Visible = False
9760  optTelephone.Visible = False
9770  txtDOB.Visible = False
9780  Label3.Visible = False
9790  txtDOInt.Visible = False
9800  Label4.Visible = False
9810  txtTOI.Visible = False
9820  Label5.Visible = False

9830  txtLocation.Visible = False
9840  txtTel.Visible = False


9850  For Each contr In frameTypeOfReport.Controls
9860     contr.Enabled = True
9870     contr.Visible = True
9880  Next
              
9890          txtDateofApp.Visible = False
9900          txtTimeofApp.Visible = False
9910          txtDept.Visible = False
9920          Label29.Visible = False
9930          Label30.Visible = False
9940          Label31.Visible = False

9950  txtWitFirst.SetFocus



9960      On Error GoTo 0
9970  Exit Sub

optContactLetter_Click_Error:

         
9980      Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

9990      zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: optContactLetter_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

10000     Print #1, zMsg

10010     Close #1

            
10020     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub optDR_Click()

10030 On Error GoTo optDR_Click_Error
      Dim zMsg As String

10040     txtSalutation = "Dr."
10050     txtWitAddress.SetFocus

10060     On Error GoTo 0
10070 Exit Sub

optDR_Click_Error:

         
10080     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

10090     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: optDR_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

10100     Print #1, zMsg

10110     Close #1

            
10120     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub optDueDiligence_Click()

10130 On Error GoTo optDueDiligence_Click_Error
      Dim zMsg As String

10140 cmdDueDiligence.Visible = True
10150 cmdPhotoReport.Visible = False
10160 cmdCreateWitness.Visible = False
10170 cmdSubService.Visible = False
10180 cmdTestify.Visible = False
10190 cmdCancel.Enabled = True
10200 optOther.Caption = "Other Report"

      Dim contr As Control

10210 For Each contr In frmWitnessEntry.Controls
10220     If TypeName(contr) = "TextBox" Or TypeName(contr) = "Label" Or TypeName(contr) = "OptionButton" Then
10230         contr.Enabled = False
10240         contr.Visible = False
10250     End If
10260 Next
10270 frameGender.Visible = False
10280 frameGender.Enabled = False
10290 frameServiceInfo.Visible = False
10300 frameContactLetter.Visible = False
10310 frameContactLetter.Enabled = False
10320 frameFaxCover.Visible = False
10330 frameMileage.Visible = False
10340 Label2.Caption = "Last name"
10350 Label1.Caption = "Witness First name"

10360 For Each contr In frameTypeOfReport.Controls
10370    contr.Enabled = True
10380    contr.Visible = True
10390 Next

10400 txtWitFirst.Enabled = True
10410 txtWitFirst.Visible = True
10420 Label1.Enabled = True
10430 Label1.Visible = True

10440 Label2.Enabled = True
10450 Label2.Visible = True
10460 txtDuration.Visible = True
10470 txtDuration.Enabled = True
10480 Label13.Enabled = True
10490 Label13.Visible = True
10500 Label13.Caption = "Duration of Report"
10510 txtWitLast.Enabled = True
10520 txtWitLast.Visible = True

10530 txtWitFirst.SetFocus


10540     On Error GoTo 0
10550 Exit Sub

optDueDiligence_Click_Error:

         
10560     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

10570     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: optDueDiligence_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

10580     Print #1, zMsg

10590     Close #1

            
10600     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub optEmailService_Click()
10610     txtLOS.SetFocus
End Sub

Private Sub optFax_Click()

10620 On Error GoTo optFax_Click_Error
      Dim zMsg As String

10630 cmdDueDiligence.Visible = False
10640 cmdPhotoReport.Visible = False
10650 cmdCreateWitness.Visible = False
10660 cmdSubService.Visible = False
10670 cmdTestify.Visible = False
10680 optOther.Caption = "Other Report"
10690 Label1.Caption = "Witness First name"
10700 Label2.Caption = "Last name"

10710 frameFaxCover.Visible = True
10720 frameFaxCover.Top = 126


      Dim contr As Control

10730 For Each contr In frmWitnessEntry.Controls
10740     If TypeName(contr) = "TextBox" Or TypeName(contr) = "Label" Or TypeName(contr) = "OptionButton" Then
10750         contr.Enabled = False
10760         contr.Visible = False
10770     End If
10780 Next

10790 For Each contr In frameFaxCover.Controls
10800     If TypeName(contr) = "TextBox" Or TypeName(contr) = "Label" Or TypeName(contr) = "OptionButton" Then
10810         contr.Enabled = True
10820         contr.Visible = True
10830     End If
10840 Next
10850 frameGender.Visible = False
10860 frameGender.Enabled = False
10870 frameServiceInfo.Visible = False
10880 frameContactLetter.Visible = False
10890 frameContactLetter.Enabled = False
10900 frameMileage.Visible = False

10910 For Each contr In frameTypeOfReport.Controls
10920    contr.Enabled = True
10930    contr.Visible = True
10940 Next

10950 txtFaxName.SetFocus


10960     On Error GoTo 0
10970 Exit Sub

optFax_Click_Error:

         
10980     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

10990     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: optFax_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

11000     Print #1, zMsg

11010     Close #1

            
11020     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub optFaxService_Click()
11030     txtLOS.SetFocus
End Sub

Private Sub optFemale_Click()

11040 If optTelephone = True Then
11050     optTelephone.SetFocus
11060     Else
11070     optInPerson.SetFocus
11080 End If
          
End Sub



Private Sub optInPerService_Click()
11090 txtLOS.SetFocus
End Sub

Private Sub optInPerson_Change()
11100     txtLocation.SetFocus
End Sub



Private Sub optInPerson_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
11110 Select Case KeyCode
              Case 9 'Tab
11120             KeyCode = 0
11130             txtLocation.SetFocus
11140         Case 13 ' Enter
11150             txtLocation.SetFocus
11160         Case 32 'Space bar
11170             optTelephone.SetFocus
11180         Case Else
                'do nothing
11190     End Select
End Sub




Private Sub optMale_Click()
11200 If optTelephone = True Then
11210     optTelephone.SetFocus
11220     Else
11230     optInPerson.SetFocus
11240 End If
End Sub

Private Sub optMale_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
11250 Select Case KeyCode
              Case 9 'Tab
11260             If optTelephone = True Then
11270                 optTelephone.SetFocus
11280              Else
11290                 optInPerson.SetFocus
11300             End If
11310         Case 13 ' Enter
11320             If optTelephone = True Then
11330                 optTelephone.SetFocus
11340              Else
11350                 optInPerson.SetFocus
11360             End If
              
11370         Case 32 'Space bar
11380             optFemale.SetFocus
11390         Case Else
                'do nothing
11400     End Select
End Sub





Private Sub optMR_Click()
11410     txtSalutation = "Mr."
11420     txtWitAddress.SetFocus
End Sub

Private Sub optMRS_Click()
11430     txtSalutation = "Mrs."
11440     txtWitAddress.SetFocus
End Sub

Private Sub optMS_Click()
11450     txtSalutation = "Ms."
11460     txtWitAddress.SetFocus
End Sub

Private Sub optOther_Click()


11470 On Error GoTo optOther_Click_Error
      Dim zMsg As String
      Dim result As Integer

      Dim Msg As String

11480 cmdDueDiligence.Visible = False
11490 cmdSubService.Visible = False
11500 cmdPhotoReport.Visible = True
11510 cmdPhotoReport.Caption = "Create Other Report"
11520 cmdCreateWitness.Visible = False
11530 cmdTestify.Visible = False
11540 cmdCancel.Enabled = True
11550 frameServiceInfo.Visible = False
11560 frameContactLetter.Visible = False
11570 frameContactLetter.Enabled = False
11580 frameMileage.Visible = True



11590 Msg = Application.InputBox("Enter the type of report, including the word report. (ie Document Report)", "Report Type")
11600   Do While ValidName(Msg) = False
11610       result = MsgBox("Can't have an illegal character such as [ | \ / ? * ( ) : ; ] in the report name.", vbCritical, "Illegal report name")
11620       Msg = Application.InputBox("Enter the type of report, including the word report. (ie Document Report)", "Report Type")
11630   Loop
11640 If Msg = vbNullString Or Msg = "False" Then
11650     cmdPhotoReport.Caption = "Create Photo Report"
11660     optOther.Caption = "Other Report"
11670     optPhotoReport = xlOn
11680     Exit Sub
11690 End If

      'Error check message
11700 optOther.Caption = Msg
11710 cmdPhotoReport.Caption = "Create " & Msg

      Dim contr As Control

11720 For Each contr In frmWitnessEntry.Controls
11730     If TypeName(contr) = "TextBox" Or TypeName(contr) = "Label" Or TypeName(contr) = "OptionButton" Then
11740         contr.Enabled = False
11750         contr.Visible = False
11760     End If
11770 Next

11780 For Each contr In frameTypeOfReport.Controls
11790    contr.Enabled = True
11800    contr.Visible = True
11810 Next
11820 For Each contr In frameMileage.Controls
11830     contr.Enabled = True
11840     contr.Visible = True
11850 Next
11860 frameGender.Visible = False
11870 frameGender.Enabled = False
11880 frameFaxCover.Visible = False
11890 Label13.Enabled = True
11900 txtDuration.Visible = True
11910 txtDuration.Enabled = True
11920 Label13.Visible = True
11930 Label13.Caption = "Duration of Report"
11940 cmdPhotoReport.SetFocus

      'Keep Version number available- also in photo
      'txtVersion.Enabled = True
      'txtVersion.Visible = True
      'Label9.Enabled = True
      'Label9.Visible = True


11950     On Error GoTo 0
11960 Exit Sub

optOther_Click_Error:

         
11970     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

11980     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: optOther_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

11990     Print #1, zMsg

12000     Close #1

            
12010     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub optOtherSalutation_Click()
12020     txtSalutation = InputBox("Please enter the salutation.", "Salutation")
12030     txtWitAddress.SetFocus
End Sub

Private Sub optOtherService_Click()
12040 txtLOS.SetFocus
End Sub

Private Sub optPhotoReport_Click()

12050 On Error GoTo optPhotoReport_Click_Error
      Dim zMsg As String

12060 cmdDueDiligence.Visible = False
12070 cmdPhotoReport.Visible = True
12080 cmdCreateWitness.Visible = False
12090 cmdSubService.Visible = False
12100 cmdTestify.Visible = False
12110 cmdCancel.Enabled = True
12120 frameServiceInfo.Visible = False
12130 frameContactLetter.Visible = False
12140 frameContactLetter.Enabled = False
12150 frameMileage.Visible = False

      'cmdOtherReport.visible = false
12160 optOther.Caption = "Other Report"

      Dim contr As Control

12170 For Each contr In frmWitnessEntry.Controls
12180     If TypeName(contr) = "TextBox" Or TypeName(contr) = "Label" Or TypeName(contr) = "OptionButton" Then
12190         contr.Enabled = False
12200         contr.Visible = False
12210     End If
12220 Next

12230 For Each contr In frameTypeOfReport.Controls
12240    contr.Enabled = True
12250    contr.Visible = True
12260 Next
12270 For Each contr In frameMileage.Controls
12280     contr.Enabled = True
12290     contr.Visible = True
12300 Next
12310 frameGender.Visible = False
12320 frameGender.Enabled = False
12330 frameFaxCover.Visible = False
12340 Label13.Enabled = True
12350 txtDuration.Visible = True
12360 txtDuration.Enabled = True
12370 Label13.Visible = True
12380 Label13.Caption = "Duration of Report"
12390 cmdPhotoReport.SetFocus



12400     On Error GoTo 0
12410 Exit Sub

optPhotoReport_Click_Error:

         
12420     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

12430     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: optPhotoReport_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

12440     Print #1, zMsg

12450     Close #1

            
12460     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub optSubpoena_Click()

12470 On Error GoTo optSubpoena_Click_Error
      Dim zMsg As String

12480 cmdDueDiligence.Visible = False
12490 cmdPhotoReport.Visible = False
12500 cmdCreateWitness.Visible = False
12510 cmdSubService.Visible = True
12520 cmdTestify.Visible = False
12530 cmdCancel.Enabled = False
12540 optOther.Caption = "Other Report"
12550 Label4.Caption = "Date of Service"
12560 Label5.Caption = "Time of Service"
12570 Label1.Caption = "Company name or First name"
12580 Label2.Caption = "Last name"
12590 txtDuration.Visible = False
12600 Label13.Visible = False

      Dim contr As Control

12610 For Each contr In frmWitnessEntry.Controls
12620     If TypeName(contr) = "TextBox" Or TypeName(contr) = "Label" Or TypeName(contr) = "OptionButton" Then
12630         contr.Enabled = True
12640         contr.Visible = True
12650     End If
12660 Next
12670 frameGender.Visible = False
12680 frameGender.Enabled = False
12690 frameContactLetter.Visible = False
12700 frameContactLetter.Enabled = False
12710 frameFaxCover.Visible = False
12720 frameServiceInfo.Visible = True
12730 frameMileage.Visible = True

12740 txtIncidentDate.Visible = False
12750 txtDuration.Visible = False
12760 optInPerson.Visible = False
12770 optTelephone.Visible = False
12780 txtLocation.Visible = False
12790 txtTel.Visible = False


12800 For Each contr In frameTypeOfReport.Controls
12810    contr.Enabled = True
12820    contr.Visible = True
12830 Next
12840 For Each contr In frameMileage.Controls
12850     contr.Enabled = True
12860     contr.Visible = True
12870 Next


12880 txtWitFirst.SetFocus



12890     On Error GoTo 0
12900 Exit Sub

optSubpoena_Click_Error:

         
12910     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

12920     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: optSubpoena_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

12930     Print #1, zMsg

12940     Close #1

            
12950     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub optTelephone_Change()
12960     txtLocation.SetFocus
End Sub

Private Sub optTelephone_Click()
12970 txtLocation.SetFocus
End Sub

Private Sub optTelephone_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
12980 Select Case KeyCode
              Case 9 'Tab
12990             KeyCode = 0
13000             txtLocation.SetFocus
13010         Case 13 ' Enter
13020             txtLocation.SetFocus
13030         Case 32 'Space bar
13040             optInPerson.SetFocus
13050         Case Else
                'do nothing
13060     End Select
End Sub

Private Sub optTestify_Click()

13070 On Error GoTo optTestify_Click_Error
      Dim zMsg As String

13080 cmdDueDiligence.Visible = False
13090 cmdPhotoReport.Visible = False
13100 cmdSubService.Visible = False
13110 cmdCreateWitness.Visible = False
13120 cmdTestify.Visible = True
13130 cmdCancel.Enabled = True

13140 optOther.Caption = "Other Report"
13150 Label4.Caption = "Date of Testimony"
13160 Label5.Caption = "Time of Testimony"




      Dim contr As Control


13170 For Each contr In frmWitnessEntry.Controls
13180     If TypeName(contr) = "TextBox" Or TypeName(contr) = "Label" Or TypeName(contr) = "OptionButton" Then
13190         contr.Enabled = False
13200         contr.Visible = False
13210     End If
13220 Next
13230 Label4.Enabled = True
13240 Label5.Enabled = True
13250 Label4.Visible = True
13260 Label5.Visible = True
13270 txtDOInt.Enabled = True
13280 txtDOInt.Visible = True
13290 txtTOI.Enabled = True
13300 txtTOI.Visible = True
13310 Label13.Enabled = True
13320 Label13.Visible = True
13330 Label13.Caption = "Duration of Testimony"
13340 txtDuration.Enabled = True
13350 txtDuration.Visible = True


13360 frameServiceInfo.Visible = False
13370 frameContactLetter.Visible = False
13380 frameContactLetter.Enabled = False
13390 frameGender.Visible = False
13400 frameGender.Enabled = False
13410 frameFaxCover.Visible = False
13420 frameMileage.Visible = False


13430 For Each contr In frameTypeOfReport.Controls
13440    contr.Enabled = True
13450    contr.Visible = True
13460 Next

13470 txtDOInt.SetFocus



13480     On Error GoTo 0
13490 Exit Sub

optTestify_Click_Error:

         
13500     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

13510     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: optTestify_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

13520     Print #1, zMsg

13530     Close #1

            
13540     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub optUSMail_Click()
13550 txtLOS.SetFocus
End Sub

Private Sub optWitnessReport_Click()

13560 On Error GoTo optWitnessReport_Click_Error
      Dim zMsg As String

13570 cmdDueDiligence.Visible = False
13580 cmdPhotoReport.Visible = False
13590 cmdSubService.Visible = False
13600 cmdTestify.Visible = False
13610 cmdCreateWitness.Visible = True
13620 cmdCancel.Enabled = True
13630 optOther.Caption = "Other Report"
13640 Label4.Caption = "Date of Interview"
13650 Label5.Caption = "Time of Interview"
13660 Label2.Visible = True
13670 Label1.Visible = True
13680 Label2.Caption = "Last name"
13690 Label1.Caption = "Witness First name"


      Dim contr As Control

13700 For Each contr In frmWitnessEntry.Controls
13710     If TypeName(contr) = "TextBox" Or TypeName(contr) = "Label" Or TypeName(contr) = "OptionButton" Then
13720         contr.Enabled = True
13730         contr.Visible = True
13740     End If
13750 Next
13760 frameGender.Visible = True
13770 frameGender.Enabled = True
13780 frameServiceInfo.Visible = False
13790 frameContactLetter.Visible = False
13800 frameContactLetter.Enabled = False
13810 frameFaxCover.Visible = False
13820 frameMileage.Visible = True


13830 txtDuration.Visible = True
13840 Label13.Visible = True
13850 Label13.Caption = "Duration of Interview"
13860 txtWitFirst.SetFocus



13870     On Error GoTo 0
13880 Exit Sub

optWitnessReport_Click_Error:

         
13890     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

13900     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: optWitnessReport_Click Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

13910     Print #1, zMsg

13920     Close #1

            
13930     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub txtDateofApp_AfterUpdate()

13940 On Error GoTo txtDateofApp_AfterUpdate_Error
      Dim zMsg As String

      Dim varDate As Date
13950 If IsDate(txtDateofApp) Then
13960     varDate = DateValue(txtDateofApp)
13970     txtDateofApp = Format(varDate, "MMMM d, yyyy")
13980 Else
13990     MsgBox "Invalid date"
14000     txtDateofApp.SetFocus
14010     Exit Sub
          
14020 End If

14030     On Error GoTo 0
14040 Exit Sub

txtDateofApp_AfterUpdate_Error:

         
14050     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

14060     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDateofApp_AfterUpdate Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

14070     Print #1, zMsg

14080     Close #1

            
14090     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub txtDateofApp_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
14100 On Error GoTo txtDateofApp_DblClick_Error
      Dim zMsg As String

14110 DatePickerForm.Show vbModal
14120     Select Case [DatePickerForm]![CallingForm].Caption
            Case "Form"
14130           If IsaDate(txtDateofApp) = False Then
14140               txtDateofApp = Format(DateValue(Now()), "MMMM d, yyyy")
14150           End If
14160        Case Else
14170           txtDateofApp = [DatePickerForm]![CallingForm].Caption
14180  End Select

14190     Cancel = True

14200     On Error GoTo 0
14210 Exit Sub

txtDateofApp_DblClick_Error:

         
14220     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

14230     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDateofApp_DblClick Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

14240     Print #1, zMsg

14250     Close #1

            
14260     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDateofApp_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

14270 On Error GoTo txtDateofApp_KeyPress_Error
      Dim zMsg As String

14280 If KeyAscii = 43 Then
14290     txtDateofApp = DateAdd("d", 1, txtDateofApp)
14300     KeyAscii = 0
14310     txtDateofApp = Format(DateValue(txtDateofApp), "MMMM d, yyyy")
14320 End If
14330 If KeyAscii = 45 Then
14340     txtDateofApp = DateAdd("d", -1, txtDateofApp)
14350     KeyAscii = 0
14360     txtDateofApp = Format(DateValue(txtDateofApp), "MMMM d, yyyy")
14370 End If


14380     On Error GoTo 0
14390 Exit Sub

txtDateofApp_KeyPress_Error:

         
14400     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

14410     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDateofApp_KeyPress Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

14420     Print #1, zMsg

14430     Close #1

            
14440     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDOB_AfterUpdate()

14450 On Error GoTo txtDOB_AfterUpdate_Error
      Dim zMsg As String

      Dim varDate As Date
14460 If IsDate(txtDOB) Then
14470     varDate = DateValue(txtDOB)
14480     txtDOB = Format(varDate, "MMMM d, yyyy")
14490 Else
14500     MsgBox "Invalid date"
14510     txtDOB.SetFocus
14520     Exit Sub
          
14530 End If

14540     On Error GoTo 0
14550 Exit Sub

txtDOB_AfterUpdate_Error:

         
14560     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

14570     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDOB_AfterUpdate Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

14580     Print #1, zMsg

14590     Close #1

            
14600     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub



Private Sub txtDOB_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtDOB_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Dim TextStr As String
    TextStr = txtDOB.Text

    If KeyCode <> 8 Then ' i.e. not a backspace or /111

        If (Len(TextStr) = 2 Or Len(TextStr) = 5) Then
            TextStr = TextStr & "/"
        End If

    End If
    If KeyCode = 8 Then
        If TextStr = "" Then TextStr = " "
        TextStr = Left(TextStr, Len(TextStr) - 1)
    End If
    txtDOB.Text = TextStr
End Sub


Private Sub txtDOInt_AfterUpdate()

14610 On Error GoTo txtDOInt_AfterUpdate_Error
      Dim zMsg As String

      Dim varDate As Date
14620 If IsDate(txtDOInt) Then
14630     varDate = DateValue(txtDOInt)
14640     txtDOInt = Format(varDate, "MMMM d, yyyy")
14650 Else
14660     MsgBox "Invalid date"
14670     txtDOInt.SetFocus
14680     Exit Sub
          
14690 End If

14700     On Error GoTo 0
14710 Exit Sub

txtDOInt_AfterUpdate_Error:

         
14720     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

14730     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDOInt_AfterUpdate Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

14740     Print #1, zMsg

14750     Close #1

            
14760     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub txtDOInt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

14770 On Error GoTo txtDOInt_DblClick_Error
      Dim zMsg As String
14780 DatePickerForm.Caption = Label4.Caption
14790 DatePickerForm.Show vbModal
14800     Select Case [DatePickerForm]![CallingForm].Caption
            Case "Form"
14810           If IsaDate(txtDOInt) = False Then
14820               txtDOInt = Format(DateValue(Now()), "MMMM d, yyyy")
14830           End If
14840        Case Else
14850           txtDOInt = [DatePickerForm]![CallingForm].Caption
14860  End Select

14870     Cancel = True

14880     On Error GoTo 0
14890 Exit Sub

txtDOInt_DblClick_Error:

         
14900     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

14910     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDOInt_DblClick Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

14920     Print #1, zMsg

14930     Close #1

            
14940     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtDOInt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

14950 On Error GoTo txtDOInt_KeyPress_Error
      Dim zMsg As String

14960 If KeyAscii = 43 Then
14970     txtDOInt = DateAdd("d", 1, txtDOInt)
14980     KeyAscii = 0
14990     txtDOInt = Format(DateValue(txtDOInt), "MMMM d, yyyy")
15000 End If
15010 If KeyAscii = 45 Then
15020     txtDOInt = DateAdd("d", -1, txtDOInt)
15030     KeyAscii = 0
15040     txtDOInt = Format(DateValue(txtDOInt), "MMMM d, yyyy")
15050 End If


15060     On Error GoTo 0
15070 Exit Sub

txtDOInt_KeyPress_Error:

         
15080     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

15090     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtDOInt_KeyPress Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

15100     Print #1, zMsg

15110     Close #1

            
15120     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub txtEndM_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Only allow numbers or a decimal
15130 Select Case KeyAscii
          Case Is = 46
15140         KeyAscii = KeyAscii
15150     Case 48 To 57
15160         KeyAscii = KeyAscii
15170     Case Else
15180         KeyAscii = 0
15190     End Select
End Sub

Private Sub txtFaxNum_Exit(ByVal Cancel As MSForms.ReturnBoolean)

15200 On Error GoTo txtFaxNum_Exit_Error
      Dim zMsg As String

      Dim result As Integer
15210     txtFaxNum.Text = Replace(txtFaxNum, "-", "")
15220     txtFaxNum.Text = Replace(txtFaxNum, "(", "")
15230     txtFaxNum.Text = Replace(txtFaxNum, ")", "")
15240     txtFaxNum.Text = Replace(txtFaxNum, " ", "")
15250     If Len(txtFaxNum) <> 10 And Len(txtFaxNum) <> 0 Then
15260         result = MsgBox("Standard telephone number uses 10 digits, are you sure you want to continue with " & Format(txtFaxNum.Text, "(000) 000-0000") & " ?", vbYesNo, "Confirm number")
15270         If result = vbNo Then
15280             Cancel = True
15290             txtFaxNum.SetFocus
15300         End If
15310     End If
15320     txtFaxNum.Text = Format(txtFaxNum.Text, "(000) 000-0000")

15330     On Error GoTo 0
15340 Exit Sub

txtFaxNum_Exit_Error:

         
15350     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

15360     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtFaxNum_Exit Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

15370     Print #1, zMsg

15380     Close #1

            
15390     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtIncidentDate_AfterUpdate()

15400 On Error GoTo txtIncidentDate_AfterUpdate_Error
      Dim zMsg As String

      Dim varDate As Date
15410 If IsDate(txtIncidentDate) Then
15420     varDate = DateValue(txtIncidentDate)
15430     txtIncidentDate = Format(varDate, "MMMM d, yyyy")
15440 Else
15450     MsgBox "Invalid date"
15460     txtIncidentDate.SetFocus
15470     Exit Sub
          
15480 End If

15490     On Error GoTo 0
15500 Exit Sub

txtIncidentDate_AfterUpdate_Error:

         
15510     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

15520     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtIncidentDate_AfterUpdate Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

15530     Print #1, zMsg

15540     Close #1

            
15550     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub


Private Sub txtIncidentDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

15560 On Error GoTo txtIncidentDate_DblClick_Error
      Dim zMsg As String
15570 DatePickerForm.Caption = "Incident Date"
15580 DatePickerForm.Show vbModal
15590     Select Case [DatePickerForm]![CallingForm].Caption
            Case "Form"
15600           If IsaDate(txtIncidentDate) = False Then
15610               txtIncidentDate = Format(DateValue(Now()), "MMMM d, yyyy")
15620           End If
15630        Case Else
15640           txtIncidentDate = [DatePickerForm]![CallingForm].Caption
15650  End Select

15660     Cancel = True

15670     On Error GoTo 0
15680 Exit Sub

txtIncidentDate_DblClick_Error:

         
15690     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

15700     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtIncidentDate_DblClick Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

15710     Print #1, zMsg

15720     Close #1

            
15730     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtLocation_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

15740 On Error GoTo txtLocation_KeyDown_Error
      Dim zMsg As String

15750 Select Case KeyCode
              Case vbKeyS And Shift = 2 'Ctrl S
15760             txtLocation = txtLocation & "Sacramento, CA "
                              
15770         Case vbKeyC And Shift = 2 'Ctrl C
15780             txtLocation = txtLocation & "Citrus Heights, CA "
                  
15790         Case vbKeyC And Shift = 3 'Ctrl Shift C
15800             txtLocation = txtLocation & "Carmichael, CA "
                  
15810         Case vbKeyE And Shift = 2 'Ctrl E
15820             txtLocation = txtLocation & "Elk Grove, CA "
                  
15830         Case vbKeyG And Shift = 2 'Ctrl G
15840             txtLocation = txtLocation & "Galt, CA "

15850         Case vbKeyR And Shift = 2 'Ctrl R
15860             txtLocation = txtLocation & "Rancho Cordova, CA "

15870         Case vbKeyF And Shift = 2 'Ctrl F
15880             txtLocation = txtLocation & "Folsom, CA "
                  
15890         Case vbKeyF And Shift = 3 'Ctrl Shift F
15900             txtLocation = txtLocation & "Fair Oaks, CA "
       
15910         Case vbKeyA And Shift = 2 'Ctrl A
15920             txtLocation = txtLocation & "Antelope, CA "

15930         Case vbKeyN And Shift = 2 'Ctrl N
15940             txtLocation = txtLocation & "North Highlands, CA "
                  
15950         Case vbKeyO And Shift = 2 'Ctrl O
15960             txtLocation = txtLocation & "Orangevale, CA "
                  
15970         Case vbKeyR And Shift = 3 'Ctrl Shift R
15980             txtLocation = txtLocation & "Roseville, CA "
                  
15990         Case Else
                'do nothing
16000     End Select

16010     On Error GoTo 0
16020     Exit Sub

txtLocation_KeyDown_Error:
16030     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1
           
16040     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtLocation_KeyDown within: Sub - frmWitnessEntry " & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

16050     Print #1, zMsg

16060     Close #1

            
16070     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
          

End Sub



Private Sub txtLOS_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
16080 On Error GoTo txtLOS_KeyDown_Error
      Dim zMsg As String

16090 Select Case KeyCode
              Case vbKeyS And Shift = 2 'Ctrl S
16100             txtLOS = txtLOS & "Sacramento, CA "
                              
16110         Case vbKeyC And Shift = 2 'Ctrl C
16120             txtLOS = txtLOS & "Citrus Heights, CA "
                  
16130         Case vbKeyC And Shift = 3 'Ctrl Shift C
16140             txtLOS = txtLOS & "Carmichael, CA "
                  
16150         Case vbKeyE And Shift = 2 'Ctrl E
16160             txtLOS = txtLOS & "Elk Grove, CA "

16170         Case vbKeyR And Shift = 2 'Ctrl R
16180             txtLOS = txtLOS & "Rancho Cordova, CA "

16190         Case vbKeyF And Shift = 2 'Ctrl F
16200             txtLOS = txtLOS & "Folsom, CA "
       
16210         Case vbKeyA And Shift = 2 'Ctrl A
16220             txtLOS = txtLOS & "Antelope, CA "

16230         Case vbKeyN And Shift = 2 'Ctrl N
16240             txtLOS = txtLOS & "North Highlands, CA "
                  
16250         Case vbKeyR And Shift = 3 'Ctrl Shift R
16260             txtLOS = txtLOS & "Roseville, CA "

16270         Case vbKeyO And Shift = 2 'Ctrl O
16280             txtLOS = txtLOS & "Orangevale, CA "
                  
16290         Case vbKeyG And Shift = 2 'Ctrl G
16300             txtLOS = txtLOS & "Galt, CA "
                  
16310         Case vbKeyF And Shift = 3 'Ctrl Shift F
16320             txtLOS = txtLOS & "Fair Oaks, CA "
                  
                  
16330         Case Else
                'do nothing
16340     End Select

16350     On Error GoTo 0
16360     Exit Sub

txtLOS_KeyDown_Error:
16370     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1
           
16380     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtLOS_KeyDown within: Sub - frmWitnessEntry " & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

16390     Print #1, zMsg

16400     Close #1

            
16410     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
End Sub



Private Sub txtMileageAddress_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
16420 On Error GoTo txtMileageAddress_KeyDown_Error
      Dim zMsg As String

16430 Select Case KeyCode
              Case vbKeyS And Shift = 2 'Ctrl S
16440             txtMileageAddress = txtMileageAddress & "Sacramento, CA "
                              
16450         Case vbKeyC And Shift = 2 'Ctrl C
16460             txtMileageAddress = txtMileageAddress & "Citrus Heights, CA "
                  
16470         Case vbKeyC And Shift = 3 'Ctrl Shift C
16480             txtMileageAddress = txtMileageAddress & "Carmichael, CA "
                  
16490         Case vbKeyE And Shift = 2 'Ctrl E
16500             txtMileageAddress = txtMileageAddress & "Elk Grove, CA "

16510         Case vbKeyR And Shift = 2 'Ctrl R
16520             txtMileageAddress = txtMileageAddress & "Rancho Cordova, CA "

16530         Case vbKeyF And Shift = 2 'Ctrl F
16540             txtMileageAddress = txtMileageAddress & "Folsom, CA "
       
16550         Case vbKeyA And Shift = 2 'Ctrl A
16560             txtMileageAddress = txtMileageAddress & "Antelope, CA "

16570         Case vbKeyN And Shift = 2 'Ctrl N
16580             txtMileageAddress = txtMileageAddress & "North Highlands, CA "
                  
16590         Case vbKeyR And Shift = 3 'Ctrl Shift R
16600             txtMileageAddress = txtMileageAddress & "Roseville, CA "

16610         Case vbKeyO And Shift = 2 'Ctrl O
16620             txtMileageAddress = txtMileageAddress & "Orangevale, CA "
                  
16630         Case vbKeyG And Shift = 2 'Ctrl G
16640             txtMileageAddress = txtMileageAddress & "Galt, CA "
                  
16650         Case vbKeyF And Shift = 3 'Ctrl Shift F
16660             txtMileageAddress = txtMileageAddress & "Fair Oaks, CA "
                  
                  
16670         Case Else
                'do nothing
16680     End Select

16690     On Error GoTo 0
16700     Exit Sub

txtMileageAddress_KeyDown_Error:
16710     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1
           
16720     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtMileageAddress_KeyDown within: Sub - frmWitnessEntry " & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

16730     Print #1, zMsg

16740     Close #1

            
16750     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"


End Sub

Private Sub txtStartM_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
16760 Select Case KeyAscii
          Case Is = 46
16770         KeyAscii = KeyAscii
16780     Case 48 To 57
16790         KeyAscii = KeyAscii
16800     Case Else
16810         KeyAscii = 0
16820     End Select
          
End Sub

Private Sub txtTel_Exit(ByVal Cancel As MSForms.ReturnBoolean)

16830 On Error GoTo txtTel_Exit_Error
      Dim zMsg As String

      Dim result As Integer

16840     txtTel.Text = Replace(txtTel, "-", "")
16850     txtTel.Text = Replace(txtTel, "(", "")
16860     txtTel.Text = Replace(txtTel, ")", "")
16870     txtTel.Text = Replace(txtTel, " ", "")
          
16880     If Len(txtTel) <> 10 And Len(txtTel) <> 0 Then
16890         result = MsgBox("Standard telephone number uses 10 digits, are you sure you want to continue with " & Format(txtTel.Text, "(000) 000-0000") & " ?", vbYesNo, "Confirm number")
16900         If result = vbNo Then
16910             Cancel = True
16920             txtTel.SetFocus
16930         End If
16940     End If
              
16950     txtTel.Text = Format(txtTel.Text, "(000) 000-0000")

16960     On Error GoTo 0
16970 Exit Sub

txtTel_Exit_Error:

         
16980     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

16990     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtTel_Exit Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

17000     Print #1, zMsg

17010     Close #1

            
17020     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtTelS_Exit(ByVal Cancel As MSForms.ReturnBoolean)

17030 On Error GoTo txtTelS_Exit_Error
      Dim zMsg As String

      Dim result As Integer
          
17040     txtTelS.Text = Replace(txtTelS, "-", "")
17050     txtTelS.Text = Replace(txtTelS, "(", "")
17060     txtTelS.Text = Replace(txtTelS, ")", "")
17070     txtTelS.Text = Replace(txtTelS, " ", "")
          
17080     If Len(txtTelS) <> 10 And Len(txtTelS) <> 0 Then
17090         result = MsgBox("Standard telephone number uses 10 digits, are you sure you want to continue with " & Format(txtTelS.Text, "(000) 000-0000") & " ?", vbYesNo, "Confirm number")
17100         If result = vbNo Then
17110             Cancel = True
17120             txtTelS.SetFocus
17130         End If
17140     End If
          
17150     txtTelS.Text = Format(txtTelS.Text, "(000) 000-0000")

17160     On Error GoTo 0
17170 Exit Sub

txtTelS_Exit_Error:

         
17180     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

17190     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtTelS_Exit Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

17200     Print #1, zMsg

17210     Close #1

            
17220     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub




Private Sub txtTimeofApp_Exit(ByVal Cancel As MSForms.ReturnBoolean)

17230 On Error GoTo txtTimeofApp_Exit_Error
      Dim zMsg As String

      Dim result As Integer
17240 If IsTime(txtTimeofApp) = False Then
17250     Cancel = True
17260     txtTimeofApp.SetFocus
17270     result = MsgBox("Invalid Time, check your entry! ", vbOKOnly, "Check time entry")
17280 End If

17290     On Error GoTo 0
17300 Exit Sub

txtTimeofApp_Exit_Error:

         
17310     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

17320     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtTimeofApp_Exit Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

17330     Print #1, zMsg

17340     Close #1

            
17350     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtTimeofApp_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)


17360 On Error GoTo txtTimeofApp_KeyPress_Error
      Dim zMsg As String

      Dim LMin, LMinNew, IntervalAdd As Integer

17370 If KeyAscii = 43 Then
17380     LMin = Minute(txtTimeofApp)
17390     LMinNew = Round(LMin / 5, 0) * 5
17400     IntervalAdd = (LMin - LMinNew) * -1
17410     txtTimeofApp = DateAdd("N", IntervalAdd, txtTimeofApp)
          
17420     txtTimeofApp = DateAdd("N", 5, txtTimeofApp)
17430     KeyAscii = 0
17440     txtTimeofApp = Format(txtTimeofApp, "h:mm AM/PM")
17450 End If

17460 If KeyAscii = 42 Then
         
17470     txtTimeofApp = DateAdd("N", 1, txtTimeofApp)
17480     KeyAscii = 0
17490     txtTimeofApp = Format(txtTimeofApp, "h:mm AM/PM")
17500 End If

17510 If KeyAscii = 47 Then
            
17520     txtTimeofApp = DateAdd("N", -1, txtTimeofApp)
17530     KeyAscii = 0
17540     txtTimeofApp = Format(txtTimeofApp, "h:mm AM/PM")
17550 End If
17560 If KeyAscii = 45 Then
17570     LMin = Minute(txtTimeofApp)
17580     LMinNew = Round(LMin / 5, 0) * 5
17590     IntervalAdd = (LMin - LMinNew) * -1
17600     txtTimeofApp = DateAdd("N", IntervalAdd, txtTimeofApp)
          
17610     txtTimeofApp = DateAdd("N", -5, txtTimeofApp)
17620     KeyAscii = 0
17630     txtTimeofApp = Format(txtTimeofApp, "h:mm AM/PM")
17640 End If

17650     On Error GoTo 0
17660 Exit Sub

txtTimeofApp_KeyPress_Error:

         
17670     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

17680     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtTimeofApp_KeyPress Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

17690     Print #1, zMsg

17700     Close #1

            
17710     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtTOI_Exit(ByVal Cancel As MSForms.ReturnBoolean)

17720 On Error GoTo txtTOI_Exit_Error
      Dim zMsg As String

      Dim result As Integer
17730 If IsTime(txtTOI) = False Then
17740     Cancel = True
17750     txtTOI.SetFocus
17760     result = MsgBox("Invalid Time, check your entry! " & txtTOI, vbOKOnly, "Check time entry")
17770 End If

17780     On Error GoTo 0
17790 Exit Sub

txtTOI_Exit_Error:

         
17800     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

17810     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtTOI_Exit Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

17820     Print #1, zMsg

17830     Close #1

            
17840     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtTOI_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)


17850 On Error GoTo txtTOI_KeyPress_Error
      Dim zMsg As String

      Dim LMin, LMinNew, IntervalAdd As Integer

17860 If KeyAscii = 43 Then
17870     LMin = Minute(txtTOI)
17880     LMinNew = Round(LMin / 5, 0) * 5
17890     IntervalAdd = (LMin - LMinNew) * -1
17900     txtTOI = DateAdd("N", IntervalAdd, txtTOI)
          
17910     txtTOI = DateAdd("N", 5, txtTOI)
17920     KeyAscii = 0
17930     txtTOI = Format(txtTOI, "h:mm AM/PM")
17940 End If
17950 If KeyAscii = 42 Then
         
17960     txtTOI = DateAdd("N", 1, txtTOI)
17970     KeyAscii = 0
17980     txtTOI = Format(txtTOI, "h:mm AM/PM")
17990 End If
18000 If KeyAscii = 45 Then
18010     LMin = Minute(txtTOI)
18020     LMinNew = Round(LMin / 5, 0) * 5
18030     IntervalAdd = (LMin - LMinNew) * -1
18040     txtTOI = DateAdd("N", IntervalAdd, txtTOI)
          
18050     txtTOI = DateAdd("N", -5, txtTOI)
18060     KeyAscii = 0
18070     txtTOI = Format(txtTOI, "h:mm AM/PM")
18080 End If
18090 If KeyAscii = 47 Then
            
18100     txtTOI = DateAdd("N", -1, txtTOI)
18110     KeyAscii = 0
18120     txtTOI = Format(txtTOI, "h:mm AM/PM")
18130 End If



18140     On Error GoTo 0
18150 Exit Sub

txtTOI_KeyPress_Error:

         
18160     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

18170     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtTOI_KeyPress Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

18180     Print #1, zMsg

18190     Close #1

            
18200     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtWitAddress_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
18210 On Error GoTo txtWitAddress_KeyDown_Error
      Dim zMsg As String

18220 Select Case KeyCode
              Case vbKeyS And Shift = 2 'Ctrl S
18230             txtWitCity = "Sacramento"
18240             txtWitState = "CA"
18250             txtWitZip.SetFocus
                  
18260         Case vbKeyC And Shift = 2 'Ctrl C
18270             txtWitCity = "Citrus Heights"
18280             txtWitState = "CA"
18290             txtWitZip.SetFocus
                  
18300         Case vbKeyC And Shift = 3 'Ctrl Shift C
18310             txtWitCity = "Carmichael"
18320             txtWitState = "CA"
18330             txtWitZip.SetFocus
                  
18340         Case vbKeyE And Shift = 2 'Ctrl E
18350             txtWitCity = "Elk Grove"
18360             txtWitState = "CA"
18370             txtWitZip.SetFocus
                  
18380         Case vbKeyR And Shift = 2 'Ctrl R
18390             txtWitCity = "Rancho Cordova"
18400             txtWitState = "CA"
18410             txtWitZip.SetFocus
                  
18420         Case vbKeyR And Shift = 3 'Ctrl Shift R
18430             txtWitCity = "Roseville"
18440             txtWitState = "CA"
18450             txtWitZip.SetFocus
                  
18460         Case vbKeyF And Shift = 2 'Ctrl F
18470             txtWitCity = "Folsom"
18480             txtWitState = "CA"
18490             txtWitZip.SetFocus
                  
18500         Case vbKeyF And Shift = 3 'Ctrl Shift F
18510             txtWitCity = "Fair Oaks"
18520             txtWitState = "CA"
18530             txtWitZip.SetFocus
                  
18540         Case vbKeyA And Shift = 2 'Ctrl A
18550             txtWitCity = "Antelope"
18560             txtWitState = "CA"
18570             txtWitZip.SetFocus
                  
18580         Case vbKeyN And Shift = 2 'Ctrl N
18590             txtWitCity = "North Highlands"
18600             txtWitState = "CA"
18610             txtWitZip.SetFocus
                  
18620         Case vbKeyO And Shift = 2 'Ctrl O
18630             txtWitCity = "Orangevale"
18640             txtWitState = "CA"
18650             txtWitZip.SetFocus
                  
18660         Case vbKeyG And Shift = 2 'Ctrl G
18670             txtWitCity = "Galt"
18680             txtWitState = "CA"
18690             txtWitZip.SetFocus
                  
18700         Case Else
                'do nothing
18710     End Select



18720     On Error GoTo 0
18730     Exit Sub

txtWitAddress_KeyDown_Error:
18740     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1
           
18750     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtWitAddress_KeyDown within: Sub - frmWitnessEntry " & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

18760     Print #1, zMsg

18770     Close #1

            
18780     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
          
End Sub

Private Sub txtWitCity_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
18790 On Error GoTo txtWitCity_KeyDown_Error
      Dim zMsg As String

18800 Select Case KeyCode
              Case vbKeyS And Shift = 2 'Ctrl S
18810             txtWitCity = "Sacramento"
18820             txtWitState = "CA"
18830             txtWitZip.SetFocus
                  
18840         Case vbKeyC And Shift = 2 'Ctrl C
18850             txtWitCity = "Citrus Heights"
18860             txtWitState = "CA"
18870             txtWitZip.SetFocus
                  
18880         Case vbKeyC And Shift = 3 'Ctrl Shift C
18890             txtWitCity = "Carmichael"
18900             txtWitState = "CA"
18910             txtWitZip.SetFocus
                  
18920         Case vbKeyE And Shift = 2 'Ctrl E
18930             txtWitCity = "Elk Grove"
18940             txtWitState = "CA"
18950             txtWitZip.SetFocus
                  
18960         Case vbKeyR And Shift = 2 'Ctrl R
18970             txtWitCity = "Rancho Cordova"
18980             txtWitState = "CA"
18990             txtWitZip.SetFocus
                  
19000         Case vbKeyR And Shift = 3 'Ctrl Shift R
19010             txtWitCity = "Roseville"
19020             txtWitState = "CA"
19030             txtWitZip.SetFocus
                  
19040         Case vbKeyF And Shift = 2 'Ctrl F
19050             txtWitCity = "Folsom"
19060             txtWitState = "CA"
19070             txtWitZip.SetFocus
                  
19080         Case vbKeyF And Shift = 3 'Ctrl Shift F
19090             txtWitCity = "Fair Oaks"
19100             txtWitState = "CA"
19110             txtWitZip.SetFocus
                  
19120         Case vbKeyA And Shift = 2 'Ctrl A
19130             txtWitCity = "Antelope"
19140             txtWitState = "CA"
19150             txtWitZip.SetFocus
                  
19160         Case vbKeyN And Shift = 2 'Ctrl N
19170             txtWitCity = "North Highlands"
19180             txtWitState = "CA"
19190             txtWitZip.SetFocus
                  
19200         Case vbKeyO And Shift = 2 'Ctrl O
19210             txtWitCity = "Orangevale"
19220             txtWitState = "CA"
19230             txtWitZip.SetFocus
                  
19240         Case vbKeyG And Shift = 2 'Ctrl G
19250             txtWitCity = "Galt"
19260             txtWitState = "CA"
19270             txtWitZip.SetFocus
                  
19280         Case Else
                'do nothing
19290     End Select



19300     On Error GoTo 0
19310     Exit Sub

txtWitCity_KeyDown_Error:
19320     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1
           
19330     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtWitCity_KeyDown within: Sub - frmWitnessEntry " & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

19340     Print #1, zMsg

19350     Close #1

            
19360     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"
          
End Sub


Private Sub txtWitFirst_Change()

19370 On Error GoTo txtWitFirst_Change_Error
      Dim zMsg As String

      ' Prevent the following characters < > : " / \ | ? *
      Dim DQ As String
19380 DQ = Chr(34)
19390 txtWitFirst = Replace(txtWitFirst, DQ, "'")
19400 txtWitFirst = Replace(txtWitFirst, "/", "-")
19410 txtWitFirst = Replace(txtWitFirst, "\", "-")



19420     On Error GoTo 0
19430 Exit Sub

txtWitFirst_Change_Error:

         
19440     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

19450     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtWitFirst_Change Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

19460     Print #1, zMsg

19470     Close #1

            
19480     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtWitFirst_Exit(ByVal Cancel As MSForms.ReturnBoolean)
          

19490 On Error GoTo txtWitFirst_Exit_Error
      Dim zMsg As String

19500     If ValidName(txtWitFirst) = False Then
19510         Cancel = True
19520         txtWitFirst.SetFocus
19530         MsgBox "Check for illegal characters!"
19540         Exit Sub
19550     End If
19560 txtWitFirst = Trim(txtWitFirst)
19570 txtWitFirst = ProperCase(txtWitFirst, 0, 0)


19580     On Error GoTo 0
19590 Exit Sub

txtWitFirst_Exit_Error:

         
19600     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

19610     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtWitFirst_Exit Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

19620     Print #1, zMsg

19630     Close #1

            
19640     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtWitLast_Change()

19650 On Error GoTo txtWitLast_Change_Error
      Dim zMsg As String

      ' Prevent the following characters < > : " / \ | ? *
      Dim DQ As String
19660 DQ = Chr(34)
19670 txtWitLast = Replace(txtWitLast, DQ, "'")
19680 txtWitLast = Replace(txtWitLast, "/", "-")
19690 txtWitLast = Replace(txtWitLast, "\", "-")



19700     On Error GoTo 0
19710 Exit Sub

txtWitLast_Change_Error:

         
19720     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

19730     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtWitLast_Change Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

19740     Print #1, zMsg

19750     Close #1

            
19760     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtWitLast_Exit(ByVal Cancel As MSForms.ReturnBoolean)

19770 On Error GoTo txtWitLast_Exit_Error
      Dim zMsg As String

19780 If ValidName(txtWitLast) = False Then
19790         Cancel = True
19800         txtWitLast.SetFocus
19810         MsgBox "Check for illegal characters!"
19820         Exit Sub
19830     End If
19840     txtWitLast = Trim(txtWitLast)
19850     txtWitLast = ProperCase(txtWitLast)
          

19860     On Error GoTo 0
19870 Exit Sub

txtWitLast_Exit_Error:

         
19880     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

19890     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: txtWitLast_Exit Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

19900     Print #1, zMsg

19910     Close #1

            
19920     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub txtWitState_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
19930 Select Case KeyAscii
      Case 65 To 90 ' Cap letters
19940         KeyAscii = KeyAscii
19950     Case 97 To 122  'Lowercase letters : convert to upper case
19960         KeyAscii = KeyAscii - 32
19970     Case Else
19980         KeyAscii = 0
19990 End Select

End Sub


Private Sub UserForm_Initialize()
          

20000 On Error GoTo UserForm_Initialize_Error
      Dim zMsg As String

20010 If IsInternetConnected = False Then
20020   MsgBox "No Network Connection Detected! Shut down and re-start", vbExclamation, "No Connection"
20030   Exit Sub
20040 End If
20050 CenterForm Me
          
20060     optMale = True
20070     optTelephone = True
20080     optWitnessReport = True
20090     cmdDueDiligence.Visible = False
20100     cmdPhotoReport.Visible = False
20110     cmdSubService.Visible = False
20120     frameServiceInfo.Visible = False
20130     frameContactLetter.Visible = False
20140     frameContactLetter.Enabled = False
20150     frameFaxCover.Visible = False
20160     frameMileage.Visible = True
20170     ckBox125 = True
20180     ckAddMileage = False
          
              
20190     txtDOInt = Format(Now(), "MMMM d, yyyy")
20200     txtTOI = Format(Now(), "h:mm AM/PM")
20210     txtWitFirst.SetFocus
20220     Label13.Caption = "Duration of Interview"
          

20230     On Error GoTo 0
20240 Exit Sub

UserForm_Initialize_Error:

         
20250     Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

20260     zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: UserForm_Initialize Within: frmWitnessEntry" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

20270     Print #1, zMsg

20280     Close #1

            
20290     MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

