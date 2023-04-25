VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------------------------------------------
' Method : Workbook_Open
' Author : gouldd
' Date   : 8/30/2016
' Purpose: Establish the defaults, Populate Combos
'---------------------------------------------------------------------------------------
Private Sub Workbook_Open()


10    On Error GoTo Workbook_Open_Error
      Dim zMsg As String


      Dim strTemplateFileName As String
      Dim strTemplatePath As String
      Dim lngRow As Long
      Dim result As Integer
      Dim strCurrentPath As String
      Dim lngTotal As Long
      Dim lngSubTotal As Long

20       Application.ScreenUpdating = False
30       Application.OnTime TimeValue("7:30 PM"), "EndofNight"
          
40        IOD.Protect UserInterfaceOnly:=True
50        MileageLog.Protect UserInterfaceOnly:=True
60        Attorneys.Protect UserInterfaceOnly:=True
70        Files.Protect UserInterfaceOnly:=True
          
          
80        InvestigationLog.Activate
90        ActiveSheet.Protect UserInterfaceOnly:=True
           
         
100      strCurrentPath = Application.ActiveWorkbook.Path
110      If ActiveWorkbook.Name <> Files.Cells(35, 2).Value Then
120           Files.Cells(35, 2).Value = ActiveWorkbook.Name
130      End If
         
         'If in the ICMS directory - need to save in user directory
         'Check hidden sheet files for parameters
         
140       If Files.Cells(20, 2).Value = "" Or Files.Cells(20, 2).Value = "New User" Or Files.Cells(37, 2).Value = False Or InStr(1, strCurrentPath, "Investigations\ICMS", vbTextCompare) > 0 Then
150           frmNewUser.Show
160       End If
170       If Not StrComp(Files.Cells(39, 2).Value, Environ("UserName"), vbTextCompare) = 0 Then
180           UserChange ("ThisWorkbook")
190       End If
200    For lngRow = 1 To 35
                      
210               strTemplatePath = Files.Cells(lngRow, 2).Value
220               strTemplateFileName = strTemplatePath
                 
230              If InStr(1, Files.Cells(lngRow, 1).Value, "Folder") > 0 And FolderExists(strTemplatePath) = False Then
240                   result = MsgBox("Please select the " & Files.Cells(lngRow, 1).Value & "!", vbOKCancel, "Need the file or path")
250                   If result = vbCancel Then GoTo Nexti
260                   strTemplatePath = PathPicked(Files.Cells(lngRow, 1).Value)
270                   Files.Cells(lngRow, 2).Value = strTemplatePath
280               End If
                 
                 
290                   If InStr(1, Files.Cells(lngRow, 1).Value, "Template") > 0 And TemplateExists(strTemplateFileName) = False Then
                          'If InStr(1, Files.Cells(lngRow, 1).Value, "Template Folder") > 0 Then GoTo Nexti:
300                       result = MsgBox("Please select the " & Files.Cells(lngRow, 1).Value & "!", vbOKCancel, "Need the file or path")
310                       If result = vbCancel Then GoTo Nexti
320                       strTemplateFileName = FilePicked(Files.Cells(lngRow, 1).Value)
330                       Files.Cells(lngRow, 2).Value = strTemplateFileName
340                   End If
                  
                         
Nexti:
350        Next lngRow
          'If Not ActiveWorkbook.ReadOnly Then
360         PopulateCombo
370         SortAttorneys
          'End If
          
380       lngTotal = Application.WorksheetFunction.Sum(Columns("W:W")) + Application.WorksheetFunction.Sum(Columns("X:X"))
          
390       AwardCriteria lngTotal, "Int"
              
400       lngSubTotal = Application.WorksheetFunction.Sum(Columns("AA:AA")) + Application.WorksheetFunction.Sum(Columns("AB:AB"))
          
410       AwardCriteria lngSubTotal, "Sub"
          
420       Application.ScreenUpdating = True


430       On Error GoTo 0
440   Exit Sub

Workbook_Open_Error:

         
450       Open "W:\Investigations\ICMS\ErrorLogs\ICMSErrorLog.txt" For Append As #1

460       zMsg = Now & " " & Files.Cells(20, 2).Value & " Line: " & _
                    Format(Erl, "###") & vbCrLf & _
                    "Procedure: Workbook_Open Within: ThisWorkbook" & vbCrLf & Err.Number & ":" & Err.Description & vbCrLf

470       Print #1, zMsg

480       Close #1

            
490       MsgBox zMsg, vbOKOnly + vbCritical, "Untrapped Error:"

End Sub

Private Sub Workbook_Activate()

End Sub

Private Sub Workbook_Deactivate()
500       Application.CommandBars("Cell").Reset
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

510   ThisWorkbook.Save
End Sub




