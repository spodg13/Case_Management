VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePickerForm 
   Caption         =   "Date picker"
   ClientHeight    =   2808
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3840
   OleObjectBlob   =   "DatePickerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatePickerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Public WithEvents Calendar1 As cCalendar
Attribute Calendar1.VB_VarHelpID = -1

Public Target As Control

Private Sub Calendar1_Click()
10        Call CloseDatePicker(True)
End Sub

Private Sub Calendar1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
20        If KeyCode = vbKeyEscape Then
30            Call CloseDatePicker(False)
40        End If
End Sub


Private Sub UserForm_Initialize()
50        If Calendar1 Is Nothing Then
60            Set Calendar1 = New cCalendar
70            With Calendar1
80                .Add_Calendar_into_Frame Me.Frame1
90                .UseDefaultBackColors = False
100               .DayLength = 3
110               .MonthLength = mlENShort
120               .Height = 120
130               .Width = 180
140               .GridFont.Size = 7
150               .DayFont.Size = 7
160               .Refresh
170           End With
180           Me.Height = 153 'Win7 Aero
190           Me.Width = 197
200       End If
        
          
End Sub

Private Sub UserForm_Activate()
          
          'If IsDate(Target.Value) Then
          '    Calendar1.Value = Target.Value
          'End If
          
          'Call MoveToTarget
210       CallingForm.Caption = Calendar1.Value
          
          
End Sub


Public Sub MoveToTarget()
          Dim dLeft As Double, dTop As Double

220       dLeft = Target.Left - ActiveWindow.VisibleRange.Left + ActiveWindow.Left
230       If dLeft > Application.Width - Me.Width Then
240           dLeft = Application.Width - Me.Width
250       End If
260       dLeft = dLeft + Application.Left
          
270       dTop = Target.Top - ActiveWindow.VisibleRange.Top + ActiveWindow.Top
280       If dTop > Application.Height - Me.Height Then
290           dTop = Application.Height - Me.Height
300       End If
310       dTop = dTop + Application.Top
          
320       Me.Left = IIf(dLeft > 0, dLeft, 0)
330       Me.Top = IIf(dTop > 0, dTop, 0)
End Sub


Sub CloseDatePicker(Save As Boolean)
      'Dim FormName As String
      'Dim ControlName As String

      'FormName = CallingForm.Caption
      'ControlName = CallingTextBox.Caption

      'FormName!Controls [ControlName] = Calendar1.Value
      '*****************************************
          
          'If Save And Not Target Is Nothing And IsDate(Calendar1.Value) Then
          '    Target.Value = Calendar1.Value
          'End If
          'Set Target = Nothing
340       CallingForm.Caption = Calendar1.Value
350       Me.Hide
End Sub


