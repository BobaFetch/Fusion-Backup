VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SysDate
   BorderStyle = 3 'Fixed Dialog
   ClientHeight = 2256
   ClientLeft = 36
   ClientTop = 36
   ClientWidth = 2676
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 2256
   ScaleWidth = 2676
   ShowInTaskbar = 0 'False
   StartUpPosition = 3 'Windows Default
   Begin MSComCtl2.MonthView MonthView
      Height = 2256
      Left = 0
      TabIndex = 0
      Top = 0
      Width = 2664
      _ExtentX = 4699
      _ExtentY = 3979
      _Version = 393216
      ForeColor = -2147483630
      BackColor = -2147483633
      Appearance = 1
      StartOfWeek = 45023233
      CurrentDate = 38986
      MinDate = 4019
   End
End
Attribute VB_Name = "SysDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Deactivate()
   Unload Me
   
End Sub

Private Sub Form_Load()
   If IsDate(MDISect.ActiveForm.ActiveControl) Then
      MonthView.Value = MDISect.ActiveForm.ActiveControl
   Else
      MonthView.Value = Format(Now, "mm/dd/yy")
   End If
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   bSysCalendar = False
   Set SysDate = Nothing
   
End Sub


Private Sub MonthView_DateClick(ByVal DateClicked As Date)
   MDISect.ActiveForm.ActiveControl = Format(MonthView.Value, "mm/dd/yy")
   
End Sub

Private Sub MonthView_DateDblClick(ByVal DateDblClicked As Date)
   Form_Deactivate
   
End Sub
