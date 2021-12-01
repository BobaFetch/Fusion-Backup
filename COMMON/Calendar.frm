VERSION 5.00
Object = "{2D0BBCE1-4187-11D2-81AA-00AA00A8932E}#1.0#0"; "msvbcldr.ocx"
Begin VB.Form Calendar
   AutoRedraw = -1 'True
   BorderStyle = 3 'Fixed Dialog
   ClientHeight = 2025
   ClientLeft = 2685
   ClientTop = 2265
   ClientWidth = 2235
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 2025
   ScaleWidth = 2235
   ShowInTaskbar = 0 'False
   Begin MSVBCalendar.Calendar Calendar1
      Height = 2055
      Left = 0
      TabIndex = 1
      Top = 0
      Width = 2295
      _ExtentX = 4048
      _ExtentY = 3625
      Day = 1
      Month = 9
      Year = 1998
      BeginProperty DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      DayNameFormat = 0
      DayColor = 8388608
      DayNameColor = 8388608
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 255
      Left = 480
      TabIndex = 0
      Top = 2040
      Width = 1215
   End
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Private Sub Calendar1_Click()
   On Error Resume Next
   MdiSect.ActiveForm.ActiveControl.List(0) = Format(Calendar1.Value, "mm/dd/yy")
   MdiSect.ActiveForm.ActiveControl.Text = Format(Calendar1.Value, "mm/dd/yy")
   
End Sub

Private Sub Calendar1_DblClick()
   Hide
   
End Sub


Private Sub Calendar1_LostFocus()
   Unload Me
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub Form_Deactivate()
   Hide
   
End Sub

Private Sub Form_Load()
   '
   Show
   
End Sub

Private Sub Form_LostFocus()
   Unload Me
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Hide
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' @@@TEL bCalendar = False
   Set Calendar = Nothing
   
End Sub
