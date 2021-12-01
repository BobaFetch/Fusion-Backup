VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form diaSql
   BorderStyle = 3 'Fixed Dialog
   Caption = "ES/2000 ERP"
   ClientHeight = 975
   ClientLeft = 1740
   ClientTop = 4890
   ClientWidth = 3345
   ControlBox = 0 'False
   BeginProperty Font
   Name = "MS Sans Serif"
   Size = 8.25
   Charset = 0
   Weight = 700
   Underline = 0 'False
   Italic = 0 'False
   Strikethrough = 0 'False
   EndProperty
   ForeColor = &H80000008&
   Icon = "Diasql.frx":0000
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 975
   ScaleWidth = 3345
   ShowInTaskbar = 0 'False
   Begin VB.Timer Timer1
      Enabled = 0 'False
      Interval = 8125
      Left = 3000
      Top = 840
   End
   Begin ComctlLib.ProgressBar prg1
      Height = 200
      Left = 360
      TabIndex = 1
      Top = 700
      Width = 2805
      _ExtentX = 4948
      _ExtentY = 344
      _Version = 327682
      Appearance = 0
   End
   Begin Threed.SSPanel Pnl
      Height = 285
      Left = 360
      TabIndex = 0
      Top = 270
      Width = 2805
      _Version = 65536
      _ExtentX = 4948
      _ExtentY = 503
      _StockProps = 15
      Caption = " Opening SQL Server"
      ForeColor = 8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      BorderWidth = 1
      BevelOuter = 1
      RoundedCorners = 0 'False
      FloodShowPct = 0 'False
   End
End
Attribute VB_Name = "diaSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim sYear As String
   sYear = "ES/" & Format$(Now, "yyyy") & " ERP"
   Caption = sYear
   BackColor = Es_FormBackColor
   If Screen.Width < 10000 Then
      Move (MdiSect.Width) - (MdiSect.Width - 2000), MdiSect.Height - 1800
   Else
      Move (MdiSect.Width) - (MdiSect.Width - 2500), MdiSect.Height - 1900
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set diaSql = Nothing
   
End Sub

Private Sub Timer1_Timer()
   'Set for 2min i case it won't go away
   Static b As Byte
   If MdiSect.WindowState = 1 Then Hide
   b = b + 1
   If b = 8 Then Unload Me
   
End Sub
