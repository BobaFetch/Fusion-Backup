VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form BompBMview
   BackColor = &H80000018&
   BorderStyle = 1 'Fixed Single
   Caption = "Current Parts List Items"
   ClientHeight = 3132
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 5004
   Icon = "BompBMview.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 3132
   ScaleWidth = 5004
   ShowInTaskbar = 0 'False
   StartUpPosition = 3 'Windows Default
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 3720
      Top = 3000
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3132
      FormDesignWidth = 5004
   End
   Begin MSFlexGridLib.MSFlexGrid grd
      Height = 2535
      Left = 120
      TabIndex = 6
      ToolTipText = "Press ""Esc"" To Close"
      Top = 360
      Width = 4780
      _ExtentX = 8446
      _ExtentY = 4466
      _Version = 393216
      Rows = 10
      Cols = 5
      FixedRows = 0
      FixedCols = 0
      ForeColor = 8404992
      FocusRect = 0
      HighLight = 0
      GridLinesFixed = 0
      ScrollBars = 2
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "&Close"
      Height = 435
      Left = 3720
      TabIndex = 0
      Top = 3240
      Width = 915
   End
   Begin VB.Label lblAdt
      BackStyle = 0 'Transparent
      Caption = "Conversion   "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 7.8
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Left = 3840
      TabIndex = 5
      Top = 120
      Width = 975
   End
   Begin VB.Label lblSdt
      BackStyle = 0 'Transparent
      Caption = "Um          "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 7.8
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Left = 3000
      TabIndex = 4
      Top = 120
      Width = 855
   End
   Begin VB.Label lblQty
      BackStyle = 0 'Transparent
      Caption = "Quantity    "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 7.8
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Left = 2160
      TabIndex = 3
      Top = 120
      Width = 735
   End
   Begin VB.Label lblPrt
      BackStyle = 0 'Transparent
      Caption = "Part Number               "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 7.8
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Left = 600
      TabIndex = 2
      Top = 120
      Width = 1455
   End
   Begin VB.Label lblitm
      BackStyle = 0 'Transparent
      Caption = "Seq   "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 7.8
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Left = 120
      TabIndex = 1
      Top = 120
      Width = 495
   End
End
Attribute VB_Name = "BompBMview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub Form_Click()
   cmdCan_Click
   
End Sub

Private Sub Form_Initialize()
   BackColor = ES_ViewBackColor
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   If iBarOnTop Then
      Move MDISect.Left + 800, BompBMe02b.Top + 1900
   Else
      Move MDISect.Left + 3600, BompBMe02b.Top + 1100
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   BompBMe02b.txtQty.SetFocus
   WindowState = 1
   Set BompBMview = Nothing
   
End Sub
