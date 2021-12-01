VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form diaTmvew
   BackColor = &H80000018&
   BorderStyle = 1 'Fixed Single
   Caption = "Current Time Card"
   ClientHeight = 3180
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 5010
   Icon = "diaTmvew.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 3180
   ScaleWidth = 5010
   ShowInTaskbar = 0 'False
   StartUpPosition = 3 'Windows Default
   Begin VB.CheckBox optRev
      Caption = "Check1"
      Height = 195
      Left = 1560
      TabIndex = 7
      Top = 3000
      Visible = 0 'False
      Width = 855
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 3720
      Top = 3000
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3180
      FormDesignWidth = 5010
   End
   Begin MSFlexGridLib.MSFlexGrid grd
      Height = 2535
      Left = 120
      TabIndex = 6
      ToolTipText = "Press ""Esc"" To Close"
      Top = 360
      Width = 4780
      _ExtentX = 8440
      _ExtentY = 4471
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
      Caption = "Hours"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
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
      Caption = "End"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
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
      Caption = "Start"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
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
      Caption = "Run Account             "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
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
      Caption = "D/I"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
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
Attribute VB_Name = "diaTmvew"
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

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub Form_Activate()
   'If optRev.Value = vbChecked Then
   If iBarOnTop Then
      Move MdiSect.Left + 800, diaHrtme.Top + 1900
   Else
      Move MdiSect.Left + 3600, diaHrtme.Top + 1100
   End If
   '    Else
   '        If iBarOnTop Then
   '            Move MdiSect.Left + 800, diaHdtme.Top + 1900
   '        Else
   '            Move MdiSect.Left + 3600, diaHdtme.Top + 1100
   '        End If
   '    End If
   
End Sub

Private Sub Form_Click()
   cmdCan_Click
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   BackColor = Es_HelpBackGroundColor
   Show
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   WindowState = 1
   Set diaTmvew = Nothing
   
End Sub
