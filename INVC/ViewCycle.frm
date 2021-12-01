VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ViewCycle
   BackColor = &H80000018&
   BorderStyle = 1 'Fixed Single
   Caption = "Current Time Card"
   ClientHeight = 3516
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 4728
   Icon = "ViewCycle.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 3516
   ScaleWidth = 4728
   ShowInTaskbar = 0 'False
   StartUpPosition = 3 'Windows Default
   Begin VB.CheckBox optRev
      Caption = "Check1"
      Height = 195
      Left = 1560
      TabIndex = 6
      Top = 3720
      Visible = 0 'False
      Width = 855
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4320
      Top = 3360
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3516
      FormDesignWidth = 4728
   End
   Begin MSFlexGridLib.MSFlexGrid grd
      Height = 2535
      Left = 120
      TabIndex = 5
      ToolTipText = "Press ""Esc"" To Close"
      Top = 360
      Width = 4395
      _ExtentX = 7747
      _ExtentY = 4466
      _Version = 393216
      Rows = 10
      Cols = 4
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
      Top = 3720
      Width = 915
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Max 1000 Rows Will be Shown)"
      Height = 255
      Index = 1
      Left = 2160
      TabIndex = 9
      Top = 3000
      Width = 2415
   End
   Begin VB.Label lblRows
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Caption = "   "
      Height = 255
      Left = 1200
      TabIndex = 8
      Top = 3000
      Width = 495
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Rows Shown"
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 7
      Top = 3000
      Width = 1215
   End
   Begin VB.Label lblAdt
      BackStyle = 0 'Transparent
      Caption = "Next Count  "
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
      Left = 3600
      TabIndex = 4
      Top = 120
      Width = 975
   End
   Begin VB.Label lblSdt
      BackStyle = 0 'Transparent
      Caption = "Last Count    "
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
      Left = 2640
      TabIndex = 3
      Top = 120
      Width = 855
   End
   Begin VB.Label lblQty
      BackStyle = 0 'Transparent
      Caption = "Std Cost    "
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
      Left = 1800
      TabIndex = 2
      Top = 120
      Width = 735
   End
   Begin VB.Label P
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
      Left = 120
      TabIndex = 1
      Top = 120
      Width = 1575
   End
End
Attribute VB_Name = "ViewCycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub Form_Activate()
   If iBarOnTop Then
      Move MdiSect.Left + 800, MdiSect.ActiveForm.Top + 1900
   Else
      Move MdiSect.Left + 3600, MdiSect.ActiveForm.Top + 1100
   End If
   
End Sub

Private Sub Form_Click()
   cmdCan_Click
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   BackColor = Es_HelpBackGroundColor
   With grd
      .ColAlignment(0) = 0
      .ColWidth(0) = 1600
      .ColWidth(1) = 850
      .ColWidth(2) = 900
      .ColWidth(3) = 900
      
   End With
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   WindowState = 1
   Set ViewCycle = Nothing
   
End Sub
