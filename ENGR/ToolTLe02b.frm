VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form ToolTLe02b
   BorderStyle = 3 'Fixed Dialog
   Caption = "Tool Quantity"
   ClientHeight = 1536
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 5208
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H8000000F&
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 1536
   ScaleWidth = 5208
   ShowInTaskbar = 0 'False
   Begin VB.TextBox txtQty
      Height = 285
      Left = 3600
      TabIndex = 3
      Tag = "1"
      Top = 720
      Width = 615
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 4200
      TabIndex = 0
      Top = 0
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 6600
      Top = 4200
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 1536
      FormDesignWidth = 5208
   End
   Begin VB.Label lblIndex
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Left = 960
      TabIndex = 4
      Top = 240
      Width = 1335
   End
   Begin VB.Label lblTool
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 960
      TabIndex = 2
      Top = 720
      Width = 2415
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Tool"
      ForeColor = &H00400000&
      Height = 255
      Index = 2
      Left = 240
      TabIndex = 1
      Top = 720
      Width = 855
   End
End
Attribute VB_Name = "ToolTLe02b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnload As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   ToolTLe02a.cmdUpd.Enabled = True
   Unload Me
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move ToolTLe02a.Left + 1000, ToolTLe02a.Top + 2000
   FormatControls
   bOnload = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ToolTLe02a.lstQty.List(Val(lblIndex)) = txtQty
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ToolTLe02b = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = Format(Abs(Val(txtQty)), "#####0")
   If Val(txtQty) < 1 Then
      Beep
      txtQty = "1"
   End If
   
End Sub
