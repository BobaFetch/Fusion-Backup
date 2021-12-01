VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form VewInvItem
   BackColor = &H8000000C&
   BorderStyle = 1 'Fixed Single
   Caption = "Selected Invoice Items"
   ClientHeight = 2850
   ClientLeft = 3000
   ClientTop = 1710
   ClientWidth = 5565
   Icon = "diaRsvew.frx":0000
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 2850
   ScaleWidth = 5565
   ShowInTaskbar = 0 'False
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5160
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2850
      FormDesignWidth = 5565
   End
   Begin VB.ListBox lstItm
      ForeColor = &H00800000&
      Height = 2400
      Left = 120
      Sorted = -1 'True
      TabIndex = 0
      Top = 360
      Width = 5340
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Quantity                "
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
      Index = 2
      Left = 4080
      TabIndex = 3
      Top = 120
      Width = 1335
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Number                                                 "
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
      Index = 1
      Left = 1635
      TabIndex = 2
      Top = 120
      Width = 2295
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Sales Order Item     "
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
      Index = 0
      Left = 120
      TabIndex = 1
      Top = 120
      Width = 1455
   End
End
Attribute VB_Name = "VewInvItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' vewInvItem - View Invoice Items
'
' Notes:
'
' Created: 12/06/02 (nth)
' Revisions:
'
'
'*********************************************************************************

Const FRMHEIGHT = 3000

Dim bOnLoad As Byte

Private Sub Form_Activate()
   MouseCursor 0
End Sub

Private Sub Form_DblClick()
   Unload Me
End Sub

Private Sub Form_Deactivate()
   Unload Me
End Sub

Private Sub Form_Load()
   On Error Resume Next
   If MdiSect.SideBar.Visible = False Then
      Move MdiSect.Left + MdiSect.ActiveForm.Left + 800, MdiSect.Top + 3200
   Else
      Move MdiSect.Left + MdiSect.ActiveForm.Left + 2600, MdiSect.Top + 3600
   End If
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   WindowState = 1
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set VewInvItem = Nothing
End Sub
