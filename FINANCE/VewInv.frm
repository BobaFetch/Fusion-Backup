VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form VewInv
   BackColor = &H8000000C&
   BorderStyle = 3 'Fixed Dialog
   Caption = "Selected Invoices"
   ClientHeight = 3375
   ClientLeft = 1620
   ClientTop = 3645
   ClientWidth = 4935
   Icon = "VewInv.frx":0000
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 3375
   ScaleWidth = 4935
   ShowInTaskbar = 0 'False
   Begin VB.ListBox lstInv
      Height = 2400
      Left = 120
      TabIndex = 0
      Top = 840
      Width = 4695
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4440
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3375
      FormDesignWidth = 4935
   End
   Begin VB.Line Line1
      X1 = 4800
      X2 = 120
      Y1 = 720
      Y2 = 720
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Amount     "
      Height = 255
      Index = 3
      Left = 3840
      TabIndex = 5
      Top = 480
      Width = 975
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Invoice #"
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 4
      Top = 480
      Width = 1095
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Customer                                               "
      Height = 255
      Index = 1
      Left = 2280
      TabIndex = 3
      Top = 480
      Width = 1215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Date           "
      Height = 255
      Index = 2
      Left = 1080
      TabIndex = 2
      Top = 480
      Width = 975
   End
   Begin VB.Label lblCaption
      BackStyle = 0 'Transparent
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 9.75
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = -1 'True
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H80000009&
      Height = 255
      Left = 120
      TabIndex = 1
      Top = 120
      Width = 4695
   End
End
Attribute VB_Name = "VewInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
   Set VewInv = Nothing
End Sub
