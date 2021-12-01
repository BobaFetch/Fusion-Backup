VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form VewPsItm
   BackColor = &H8000000C&
   BorderStyle = 1 'Fixed Single
   Caption = "Selected Invoice Items"
   ClientHeight = 3030
   ClientLeft = 3000
   ClientTop = 1710
   ClientWidth = 5565
   ForeColor = &H00800000&
   Icon = "diaPsvew.frx":0000
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 3030
   ScaleWidth = 5565
   ShowInTaskbar = 0 'False
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5040
      Top = 120
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3030
      FormDesignWidth = 5565
   End
   Begin VB.ListBox lstItm
      ForeColor = &H80000012&
      Height = 2400
      Left = 120
      Sorted = -1 'True
      TabIndex = 0
      Top = 480
      Width = 5340
   End
   Begin VB.Line Line1
      X1 = 120
      X2 = 5400
      Y1 = 360
      Y2 = 360
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Quantity                "
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
      Height = 255
      Index = 1
      Left = 1635
      TabIndex = 2
      Top = 120
      Width = 2295
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Item                       "
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 1
      Top = 120
      Width = 1455
   End
End
Attribute VB_Name = "VewPsItm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnLoad As Byte

Private Sub Form_Activate()
   MouseCursor 0
   
End Sub

Private Sub Form_Click()
   Unload Me
   
End Sub


Private Sub Form_Deactivate()
   Unload Me
End Sub

Private Sub Form_Load()
   SetFormSize Me
   Move MdiSect.Left + 2000 + diaARe02a.Left + 400, MdiSect.Top + 400 + diaARe02a.Top + 400
   bOnLoad = True
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   WindowState = 1
   Set VewPsItm = Nothing
   
End Sub



Public Sub FillCombo()
   
End Sub
