VERSION 5.00
Begin VB.Form VewAPInv
   Caption = "Selected Invoices"
   ClientHeight = 4875
   ClientLeft = 1635
   ClientTop = 3660
   ClientWidth = 8640
   Icon = "VewAPInv.frx":0000
   LinkTopic = "Form1"
   ScaleHeight = 4875
   ScaleWidth = 8640
   Begin VB.ListBox lstInv
      BeginProperty Font
      Name = "Courier New"
      Size = 9.75
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 3900
      Left = 120
      TabIndex = 0
      Top = 840
      Width = 8415
   End
   Begin VB.Line Line1
      X1 = 8100
      X2 = 120
      Y1 = 720
      Y2 = 720
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Name        Invoice #             Date      Posted          Amount"
      BeginProperty Font
      Name = "Courier"
      Size = 9.75
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Index = 0
      Left = 180
      TabIndex = 2
      Top = 480
      Width = 7935
   End
   Begin VB.Label lblCaption
      BackStyle = 0 'Transparent
      Caption = "Caption Goes Here"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 9.75
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00000000&
      Height = 255
      Left = 120
      TabIndex = 1
      Top = 120
      Width = 4695
   End
End
Attribute VB_Name = "VewAPInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                    ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

Option Explicit

'*************************************************************************************
' vewAPInv - View QuickBooks AP Export
' Called by QuickBooks AP Export ( diaAPf06a ) to view a list of invoices
'
' Notes:
'
' Created: 05/04/05 (TEL)
'*************************************************************************************

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
   If Me.ScaleHeight > 0 Then
      lstInv.Height = Me.ScaleHeight - lstInv.Top - 60
   End If
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set VewAPInv = Nothing
End Sub
