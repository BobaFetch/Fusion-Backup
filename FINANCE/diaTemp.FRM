VERSION 5.00
Begin VB.Form diaTemp
   BorderStyle = 3 'Fixed Dialog
   Caption = "Temp"
   ClientHeight = 3720
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 5775
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H80000007&
   LinkTopic = "Form1"
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 3720
   ScaleWidth = 5775
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 4800
      TabIndex = 0
      TabStop = 0 'False
      Top = 120
      Width = 915
   End
End
Attribute VB_Name = "diaTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' dia -
'
' Notes:
'
' Created: (xxx) mm/dd/yy
' Revisions:
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set dia = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   FillCustomers Me
   Exit Sub
   DiaErr1:
   sProcName = "fillcomb"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
