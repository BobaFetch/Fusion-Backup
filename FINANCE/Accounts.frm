VERSION 5.00
Begin VB.Form Accounts
   BorderStyle = 3 'Fixed Dialog
   ClientHeight = 2400
   ClientLeft = 45
   ClientTop = 45
   ClientWidth = 4320
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 2400
   ScaleWidth = 4320
   ShowInTaskbar = 0 'False
   StartUpPosition = 3 'Windows Default
   Begin VB.CommandButton cmdCan
      Caption = "&Close"
      Height = 495
      Left = 3000
      TabIndex = 1
      Top = 360
      Visible = 0 'False
      Width = 975
   End
   Begin VB.ListBox Accounts1
      Height = 2400
      Left = 0
      TabIndex = 0
      Top = 0
      Width = 4305
   End
End
Attribute VB_Name = "Accounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

Option Explicit

' See the UpdateTables prodecure for database revisions

' Width  = 4000
' Height = 2500

Private Sub Accounts1_Click()
   Dim sAcct As String
   sAcct = Trim(Left(Accounts1.List(Accounts1.ListIndex), _
           InStr(1, Accounts1.List(Accounts1.ListIndex), Chr(9)) - 1))
   On Error Resume Next
   MdiSect.ActiveForm.ActiveControl.List(0) = sAcct
   MdiSect.ActiveForm.ActiveControl.Text = sAcct
End Sub

Private Sub Accounts1_DblClick()
   Hide
End Sub

Private Sub Accounts1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Dim rdoAct As rdoResultset
   sSql = "Qry_FillAccountCombo"
   bSqlRows = GetDataSet(rdoAct, ES_FORWARD)
   If bSqlRows Then
      Accounts1.Clear
      With rdoAct
         While Not .EOF
            'AddComboStr Accounts1.hWnd, "" & Trim(!GLACCTNO)
            Accounts1.AddItem "" & Trim(!GLACCTNO) & Chr(9) & Chr(9) & "" & Trim(!GLDESCR)
            .MoveNext
         Wend
         .Cancel
      End With
   End If
   Set rdoAct = Nothing
End Sub

Private Sub Form_Deactivate()
   Hide
End Sub

Private Sub Form_LostFocus()
   Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Hide
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Accounts = Nothing
End Sub
