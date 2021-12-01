VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form VewParts
   BackColor = &H8000000C&
   BorderStyle = 3 'Fixed Dialog
   Caption = "Part Number Search"
   ClientHeight = 3495
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 5625
   Icon = "PartSrch.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 3495
   ScaleWidth = 5625
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton Command1
      Cancel = -1 'True
      Caption = "Command1"
      Height = 255
      Left = 1800
      TabIndex = 5
      TabStop = 0 'False
      Top = 3600
      Width = 1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1
      Height = 2055
      Left = 240
      TabIndex = 2
      ToolTipText = "Select Then Double Click To Insert Part Number Or Select And Press Enter"
      Top = 1200
      Width = 5175
      _ExtentX = 9128
      _ExtentY = 3625
      _Version = 393216
      FixedCols = 0
      AllowBigSelection = 0 'False
      HighLight = 0
      ScrollBars = 2
      SelectionMode = 1
   End
   Begin VB.CommandButton cmdGo
      Caption = "Select"
      Height = 255
      Index = 0
      Left = 4680
      TabIndex = 1
      ToolTipText = "Selects A Maximum Of 300 Items"
      Top = 480
      Width = 735
   End
   Begin VB.TextBox txtPrt
      Height = 285
      Left = 1560
      TabIndex = 0
      Tag = "3"
      ToolTipText = "Leading Characters (1 min). Returns Up To 300 Numbers"
      Top = 480
      Width = 3075
   End
   Begin VB.Label lblSelected
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 4560
      TabIndex = 8
      ToolTipText = "Selects A Maximum Of 300 Items"
      Top = 840
      Width = 855
   End
   Begin VB.Label Label1
      BackStyle = 0 'Transparent
      Caption = "Selected"
      ForeColor = &H80000008&
      Height = 255
      Index = 2
      Left = 3600
      TabIndex = 7
      ToolTipText = "Selects A Maximum Of 300 Items"
      Top = 840
      Width = 1095
   End
   Begin VB.Label Label1
      BackStyle = 0 'Transparent
      Caption = "Enter At Least (1) Leading Character"
      ForeColor = &H80000008&
      Height = 255
      Index = 1
      Left = 240
      TabIndex = 6
      Top = 840
      Width = 3015
   End
   Begin VB.Label Label1
      BackStyle = 0 'Transparent
      Caption = "Part Number Search"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 9.75
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H80000008&
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 4
      Top = 120
      Width = 3015
   End
   Begin VB.Label P
      BackStyle = 0 'Transparent
      Caption = "Part Number(s)"
      Height = 285
      Index = 0
      Left = 240
      TabIndex = 3
      Top = 480
      Width = 1425
   End
End
Attribute VB_Name = "VewParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private Sub cmdGo_Click(Index As Integer)
   GetParts
   
End Sub

Private Sub Command1_Click()
   Unload Me
   
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then GetParts
   bOnLoad = 0
   
End Sub

Private Sub Form_DblClick()
   Unload Me
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   MdiSect.ActiveForm.optVew.Value = vbUnchecked
   Unload Me
   
End Sub


Private Sub Form_Load()
   On Error Resume Next
   If MdiSect.SideBar.Visible = False Then
      Move MdiSect.Left + MdiSect.ActiveForm.Left + 800, MdiSect.Top + 3200
   Else
      Move MdiSect.Left + MdiSect.ActiveForm.Left + 2600, MdiSect.Top + 3600
   End If
   With Grid1
      .Rows = 2
      .ColWidth(0) = 2450
      .ColWidth(1) = 2450
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .Row = 0
      .Col = 0
      .Text = "Part Number"
      .Col = 1
      .Text = "Part Description"
   End With
   bOnLoad = 1
   
End Sub


Public Sub GetParts()
   Dim RdoGet As rdoResultset
   Dim i As Integer
   
   On Error Resume Next
   Grid1.Rows = 1
   Grid1.Row = 1
   sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable WHERE " _
          & "PARTREF LIKE '" & Compress(txtPrt) & "%' "
   bSqlRows = GetDataSet(RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         Do Until .EOF
            If i >= 300 Then Exit Do
            i = i + 1
            Grid1.Rows = i + 1
            Grid1.Col = 0
            Grid1.Row = i
            Grid1.Text = "" & Trim(!PARTNUM)
            Grid1.Col = 1
            Grid1.Text = "" & Trim(!PADESC)
            .MoveNext
         Loop
         .Cancel
      End With
      On Error Resume Next
      lblSelected = i
      If bOnLoad = 1 Then txtPrt.SetFocus _
                   Else Grid1.SetFocus
      bOnLoad = 0
   End If
   Set RdoGet = Nothing
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   MdiSect.ActiveForm.optVew.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub



Private Sub Form_Unload(Cancel As Integer)
   Set VewParts = Nothing
   
End Sub

Private Sub Grid1_DblClick()
   On Error Resume Next
   Err = 0
   Grid1.Col = 0
   MdiSect.ActiveForm.cmbPrt = Grid1.Text
   If Err > 0 Then MdiSect.ActiveForm.txtPrt = Grid1.Text
   Unload Me
   
End Sub


Private Sub Grid1_GotFocus()
   Grid1.Col = 0
   Grid1.Row = 1
   
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error Resume Next
      Err = 0
      Grid1.Col = 0
      MdiSect.ActiveForm.cmbPrt = Grid1.Text
      If Err > 0 Then MdiSect.ActiveForm.txtPrt = Grid1.Text
      Unload Me
   End If
   
End Sub


Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   Err = 0
   If Button = 1 Then
      Grid1.Col = 0
      MdiSect.ActiveForm.cmbPrt = Grid1.Text
      If Err > 0 Then MdiSect.ActiveForm.txtPrt = Grid1.Text
   End If
   
End Sub

Private Sub txtPrt_KeyPress(KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   GetParts
   
End Sub
