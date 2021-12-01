VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ViewParts 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Number Search"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "PartSrch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleMode       =   0  'User
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbLvl 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   4920
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Part Type (Level) 0 For All"
      Top             =   720
      Width           =   612
   End
   Begin VB.CheckBox optDsc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "Search By &Description"
      Height          =   195
      Left            =   220
      TabIndex        =   9
      ToolTipText     =   "Click Or ALT ""D"""
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2532
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Select Then Double Click To Insert Part Number Or Select And Press Enter (Exit)"
      Top             =   1080
      Width           =   6500
      _ExtentX        =   11456
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "S&elect"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Selects A Maximum Of 300 Items"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Leading Characters (1 min). Returns Up To 300 Numbers"
      Top             =   720
      Width           =   2835
   End
   Begin VB.Label lblControl 
      Height          =   252
      Left            =   480
      TabIndex        =   11
      Top             =   4200
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "The Actual Form.BackColor uses ES_ViewBackColor"
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblSelected 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   6000
      TabIndex        =   8
      ToolTipText     =   "Selects A Maximum Of 300 Items"
      Top             =   360
      Width           =   732
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   4920
      TabIndex        =   7
      ToolTipText     =   "Selects A Maximum Of 300 Items"
      Top             =   360
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1665
   End
End
Attribute VB_Name = "ViewParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/17/04 added passing search and formated backcolor
'11/19/04 added checking and passing a description
'3/20/06 Fixed the selection
'   Added Part Type (PALEVEL)
'1/10/07 Expanded Part and Desc width 7.1.6
Option Explicit
Dim bOnLoad As Byte
Dim bDesc As Byte
Dim bText As Byte

Private Sub cmbLvl_KeyPress(KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub cmbLvl_LostFocus()
   If Val(cmbLvl) > 8 Then
      Beep
      cmbLvl = cmbLvl.List(0)
   End If
   
End Sub


Private Sub cmdGo_Click(Index As Integer)
   If bOnLoad = 0 Then GetParts
   
End Sub

Private Sub Command1_Click()
   Unload Me
   
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then
      GetParts
      P(0).Caption = "Part Number(s)"
      bOnLoad = 0
      
   End If
   
End Sub

Private Sub Form_DblClick()
   Unload Me
   
End Sub

Private Sub Form_Deactivate()
   'MdiSect.ActiveForm.optVew.Value = vbUnchecked
   On Error Resume Next
   MdiSect.ActiveForm.Refresh
   Unload Me
   
End Sub


Private Sub Form_Initialize()
   Dim iLeft As Integer
   BackColor = ES_ViewBackColor
   optDsc.BackColor = BackColor
   If iBarOnTop = 0 Then
      Left = MdiSect.SideBar.Width + MdiSect.ActiveForm.ActiveControl.Left - 3075
      Top = (640) + (MdiSect.ActiveForm.Top + MdiSect.ActiveForm.ActiveControl.Top + 720)
   Else
      Left = MdiSect.ActiveForm.ActiveControl.Left - 3075
      Top = (MdiSect.TopBar.Height + 640) + (MdiSect.ActiveForm.Top + MdiSect.ActiveForm.ActiveControl.Top + 720)
   End If
   
   
End Sub

Private Sub Form_Load()
   AlwaysOnTop hWnd, True
   Dim bByte As Byte
   
   On Error Resume Next
   '11/04/04 see Initialize
   '    If MdiSect.SideBar.Visible = False Then
   '        Move MdiSect.Left + MdiSect.ActiveForm.Left + 800, MdiSect.Top + 2200
   '    Else
   '        Move MdiSect.Left + MdiSect.ActiveForm.Left + 2600, MdiSect.Top + 3200
   '    End If
   With Grid1
      .Rows = 2
      .ColWidth(0) = 2540
      .ColWidth(1) = 2650
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .Row = 0
      .Col = 0
      .Text = "Part Number"
      .Col = 1
      .Text = "Part Description"
      .Col = 2
      .Text = "Qoh"
      .Col = 0
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
   End With
   For bByte = 0 To 7
      cmbLvl.AddItem str$(bByte)
   Next
   cmbLvl.AddItem str$(bByte)
   cmbLvl.ListIndex = 0
   bDesc = CheckDescription()
   bOnLoad = 1
   
End Sub


Private Sub GetParts()
   Dim RdoGet As rdoResultset
   Dim iRow As Integer
   Dim sSearchStr As String
   
   On Error Resume Next
   Grid1.Enabled = False
   Grid1.Rows = 1
   Grid1.Row = 1
   
'   If txtPrt <> "ALL" Then sSearchStr = UCase$(Compress(txtPrt))
   If optDsc.value = vbUnchecked Then
      If txtPrt <> "ALL" Then sSearchStr = UCase$(Compress(txtPrt))
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAQOH FROM PartTable WHERE " _
             & "(PARTREF LIKE '" & sSearchStr & "%' "
      If Val(cmbLvl) > 0 Then sSql = sSql & "AND PALEVEL=" & cmbLvl
      sSql = sSql & ") ORDER BY PARTREF"
   Else
      If txtPrt <> "ALL" Then sSearchStr = UCase$(txtPrt)   'We don't want to compress this if it's a description
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAQOH FROM PartTable WHERE " _
             & "(UPPER(PADESC) LIKE '" & sSearchStr & "%' "
      If Val(cmbLvl) > 0 Then sSql = sSql & "AND PALEVEL=" & cmbLvl
      sSql = sSql & ") ORDER BY PADESC,PARTREF"
   End If
   bSqlRows = GetDataSet(RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         Do Until .EOF
            If iRow >= 300 Then Exit Do
            iRow = iRow + 1
            Grid1.Rows = iRow + 1
            Grid1.Row = iRow
            Grid1.Col = 0
            If optDsc.value = vbUnchecked Then
               Grid1.Text = "" & Trim(!PartNum)
               Grid1.Col = 1
               Grid1.Text = "" & Trim(!PADESC)
            Else
               Grid1.Text = "" & Trim(!PADESC)
               Grid1.Col = 1
               Grid1.Text = "" & Trim(!PartNum)
            End If
            Grid1.Col = 2
            'Grid1.Text = Format(!PAQOH, "######0.000")
            Grid1.Text = !PAQOH
            .MoveNext
         Loop
         ClearResultSet RdoGet
      End With
      On Error Resume Next
      lblSelected = iRow
      Grid1.Col = 0
      bOnLoad = 0
   End If
   If Grid1.Rows > 1 Then
      Grid1.Enabled = True
      Grid1.SetFocus
      Grid1.Row = 1
   End If
   Set RdoGet = Nothing
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   AlwaysOnTop hWnd, False
   MdiSect.ActiveForm.optVew.value = vbUnchecked
   If lblControl <> "" Then
      If UCase$(lblControl) = "TXTPRT" Then
         MdiSect.ActiveForm.txtPrt.SetFocus
      Else
         MdiSect.ActiveForm.cmbPrt.SetFocus
      End If
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub



Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   MdiSect.ActiveForm.optPrn.Refresh
   MdiSect.ActiveForm.optDis.Refresh
   Err.Clear
   Set ViewParts = Nothing
   
End Sub

Private Sub Grid1_Click()
   On Error Resume Next
   If Grid1.Rows > 1 Then
      If MdiSect.ActiveForm.Name = "InvcINp02a" Then
         If optDsc.value = vbChecked Then
            Grid1.Col = 0
         Else
            Grid1.Col = 1
         End If
         MdiSect.ActiveForm.txtPrt = Grid1.Text
         Exit Sub
      End If
      If optDsc = vbUnchecked Then Grid1.Col = 0 _
                  Else Grid1.Col = 1
      If bText = 0 Then
         MdiSect.ActiveForm.cmbPrt = Grid1.Text
      Else
         MdiSect.ActiveForm.txtPrt = Grid1.Text
      End If
      If bDesc = 1 Then
         If optDsc.value = vbUnchecked Then
            Grid1.Col = 1
         Else
            Grid1.Col = 0
         End If
         MdiSect.ActiveForm.lblDsc = Trim(Grid1.Text)
      End If
      Grid1.Col = 0
   End If
   
End Sub

Private Sub Grid1_DblClick()
   Dim bByte As Byte
   On Error Resume Next
   If Grid1.Rows > 1 Then
      If MdiSect.ActiveForm.Name = "InvcINp02a" Then
         If optDsc.value = vbChecked Then
            Grid1.Col = 0
         Else
            Grid1.Col = 1
         End If
         MdiSect.ActiveForm.txtPrt = Grid1.Text
         bByte = 1
         GoTo DiaErr1
      End If
      Grid1.Col = 1
      MdiSect.ActiveForm.lblDsc = Trim(Grid1.Text)
      Unload Me
   End If
   Exit Sub
DiaErr1:
   If bByte = 1 Then Unload Me
   
End Sub


Private Sub Grid1_GotFocus()
   If Grid1.Rows > 1 Then
      Grid1.Col = 0
      Grid1.Row = 1
   End If
   
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
   Dim bByte As Byte
   If Grid1.Rows > 1 Then
      If KeyAscii = vbKeyReturn Then
         On Error Resume Next
         If MdiSect.ActiveForm.Name = "InvcINp02a" Then
            If optDsc.value = vbChecked Then
               Grid1.Col = 0
            Else
               Grid1.Col = 1
            End If
            MdiSect.ActiveForm.txtPrt = Grid1.Text
            bByte = 1
            GoTo DiaErr1
         End If
         If optDsc = vbUnchecked Then Grid1.Col = 0 _
                     Else Grid1.Col = 1
         If bText = 0 Then
            MdiSect.ActiveForm.cmbPrt = Grid1.Text
         Else
            MdiSect.ActiveForm.txtPrt = Grid1.Text
         End If
         If bDesc = 1 Then
            If optDsc.value = vbUnchecked Then
               Grid1.Col = 1
            Else
               Grid1.Col = 0
            End If
            MdiSect.ActiveForm.lblDsc = Trim(Grid1.Text)
         End If
         Unload Me
      End If
   End If
   Exit Sub
DiaErr1:
   If bByte = 1 Then Unload Me
   
End Sub


Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '    On Error Resume Next
   '    If Grid1.Rows > 1 Then
   '        If Button = 1 Then
   '            If MdiSect.ActiveForm.Name = "InvcINp02a" Then
   '                If optDsc.Value = vbChecked Then
   '                    Grid1.Col = 0
   '                Else
   '                    Grid1.Col = 1
   '                End If
   '                MdiSect.ActiveForm.txtPrt = Grid1.Text
   '                Exit Sub
   '            End If
   '            If optDsc = vbUnchecked Then Grid1.Col = 0 _
   '                Else Grid1.Col = 1
   '            Hide
   '            If bText = 0 Then
   '                MdiSect.ActiveForm.cmbPrt = Grid1.Text
   '            Else
   '                MdiSect.ActiveForm.txtPrt = Grid1.Text
   '            End If
   '            If bDesc = 1 Then
   '                If optDsc.Value = vbUnchecked Then
   '                    Grid1.Col = 1
   '                Else
   '                    Grid1.Col = 0
   '                 End If
   '                 MdiSect.ActiveForm.lblDsc = Trim(Grid1.Text)
   '            End If
   '        End If
   '        Grid1.Col = 0
   '    End If
   '
End Sub

Private Sub optDsc_Click()
   Grid1.Rows = 1
   Grid1.Row = 0
   If Trim(P(0).Caption) = "Part Number(s)" Then
      P(0).Caption = "Part Description(s)"
      With Grid1
         .Col = 0
         .Text = "Part Description"
         .Col = 1
         .Text = "Part Number"
      End With
   Else
      P(0).Caption = "Part Number(s)"
      With Grid1
         .Col = 0
         .Text = "Part Number"
         .Col = 1
         .Text = "Part Description"
      End With
   End If
   
End Sub

Private Sub txtPrt_GotFocus()
   SelectFormat Me
   
End Sub

Private Sub txtPrt_KeyPress(KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If bOnLoad = 0 Then GetParts
   
End Sub



'11/19/04

Private Function CheckDescription()
   On Error GoTo DiaErr1
   Dim iControl As Integer
   CheckDescription = 0
   For iControl = 0 To MdiSect.ActiveForm.Controls.Count - 1
      If MdiSect.ActiveForm.Controls(iControl).Name = "lblDsc" Then
         CheckDescription = 1
      Else
         If MdiSect.ActiveForm.Controls(iControl).Name = "txtPrt" Then bText = 1
      End If
   Next
   Exit Function
   
DiaErr1:
   CheckDescription = 0
   
End Function
