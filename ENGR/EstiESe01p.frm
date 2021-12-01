VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESe01p 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add A Part Number"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   ClipControls    =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "EstiESe01p.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESe01p.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtMatAlloy 
      Height          =   288
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "2"
      Top             =   3000
      Width           =   3252
   End
   Begin VB.ComboBox txtMatType 
      Height          =   288
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "2"
      Top             =   2640
      Width           =   3252
   End
   Begin VB.CheckBox optClose 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2700
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Copies New Parts To The Active Form"
      Top             =   600
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.TextBox txtDmy 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      TabIndex        =   16
      Top             =   720
      Width           =   75
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   285
      Left            =   4800
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Add A New Part Number"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtTyp 
      Height          =   285
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtMbe 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "M, B or E"
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Part Description (30 chars)"
      Top             =   1080
      Width           =   2985
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "Part Description (30 chars)"
      Top             =   1440
      Width           =   2985
   End
   Begin VB.CheckBox optFull 
      Caption         =   "Full Bid"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox optQwik 
      Caption         =   "Qwik Bid"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3840
      FormDesignWidth =   5790
   End
   Begin VB.Label lblExists 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "*** This Part Has Been Previously Recorded ***"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   960
      TabIndex        =   22
      Top             =   3480
      Width           =   4092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alloy/Temper"
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   21
      Top             =   3000
      Width           =   1272
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material"
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   20
      Top             =   2640
      Width           =   1152
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Paste And Close On New Part"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   19
      ToolTipText     =   "Copies New Parts To The Active Form"
      Top             =   600
      Width           =   2355
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BID"
      Height          =   285
      Left            =   4800
      TabIndex        =   17
      ToolTipText     =   "Default Product Code For Estimating Created Parts"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valid (1 Thru 8)"
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   14
      Top             =   2160
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make, Buy or Either"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(M,B or E)"
      Height          =   285
      Index           =   8
      Left            =   2640
      TabIndex        =   11
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   1155
   End
End
Attribute VB_Name = "EstiESe01p"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/18/06 Option combos and passed Part Number
Option Explicit
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bPartExists As Byte
Dim bIdle As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function CheckPart() As Byte
   Dim RdoChk As ADODB.Recordset
   txtdmy.Enabled = True
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAMAKEBUY," _
          & "PAPRODCODE,PAMATERIALTYPE,PAMATERIALALLOY " _
          & "FROM PartTable WHERE PARTREF='" & Compress(txtPrt) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD)
   If bSqlRows Then
      With RdoChk
         txtPrt = "" & Trim(!PartNum)
         txtDsc = "" & Trim(!PADESC)
         txtTyp = !PALEVEL
         txtMbe = "" & Trim(!PAMAKEBUY)
         txtMatType = "" & Trim(!PAMATERIALTYPE)
         txtMatAlloy = "" & Trim(!PAMATERIALALLOY)
         ClearResultSet RdoChk
      End With
      CheckPart = 1
      txtDsc.Enabled = False
      txtTyp.Enabled = False
      txtMbe.Enabled = False
      txtMatType.Enabled = False
      txtMatAlloy.Enabled = False
      cmdAdd.Enabled = False
      lblExists.ForeColor = ES_RED
      lblExists.Visible = True
   Else
      CheckPart = 0
      txtdmy.Enabled = False
      lblExists.Visible = False
      If bOnLoad = 1 Then
         txtDsc = ""
         txtTyp = "3"
         txtMbe = "M"
         txtMatType = ""
         txtMatAlloy = ""
         bOnLoad = 0
      End If
      txtDsc.Enabled = True
      txtTyp.Enabled = True
      txtMbe.Enabled = True
      txtMatType.Enabled = True
      txtMatAlloy.Enabled = True
      On Error Resume Next
      txtDsc.SetFocus
      cmdAdd.Enabled = True
   End If
   Set RdoChk = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub cmdAdd_Click()
   bIdle = False
   bPartExists = CheckPart()
   If bPartExists = 0 Then AddNewPart
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      'SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   If bOnLoad = 1 Then
      FillCombos
      bPartExists = CheckPart()
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   If iBarOnTop Then
      Move 1000, 2500
   Else
      Move 2500, 2500
   End If
   FormatControls
   txtdmy.BackColor = Es_FormBackColor
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If optQwik.value = vbChecked Then EstiESe01a.optPart.value = vbUnchecked
   On Error Resume Next
   If optQwik.value = vbChecked Then
      EstiESe01a.txtPrt.SetFocus
   Else
      EstiESe02a.txtPrt.SetFocus
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set EstiESe01p = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtMbe = "M"
   txtTyp = "3"
   
End Sub



Private Sub txtDmy_LostFocus()
   bIdle = False
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   
End Sub


Private Sub txtMatAlloy_LostFocus()
   txtMatAlloy = CheckLen(txtMatAlloy, 30)
   
End Sub


Private Sub txtMatType_LostFocus()
   txtMatType = CheckLen(txtMatType, 30)
   txtMatType = StrCase(txtMatType)
   
End Sub


Private Sub txtMbe_LostFocus()
   Dim bByte As Byte
   txtMbe = CheckLen(txtMbe, 1)
   Select Case txtMbe
      Case "M", "B", "E"
         bByte = True
      Case Else
         bByte = False
   End Select
   If Not bByte Then
      MsgBox "Must Be M, B or E.", vbInformation, Caption
      txtMbe = "M"
   End If
   bIdle = False
   
End Sub


Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If bCancel = 1 Then Exit Sub
   bPartExists = CheckPart()
   bIdle = False
   
End Sub


Private Sub txtTyp_LostFocus()
   txtTyp = CheckLen(txtTyp, 1)
   If Val(txtTyp) < 1 Or Val(txtTyp) > 8 Then
      Beep
      txtTyp = "3"
   End If
   bIdle = False
   
End Sub







Private Sub AddNewPart()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sNewPart As String
   
   If Trim(txtDsc) = "" Then
      MsgBox "Please Enter A Part Description.", _
         vbInformation, Caption
      'On Error Resume Next
      txtDsc.SetFocus
      Exit Sub
   End If
   sMsg = "Do You Want To Add This Part Number?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   'On Error Resume Next
   If bResponse = vbYes Then
      MouseCursor 13
      sNewPart = Compress(txtPrt)
'      sSql = "INSERT INTO PartTable (PARTREF,PARTNUM,PADESC,PALEVEL,PAMAKEBUY," _
'             & "PAPRODCODE,PAMATERIALTYPE,PAMATERIALALLOY) " _
'             & "VALUES('" & sNewPart & "','" & txtPrt & "','" _
'             & txtDsc & "'," & Val(txtTyp) & ",'" & txtMbe & "','" & lblCode _
'             & "','" & txtMatType & "','" & txtMatAlloy & "') "
'      clsADOCon.ExecuteSQL sSql ' rdExecDirect

      clsADOCon.BeginTrans
      
      Dim part As New ClassPart
      If part.CreateNewPart(txtPrt, Val(txtTyp), txtDsc, txtMbe) Then
      
         sSql = "UPDATE PartTable" & vbCrLf _
                & "SET PAPRODCODE = '" & lblCode & "'," & vbCrLf _
                & "PAMATERIALTYPE = '" & txtMatType & "'," & vbCrLf _
                & "PAMATERIALALLOY = '" & txtMatAlloy & "'" & vbCrLf _
                & "WHERE PARTREF = '" & sNewPart & "'"
         clsADOCon.ExecuteSQL sSql ' rdExecDirect
      
         clsADOCon.CommitTrans
         MouseCursor 0
         
      Else
         clsADOCon.RollbackTrans
         MouseCursor 0
         MsgBox "Part creation failed"
         Exit Sub
      End If
      
      MouseCursor 0
      'If Err = 0 Then
         SysMsg "Part Was Added.", True
         If optClose.value = vbChecked Then
            If optQwik.value = vbChecked Then
               EstiESe01a.txtPrt.Text = txtPrt
               EstiESe01a.txtPrt.ToolTipText = txtDsc & "     "
               EstiESe01a.txtPrt.SetFocus
            Else
               EstiESe02a.txtPrt.Text = txtPrt
               EstiESe02a.txtPrt.ToolTipText = txtDsc & "     "
               EstiESe02a.txtPrt.SetFocus
            End If
            txtPrt = ""
            txtDsc = ""
            On Error Resume Next
            txtPrt.SetFocus
         End If
      'Else
      '   MsgBox "Unable To Add That Part Number.", _
      '      vbExclamation, Caption
      'End If
   Else
      CancelTrans
   End If
   Unload Me
   
End Sub

Private Sub FillCombos()
   On Error Resume Next
   sSql = "SELECT DISTINCT PAMATERIALTYPE FROM PartTable"
   LoadComboBox txtMatType, -1
   If txtMatType.ListCount > 0 Then txtMatType = txtMatType.List(0)
   
   sSql = "SELECT DISTINCT PAMATERIALALLOY FROM PartTable"
   LoadComboBox txtMatAlloy, -1
   If txtMatAlloy.ListCount > 0 Then txtMatAlloy = txtMatAlloy.List(0)
   
End Sub
