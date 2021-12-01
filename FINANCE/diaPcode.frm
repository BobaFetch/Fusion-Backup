VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form diaPcode
   BorderStyle = 3 'Fixed Dialog
   Caption = "Product Codes"
   ClientHeight = 4845
   ClientLeft = 2055
   ClientTop = 330
   ClientWidth = 7260
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 4845
   ScaleWidth = 7260
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdUpd
      Caption = "P&arts"
      Height = 315
      Left = 6240
      TabIndex = 17
      TabStop = 0 'False
      ToolTipText = "Update Part Accounts"
      Top = 480
      Width = 875
   End
   Begin VB.ComboBox cmbCde
      Height = 315
      Left = 2160
      Sorted = -1 'True
      TabIndex = 0
      Tag = "3"
      ToolTipText = "Enter/Revise Product Code (6 Char)"
      Top = 600
      Width = 1215
   End
   Begin VB.TextBox txtDsc
      Height = 285
      Left = 2160
      TabIndex = 1
      Tag = "2"
      Top = 960
      Width = 3075
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 6240
      TabIndex = 16
      TabStop = 0 'False
      Top = 0
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5640
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 4845
      FormDesignWidth = 7260
   End
   Begin TabDlg.SSTab Tab1
      Height = 3375
      Left = 0
      TabIndex = 18
      Top = 1440
      Width = 7215
      _ExtentX = 12726
      _ExtentY = 5953
      _Version = 393216
      Style = 1
      TabHeight = 476
      Enabled = 0 'False
      TabCaption(0) = "&Accounts        "
      TabPicture(0) = "diaPcode.frx":0000
      Tab(0).ControlEnabled = -1 'True
      Tab(0).Control(0) = "z1(6)"
      Tab(0).Control(0).Enabled = 0 'False
      Tab(0).Control(1) = "z1(7)"
      Tab(0).Control(1).Enabled = 0 'False
      Tab(0).Control(2) = "lblTrv"
      Tab(0).Control(2).Enabled = 0 'False
      Tab(0).Control(3) = "lblTcg"
      Tab(0).Control(3).Enabled = 0 'False
      Tab(0).Control(4) = "z1(2)"
      Tab(0).Control(4).Enabled = 0 'False
      Tab(0).Control(5) = "z1(3)"
      Tab(0).Control(5).Enabled = 0 'False
      Tab(0).Control(6) = "lblRev"
      Tab(0).Control(6).Enabled = 0 'False
      Tab(0).Control(7) = "lblDis"
      Tab(0).Control(7).Enabled = 0 'False
      Tab(0).Control(8) = "z1(16)"
      Tab(0).Control(8).Enabled = 0 'False
      Tab(0).Control(9) = "txtTrv"
      Tab(0).Control(9).Enabled = 0 'False
      Tab(0).Control(10) = "txtTcg"
      Tab(0).Control(10).Enabled = 0 'False
      Tab(0).Control(11) = "txtRev"
      Tab(0).Control(11).Enabled = 0 'False
      Tab(0).Control(12) = "txtDis"
      Tab(0).Control(12).Enabled = 0 'False
      Tab(0).ControlCount = 13
      TabCaption(1) = "&Inventory/Expense"
      TabPicture(1) = "diaPcode.frx":001C
      Tab(1).ControlEnabled = 0 'False
      Tab(1).Control(0) = "txtWex"
      Tab(1).Control(1) = "txtWoh"
      Tab(1).Control(2) = "txtWma"
      Tab(1).Control(3) = "txtWla"
      Tab(1).Control(4) = "z1(14)"
      Tab(1).Control(5) = "lblWex"
      Tab(1).Control(6) = "lblWoh"
      Tab(1).Control(7) = "lblWma"
      Tab(1).Control(8) = "lblWla"
      Tab(1).Control(9) = "z1(12)"
      Tab(1).Control(10) = "z1(11)"
      Tab(1).Control(11) = "z1(10)"
      Tab(1).Control(12) = "z1(9)"
      Tab(1).ControlCount = 13
      TabCaption(2) = "&Cost Of Goods"
      Tab(2).ControlEnabled = 0 'False
      Tab(2).Control(0) = "txtGex"
      Tab(2).Control(1) = "txtGoh"
      Tab(2).Control(2) = "txtGma"
      Tab(2).Control(3) = "txtGla"
      Tab(2).Control(4) = "lblGex"
      Tab(2).Control(5) = "lblGoh"
      Tab(2).Control(6) = "lblGma"
      Tab(2).Control(7) = "lblGla"
      Tab(2).Control(8) = "z1(15)"
      Tab(2).Control(9) = "z1(13)"
      Tab(2).Control(10) = "z1(8)"
      Tab(2).Control(11) = "z1(5)"
      Tab(2).Control(12) = "z1(4)"
      Tab(2).ControlCount = 13
      Begin VB.ComboBox txtGex
         Height = 315
         Left = -72840
         TabIndex = 14
         Tag = "3"
         Top = 2160
         Width = 1935
      End
      Begin VB.ComboBox txtGoh
         Height = 315
         Left = -72840
         TabIndex = 13
         Tag = "3"
         Top = 1800
         Width = 1935
      End
      Begin VB.ComboBox txtGma
         Height = 315
         Left = -72840
         TabIndex = 12
         Tag = "3"
         Top = 1440
         Width = 1935
      End
      Begin VB.ComboBox txtGla
         Height = 315
         Left = -72840
         TabIndex = 11
         Tag = "3"
         Top = 1080
         Width = 1935
      End
      Begin VB.ComboBox txtDis
         Height = 315
         Left = 2160
         TabIndex = 4
         Top = 1440
         Width = 1935
      End
      Begin VB.ComboBox txtRev
         Height = 315
         Left = 2160
         TabIndex = 3
         Top = 1080
         Width = 1935
      End
      Begin VB.ComboBox txtTcg
         Height = 315
         Left = 2160
         TabIndex = 6
         Top = 2160
         Width = 1935
      End
      Begin VB.ComboBox txtTrv
         Height = 315
         Left = 2160
         TabIndex = 5
         Top = 1800
         Width = 1935
      End
      Begin VB.ComboBox txtWex
         Height = 315
         Left = -72840
         TabIndex = 10
         Tag = "3"
         Top = 2160
         Width = 1935
      End
      Begin VB.ComboBox txtWoh
         Height = 315
         Left = -72840
         TabIndex = 9
         Tag = "3"
         Top = 1800
         Width = 1935
      End
      Begin VB.ComboBox txtWma
         Height = 315
         Left = -72840
         TabIndex = 8
         Tag = "3"
         Top = 1440
         Width = 1935
      End
      Begin VB.ComboBox txtWla
         Height = 315
         Left = -72840
         TabIndex = 7
         Tag = "3"
         Top = 1080
         Width = 1935
      End
      Begin VB.Label lblGex
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = -70800
         TabIndex = 45
         Top = 2160
         Width = 2775
      End
      Begin VB.Label lblGoh
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = -70800
         TabIndex = 44
         Top = 1800
         Width = 2775
      End
      Begin VB.Label lblGma
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = -70800
         TabIndex = 43
         Top = 1440
         Width = 2775
      End
      Begin VB.Label lblGla
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = -70800
         TabIndex = 42
         Top = 1080
         Width = 2775
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Miscellaneous Accounts:"
         Height = 255
         Index = 16
         Left = 240
         TabIndex = 41
         Top = 720
         Width = 3075
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Cost Of Goods Sold Accounts:"
         Height = 255
         Index = 15
         Left = -74760
         TabIndex = 40
         Top = 720
         Width = 3075
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Inventory/Expense Accounts:"
         Height = 255
         Index = 14
         Left = -74760
         TabIndex = 39
         Top = 720
         Width = 3075
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Expense Account"
         Height = 255
         Index = 13
         Left = -74760
         TabIndex = 38
         Top = 2160
         Width = 1995
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Overhead Account"
         Height = 255
         Index = 8
         Left = -74760
         TabIndex = 37
         Top = 1800
         Width = 1995
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Material Account"
         Height = 255
         Index = 5
         Left = -74760
         TabIndex = 36
         Top = 1440
         Width = 1995
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Labor Account"
         Height = 255
         Index = 4
         Left = -74760
         TabIndex = 35
         Top = 1080
         Width = 1995
      End
      Begin VB.Label lblDis
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = 4200
         TabIndex = 34
         Top = 1470
         Width = 2775
      End
      Begin VB.Label lblRev
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = 4200
         TabIndex = 33
         Top = 1110
         Width = 2775
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Discount Account"
         Height = 255
         Index = 3
         Left = 240
         TabIndex = 32
         Top = 1470
         Width = 1995
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Revenue Account"
         Height = 255
         Index = 2
         Left = 240
         TabIndex = 31
         Top = 1080
         Width = 1995
      End
      Begin VB.Label lblTcg
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = 4200
         TabIndex = 30
         Top = 2160
         Width = 2775
      End
      Begin VB.Label lblTrv
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = 4200
         TabIndex = 29
         Top = 1800
         Width = 2775
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Transfer CGS Acct"
         Height = 255
         Index = 7
         Left = 240
         TabIndex = 28
         Top = 2160
         Width = 1995
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Transfer Revenue Acct"
         Height = 255
         Index = 6
         Left = 240
         TabIndex = 27
         Top = 1800
         Width = 1995
      End
      Begin VB.Label lblWex
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = -70800
         TabIndex = 26
         Top = 2160
         Width = 2775
      End
      Begin VB.Label lblWoh
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = -70800
         TabIndex = 25
         Top = 1800
         Width = 2775
      End
      Begin VB.Label lblWma
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = -70800
         TabIndex = 24
         Top = 1440
         Width = 2775
      End
      Begin VB.Label lblWla
         BackStyle = 0 'Transparent
         BorderStyle = 1 'Fixed Single
         Height = 285
         Left = -70800
         TabIndex = 23
         Top = 1080
         Width = 2775
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Expense Account"
         Height = 375
         Index = 12
         Left = -74760
         TabIndex = 22
         Top = 2160
         Width = 1995
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Overhead Account"
         Height = 375
         Index = 11
         Left = -74760
         TabIndex = 21
         Top = 1800
         Width = 1995
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Material Account"
         Height = 375
         Index = 10
         Left = -74760
         TabIndex = 20
         Top = 1440
         Width = 1995
      End
      Begin VB.Label z1
         BackStyle = 0 'Transparent
         Caption = "Labor Account"
         Height = 255
         Index = 9
         Left = -74760
         TabIndex = 19
         Top = 1080
         Width = 1995
      End
   End
   Begin Threed.SSFrame fra2
      Height = 30
      Left = 0
      TabIndex = 46
      Top = 1320
      Width = 7215
      _Version = 65536
      _ExtentX = 12726
      _ExtentY = 53
      _StockProps = 14
      ForeColor = 12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 47
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaPcode.frx":0038
      PictureDn = "diaPcode.frx":017E
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Description"
      Height = 255
      Index = 1
      Left = 240
      TabIndex = 15
      Top = 960
      Width = 1695
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Product Code"
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 2
      Top = 600
      Width = 1575
   End
End
Attribute VB_Name = "diaPcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bGoodCode As Byte
Dim bOnLoad As Byte
Dim bNoAccts As Boolean
Dim rdoRes As rdoResultset

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCde_Click()
   bGoodCode = GetCode()
   Tab1.Enabled = False
   
End Sub

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If Len(cmbCde) Then
      bGoodCode = GetCode(True)
      If Not bGoodCode Then
         AddProductCode
      Else
         Tab1.Enabled = True
      End If
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   Dim l&
   If cmdHlp Then
      MouseCursor 13
      l& = WinHelp(hwnd, sReportPath & "Esiadmn.hlp", HELP_KEY, "Product Codes")
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sPcode As String
   Dim sMsg As String
   Dim sPart As String
   Dim sProdAccount(14) As String
   
   sMsg = "Do You Want To Update Accounts Of " & vbCr _
          & "All Parts With Product Code " & cmbCde & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdUpd.Enabled = False
      sPcode = Compress(cmbCde)
      sProdAccount(0) = Compress(txtRev)
      sProdAccount(1) = Compress(txtDis)
      sProdAccount(4) = Compress(txtTrv)
      sProdAccount(5) = Compress(txtTcg)
      'INV/EXP
      sProdAccount(6) = Compress(txtWla)
      sProdAccount(7) = Compress(txtWma)
      sProdAccount(8) = Compress(txtWoh)
      sProdAccount(9) = Compress(txtWex)
      'CGS
      sProdAccount(10) = Compress(txtGla)
      sProdAccount(11) = Compress(txtGma)
      sProdAccount(12) = Compress(txtGoh)
      sProdAccount(13) = Compress(txtGex)
      
      RdoCon.BeginTrans
      sSql = "UPDATE PartTable SET " _
             & "PAREVACCT='" & sProdAccount(0) & "'," _
             & "PADISACCT='" & sProdAccount(1) & "'," _
             & "PACGSACCT='" & sProdAccount(2) & "'," _
             & "PAACCTNO='" & sProdAccount(3) & "'," _
             & "PATFRREVACCT='" & sProdAccount(4) & "'," _
             & "PATFRCGSACCT='" & sProdAccount(5) & "'," _
             & "PAINVLABACCT='" & sProdAccount(6) & "'," _
             & "PAINVMATACCT='" & sProdAccount(7) & "'," _
             & "PAINVOHDACCT='" & sProdAccount(8) & "'," _
             & "PAINVEXPACCT='" & sProdAccount(9) & "'," _
             & "PACGSLABACCT='" & sProdAccount(10) & "'," _
             & "PACGSMATACCT='" & sProdAccount(11) & "'," _
             & "PACGSOHDACCT='" & sProdAccount(12) & "'," _
             & "PACGSEXPACCT='" & sProdAccount(13) & "' " _
             & "WHERE PAPRODCODE='" & sPcode & "' "
      RdoCon.Execute sSql, rdExecDirect
      MouseCursor 0
      cmdUpd.Enabled = True
      If RdoCon.RowsAffected > 0 Then
         sMsg = Trim(Str(RdoCon.RowsAffected)) & " Parts Are Selected To Be Updated." & vbCr _
                & "You Wish To Continue Updating?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            RdoCon.CommitTrans
            MsgBox Str(RdoCon.RowsAffected) & " Parts Were Updated.", _
                       vbInformation, Caption
         Else
            RdoCon.RollbackTrans
            CancelTrans
         End If
      Else
         RdoCon.RollbackTrans
         MsgBox "No Parts Were Updated.", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductCodes Me
      FillAccounts
      If cmbCde.ListCount > 0 Then
         cmbCde = cmbCde.List(0)
         bGoodCode = GetCode()
      End If
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   Tab1.Tab = 0
   bOnLoad = True
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaPcode = Nothing
   
End Sub







Private Sub tab1_Click(PreviousTab As Integer)
   On Error Resume Next
   Select Case Tab1.Tab
      Case 1
         txtWla.SetFocus
      Case 2
         txtGla.SetFocus
      Case Else
         txtRev.SetFocus
   End Select
   
End Sub

Private Sub txtDis_Click()
   GetAccount txtDis, "txtDis"
   
End Sub

Private Sub txtDis_LostFocus()
   txtDis = CheckLen(txtDis, 12)
   GetAccount txtDis, "txtDis"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCDISCACCT = "" & Compress(txtDis)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCDESC = "" & txtDsc
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub

Private Sub txtGex_Click()
   GetAccount txtGex, "txtGex"
   
End Sub


Private Sub txtGex_LostFocus()
   txtGex = CheckLen(txtGex, 12)
   GetAccount txtGex, "txtGex"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCCGSEXPACCT = "" & Compress(txtGex)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtGla_Click()
   GetAccount txtGla, "txtGla"
   
End Sub


Private Sub txtGla_LostFocus()
   txtGla = CheckLen(txtGla, 12)
   GetAccount txtGla, "txtGla"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCCGSLABACCT = "" & Compress(txtGla)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtGma_Click()
   GetAccount txtGma, "txtGma"
   
End Sub


Private Sub txtGma_LostFocus()
   txtGma = CheckLen(txtGma, 12)
   GetAccount txtGma, "txtGma"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCCGSMATACCT = "" & Compress(txtGma)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtGoh_Click()
   GetAccount txtGoh, "txtGoh"
   
End Sub


Private Sub txtGoh_LostFocus()
   txtGoh = CheckLen(txtGoh, 12)
   GetAccount txtGoh, "txtGoh"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCCGSOHDACCT = "" & Compress(txtGoh)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtRev_Click()
   GetAccount txtRev, "txtRev"
   
End Sub

Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 12)
   GetAccount txtRev, "txtRev"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCREVACCT = "" & Compress(txtRev)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtTcg_Click()
   GetAccount txtTcg, "txtTcg"
   
End Sub

Private Sub txtTcg_LostFocus()
   txtTcg = CheckLen(txtTcg, 12)
   GetAccount txtTcg, "txtTcg"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCDCGSXFERAC = "" & Compress(txtTcg)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtTrv_Click()
   GetAccount txtTrv, "txtTrv"
   
End Sub

Private Sub txtTrv_LostFocus()
   txtTrv = CheckLen(txtTrv, 12)
   GetAccount txtTrv, "txtTrv"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCDREVXFERAC = "" & Compress(txtTrv)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub


Public Function GetCode(Optional MoveTab As Byte) As Byte
   Dim sPcode As String
   
   sPcode = Compress(cmbCde)
   If MoveTab Then Tab1.Tab = 0
   On Error GoTo DiaErr1
   
   sSql = "SELECT * FROM PcodTable WHERE PCREF='" & sPcode & "' "
   
   
   bSqlRows = GetDataSet(rdoRes, ES_KEYSET)
   If bSqlRows Then
      With rdoRes
         cmbCde = "" & Trim(!PCCODE)
         txtDsc = "" & Trim(!PCDESC)
         
         txtRev = "" & Trim(!PCREVACCT)
         GetAccount txtRev, "txtRev"
         
         txtDis = "" & Trim(!PCDISCACCT)
         GetAccount txtDis, "txtDis"
         
         txtTrv = "" & Trim(!PCDREVXFERAC)
         GetAccount txtTrv, "txtTrv"
         
         txtTcg = "" & Trim(!PCDCGSXFERAC)
         GetAccount txtTcg, "txtTcg"
         'Inv/Exp
         txtWla = "" & Trim(!PCINVLABACCT)
         GetAccount txtWla, "txtWla"
         
         txtWma = "" & Trim(!PCINVMATACCT)
         GetAccount txtWma, "txtWma"
         
         txtWoh = "" & Trim(!PCINVOHDACCT)
         GetAccount txtWoh, "txtWoh"
         
         txtWex = "" & Trim(!PCINVEXPACCT)
         GetAccount txtWex, "txtWex"
         'CGS
         txtGla = "" & Trim(!PCCGSLABACCT)
         GetAccount txtGla, "txtGla"
         
         txtGma = "" & Trim(!PCCGSMATACCT)
         GetAccount txtGma, "txtGma"
         
         txtGoh = "" & Trim(!PCCGSOHDACCT)
         GetAccount txtGoh, "txtGoh"
         
         txtGex = "" & Trim(!PCCGSEXPACCT)
         GetAccount txtGex, "txtGex"
         
      End With
      GetCode = True
   Else
      txtDsc = ""
      txtRev = ""
      txtDis = ""
      txtTrv = ""
      txtTcg = ""
      txtWla = ""
      txtWma = ""
      txtWoh = ""
      txtWex = ""
      txtGla = ""
      txtGma = ""
      txtGoh = ""
      txtGex = ""
      GetCode = False
   End If
   Exit Function
   
   DiaErr1:
   sProcName = "getcode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Sub AddProductCode()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sCode As String
   
   sCode = Compress(cmbCde)
   If sCode = "ALL" Then
      MsgBox "Illegal Product Code Name.", vbExclamation, Caption
      Exit Sub
   End If
   sMsg = cmbCde & " Wasn't Found. Add The Product Code?"
   On Error GoTo DiaErr1
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      sSql = "INSERT INTO PcodTable (PCREF,PCCODE) " _
             & "VALUES('" & sCode & "','" & cmbCde & "')"
      RdoCon.Execute sSql, rdExecDirect
      If RdoCon.RowsAffected Then
         Sysmsg "Product Code Added.", True
         cmbCde.AddItem cmbCde
         bGoodCode = GetCode()
         Tab1.Enabled = True
         On Error Resume Next
         txtDsc.SetFocus
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "addproduc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub FillAccounts()
   Dim rdoGlm As rdoResultset
   On Error GoTo DiaErr1
   sSql = "Qry_FillLowAccounts"
   bSqlRows = GetDataSet(rdoGlm, ES_FORWARD)
   If bSqlRows Then
      With rdoGlm
         Do Until .EOF
            AddComboStr txtRev.hwnd, "" & Trim(!GLACCTNO)
            txtDis.AddItem "" & Trim(!GLACCTNO)
            txtTrv.AddItem "" & Trim(!GLACCTNO)
            txtTcg.AddItem "" & Trim(!GLACCTNO)
            'Inv/Exp
            txtWla.AddItem "" & Trim(!GLACCTNO)
            txtWma.AddItem "" & Trim(!GLACCTNO)
            txtWoh.AddItem "" & Trim(!GLACCTNO)
            txtWex.AddItem "" & Trim(!GLACCTNO)
            'Inv/Exp
            txtGla.AddItem "" & Trim(!GLACCTNO)
            txtGma.AddItem "" & Trim(!GLACCTNO)
            txtGoh.AddItem "" & Trim(!GLACCTNO)
            txtGex.AddItem "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      bNoAccts = True
      CloseBoxes
   End If
   Set rdoGlm = Nothing
   Exit Sub
   
   DiaErr1:
   bNoAccts = True
   CloseBoxes
   
End Sub

Public Sub GetAccount(sAccount As String, sBox As String)
   Dim rdoGlm As rdoResultset
   On Error GoTo DiaErr1
   If bNoAccts Then Exit Sub
   sAccount = Compress(sAccount)
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR FROM GlacTable " _
          & "WHERE GLACCTREF='" & sAccount & "'"
   bSqlRows = GetDataSet(rdoGlm, ES_FORWARD)
   If bSqlRows Then
      With rdoGlm
         Select Case sBox
            Case "txtRev"
               txtRev = "" & Trim(!GLACCTNO)
               lblRev = "" & Trim(!GLDESCR)
            Case "txtDis"
               txtDis = "" & Trim(!GLACCTNO)
               lblDis = "" & Trim(!GLDESCR)
            Case "txtTrv"
               txtTrv = "" & Trim(!GLACCTNO)
               lblTrv = "" & Trim(!GLDESCR)
            Case "txtTcg"
               txtTcg = "" & Trim(!GLACCTNO)
               lblTcg = "" & Trim(!GLDESCR)
               'Inv/Exp
            Case "txtWla"
               txtWla = "" & Trim(!GLACCTNO)
               lblWla = "" & Trim(!GLDESCR)
            Case "txtWma"
               txtWma = "" & Trim(!GLACCTNO)
               lblWma = "" & Trim(!GLDESCR)
            Case "txtWoh"
               txtWoh = "" & Trim(!GLACCTNO)
               lblWoh = "" & Trim(!GLDESCR)
            Case "txtWex"
               txtWex = "" & Trim(!GLACCTNO)
               lblWex = "" & Trim(!GLDESCR)
               'Cogs
            Case "txtGla"
               txtGla = "" & Trim(!GLACCTNO)
               lblGla = "" & Trim(!GLDESCR)
            Case "txtGma"
               txtGma = "" & Trim(!GLACCTNO)
               lblGma = "" & Trim(!GLDESCR)
            Case "txtGoh"
               txtGoh = "" & Trim(!GLACCTNO)
               lblGoh = "" & Trim(!GLDESCR)
            Case "txtGex"
               txtGex = "" & Trim(!GLACCTNO)
               lblGex = "" & Trim(!GLDESCR)
         End Select
         .Cancel
      End With
   Else
      Select Case sBox
         Case "txtRev"
            txtRev = ""
            lblRev = ""
         Case "txtDis"
            txtDis = ""
            lblDis = ""
         Case "txtTrv"
            txtTrv = ""
            lblTrv = ""
         Case "txtTcg"
            txtTcg = ""
            lblTcg = ""
         Case "txtWla"
            txtWla = ""
            lblWla = ""
         Case "txtWma"
            txtWma = ""
            lblWma = ""
         Case "txtWoh"
            txtWoh = ""
            lblWoh = ""
         Case "txtWex"
            txtWex = ""
            lblWex = ""
      End Select
   End If
   Set rdoGlm = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "getaccount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Public Sub CloseBoxes()
   txtRev.Enabled = False
   txtDis.Enabled = False
   txtTrv.Enabled = False
   txtTcg.Enabled = False
   txtWex.Enabled = False
   txtWla.Enabled = False
   txtWma.Enabled = False
   txtWoh.Enabled = False
   txtGex.Enabled = False
   txtGla.Enabled = False
   txtGma.Enabled = False
   txtGoh.Enabled = False
   cmdUpd.Enabled = False
   MsgBox "You Have No Accounts Established." & vbCrLf _
      & "Please Open Finance And Enter Accounts.", _
      vbInformation, Caption
   
End Sub

Private Sub txtWex_Click()
   GetAccount txtWex, "txtWex"
   
End Sub


Private Sub txtWex_LostFocus()
   txtWex = CheckLen(txtWex, 12)
   GetAccount txtWex, "txtWex"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCINVEXPACCT = "" & Compress(txtWex)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtWla_Click()
   GetAccount txtWla, "txtWla"
   
End Sub


Private Sub txtWla_LostFocus()
   txtWla = CheckLen(txtWla, 12)
   GetAccount txtWla, "txtWla"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCINVLABACCT = "" & Compress(txtWla)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtWma_Click()
   GetAccount txtWma, "txtWma"
   
End Sub


Private Sub txtWma_LostFocus()
   txtWma = CheckLen(txtWma, 12)
   GetAccount txtWma, "txtWma"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCINVMATACCT = "" & Compress(txtWma)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtWoh_Click()
   GetAccount txtWoh, "txtWoh"
   
End Sub


Private Sub txtWoh_LostFocus()
   txtWoh = CheckLen(txtWoh, 12)
   GetAccount txtWoh, "txtWoh"
   If bGoodCode Then
      On Error Resume Next
      rdoRes.Edit
      rdoRes!PCINVOHDACCT = "" & Compress(txtWoh)
      rdoRes.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub
