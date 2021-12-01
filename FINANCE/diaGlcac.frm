VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGlcac
   BorderStyle = 3 'Fixed Dialog
   Caption = "Chart Of Accounts"
   ClientHeight = 4155
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 6840
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 4155
   ScaleWidth = 6840
   ShowInTaskbar = 0 'False
   Begin VB.ComboBox cmbvnd
      Height = 315
      Left = 3480
      TabIndex = 5
      Tag = "3"
      ToolTipText = "Select A Vendor For Cash Accounts Only"
      Top = 3120
      Width = 1455
   End
   Begin VB.CheckBox optCash
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Left = 1800
      TabIndex = 4
      Top = 3120
      Width = 855
   End
   Begin VB.CheckBox optVew
      Caption = "vew"
      Height = 255
      Left = 960
      TabIndex = 23
      Top = 120
      Visible = 0 'False
      Width = 735
   End
   Begin VB.CommandButton cmdFnd
      Height = 315
      Left = 4800
      Picture = "diaGlcac.frx":0000
      Style = 1 'Graphical
      TabIndex = 22
      TabStop = 0 'False
      ToolTipText = "Find An Account"
      Top = 480
      Width = 350
   End
   Begin VB.CheckBox optChg
      Height = 255
      Left = 360
      TabIndex = 21
      Top = 120
      Visible = 0 'False
      Width = 735
   End
   Begin VB.CommandButton cmdChg
      Caption = "C&hange"
      Height = 315
      Left = 5880
      TabIndex = 20
      TabStop = 0 'False
      ToolTipText = "Change Current Account Number"
      Top = 720
      Width = 875
   End
   Begin VB.CommandButton cmdDel
      Caption = "&Delete"
      Height = 315
      Left = 5880
      TabIndex = 19
      TabStop = 0 'False
      ToolTipText = "Delete Current Account Number"
      Top = 1080
      Width = 875
   End
   Begin VB.CommandButton cmdVew
      DownPicture = "diaGlcac.frx":0342
      Height = 320
      Left = 4320
      Picture = "diaGlcac.frx":0CB4
      Style = 1 'Graphical
      TabIndex = 18
      TabStop = 0 'False
      ToolTipText = "Show Existing Accounts"
      Top = 480
      Width = 350
   End
   Begin VB.ComboBox cmbMst
      Height = 315
      Left = 1800
      TabIndex = 2
      Tag = "3"
      Top = 1680
      Width = 1935
   End
   Begin VB.CheckBox optAct
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Left = 1800
      TabIndex = 6
      Top = 3480
      Width = 735
   End
   Begin VB.TextBox txtLvl
      Height = 285
      Left = 1800
      TabIndex = 3
      Tag = "1"
      Top = 2520
      Width = 255
   End
   Begin VB.ComboBox cmbAct
      Height = 315
      Left = 1800
      Sorted = -1 'True
      TabIndex = 0
      Tag = "3"
      Top = 480
      Width = 1815
   End
   Begin VB.TextBox txtDsc
      Height = 285
      Left = 1800
      TabIndex = 1
      Tag = "2"
      Top = 840
      Width = 3360
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 5880
      TabIndex = 9
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 10
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaGlcac.frx":1626
      PictureDn = "diaGlcac.frx":176C
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5400
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 4155
      FormDesignWidth = 6840
   End
   Begin Threed.SSFrame z2
      Height = 30
      Left = 120
      TabIndex = 12
      Top = 1440
      Width = 6675
      _Version = 65536
      _ExtentX = 11774
      _ExtentY = 53
      _StockProps = 14
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
   Begin VB.Label lblTyp
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Index = 1
      Left = 3480
      TabIndex = 28
      Top = 3840
      Visible = 0 'False
      Width = 375
   End
   Begin VB.Label lblTyp
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Index = 0
      Left = 4320
      TabIndex = 27
      Top = 1680
      Width = 375
   End
   Begin VB.Label lblnme
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 3480
      TabIndex = 26
      Top = 3480
      Width = 3135
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Vendor"
      Height = 255
      Index = 9
      Left = 2760
      TabIndex = 25
      Top = 3120
      Width = 855
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Cash Account?"
      Height = 255
      Index = 7
      Left = 120
      TabIndex = 24
      Top = 3120
      Width = 1575
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Deactivate Account?"
      Height = 255
      Index = 6
      Left = 120
      TabIndex = 17
      Top = 3480
      Width = 1695
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Accounts With Level Equal Or Less Than Requested Will Appear On Statements)"
      Height = 495
      Index = 5
      Left = 2280
      TabIndex = 16
      Top = 2520
      Width = 4455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Detail Level "
      Height = 255
      Index = 4
      Left = 120
      TabIndex = 15
      Top = 2520
      Width = 1455
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 1800
      TabIndex = 14
      Top = 2040
      Width = 3255
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Type"
      Height = 255
      Index = 3
      Left = 3840
      TabIndex = 13
      Top = 1680
      Visible = 0 'False
      Width = 495
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Sub Account Of Master Account"
      Height = 495
      Index = 2
      Left = 120
      TabIndex = 11
      Top = 1680
      Width = 1575
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Description"
      Height = 255
      Index = 1
      Left = 120
      TabIndex = 8
      Top = 840
      Width = 1455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Account Number"
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 7
      Top = 480
      Width = 1575
   End
End
Attribute VB_Name = "diaGlcac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' diaGlcac - Add/Revise Chart Of Accounts
'
' Created: (cjs)
' Revisions:
'   9/05/01 Add support for the cash account option / checking accounts (nth)
'   9/10/01 Add vendors to cash accounts (nth)
'  11/21/01 fixed error with lblTyp() (nth)
'
'***************************************************************************
Option Explicit
Dim rdoAct As rdoResultset
Dim bCancel As Boolean
Dim bOnLoad As Byte
Dim bGoodAccount As Byte
Dim iTotal As Integer
Dim sOldAccount As String
Dim sAccount As String
Dim vAccounts(1000, 4) As Variant
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbact_Click()
   bGoodAccount = GetAccount()
   sAccount = Compress(cmbAct)
End Sub

Private Sub cmbact_LostFocus()
   cmbAct = CheckLen(cmbAct, 12)
   If Len(cmbAct) And Not bCancel Then
      bGoodAccount = GetAccount()
      If Not bGoodAccount Then AddAccount
   Else
      bGoodAccount = False
   End If
   sAccount = Compress(cmbAct)
End Sub

Private Sub cmbMst_Click()
   If cmbMst.ListIndex >= 0 Then
      lblTyp(0) = vAccounts(cmbMst.ListIndex, 3)
      lblDsc = vAccounts(cmbMst.ListIndex, 2)
   End If
End Sub

Private Sub cmbMst_LostFocus()
   Dim b As Byte
   Dim i As Integer
   
   On Error Resume Next
   
   If cmbMst = cmbAct Then
      MsgBox "Cannot Use An Account On Itself."
      For i = 0 To cmbMst.ListCount - 1
         If sOldAccount = vAccounts(i, 0) Then
            cmbMst = cmbMst.List(i)
            lblDsc = vAccounts(i, 2)
            lblTyp(0) = vAccounts(i, 3)
            cmbMst.ListIndex = i
            Exit For
         End If
      Next
   End If
   cmbMst = CheckLen(cmbMst, 12)
   For i = 0 To cmbMst.ListCount - 1
      If cmbMst = cmbMst.List(i) Then
         b = 1
         cmbMst.ListIndex = i
      End If
   Next
   If b <> 1 Then
      Beep
      cmbMst.ListIndex = 0
      cmbMst = cmbMst.List(0)
   End If
   If cmbMst.ListIndex > 0 Then
      lblTyp(0) = vAccounts(cmbMst.ListIndex, 3)
      lblDsc = vAccounts(cmbMst.ListIndex, 2)
   End If
   sOldAccount = Compress(cmbMst)
   lblTyp(1) = lblTyp(0)
   
   If bGoodAccount Then
      rdoAct.Edit
      rdoAct!GLMASTER = Compress(cmbMst)
      rdoAct!GLTYPE = Val(lblTyp(0))
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub

Private Sub cmbVnd_Click()
   FindVendor Me
End Sub

Private Sub cmbVnd_LostFocus()
   FindVendor Me
   rdoAct.Edit
   rdoAct!GLVENDOR = Compress(cmbVnd)
   rdoAct.Update
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
   
End Sub


Private Sub cmdChg_Click()
   optChg.Value = vbChecked
   diaGlchg.Show
   
End Sub

Private Sub cmdDel_Click()
   Dim bResponse As Byte
   Dim smsg As String
   smsg = "This Function Removes The Account.  " & vbCrLf _
          & "Cannot Delete Account With References.  " & vbCrLf _
          & "Do You Still Want To Delete This Account?"
   bResponse = MsgBox(smsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      DeleteAccount
   Else
      CancelTrans
   End If
End Sub

Private Sub cmdFnd_Click()
   If cmdFnd Then
      VewAcct.Show
      optVew.Value = vbChecked
      cmdFnd = False
   End If
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Chart Of Accounts"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdVew_Click()
   If cmdVew.Value = True Then
      ActTree.Show
      cmdVew.Value = False
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If optChg.Value = vbChecked Then
      Unload diaGlchg
      optChg.Value = vbUnchecked
   End If
   If bOnLoad Then
      'CheckAccounts
      FillAccounts True
      bOnLoad = False
   End If
   If optVew.Value = vbChecked Then
      Unload VewAcct
      optVew.Value = vbUnchecked
   End If
   
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim RdoTst As rdoResultset
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   On Error Resume Next
   MouseCursor 13
   bOnLoad = True
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Erase vAccounts
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set rdoAct = Nothing
   Set diaGlcac = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optAct_Click()
   If bGoodAccount Then
      On Error Resume Next
      rdoAct.Edit
      rdoAct!GLINACTIVE = optAct.Value
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub optAct_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optCash_Click()
   If bGoodAccount Then
      On Error Resume Next
      rdoAct.Edit
      rdoAct!GLCASH = optCash.Value
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
   If optCash Then
      cmbVnd.Enabled = True
   Else
      cmbVnd = ""
      lblNme = ""
      cmbVnd.Enabled = False
      
      rdoAct.Edit
      rdoAct!GLVENDOR = ""
      rdoAct.Update
   End If
   
End Sub

Private Sub optCash_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub

Private Sub optChg_Click()
   'Never visible. Checks for diaGlchg
End Sub

Private Sub optVew_Click()
   'never visible - view showing
   If optVew.Value = vbUnchecked Then
      If Len(Trim(cmbAct)) Then bGoodAccount = GetAccount()
   End If
   
End Sub

Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   If Len(Trim(txtDsc)) Then
      If bGoodAccount Then
         With rdoAct
            On Error Resume Next
            .Edit
            !GLDESCR = txtDsc
            .Update
            If Err = 0 Then ValidateEdit Me
         End With
      End If
   End If
End Sub

Private Sub txtLvl_LostFocus()
   txtLvl = CheckLen(txtLvl, 1)
   txtLvl = Format(Abs(Val(txtLvl)), "0")
   If bGoodAccount Then
      On Error Resume Next
      rdoAct.Edit
      rdoAct!GLFSLEVEL = Val(txtLvl)
      rdoAct!GLMASTER = Compress(cmbMst)
      rdoAct!GLTYPE = Val(lblTyp(0))
      rdoAct.Update
      If Err = 0 Then
         '    RdoCon.CommitTrans
      Else
         '    RdoCon.RollbackTrans
         ValidateEdit Me
      End If
   End If
   
End Sub

Public Sub ManageBoxes(bOpen As Boolean)
   If Not bOpen Then
      cmbAct.Enabled = False
      txtLvl.Enabled = False
      optAct.Enabled = False
      cmdDel.Enabled = False
   Else
      cmbAct.Enabled = True
      txtLvl.Enabled = True
      optAct.Enabled = True
      cmdDel.Enabled = True
   End If
   
End Sub

Public Sub FillAccounts(bGetAccount As Boolean)
   Dim i As Integer
   Dim rdoGlm As rdoResultset
   Dim RdoVnd As rdoResultset
   Dim sAccount As String
   
   On Error GoTo diaErr1
   On Error GoTo 0
   MouseCursor 13
   
   cmbAct.Clear
   cmbMst.Clear
   sSql = "SELECT * FROM GlmsTable WHERE COACCTREC=1"
   bSqlRows = GetDataSet(rdoGlm, ES_FORWARD)
   If bSqlRows Then
      With rdoGlm
         i = 0
         vAccounts(i, 0) = "" & Trim(!COASSTREF)
         vAccounts(i, 1) = "" & Trim(!COASSTACCT)
         vAccounts(i, 2) = "" & Trim(!COASSTDESC)
         vAccounts(i, 3) = Format(!COASSTTYPE, "0")
         
         i = i + 1
         vAccounts(i, 0) = "" & Trim(!COLIABREF)
         vAccounts(i, 1) = "" & Trim(!COLIABACCT)
         vAccounts(i, 2) = "" & Trim(!COLIABDESC)
         vAccounts(i, 3) = Format(!COLIABTYPE, "0")
         
         i = i + 1
         vAccounts(i, 0) = "" & Trim(!COEQTYREF)
         vAccounts(i, 1) = "" & Trim(!COEQTYACCT)
         vAccounts(i, 2) = "" & Trim(!COEQTYDESC)
         vAccounts(i, 3) = Format(!COEQTYTYPE, "0")
         
         i = i + 1
         vAccounts(i, 0) = "" & Trim(!COINCMREF)
         vAccounts(i, 1) = "" & Trim(!COINCMACCT)
         vAccounts(i, 2) = "" & Trim(!COINCMDESC)
         vAccounts(i, 3) = Format(!COINCMTYPE, "0")
         
         i = i + 1
         vAccounts(i, 0) = "" & Trim(!COEXPNREF)
         vAccounts(i, 1) = "" & Trim(!COEXPNACCT)
         vAccounts(i, 2) = "" & Trim(!COEXPNDESC)
         vAccounts(i, 3) = Format(!COEXPNTYPE, "0")
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = i + 1 '5
            vAccounts(i, 0) = "" & Trim(!COCOGSREF)
            vAccounts(i, 1) = "" & Trim(!COCOGSACCT)
            vAccounts(i, 2) = "" & Trim(!COCOGSDESC)
            vAccounts(i, 3) = Format(!COCOGSTYPE, "0")
         End If
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = i + 1 '6
            vAccounts(i, 0) = "" & Trim(!COOINCREF)
            vAccounts(i, 1) = "" & Trim(!COOINCACCT)
            vAccounts(i, 2) = "" & Trim(!COOINCDESC)
            vAccounts(i, 3) = Format(!COOINCTYPE, "0")
         End If
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = i + 1 '7
            vAccounts(i, 0) = "" & Trim(!COOEXPREF)
            vAccounts(i, 1) = "" & Trim(!COOEXPACCT)
            vAccounts(i, 2) = "" & Trim(!COOEXPDESC)
            vAccounts(i, 3) = Format(!COOEXPTYPE, "0")
         End If
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = i + 1 '8
            vAccounts(i, 0) = "" & Trim(!COFDTXREF)
            vAccounts(i, 1) = "" & Trim(!COFDTXACCT)
            vAccounts(i, 2) = "" & Trim(!COFDTXDESC)
            vAccounts(i, 3) = Format(!COFDTXTYPE, "0")
         End If
         .Cancel
      End With
   End If
   iTotal = i
   For i = 0 To iTotal
      'cmbMst.AddItem vAccounts(i, 1)
      AddComboStr cmbMst.hWnd, Format$(vAccounts(i, 1))
   Next
   If cmbMst.ListCount > 0 Then
      cmbMst = cmbMst.List(0)
      lblTyp(0) = vAccounts(0, 3)
      lblDsc = vAccounts(0, 2)
   End If
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR,GLTYPE FROM GlacTable "
   bSqlRows = GetDataSet(rdoGlm)
   If bSqlRows Then
      With rdoGlm
         Do Until .EOF
            iTotal = iTotal + 1
            vAccounts(iTotal, 0) = "" & Trim(!GLACCTREF)
            vAccounts(iTotal, 1) = "" & Trim(!GLACCTNO)
            vAccounts(iTotal, 2) = "" & Trim(!GLDESCR)
            vAccounts(iTotal, 3) = Format(!GLTYPE, "0")
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr cmbMst.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
      End With
   End If
   Set rdoGlm = Nothing
   If cmbAct.ListCount > 0 Then
      If bGetAccount Then
         cmbAct = cmbAct.List(0)
         bGoodAccount = GetAccount()
      End If
   End If
   
   'Now fill vendor combo box
   sSql = "SELECT DISTINCT VENICKNAME FROM VndrTable"
   bSqlRows = GetDataSet(RdoVnd)
   If bSqlRows Then
      With RdoVnd
         While Not .EOF
            AddComboStr cmbVnd.hWnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Wend
      End With
   End If
   Set RdoVnd = Nothing
   
   MouseCursor 0
   Exit Sub
   
   diaErr1:
   sProcName = "fillaccou"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetAccount() As Byte
   Dim bMaster As Boolean
   Dim i As Integer
   Dim sAccount As String
   Dim sMaster As String
   Dim RdoVnd As rdoResultset
   
   sAccount = Compress(cmbAct)
   On Error GoTo diaErr1
   
   MouseCursor 13
   
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR,GLMASTER,GLTYPE,GLCASH,GLVENDOR," _
          & "GLFSLEVEL,GLINACTIVE FROM GlacTable WHERE GLACCTREF='" & sAccount & "' "
   bSqlRows = GetDataSet(rdoAct, ES_KEYSET)
   If bSqlRows Then
      With rdoAct
         
         cmbAct = "" & Trim(!GLACCTNO)
         txtDsc = "" & Trim(!GLDESCR)
         lblTyp(0) = Format(!GLTYPE, "0")
         txtLvl = Format(!GLFSLEVEL, "0")
         optAct.Value = !GLINACTIVE
         
         sMaster = "" & Trim(!GLMASTER)
         
         If !GLCASH Then
            sSql = "SELECT VENICKNAME FROM VndrTable WHERE VEREF = '" _
                   & !GLVENDOR & "'"
            bSqlRows = GetDataSet(RdoVnd)
            If bSqlRows Then
               cmbVnd = RdoVnd!VENICKNAME
               Set RdoVnd = Nothing
               FindVendor Me
            End If
            cmbVnd.Enabled = True
         Else
            cmbVnd = ""
            cmbVnd.Enabled = False
         End If
         optCash.Value = !GLCASH
         .Cancel
      End With
      
      For i = 0 To 8
         If sMaster = vAccounts(i, 0) Then
            bMaster = True
            Exit For
         End If
      Next
      If bMaster Then
         sOldAccount = sMaster
         cmbMst.ListIndex = i
         lblTyp(0) = vAccounts(i, 3)
         cmbMst = vAccounts(i, 1)
         lblDsc = vAccounts(i, 2)
      Else
         FindMasterAccount sMaster
      End If
      GetAccount = True
   Else
      bGoodAccount = False
      txtDsc = ""
      txtLvl = "0"
      optAct.Value = vbUnchecked
      GetAccount = False
   End If
   MouseCursor 0
   Exit Function
   
   diaErr1:
   sProcName = "getaccount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Sub AddAccount()
   Dim b As Byte
   Dim i As Integer
   Dim iLevel As Integer
   Dim bResponse As Byte
   Dim smsg As String
   Dim sAccount As String
   Dim sNewAcct As String
   Dim sMaster As String
   
   On Error GoTo diaErr1
   sNewAcct = cmbAct
   sAccount = Compress(cmbAct)
   sMaster = Compress(cmbMst)
   smsg = cmbAct & " Wasn't Found. Add The Account?"
   bResponse = MsgBox(smsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      iLevel = Val(txtLvl) + 1
      If iLevel > 8 Then iLevel = 9
      For i = 0 To 8
         If sAccount = vAccounts(i, 0) Then b = 1
      Next
      If b = 1 Then
         MsgBox "That Account Number Is In Use By A Master Account.", _
            vbInformation, Caption
         Exit Sub
      End If
      On Error Resume Next
      sSql = "INSERT INTO GlacTable(GLACCTREF,GLACCTNO," _
             & "GLMASTER,GLTYPE,GLFSLEVEL,GLINACTIVE) VALUES('" & sAccount & "','" _
             & cmbAct & "','" & sMaster & "'," & Val(lblTyp(0)) & "," _
             & iLevel & ",0)"
      RdoCon.Execute sSql, rdExecDirect
      If Err = 0 Then
         RdoCon.CommitTrans
         Sysmsg "Account Was Added.", True
         FillAccounts False
         cmbAct = sNewAcct
         bGoodAccount = GetAccount()
      Else
         RdoCon.RollbackTrans
         MsgBox "Couldn't Add Account.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
   diaErr1:
   sProcName = "addaccount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub FindMasterAccount(sMaster As String)
   Dim RdoMst As rdoResultset
   On Error GoTo diaErr1
   sMaster = Compress(sMaster)
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR,GLMASTER,GLTYPE," _
          & "GLINACTIVE FROM GlacTable WHERE GLACCTREF='" & sMaster & "' " _
          & "AND GLINACTIVE=0"
   bSqlRows = GetDataSet(RdoMst)
   If bSqlRows Then
      With RdoMst
         sOldAccount = "" & Trim(!GLACCTNO)
         cmbMst = "" & Trim(!GLACCTNO)
         lblDsc = "" & Trim(!GLDESCR)
         lblTyp(0) = Format(!GLTYPE, "0")
         .Cancel
      End With
   Else
      sOldAccount = ""
      lblTyp(0) = "0"
      lblDsc = ""
   End If
   Set RdoMst = Nothing
   Exit Sub
   
   diaErr1:
   sProcName = "findmaste"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub DeleteAccount()
   Dim RdoMst As rdoResultset
   Dim bUsed As Boolean
   Dim sMaster As String
   Dim sUsedOn As String
   
   On Error GoTo diaErr1
   sMaster = Compress(cmbAct)
   sSql = "SELECT GLACCTREF,GLACCTNO,GLMASTER FROM " _
          & "GlacTable WHERE GLMASTER='" & sMaster & "' "
   bSqlRows = GetDataSet(RdoMst, ES_FORWARD)
   If bSqlRows Then
      With RdoMst
         bUsed = True
         sUsedOn = Trim(!GLACCTNO)
         .Cancel
      End With
   End If
   If Not bUsed Then
      sSql = "SELECT DISTINCT SHPACCT FROM ShopTable " _
             & "WHERE SHPACCT='" & sMaster & "' "
      RdoCon.Execute sSql, rdExecDirect
      If RdoCon.RowsAffected Then
         MsgBox "This Account Is Used On A Shop." & vbCrLf _
            & "ES/2000 Cannot Delete Account " & cmbAct & "..", _
            vbExclamation, Caption
         Exit Sub
      Else
         bUsed = False
      End If
   End If
   If Not bUsed Then
      sSql = "SELECT  WCNACCT FROM WcntTable " _
             & "WHERE WCNACCT='" & sMaster & "' "
      RdoCon.Execute sSql, rdExecDirect
      If RdoCon.RowsAffected Then
         MsgBox "This Account Is Used On A Work Center." & vbCrLf _
            & "ES/2000 Cannot Delete Account " & cmbAct & "..", _
            vbExclamation, Caption
         Exit Sub
      Else
         bUsed = False
      End If
   End If
   If Not bUsed Then
      sSql = "SELECT * FROM PcodTable Where " _
             & "PCREVACCT='" & sMaster & "' OR " _
             & "PCDISCACCT='" & sMaster & "' OR " _
             & "PCDREVXFERAC='" & sMaster & "' OR " _
             & "PCDCGSXFERAC='" & sMaster & "' OR " _
             & "PCINVEXPAC='" & sMaster & "' OR " _
             & "PCCGSAC='" & sMaster & "'"
      bSqlRows = GetDataSet(RdoMst, ES_FORWARD)
      If bSqlRows Then
         sUsedOn = "" & Trim(RdoMst!PCCODE)
         bUsed = True
      End If
      If bUsed Then
         MsgBox "This Account Is Used On Product Code " & sUsedOn & vbCrLf _
            & "ES/2000 Cannot Delete Account " & cmbAct & "..", _
            vbExclamation, Caption
      Else
         sSql = "SELECT * FROM PartTable Where " _
                & "PAACCTNO='" & sMaster & "' OR " _
                & "PAREVACCT='" & sMaster & "' OR " _
                & "PACGSACCT='" & sMaster & "' OR " _
                & "PADISACCT='" & sMaster & "' OR " _
                & "PATFRREVACCT='" & sMaster & "' OR " _
                & "PATFRCGSACCT='" & sMaster & "' OR " _
                & "PAREJACCT='" & sMaster & "' "
         RdoCon.Execute sSql
         bUsed = RdoCon.RowsAffected
         If bUsed Then
            MsgBox "This Account Is Used On A Part Number." & vbCrLf _
               & "ES/2000 Cannot Delete Account " & cmbAct & "..", _
               vbExclamation, Caption
         Else
            sSql = "DELETE FROM GlacTable WHERE " _
                   & "GLACCTREF='" & sMaster & "' "
            RdoCon.Execute sSql, rdExecDirect
            bUsed = RdoCon.RowsAffected
            If bUsed Then
               Sysmsg "Account Was Deleted.", True
               FillAccounts True
            Else
               MsgBox "Account In Use. Couldn't Delete.", _
                  vbExclamation, Caption
            End If
         End If
      End If
   Else
      MsgBox "This Account Is Used On Lower Level " & sUsedOn & vbCrLf _
         & "ES/2000 Cannot Delete Account " & cmbAct & "..", _
         vbExclamation, Caption
   End If
   Set RdoMst = Nothing
   Exit Sub
   
   diaErr1:
   sProcName = "deleteacc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub

'This code is temporary and can go away
'Adjusts types that were backward in MCS and copied
'here  6/24/99

Public Sub CheckAccounts()
   Dim rdoChk As rdoResultset
   On Error GoTo diaErr1
   MouseCursor 13
   sSql = "SELECT COCOGSTYPE FROM GlmsTable " _
          & "WHERE COREF=1"
   bSqlRows = GetDataSet(rdoChk, ES_FORWARD)
   If bSqlRows Then
      With rdoChk
         If !COCOGSTYPE = 6 Then
            sSql = "UPDATE GlmsTable SET COCOGSTYPE=5," _
                   & "COEXPNTYPE=6 WHERE COREF=1"
            RdoCon.Execute sSql, rdExecDirect
            
            sSql = "UPDATE GlacTable SET GLTYPE=10 WHERE GLTYPE=5 "
            RdoCon.Execute sSql, rdExecDirect
            
            sSql = "UPDATE GlacTable SET GLTYPE=5 WHERE GLTYPE=6 "
            RdoCon.Execute sSql, rdExecDirect
            
            sSql = "UPDATE GlacTable SET GLTYPE=6 WHERE GLTYPE=10"
            RdoCon.Execute sSql, rdExecDirect
         End If
      End With
   End If
   
   MouseCursor 0
   Exit Sub
   
   diaErr1:
   sProcName = "checkacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
