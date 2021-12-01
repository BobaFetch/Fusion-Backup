VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chart Of Accounts"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbvnd 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Select A Vendor For Cash Accounts Only"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox optCash 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.CheckBox optVew 
      Caption         =   "vew"
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   315
      Left            =   4800
      Picture         =   "diaGLe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Find An Account"
      Top             =   480
      Width           =   350
   End
   Begin VB.CheckBox optChg 
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "C&hange"
      Height          =   315
      Left            =   5880
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Change Current Account Number"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   5880
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Delete Current Account Number"
      Top             =   960
      Width           =   875
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "diaGLe01a.frx":0342
      Height          =   320
      Left            =   4320
      MaskColor       =   &H00808080&
      Picture         =   "diaGLe01a.frx":0CB4
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Show Existing Accounts"
      Top             =   480
      Width           =   350
   End
   Begin VB.ComboBox cmbMst 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CheckBox optAct 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtLvl 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Tag             =   "1"
      Top             =   2520
      Width           =   255
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "2"
      Top             =   840
      Width           =   3360
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLe01a.frx":1626
      PictureDn       =   "diaGLe01a.frx":176C
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   3840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4215
      FormDesignWidth =   6840
   End
   Begin Threed.SSFrame z2 
      Height          =   30
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   6675
      _Version        =   65536
      _ExtentX        =   11774
      _ExtentY        =   53
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   28
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   4200
      TabIndex        =   27
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblnme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3480
      TabIndex        =   26
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Account?"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Deactivate Account?"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Financial Statment Detail Level For This Account)"
      Height          =   495
      Index           =   5
      Left            =   2280
      TabIndex        =   16
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Level "
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Account Of Master Account"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "diaGLe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'************************************************************************************
' diaGLe01a - Add/Revise Chart Of Accounts
'
' Created: (cjs)
' Revisions:
'   11/21/01 (nth) Fixed error with lblTyp()
'   09/05/01 (nth) Add support for the cash account option / checking accounts
'   09/10/01 (nth) Add vendors to cash accounts
'   09/20/02 (nth) Fixed minor bug with delete account
'   11/05/02 (nth) Add a check for master accounts.  Will not let to create GL
'                  accounts without master account being created first.
'   04/17/03 (nth) Refresh the master account description and level.
'   06/16/04 (nth) Check for account activity before deleting.
'   09/17/04 (nth) Add IsMaster funtion to prevent user entering in a master account.
'
'************************************************************************************

Dim rdoAct As ADODB.Recordset
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodAccount As Byte
Dim iTotal As Integer
Dim sOldAccount As String
Dim sAccount As String
Dim sMsg As String
Dim vAccounts(2000, 4) As Variant
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Private Sub cmbAct_Click()
   bGoodAccount = GetAccount()
   sAccount = Compress(cmbAct)
End Sub

Private Sub cmbAct_LostFocus()
   cmbAct = CheckLen(cmbAct, 12)
   If Len(cmbAct) And Not bCancel Then
      sAccount = Compress(cmbAct)
      If Not IsMaster Then
         bGoodAccount = GetAccount()
         If Not bGoodAccount Then
            AddAccount
         End If
      Else
         bGoodAccount = False
      End If
   Else
      bGoodAccount = False
   End If
   sAccount = Compress(cmbAct)
End Sub


Private Sub cmbMst_Click()
    
    If cmbMst.ListIndex >= 0 Then
        'lblTyp(0) = vAccounts(cmbMst.ListIndex, 3)
        'lblDsc = vAccounts(cmbMst.ListIndex, 2)
        Dim sMaster As String
        Dim i As Integer
        ' get the option master account number
        sMaster = cmbMst
        
        For i = 0 To cmbMst.ListCount - 1
         If sMaster = vAccounts(i, 0) Then
            lblDsc = vAccounts(i, 2)
            lblTyp(0) = vAccounts(i, 3)
            Exit For
         End If
        Next
        
   End If
End Sub

Private Sub cmbMst_LostFocus()
   Dim b As Byte
   Dim i As Integer
   Dim sMaster As String
   
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
    ' get the option master account number
    sMaster = cmbMst
   
    For i = 0 To cmbMst.ListCount - 1
     If sMaster = vAccounts(i, 0) Then
        lblDsc = vAccounts(i, 2)
        lblTyp(0) = vAccounts(i, 3)
        Exit For
     End If
    Next
   
   sOldAccount = Compress(cmbMst)
   lblTyp(1) = lblTyp(0)
   
   If bGoodAccount Then
      rdoAct!GLMASTER = Compress(cmbMst)
      rdoAct!GLTYPE = Val(lblTyp(0))
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub

Private Sub cmbMst_Validate(Cancel As Boolean)
    If GLAccountActivityExists(cmbMst) Then
        MsgBox "Please select an account with no activity/balance", vbOKOnly
        Cancel = True
    End If
    
End Sub

Private Sub cmbVnd_Click()
   FindVendor Me
End Sub

Private Sub cmbVnd_LostFocus()
   FindVendor Me
   rdoAct!GLVENDOR = Compress(cmbVnd)
   rdoAct.Update
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdChg_Click()
   optChg.Value = vbChecked
   diaGLf05a.Show
End Sub

Private Sub cmdDel_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "This Function Removes The Account.  " & vbCrLf _
          & "Cannot Delete Account With References.  " & vbCrLf _
          & "Do You Still Want To Delete This Account?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
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
      Unload diaGLf05a
      optChg.Value = vbUnchecked
   End If
   
   If optVew.Value = vbChecked Then
      Unload VewAcct
      optVew.Value = vbUnchecked
   End If
   
   If bOnLoad Then
      If Not FillAccounts(True) Then Unload Me
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   'Dim RdoTst As ADODB.Recordset
   FormLoad Me
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
   Set diaGLe01a = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optAct_Click()
   If bGoodAccount Then
      On Error Resume Next
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

      rdoAct!GLCASH = optCash.Value
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   If optCash Then
      cmbVnd.enabled = True
   Else
      cmbVnd = ""
      lblNme = ""
      cmbVnd.enabled = False

      rdoAct!GLVENDOR = ""
      rdoAct.Update
   End If
End Sub

Private Sub optCash_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub

Private Sub optCash_MouseDown(Button As Integer, Shift As Integer, _
                              x As Single, y As Single)
   
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

      rdoAct!GLFSLEVEL = Val(txtLvl)
      rdoAct!GLMASTER = Compress(cmbMst)
      rdoAct!GLTYPE = Val(lblTyp(0))
      rdoAct.Update
      If Err = 0 Then
         '    clsADOCon.CommitTrans
      Else
         '    clsADOCon.RollBackTrans
         ValidateEdit Me
      End If
   End If
   
End Sub

Public Sub ManageBoxes(bOpen As Boolean)
   If Not bOpen Then
      cmbAct.enabled = False
      txtLvl.enabled = False
      optAct.enabled = False
      cmdDel.enabled = False
   Else
      cmbAct.enabled = True
      txtLvl.enabled = True
      optAct.enabled = True
      cmdDel.enabled = True
   End If
   
End Sub

Public Function FillAccounts(bGetAccount As Boolean) As Boolean
   Dim i As Integer
   Dim RdoGlm As ADODB.Recordset
   Dim RdoVnd As ADODB.Recordset
   Dim sAccount As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   cmbAct.Clear
   cmbMst.Clear
   sSql = "SELECT * FROM GlmsTable WHERE COACCTREC=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      With RdoGlm
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
      AddComboStr cmbMst.hWnd, Format$(vAccounts(i, 1))
   Next
   
   If cmbMst.ListCount > 0 Then
      cmbMst = cmbMst.List(0)
      lblTyp(0) = vAccounts(0, 3)
      lblDsc = vAccounts(0, 2)
   Else
      'No master accounts.  Notify and exit
      sMsg = "Fina" & vbCrLf _
             & "Please Create GL Master Accounts" & vbCrLf _
             & "In Financial Statement Structure."
      MsgBox sMsg, vbInformation, Caption
      FillAccounts = False
   End If
   Set RdoGlm = Nothing
   
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR,GLTYPE FROM GlacTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm)
   If bSqlRows Then
      With RdoGlm
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
   Set RdoGlm = Nothing
   If cmbAct.ListCount > 0 Then
      If bGetAccount Then
         cmbAct = cmbAct.List(0)
         bGoodAccount = GetAccount()
      End If
   End If
   
   'Now fill vendor combo box
   sSql = "SELECT DISTINCT VENICKNAME FROM VndrTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd)
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
   FillAccounts = True
   Exit Function
   
DiaErr1:
   sProcName = "fillaccou"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Function GetAccount() As Byte
   Dim bMaster As Boolean
   Dim i As Integer
   Dim sMaster As String
   Dim RdoVnd As ADODB.Recordset
   
   sAccount = Compress(cmbAct)
   On Error GoTo DiaErr1
   MouseCursor 13
   
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR,GLMASTER,GLTYPE,GLCASH,GLVENDOR," _
          & "GLFSLEVEL,GLINACTIVE FROM GlacTable WHERE GLACCTREF='" & sAccount & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_KEYSET)
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
            bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd)
            If bSqlRows Then
               cmbVnd = RdoVnd!VENICKNAME
               Set RdoVnd = Nothing
               FindVendor Me
            End If
            cmbVnd.enabled = True
         Else
            cmbVnd = ""
            cmbVnd.enabled = False
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
         lblTyp(0) = "" & vAccounts(i, 3)
         cmbMst = "" & vAccounts(i, 1)
         lblDsc = "" & vAccounts(i, 2)
      Else
         FindMasterAccount sMaster
      End If
      GetAccount = True
      
      ' Refesh the master account detail
        For i = 0 To cmbMst.ListCount - 1
         If sMaster = vAccounts(i, 0) Then
            cmbMst = sMaster
            lblDsc = vAccounts(i, 2)
            lblTyp(0) = vAccounts(i, 3)
            Exit For
         End If
        Next

      
   Else
      bGoodAccount = False
      txtDsc = ""
      txtLvl = "0"
      optAct.Value = vbUnchecked
      GetAccount = False
   End If
   MouseCursor 0
   Exit Function
   
DiaErr1:
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
   Dim sMsg As String
   Dim sAccount As String
   Dim sNewAcct As String
   Dim sMaster As String
   
   On Error GoTo DiaErr1
   
   sNewAcct = cmbAct
   sAccount = Compress(cmbAct)
   sMaster = Compress(cmbMst)
   sMsg = cmbAct & " Wasn't Found. Add The Account?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
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
      clsADOCon.ADOErrNum = 0
      sSql = "INSERT INTO GlacTable(GLACCTREF,GLACCTNO," _
             & "GLMASTER,GLTYPE,GLFSLEVEL,GLINACTIVE) VALUES('" & sAccount & "','" _
             & cmbAct & "','" & sMaster & "'," & Val(lblTyp(0)) & "," _
             & iLevel & ",0)"
             
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "Account Was Added.", True
         FillAccounts False
         cmbAct = sNewAcct
         bGoodAccount = GetAccount()
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Couldn't Add Account.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addaccount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub FindMasterAccount(sMaster As String)
   Dim rdoMst As ADODB.Recordset
   Dim sMsg As String
   On Error GoTo DiaErr1
   sMaster = Compress(sMaster)
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR,GLMASTER,GLTYPE," _
          & "GLINACTIVE FROM GlacTable WHERE GLACCTREF='" & sMaster & "' " _
          & "AND GLINACTIVE=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoMst)
   If bSqlRows Then
      With rdoMst
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
      
      MsgBox "Master Account is InActive.", vbExclamation, Caption
      
   End If
   Set rdoMst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "findmaste"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub DeleteAccount()
   Dim rdoMst As ADODB.Recordset
   Dim rdoJrn As ADODB.Recordset
   Dim rdoGL As ADODB.Recordset
   Dim bUsed As Boolean
   Dim sMaster As String
   Dim sUsedOn As String
   
   On Error GoTo DiaErr1
   sMaster = Compress(cmbAct)
   sSql = "SELECT GLACCTREF,GLACCTNO,GLMASTER FROM " _
          & "GlacTable WHERE GLMASTER='" & sMaster & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoMst, ES_FORWARD)
   If bSqlRows Then
      With rdoMst
         bUsed = True
         sUsedOn = Trim(!GLACCTNO)
         .Cancel
      End With
   End If
   If Not bUsed Then
      sSql = "SELECT COUNT(DCHEAD) FROM JritTable " _
             & "WHERE DCACCTNO='" & sMaster & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn)
      If rdoJrn.Fields(0) > 0 Then
         MsgBox "This Account Has Journal Activity." & vbCrLf _
            & "ES/2000 Cannot Delete Account " & cmbAct & "..", _
            vbExclamation, Caption
         Set rdoJrn = Nothing
         Exit Sub
      Else
         Set rdoJrn = Nothing
         bUsed = False
      End If
   End If
   If Not bUsed Then
      sSql = "SELECT COUNT(JINAME) FROM GjitTable " _
             & "WHERE JIACCOUNT='" & sMaster & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoGL)
      If rdoGL.Fields(0) > 0 Then
         MsgBox "This Account Has General Ledger Activity." & vbCrLf _
            & "ES/2000 Cannot Delete Account " & cmbAct & ".", _
            vbInformation, Caption
         Set rdoGL = Nothing
         Exit Sub
      Else
         Set rdoGL = Nothing
         bUsed = False
      End If
   End If
   If Not bUsed Then
      sSql = "SELECT DISTINCT SHPACCT FROM ShopTable " _
             & "WHERE SHPACCT='" & sMaster & "' "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected Then
         MsgBox "This Account Is Used On A Shop." & vbCrLf _
            & "ES/2000 Cannot Delete Account " & cmbAct & ".", _
            vbInformation, Caption
         Exit Sub
      Else
         bUsed = False
      End If
   End If
   If Not bUsed Then
      sSql = "SELECT  WCNACCT FROM WcntTable " _
             & "WHERE WCNACCT='" & sMaster & "' "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.RowsAffected Then
         MsgBox "This Account Is Used On A Work Center." & vbCrLf _
            & "ES/2000 Cannot Delete Account " & cmbAct & "..", _
            vbInformation, Caption
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
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoMst, ES_FORWARD)
      If bSqlRows Then
         sUsedOn = "" & Trim(rdoMst!PCCODE)
         bUsed = True
      End If
      If bUsed Then
         MsgBox "This Account Is Used On Product Code " & sUsedOn & vbCrLf _
            & "ES/2000 Cannot Delete Account " & cmbAct & "..", _
            vbInformation, Caption
      Else
         sSql = "SELECT * FROM PartTable Where " _
                & "PAACCTNO='" & sMaster & "' OR " _
                & "PAREVACCT='" & sMaster & "' OR " _
                & "PACGSACCT='" & sMaster & "' OR " _
                & "PADISACCT='" & sMaster & "' OR " _
                & "PATFRREVACCT='" & sMaster & "' OR " _
                & "PATFRCGSACCT='" & sMaster & "' OR " _
                & "PAREJACCT='" & sMaster & "' "
         clsADOCon.ExecuteSQL sSql
         bUsed = clsADOCon.RowsAffected
         If bUsed Then
            MsgBox "This Account Is Used On A Part Number." & vbCrLf _
               & "ES/2000 Cannot Delete Account " & cmbAct & "..", _
               vbInformation, Caption
         Else
            sSql = "DELETE FROM GlacTable WHERE " _
                   & "GLACCTREF='" & sMaster & "' "
            clsADOCon.ExecuteSQL sSql
            bUsed = clsADOCon.RowsAffected
            If bUsed Then
               cmbAct = ""
               SysMsg "Account Was Deleted.", True
               FillAccounts True
            Else
               MsgBox "Account In Use. Couldn't Delete.", _
                  vbExclamation, Caption
            End If
         End If
      End If
   Else
      MsgBox "This Account Is Used On Lower Level " & sUsedOn & vbCrLf _
         & "ES/2000 Cannot Delete Account " & cmbAct & ".", _
         vbExclamation, Caption
   End If
   Set rdoMst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "deleteacc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function IsMaster() As Byte
   Dim rdoMst As ADODB.Recordset
   Dim i As Integer
   sProcName = "ismaster"
   ' First check if account attempting to revise is a master.
   ' Prevent user from adding or revising
   sSql = "SELECT COASSTACCT,COLIABACCT,COEQTYACCT,COINCMACCT,COEXPNACCT," _
          & "COCOGSACCT, COOINCACCT, COOEXPACCT, COFDTXACCT FROM GlmsTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoMst)
   If bSqlRows Then
      With rdoMst
         For i = 0 To 8
            If Compress(.Fields(i)) = sAccount Then
               IsMaster = True
               Exit For
            End If
         Next
         .Cancel
      End With
   End If
   Set rdoMst = Nothing
   If IsMaster Then
      sMsg = cmbAct & " Is A Master Account And Resides" & vbCrLf _
             & "In The Financial Statement Structure." & vbCrLf _
             & "Cannot Add Or Revise Account."
      MsgBox sMsg, vbInformation, Caption
      cmbAct.SetFocus
      Exit Function
   End If
End Function

Private Function GLAccountActivityExists(sGLAcctNo As String) As Byte
    Dim rdoGL As ADODB.Recordset
    
    GLAccountActivityExists = 0
    sSql = "SELECT TOP(1) DCACCTNO FROM JritTable WHERE DCACCTNO='" & sGLAcctNo & "'"
    On Error Resume Next
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoGL)
    If bSqlRows Then GLAccountActivityExists = 1
    Set rdoGL = Nothing
End Function
