VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AdmnQAe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Information"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnQAe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtQaIntf 
      Height          =   285
      Left            =   5640
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "International Prefix (Country Code)"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox txtQaIntp 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "International Prefix (Country Code)"
      Top             =   2880
      Width           =   375
   End
   Begin VB.ComboBox cmbSte 
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Tag             =   "3"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtAdr 
      Height          =   855
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   2
      Tag             =   "9"
      Top             =   1200
      Width           =   3475
   End
   Begin VB.TextBox txtCty 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Tag             =   "2"
      Top             =   2160
      Width           =   2085
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      Tag             =   "1"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtNme 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "2"
      Top             =   840
      Width           =   3495
   End
   Begin MSMask.MaskEdBox txtFax 
      Height          =   285
      Left            =   6020
      TabIndex        =   11
      Top             =   2880
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   12
      Mask            =   "###-###-####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtRep 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Tag             =   "2"
      Top             =   2520
      Width           =   2085
   End
   Begin VB.TextBox txtEml 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Tag             =   "2"
      ToolTipText     =   "Double Click Here To Send E-Mail (Requires An Entry)"
      Top             =   3240
      Width           =   5050
   End
   Begin VB.TextBox txtExt 
      Height          =   285
      Left            =   4320
      TabIndex        =   9
      Tag             =   "1"
      Top             =   2880
      Width           =   615
   End
   Begin VB.ComboBox cmbCst 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   480
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
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
      FormDesignHeight=   3900
      FormDesignWidth =   7410
   End
   Begin MSMask.MaskEdBox txtPhn 
      Height          =   285
      Left            =   1940
      TabIndex        =   8
      Top             =   2880
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   12
      Mask            =   "###-###-####"
      PromptChar      =   "_"
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ZIP"
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   25
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   285
      Index           =   4
      Left            =   3720
      TabIndex        =   22
      Top             =   2160
      Width           =   675
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quality Rep"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   21
      Top             =   2520
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      Height          =   285
      Index           =   9
      Left            =   5160
      TabIndex        =   20
      Top             =   2880
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   19
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   18
      Top             =   3240
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ext"
      Height          =   285
      Index           =   62
      Left            =   3720
      TabIndex        =   17
      Top             =   2880
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   495
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label lblNum 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3240
      TabIndex        =   14
      Top             =   480
      Width           =   450
   End
End
Attribute VB_Name = "AdmnQAe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/8/03 International Calling Codes
Option Explicit
Dim RdoCst As ADODB.Recordset
Dim bOnLoad As Byte
Dim bGoodCust As Byte
Dim bCanceled As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub BuildStateCodes()
   Dim iList As Integer
   Dim sStates(50, 2) As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   sStates(0, 0) = "WA"
   sStates(0, 1) = "Washington"
   
   sStates(1, 0) = "OR"
   sStates(1, 1) = "Oregon"
   
   sStates(2, 0) = "CA"
   sStates(2, 1) = "California"
   
   sStates(3, 0) = "ID"
   sStates(3, 1) = "Idaho"
   
   sStates(4, 0) = "NV"
   sStates(4, 1) = "Nevada"
   
   sStates(5, 0) = "UT"
   sStates(5, 1) = "Utah"
   
   sStates(6, 0) = "AZ"
   sStates(6, 1) = "Arizona"
   
   sStates(7, 0) = "MT"
   sStates(7, 1) = "Montana"
   
   sStates(8, 0) = "wy"
   sStates(8, 1) = "Wyoming"
   
   sStates(9, 0) = "co"
   sStates(9, 1) = "Colorado"
   
   sStates(10, 0) = "nm"
   sStates(10, 1) = "New Mexico"
   
   sStates(11, 0) = "nd"
   sStates(11, 1) = "north Dakota"
   
   sStates(12, 0) = "sd"
   sStates(12, 1) = "South Dakota"
   
   sStates(13, 0) = "ne"
   sStates(13, 1) = "Nebraska"
   
   sStates(14, 0) = "ks"
   sStates(14, 1) = "Kansas"
   
   sStates(15, 0) = "ok"
   sStates(15, 1) = "Oklahoma"
   
   sStates(16, 0) = "tx"
   sStates(16, 1) = "Texas"
   
   sStates(17, 0) = "mn"
   sStates(17, 1) = "Minnesota"
   
   sStates(18, 0) = "ia"
   sStates(18, 1) = "Iowa"
   
   sStates(19, 0) = "mo"
   sStates(19, 1) = "Missouri"
   
   sStates(20, 0) = "ar"
   sStates(20, 1) = "Arkansas"
   
   sStates(21, 0) = "la"
   sStates(21, 1) = "Louisiana"
   
   sStates(22, 0) = "wi"
   sStates(22, 1) = "Wisconsin"
   
   sStates(23, 0) = "il"
   sStates(23, 1) = "Illinois"
   
   sStates(24, 0) = "mi"
   sStates(24, 1) = "Michigan"
   
   sStates(25, 0) = "in"
   sStates(25, 1) = "Indiana"
   
   sStates(26, 0) = "ky"
   sStates(26, 1) = "Kentucky"
   
   sStates(27, 0) = "TN"
   sStates(27, 1) = "Tennessee"
   
   sStates(28, 0) = "ms"
   sStates(28, 1) = "Mississippi"
   
   sStates(29, 0) = "al"
   sStates(29, 1) = "Alabama"
   
   sStates(30, 0) = "ga"
   sStates(30, 1) = "Georgia"
   
   sStates(31, 0) = "fl"
   sStates(31, 1) = "Florida"
   
   sStates(32, 0) = "me"
   sStates(32, 1) = "Maine"
   
   sStates(33, 0) = "nh"
   sStates(33, 1) = "New Hampshire"
   
   sStates(34, 0) = "vt"
   sStates(34, 1) = "Vermont"
   
   sStates(35, 0) = "ma"
   sStates(35, 1) = "Maine"
   
   sStates(36, 0) = "ri"
   sStates(36, 1) = "Rhode Island"
   
   sStates(37, 0) = "ct"
   sStates(37, 1) = "Connecticut"
   
   sStates(38, 0) = "nj"
   sStates(38, 1) = "New Jersey"
   
   sStates(39, 0) = "de"
   sStates(39, 1) = "Delaware"
   
   sStates(40, 0) = "ny"
   sStates(40, 1) = "New York"
   
   sStates(41, 0) = "pa"
   sStates(41, 1) = "Pennsylvania"
   
   sStates(42, 0) = "oh"
   sStates(42, 1) = "Ohio"
   
   sStates(43, 0) = "md"
   sStates(43, 1) = "Maryland"
   
   sStates(44, 0) = "va"
   sStates(44, 1) = "Virginia"
   
   sStates(45, 0) = "nc"
   sStates(45, 1) = "North Carolina"
   
   sStates(46, 0) = "sc"
   sStates(46, 1) = "South Carolina"
   
   sStates(47, 0) = "ak"
   sStates(47, 1) = "Alaska"
   
   sStates(48, 0) = "hi"
   sStates(48, 1) = "Hawaii"
   
   sStates(49, 0) = "wv"
   sStates(49, 1) = "West Virginia"
   For iList = 0 To 49
      sSql = "INSERT INTO CsteTable (STATECODE,STATEDESC) " _
             & "VALUES('" & UCase$(sStates(iList, 0)) & "','" _
             & sStates(iList, 1) & "')"
      clsADOCon.ExecuteSQL sSql
   Next
   On Error Resume Next
   FillStates Me
   Exit Sub
   
DiaErr1:
   sProcName = "buildstate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbCst_Click()
   bGoodCust = GetCustomer()
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If bCanceled Then Exit Sub
   If Len(cmbCst) Then
      bGoodCust = GetCustomer()
      If bGoodCust = 0 Then AddCustomer
   End If
   
End Sub


Private Sub cmbSte_LostFocus()
   cmbSte = CheckLen(cmbSte, 4)
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUSTATE = "" & cmbSte
      RdoCst.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6501
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillStates Me
      If cmbSte.ListCount = 0 Then BuildStateCodes
      FillCustomers
      If cmbCst.ListCount > 0 Then
         cmbCst = cmbCst.List(0)
         bGoodCust = GetCustomer()
      End If
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoCst = Nothing
   Set AdmnQAe01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Function GetCustomer() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT CUREF,CUNICKNAME,CUNUMBER,CUNAME,CUADR,CUSTATE," _
          & "CUCITY,CUZIP,CUQAFAX,CUQAEMAIL,CUQAREP,CUQAPHONE,CUQAPHONEEXT," _
          & "CUQAINTPHONE,CUQAINTFAX FROM CustTable " _
          & "WHERE CUREF='" & Compress(cmbCst) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_KEYSET)
   If bSqlRows Then
      With RdoCst
         GetCustomer = 1
         On Error Resume Next
         cmbCst = "" & Trim(!CUNICKNAME)
         lblNum = Format(0 + !CUNUMBER, "##0")
         txtNme = "" & Trim(!CUNAME)
         txtAdr = "" & Trim(!CUADR)
         txtCty = "" & Trim(!CUCITY)
         cmbSte = "" & Trim(!CUSTATE)
         txtZip = "" & Trim(!CUZIP)
         txtRep = "" & Trim(!CUQAREP)
         txtPhn.Mask = ""
         txtFax.Mask = ""
         If Len(Trim(!CUQAPHONE)) Then
            txtPhn = "" & Trim(!CUQAPHONE)
         Else
            txtPhn = ""
         End If
         If Len(Trim(!CUQAFAX)) Then
            txtFax = "" & Trim(!CUQAFAX)
         Else
            txtFax = ""
         End If
         txtFax = "" & Trim(!CUQAFAX)
         txtPhn.Mask = "###-###-####"
         txtFax.Mask = "###-###-####"
         txtExt = Format(0 + !CUQAPHONEEXT, "###0")
         txtEml = "" & Trim(!CUQAEMAIL)
         txtQaIntp = "" & Trim(!CUQAINTPHONE)
         txtQaIntf = "" & Trim(!CUQAINTFAX)
      End With
   Else
      On Error Resume Next
      GetCustomer = 0
      lblNum = "0"
      txtNme = "*** No Current Customer ***"
      txtAdr = ""
      txtCty = ""
      txtZip = ""
      txtRep = ""
      txtExt = ""
      txtEml = ""
      txtPhn.Mask = ""
      txtFax.Mask = ""
      txtPhn = ""
      txtFax = ""
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getcustom"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 160)
   txtAdr = StrCase(txtAdr, ES_FIRSTWORD)
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUADR = "" & txtAdr
      RdoCst.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtCty_LostFocus()
   txtCty = CheckLen(txtCty, 18)
   txtCty = StrCase(txtCty)
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUCITY = "" & txtCty
      RdoCst.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtEml_DblClick()
   If Trim(txtEml) <> "" Then SendEMail Trim(txtEml)
   
End Sub


Private Sub txtEml_LostFocus()
   txtEml = CheckLen(txtEml, 60)
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUQAEMAIL = "" & txtEml
      RdoCst.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtExt_LostFocus()
   txtExt = CheckLen(txtExt, 4)
   txtExt = Format(Abs(Val(txtExt)), "##0")
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUQAPHONEEXT = Val(txtExt)
      RdoCst.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtFax_LostFocus()
   txtFax = CheckLen(txtFax, 12)
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUQAFAX = "" & txtFax
      RdoCst.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtNme_Change()
   If Left(txtNme, 9) = "*** No Cu" Then
      txtNme.ForeColor = ES_RED
      txtNme.BackColor = vbButtonFace
   Else
      txtNme.ForeColor = vbBlack
      txtNme.BackColor = vbWhite
   End If
   
End Sub

Private Sub txtNme_LostFocus()
   txtNme = CheckLen(txtNme, 40)
   txtNme = StrCase(txtNme)
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUNAME = "" & txtNme
      RdoCst.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtPhn_LostFocus()
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUQAPHONE = "" & txtPhn
      RdoCst.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtQaIntf_Change()
   If Len(txtQaIntf) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtQaIntf_LostFocus()
   txtQaIntf = CheckLen(txtQaIntf, 4)
   txtQaIntf = Format(Abs(Val(txtQaIntf)), "###")
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUQAINTFAX = txtQaIntf
      RdoCst.Update
   End If
   
End Sub


Private Sub txtQaIntp_Change()
   If Len(txtQaIntp) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtQaIntp_LostFocus()
   txtQaIntp = CheckLen(txtQaIntp, 4)
   txtQaIntp = Format(Abs(Val(txtQaIntp)), "###")
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUQAINTPHONE = txtQaIntp
      RdoCst.Update
   End If
   
End Sub


Private Sub txtRep_LostFocus()
   txtRep = CheckLen(txtRep, 20)
   txtRep = StrCase(txtRep)
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUQAREP = "" & txtRep
      RdoCst.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub AddCustomer()
   Dim bResponse As Byte
   Dim lNextCust As Long
   Dim sMsg As String
   Dim sCust As String
   lNextCust = GetNextNumber()
   
   On Error GoTo DiaErr1
   sMsg = cmbCst & " Hasn't Been Established." & vbCr _
          & "Add The New Customer?...        "
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      sCust = Compress(cmbCst)
      If sCust = "ALL" Then
         MsgBox "ALL Is An Illegal Customer Nickname.", _
            vbExclamation, Caption
         Exit Sub
      Else
         On Error Resume Next
         sSql = "INSERT INTO CustTable (CUREF,CUNICKNAME,CUNUMBER,CUSTATE) " _
                & "VALUES ('" & sCust & "','" & Trim(cmbCst) & "'," _
                & lNextCust & ",'" & cmbSte & "')"
         
         clsADOCon.ADOErrNum = 0
         clsADOCon.BeginTrans
         clsADOCon.ExecuteSQL sSql
         
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            MsgBox "The Customer Was Successfully Added.", _
               vbInformation, Caption
            AddComboStr cmbCst.hwnd, cmbCst
            bGoodCust = GetCustomer()
         Else
            bGoodCust = False
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            
            MsgBox "Couldn't Successfully Add The Customer.", _
               vbInformation, Caption
         End If
      End If
   Else
      CancelTrans
      On Error Resume Next
      cmbCst.SetFocus
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addcustom"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetNextNumber() As Long
   Dim RdoNxt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MAX(CUNUMBER) FROM CustTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNxt, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoNxt.Fields(0)) Then
         GetNextNumber = RdoNxt.Fields(0) + 1
      Else
         GetNextNumber = 1
      End If
      ClearResultSet RdoNxt
   Else
      GetNextNumber = 1
   End If
   Set RdoNxt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getnextnum"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtZip_LostFocus()
   txtZip = CheckLen(txtZip, 10)
   If bGoodCust Then
      On Error Resume Next
      RdoCst!CUZIP = "" & txtZip
      RdoCst.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub
