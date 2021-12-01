VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change A Vendor Nickname"
   ClientHeight    =   2625
   ClientLeft      =   3000
   ClientTop       =   1710
   ClientWidth     =   6255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdDel 
      Cancel          =   -1  'True
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      ToolTipText     =   "Update And Apply Change To Vendor And Related Columns"
      Top             =   600
      Width           =   915
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      Top             =   1080
      Width           =   1555
   End
   Begin VB.TextBox txtNme 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   1440
      Width           =   3475
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   1680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2625
      FormDesignWidth =   6255
   End
   Begin VB.Label lblWrn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "A Vendor Is Recorded With That Nickname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Nickname"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label lblWrn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblWrn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please Close All Other Sections Before Proceeding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Nickname"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1095
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Name"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1425
   End
End
Attribute VB_Name = "PurcPRf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'7/7/05 Added SpapTable to list
'5/26/06 Added VnapTable, BuyvTable
Option Explicit
Dim bOnLoad As Byte
Dim sNewVendor As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtNme.BackColor = BackColor
   
End Sub

Private Function CheckWindows() As Byte
   Dim b As Byte
   b = Val(GetSetting("Esi2000", "Sections", "admn", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "fina", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "qual", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "invc", 0))
   If b > 0 Then
      lblWrn(0) = sSysCaption & " Has Determined " & b & " Other Open Section(s)"
      lblWrn(0).Visible = True
      lblWrn(1).Visible = True
      cmdDel.Enabled = False
   End If
   CheckWindows = b
   
End Function















Private Sub cmbCst_Click()
   GetDelVendor
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(Trim(cmbCst)) Then GetDelVendor
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdDel_Click()
   Dim b As Byte
   If Trim(txtNew) = "" Then Exit Sub
   b = TestNewVendor()
   If b = 0 Then
      MsgBox "That Vendor Has Been Previously Installed.", _
         vbInformation, Caption
      Exit Sub
   End If
   If txtNme.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Vendor.", _
         vbInformation, Caption
   Else
      ChangeTheVendor
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4352
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      CheckWindows
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   
   lblWrn(0).ForeColor = ES_RED
   lblWrn(1).ForeColor = ES_RED
   lblWrn(2).ForeColor = ES_RED
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PurcPRf03a = Nothing
   
End Sub



Private Sub FillCombo()
   On Error GoTo DiaErr1
   MouseCursor 13
   cmbCst.Clear
   sSql = "Qry_FillVendorsNone"
   LoadComboBox cmbCst
   MouseCursor 0
   If cmbCst.ListCount > 0 Then
      cmbCst = cmbCst.List(0)
      'GetDelVendor
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub GetDelVendor()
   Dim RdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT VEREF,VENICKNAME,VEBNAME FROM VndrTable WHERE " _
          & "VEREF='" & Compress(cmbCst) & "' AND VEREF<>'NONE'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         cmbCst = "" & Trim(!VENICKNAME)
         txtNme = "" & Trim(!VEBNAME)
         ClearResultSet RdoCst
      End With
   Else
      txtNme = "*** Vendor Wasn't Found ***"
   End If
   Set RdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdelcust"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub txtNew_Change()
   lblWrn(2).Visible = False
   
End Sub

Private Sub txtNew_LostFocus()
   txtNew = CheckLen(txtNew, 10)
   
End Sub


Private Sub txtNme_Change()
   If Left(txtNme, 6) = "*** Ve" Then
      txtNme.ForeColor = ES_RED
      cmdDel.Enabled = False
   Else
      txtNme.ForeColor = Es_TextForeColor
      cmdDel.Enabled = True
   End If
   
End Sub



Private Sub ChangeTheVendor()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sVendor As String
   
   
   sVendor = Compress(cmbCst)
   If sVendor = "ALL" Then
      Beep
      MsgBox "ALL Is An Illegal Vendor Nickname.", vbExclamation, Caption
      Exit Sub
   End If
   
   If sVendor = "NONE" Then
      Beep
      MsgBox "NONE Is An Illegal Vendor Nickname.", vbExclamation, Caption
      Exit Sub
   End If
   
   bResponse = IllegalCharacters(cmbCst)
   If bResponse > 0 Then
      MsgBox "The Nickname Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   sMsg = "It Is Not A Good Idea To Change A Vendor's Nickname " & vbCr _
          & "If There Is Any Chance That It Is In Use Right Now."
   MsgBox sMsg, vbExclamation, Caption
   
   sMsg = "This Function Permanently Changes The Vendor " & vbCr _
          & "Nickname Are You Sure That You Want To Continue?      "
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   DoEvents
   If bResponse = vbYes Then
      'start checking
      'Sales Orders
      On Error GoTo 0
      Screen.MousePointer = vbHourglass
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      'VENDOR
      'sSql = "UPDATE VndrTable SET VEREF='" & sNewVendor & "'," _
      '       & "VENICKNAME='" & Trim(txtNew) & "' WHERE " _
      '       & "VEREF='" & sVendor & "'"
      
      sSql = "INSERT INTO VndrTable " _
        & "(VEREF, VENICKNAME, VENUMBER, VETYPE, VEBNAME, VEBADR, VEBCITY, VEBSTATE, VEBZIP, VECNAME, VECADR, VECCITY, VECSTATE, VECZIP, " _
        & "VEFAX, VEBCONTACT, VEBPHONE, VEBEXT, VESCONTACT, VESPHONE, VESEXT, VENETDAYS, VEDDAYS, VEDISCOUNT, VEPROXDT, VEPROXDUE, " _
        & "VEFOB, VE1099, VETAXIDNO, VEDATEREV, VEAUTOAPINV, VECOMT, VEEMAIL, VEBUYER, VEINTFAX, VEINTPHONE, VEBINTPHONE, VESINTPHONE, " _
        & "VEACCOUNT, VEATTORNEY, VEBCOUNTRY, VECCOUNTRY, VEBEOM, VEAREMAIL, VEAPPROVREQ, VESURVEY, VESURVSENT, VESURVREC, VEAPPDATE, " _
        & "VEREVIEWDT, VEACCTNO) SELECT '" & txtNew & "','" & txtNew & "', VENUMBER, VETYPE, VEBNAME, VEBADR, VEBCITY, VEBSTATE, VEBZIP, " _
        & "VECNAME, VECADR, VECCITY, VECSTATE, VECZIP, VEFAX, VEBCONTACT, VEBPHONE, VEBEXT, VESCONTACT, VESPHONE, VESEXT, VENETDAYS, VEDDAYS, " _
        & "VEDISCOUNT, VEPROXDT, VEPROXDUE, VEFOB, VE1099, VETAXIDNO, VEDATEREV, VEAUTOAPINV, VECOMT, VEEMAIL, VEBUYER, VEINTFAX, VEINTPHONE, " _
        & "VEBINTPHONE, VESINTPHONE, VEACCOUNT, VEATTORNEY, VEBCOUNTRY, VECCOUNTRY, VEBEOM, VEAREMAIL, VEAPPROVREQ, VESURVEY, VESURVSENT, " _
        & "VESURVREC, VEAPPDATE, VEREVIEWDT , VEACCTNO FROM VndrTable WHERE VEREF='" & sVendor & "' "
      clsADOCon.ExecuteSQL sSql
      
      
      'Invoiced Header
      sSql = "INSERT INTO VihdTable " _
        & "(VINO, VIVENDOR, VIDATE, VIDTRECD, VIFREIGHT, VIFREIGHTINV, VITAX, VIBATCH, VICASHDISB, VIPIF, " _
        & " VIDUE, VIPAY, VIDISCOUNT, VISPAY, VISDISCOUNT, VIDUEDATE, VIUSETAXSTATE, VIUSETAXLOCAL, VIUSER," _
        & " VIREVDATE, VICOMT, VICHECKNO, VITYPE, VIBATCHNO, VICHKACCT) " _
        & "SELECT VINO, '" & txtNew & "', VIDATE, VIDTRECD, VIFREIGHT, VIFREIGHTINV, VITAX, VIBATCH, VICASHDISB, VIPIF, " _
        & " VIDUE, VIPAY, VIDISCOUNT, VISPAY, VISDISCOUNT, VIDUEDATE, VIUSETAXSTATE, VIUSETAXLOCAL, VIUSER, " _
        & " VIREVDATE, VICOMT, VICHECKNO, VITYPE, VIBATCHNO, VICHKACCT FROM VihdTable WHERE VIVENDOR = '" & sVendor & "' "
      clsADOCon.ExecuteSQL sSql

      'Invoice Items
      sSql = "INSERT INTO ViitTable " _
        & "(VITNO, VITVENDOR, VITITEM, VITPO, VITPORELEASE, VITPOITEM, VITPOITEMREV, VITQTY, VITCOST, VITMO, VITMORUN, " _
        & " VITACCOUNT, VITNOTE, VITCHECKNO, VITCHECKDT, VITCASHACCOUNT, VITDISCOUNT, VITPAID, VITTOTPAID, VITADDERS) " _
        & " SELECT VITNO, '" & txtNew & "', VITITEM, VITPO, VITPORELEASE, VITPOITEM, VITPOITEMREV, VITQTY, VITCOST, VITMO, VITMORUN, " _
        & " VITACCOUNT , VITNOTE, VITCHECKNO, VITCHECKDT, VITCASHACCOUNT, VITDISCOUNT, VITPAID, VITTOTPAID, VITADDERS " _
        & "FROM ViitTable WHERE VITVENDOR='" & sVendor & "' "
      clsADOCon.ExecuteSQL sSql
      
      'PO'S
      sSql = "UPDATE PohdTable SET POVENDOR='" & sNewVendor & "' " _
             & "WHERE POVENDOR='" & sVendor & "'"
      clsADOCon.ExecuteSQL sSql
      
      'invoice items
'      sSql = "UPDATE ViitTable SET VITVENDOR='" & sNewVendor & "' " _
'             & "WHERE VITVENDOR='" & sVendor & "'"
'      RdoCon.Execute sSql, rdExecDirect
'
'      'invoiced
'      sSql = "UPDATE VihdTable SET VIVENDOR='" & sNewVendor & "' " _
'             & "WHERE VIVENDOR='" & sVendor & "'"
'      RdoCon.Execute sSql, rdExecDirect
      

      
      'Journal
      sSql = "UPDATE JritTable SET DCVENDOR='" & sNewVendor & "' " _
             & "WHERE DCVENDOR='" & sVendor & "'"
      clsADOCon.ExecuteSQL sSql
      
      'RejTag?
      sSql = "UPDATE RjhdTable SET REJVENDOR='" & sNewVendor & "' " _
             & "WHERE REJVENDOR='" & sVendor & "'"
      clsADOCon.ExecuteSQL sSql
      
      'Checks
      sSql = "UPDATE ChksTable SET CHKVENDOR='" & sNewVendor & "' " _
             & "WHERE CHKVENDOR='" & sVendor & "'"
      clsADOCon.ExecuteSQL sSql
      
      'Gl
      sSql = "UPDATE GlacTable SET GLVENDOR='" & sNewVendor & "' " _
             & "WHERE GLVENDOR='" & sVendor & "'"
      clsADOCon.ExecuteSQL sSql
      
      'Chs
      sSql = "UPDATE ChseTable SET CHKVND='" & sNewVendor & "' " _
             & "WHERE CHKVND='" & sVendor & "'"
      clsADOCon.ExecuteSQL sSql
      
      'PO Items 10/20/03
      sSql = "UPDATE PoitTable SET PIVENDOR='" & sNewVendor _
             & "' WHERE PIVENDOR='" & sVendor & "' "
      clsADOCon.ExecuteSQL sSql
      
      'Alias 11/18/03
      sSql = "UPDATE PaalTable SET ALVENDOR='" & sNewVendor _
             & "' WHERE ALVENDOR='" & sVendor & "' "
      clsADOCon.ExecuteSQL sSql
      
      'Commissions 11/18/03
      sSql = "UPDATE SprsTable SET SPVENDOR='" & sNewVendor _
             & "' WHERE SPVENDOR='" & sVendor & "' "
      clsADOCon.ExecuteSQL sSql
      
      'Added 7/7/05
      sSql = "UPDATE SpapTable SET COAPVENDOR='" & sNewVendor _
             & "' WHERE COAPVENDOR='" & sVendor & "' "
      clsADOCon.ExecuteSQL sSql
      
      '5/26/06
      sSql = "UPDATE VnapTable SET AVVENDOR='" & sNewVendor _
             & "' WHERE AVVENDOR='" & sVendor & "' "
      clsADOCon.ExecuteSQL sSql
      
      '5/26/06
      sSql = "UPDATE BuyvTable SET BYVENDOR='" & sNewVendor _
             & "' WHERE BYVENDOR='" & sVendor & "' "
      clsADOCon.ExecuteSQL sSql
      
      sMsg = "Last Chance. Are You Sure That You Want" & vbCr _
             & "To Change Vendor " & cmbCst & "?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         sSql = "DELETE FROM ViitTable WHERE VITVENDOR='" & sVendor & "' "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "DELETE FROM VihdTable WHERE VIVENDOR = '" & sVendor & "' "
         clsADOCon.ExecuteSQL sSql
         
         
         sSql = "DELETE FROM  VndrTable WHERE VEREF='" & sVendor & "' "
         clsADOCon.ExecuteSQL sSql
         
         
         
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            SysMsg "Nickname Was Changed.", True
            txtNew = ""
            FillCombo
         Else
            clsADOCon.RollbackTrans
            MsgBox "Could Not Change The Vendor Nickname.", _
               vbExclamation, Caption
         End If
      Else
         clsADOCon.RollbackTrans
         CancelTrans
      End If
   Else
      CancelTrans
   End If
   Screen.MousePointer = vbNormal
End Sub

Private Function TestNewVendor() As Byte
   Dim RdoTst As ADODB.Recordset
   sNewVendor = Compress(txtNew)
   If sNewVendor = "" Then
      TestNewVendor = 0
      Exit Function
   End If
   On Error GoTo DiaErr1
   sSql = "SELECT VEREF FROM VndrTable WHERE VEREF='" & sNewVendor & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTst, ES_FORWARD)
   If bSqlRows Then
      ClearResultSet RdoTst
      lblWrn(2).Visible = True
      TestNewVendor = 0
   Else
      lblWrn(2).Visible = False
      TestNewVendor = 1
   End If
   Set RdoTst = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "testnewcu"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
