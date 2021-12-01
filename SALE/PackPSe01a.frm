VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PackPSe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Packing Slip"
   ClientHeight    =   2475
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   5865
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   2201
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      MaskColor       =   &H8000000F&
      Picture         =   "PackPSe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbSon 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Select Primary Sales Order From List (Contains Sales Orders With Items)"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   0
      Top             =   2880
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Tag             =   "4"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "PackPSe01a.frx":07AE
      Height          =   315
      Left            =   4960
      Picture         =   "PackPSe01a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "View Pack Slip List"
      Top             =   840
      Width           =   375
   End
   Begin VB.CheckBox optNew 
      Caption         =   "New"
      Height          =   195
      Left            =   5040
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cmbCst 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List (Contains Valid Customers)"
      Top             =   1080
      Width           =   1555
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      ToolTipText     =   "Add This New Packing Slip"
      Top             =   480
      Width           =   915
   End
   Begin VB.TextBox txtPsl 
      Height          =   285
      Left            =   1500
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "New Pack Slip Number (6 char max)"
      Top             =   720
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   480
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2475
      FormDesignWidth =   5865
   End
   Begin VB.Label lblPrefix 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   22
      Top             =   720
      Width           =   300
   End
   Begin VB.Label Info 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stuff Below here V"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Primary Sales Order)"
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   19
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblTerms 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label lblVia 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label lblStnme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label lblStadr 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   1200
      TabIndex        =   11
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label lblLst 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last "
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label txtNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1440
      Width           =   3495
   End
End
Attribute VB_Name = "PackPSe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/11/02 added the Primary SO number to control shipto
'4/27/05 Added last trap for duplicates (AddPackSlip)
'7/6/06 Blocked PS Number trap. Added additional test.
'7/25/07 TEL Don't show SOs with invoiced or cancelled items

Option Explicit
'Dim rdoQry As rdoQuery

Dim cmdObj As ADODB.Command
Dim RdoShp As ADODB.Recordset

Dim bOnLoad As Byte
Dim bGoodCust As Byte
Dim bGoodPs As Byte
Dim bSlipAdded As Byte

Dim lSoNum As Long
Dim isItarEarSo As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbCst_Click()
   bGoodCust = GetCustomer(False)
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   bGoodCust = GetCustomer(False)
End Sub

Private Sub cmbSon_Click()
   GetSoShipTo
End Sub

Private Sub cmbSon_LostFocus()
   ' make sure customer dropdown matches so
   Dim sostring As String, so As Long
   sostring = Trim(cmbSon)
   If Not IsNumeric(sostring) Then
      sostring = Right(sostring, Len(sostring) - 1)
   End If
   
   so = CLng(sostring)

   Dim rs As ADODB.Recordset
   sSql = "SELECT SOCUST FROM SohdTable " _
          & "WHERE SONUMBER=" & so
   bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
   If bSqlRows Then
      If Trim(rs!SOCUST) <> cmbCst Then
         cmbCst = Trim(rs!SOCUST)
      End If
   Else
      cmbCst = ""
   End If
   Set rs = Nothing
   
   GetSoShipTo False
End Sub

Private Sub cmdAdd_Click()
   bGoodPs = CheckPackSlip()
   If bGoodPs = 0 Then Exit Sub
   Timer1.Enabled = False
   ' The user can post PS for previous date
   'If Format(txtDte, "mm/dd/yy") < Format(ES_SYSDATE, "mm/dd/yy") Then txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   txtDte = Format(txtDte, "mm/dd/yyyy")
   If Len(Trim(txtPsl)) > 0 Then
      bGoodPs = CheckPackSlip()
   Else
      MsgBox "Requires A Pack Slip Number.", vbExclamation, Caption
      bGoodPs = False
   End If
   bGoodCust = GetCustomer(True, True)
   If bGoodPs And bGoodCust Then
      If Not GetSoShipTo() Then Exit Sub
      If isItarEarSo Then
         MsgBox "SO " & lSoNum & " is an ITAR/EAR sales order"
      End If
      
      AddPackslip
      If bSlipAdded Then
         MsgBox "Packing Slip Added.", vbInformation, Caption
         PackPSe02a.optNew.Value = vbChecked
         PackPSe02a.Show
      End If
   End If
End Sub

Private Sub cmdCan_Click()
   Timer1.Enabled = False
   optNew.Value = vbUnchecked
   Unload Me
   
End Sub




Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2201
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdVew_Click()
   Dim RdoVew As ADODB.Recordset
   Dim iList As Integer
   Dim iCol As Integer
   Dim iRows As Integer
   Dim sStatus As String
   
   On Error Resume Next
   iRows = 10
   With PackPSview.Grd
      .Rows = iRows
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      If Screen.Width > 9999 Then
         .ColWidth(0) = 1100 * 1.25
         .ColWidth(1) = 1550 * 1.25
         .ColWidth(2) = 1400 * 1.25
         .ColWidth(3) = 900 * 1.25
      Else
         .ColWidth(0) = 1100
         .ColWidth(1) = 1550
         .ColWidth(2) = 1400
         .ColWidth(3) = 900
      End If
   End With
   sSql = "SELECT PSNUMBER,PSCUST,PSDATE,PSSHIPPRINT,PSCANCELED," _
          & "CUREF,CUNICKNAME FROM PshdTable,CustTable WHERE " _
          & "(PSCUST=CUREF) AND PSNUMBER NOT LIKE 'S%' " _
          & "ORDER BY PSNUMBER DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVew, ES_FORWARD)
   If bSqlRows Then
      MouseCursor 13
      With RdoVew
         Do Until .EOF
            DoEvents
            iList = iList + 1
            iRows = iRows + 1
            PackPSview.Grd.Rows = iRows
            PackPSview.Grd.Row = iRows - 11
            
            PackPSview.Grd.Col = 0
            PackPSview.Grd = "" & Trim(!PsNumber)
            
            PackPSview.Grd.Col = 1
            PackPSview.Grd = "" & Format(!PSDATE, "mm/dd/yy")
            
            PackPSview.Grd.Col = 2
            PackPSview.Grd = Trim(!CUNICKNAME)
            
            If !PSSHIPPRINT = 1 Then
               sStatus = "P"
            Else
               If !PSCANCELED = 1 Then
                  sStatus = "C"
               Else
                  sStatus = ""
               End If
            End If
            PackPSview.Grd.Col = 3
            PackPSview.Grd = sStatus
            .MoveNext
         Loop
         ClearResultSet RdoVew
      End With
      If iList > 9 Then PackPSview.Grd.Rows = iList + 1
   End If
   MouseCursor 0
   Set RdoVew = Nothing
   On Error GoTo 0
   PackPSview.Show
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then
      Dim ps As New ClassPackSlip
      lblPrefix = ps.GetPackSlipPrefix
      txtPsl = ""
      txtPsl.MaxLength = 8 - Len(lblPrefix)
      
      GetPackslip True
      FillCombo
      Timer1.Enabled = True
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   
   sSql = "SELECT DISTINCT CUREF,CUNICKNAME,CUNAME," _
          & "CUSTNAME,CUSTADR,CUSTCITY,CUSTSTATE,CUSTZIP," _
          & "CUVIA,CUSTERMS,CUCUTOFF,SOCUST " _
          & "FROM CustTable,SohdTable WHERE (CUREF=SOCUST AND " _
          & "CUREF = ? )"
 '  Set rdoQry = RdoCon.CreateQuery("", sSql)
 '  rdoQry.MaxRows = 1
   
   Dim prmObj As ADODB.Parameter
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql

   Set prmObj = New ADODB.Parameter
   prmObj.Type = adChar
   prmObj.Size = 10
   cmdObj.parameters.Append prmObj
  
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   cUR.CurrentCustomer = cmbCst
   SaveCurrentSelections
   If optNew.Value = vbUnchecked Then FormUnload
   Set cmdObj = Nothing
   Set RdoShp = Nothing
   
   Set PackPSe01a = Nothing
   
End Sub

Private Sub FillCombo()
   ' fill customer combo box
   On Error GoTo DiaErr1
   '    sSql = "SELECT DISTINCT CUREF,CUNICKNAME,CUCUTOFF,SOCUST " _
   '        & "FROM CustTable,SohdTable WHERE (CUREF=SOCUST " _
   '        & "AND SOCANCELED=0)"
   
   sSql = "SELECT DISTINCT CUREF,CUNICKNAME,CUCUTOFF,SOCUST" & vbCrLf _
          & "FROM CustTable cust" & vbCrLf _
          & "join SohdTable so on cust.CUREF = so.SOCUST" & vbCrLf _
          & "join SoitTable item on item.ITSO = so.SONUMBER" & vbCrLf _
          & "where SOCANCELED = 0" & vbCrLf _
          & "AND ITACTUAL IS NULL AND ITQTY>0 AND ITPSNUMBER='' AND ITINVOICE=0 AND ITCANCELED=0" & vbCrLf _
          & "ORDER BY CUREF"
   
   LoadComboBox cmbCst
   If cmbCst.ListCount = 0 Then
      '    cmbCst = cmbCst.List(0)
      'Else
      MsgBox "No Customers With Open Sales Orders Found.", vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Timer1_Timer()
   GetPackslip False
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDte_LostFocus()

   Dim strJournalID As String
   txtDte = CheckDateEx(txtDte)
   '   Larry 5/13/00
   '   If Format(txtDte, "mm/dd/yy") < format(es_sysdate, "mm/dd/yy") Then
   '       MsgBox "The Date May Not Be Retroactive.", vbExclamation, Caption
   '       txtDte = format(es_sysdate, "mm/dd/yy")
   '   End If
' 4/16/2010 Users are allowed to create PS for previous week if journal is open.
   'CheckPeriodDate
   
   strJournalID = GetOpenJournal("IJ", Format(txtDte, "mm/dd/yyyy"))
   If Left(strJournalID, 4) = "None" Or (strJournalID = "") Then
      MsgBox "There Is No Open Inventory Journal For This" & vbCrLf _
         & "Period. Cannot Set Pack Slip date.", _
         vbExclamation, Caption
      ' No need to set the date to current date.
      'txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   End If
   
End Sub


Private Sub txtNme_Change()
   If txtNme = "*** Invalid Customer ***" Then
      txtNme.ForeColor = ES_RED
   Else
      txtNme.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub txtPsl_LostFocus()
   'txtPsl = CheckLen(txtPsl, txtpsl.maxlength)
   txtPsl = Format(Abs(Val(txtPsl)), String(txtPsl.MaxLength, "0"))
   '7/6/06
   '    If Val(txtPsl) <= Val(lblLst) Then
   '        Beep
   '        txtPsl = sOldPs
   '    Else
   '        sOldPs = txtPsl
   '    End If
   
End Sub


Private Sub GetPackslip(bFillText As Boolean)
'   Dim RdoGet As ADODB.Recordset
'   Dim l As Long
'   Dim n As Long
'   On Error GoTo DiaErr1
'   sSql = "SELECT CURPSNUMBER FROM ComnTable WHERE COREF=1"
'   bSqlRows = GetDataSet(RdoGet, ES_FORWARD)
'   If bSqlRows Then
'      With RdoGet
'         If Len(Trim(!CURPSNUMBER)) <= 6 Then
'            l = Val(!CURPSNUMBER)
'         Else
'            l = Val(Right$(!CURPSNUMBER, 6))
'         End If
'         ClearResultSet RdoGet
'      End With
'   Else
'      l = 0
'   End If
'   If Trim(lblLst) = "" Then
'      sSql = "SELECT PSNUMBER,PSTYPE,PSDATE FROM PshdTable " _
'             & "WHERE PSTYPE=1 AND PSNUMBER LIKE 'PS%' ORDER BY PSDATE,PSNUMBER DESC"
'      bSqlRows = GetDataSet(RdoGet)
'      If bSqlRows Then
'         With RdoGet
'            n = Val(Trim(Right$(!PsNumber, 6)))
'            ClearResultSet RdoGet
'         End With
'      End If
'   End If
'   'On Error Resume Next
'   If n > l Then
'      lblLst = Format(n, "000000")
'   Else
'      lblLst = Format(l, "000000")
'   End If
'   sOldPs = Format(Val(lblLst) + 1, "000000")
'   If bFillText Then txtPsl = sOldPs
'   Set RdoGet = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getpacksl"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
   Dim ps As New ClassPackSlip
   lblLst = ps.GetLastPackSlipNumber

   If bFillText Then
      txtPsl = Right(ps.GetNextPackSlipNumber, txtPsl.MaxLength)
   End If
End Sub

Private Function GetCustomer(bSendMessage As Boolean, Optional bDontFill As Boolean) As Byte
   Dim sCust As String
   sCust = Compress(cmbCst)
   On Error GoTo DiaErr1
   MouseCursor 13
'   rdoQry.RowsetSize = 1
'   rdoQry(0) = sCust
'   bSqlRows = GetQuerySet(RdoShp, rdoQry)
    cmdObj.parameters(0).Value = sCust
    bSqlRows = clsADOCon.GetQuerySet(RdoShp, cmdObj, ES_FORWARD, True)
    
   If bSqlRows Then
      With RdoShp
         cmbCst = "" & Trim(!CUNICKNAME)
         txtNme = "" & Trim(!CUNAME)
         'The SHIP VIA information should be obtained from the sales order
         'lblVia = "" & Trim(!CUVIA)
         If !CUCUTOFF = 0 Then
            cmdAdd.Enabled = True
            txtNme.ForeColor = vbBlack
         Else
            cmdAdd.Enabled = False
            txtNme = "*** This Customer's Credit Is On Hold ***"
            txtNme.ForeColor = ES_RED
         End If
         lblTerms = "" & Trim(!CUSTERMS)
         ClearResultSet RdoShp
      End With
      lblStnme = ReplaceString(lblStnme)
      lblStadr = ReplaceString(lblStadr)
      GetCustomer = 1
   Else
      lblStnme = ""
      lblStadr = ""
      'lblVia = ""
      lblTerms = ""
      txtNme = "*** Invalid Customer ***"
      If bSendMessage Then MsgBox "That Customer Has No Sales Orders Or Isn't Valid.", vbInformation, Caption
      GetCustomer = 0
   End If
   If Not bDontFill Then FillSalesOrders
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getcustom"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function CheckPackSlip() As Byte
   Dim RdoChk As ADODB.Recordset
   Dim sPackSlip As String
   txtPsl = Compress(txtPsl)
   Dim psno As String
   psno = lblPrefix & txtPsl
   
   On Error GoTo DiaErr1
   sSql = "SELECT PSNUMBER,PSTYPE FROM PshdTable " _
          & "WHERE PSNUMBER='" & psno & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD)
   If bSqlRows Then
      CheckPackSlip = 0
      MsgBox "Pack Slip Number " & psno & " Is In Use.", _
         vbInformation, Caption
      ClearResultSet RdoChk
      GetPackslip True
   Else
      CheckPackSlip = 1
   End If
   Set RdoChk = Nothing
   Exit Function
   
DiaErr1:
   CheckPackSlip = 0
   sProcName = "checkpack"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddPackslip()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sCust As String
   
   On Error GoTo DiaErr1
   Dim psno As String
   psno = lblPrefix & txtPsl
   sMsg = "Do You Wish To Record Packing Slip " & psno & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      bGoodPs = CheckPackSlip()
      If bGoodPs = 0 Then
         clsADOCon.RollbackTrans
         MsgBox "Packing Slip Number " & psno & " Is In Use.", _
            vbInformation, Caption
         GetPackslip False
         Exit Sub
      End If
      sCust = Compress(cmbCst)
      sSql = "INSERT INTO PshdTable (PSNUMBER,PSTYPE,PSDATE," _
             & "PSCUST,PSVIA,PSSTNAME,PSSTADR,PSTERMS,PSPRIMARYSO) " _
             & "VALUES('" & psno & "',1,'" & txtDte & "','" & sCust & "','" _
             & lblVia & "','" & Trim(lblStnme) & "','" & Trim(lblStadr) & "','" _
             & lblTerms & "'," & lSoNum & ")"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
      If clsADOCon.RowsAffected > 0 Then
         bSlipAdded = True
'         sSql = "UPDATE ComnTable SET CURPSNUMBER='" & Trim(txtPsl) & "' WHERE COREF=1"
'         RdoCon.Execute sSql, rdExecDirect
         Dim ps As New ClassPackSlip
         ps.SaveLastPSNumber psno
      End If
      clsADOCon.CommitTrans
      optNew.Value = vbChecked
   Else
      CancelTrans
      bSlipAdded = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addpacksl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'fundamental dates for Larry 5/13/00


Private Sub FillSalesOrders()
   Dim RdoCmb As ADODB.Recordset
   Dim sCust As String
   
   cmbSon.Clear
   sCust = Compress(cmbCst)
   On Error GoTo DiaErr1
   
   '    sSql = "SELECT DISTINCT SONUMBER,SOTYPE,ITSO FROM SohdTable,SoitTable " _
   '        & "WHERE SONUMBER=ITSO AND (SOCUST='" & sCust & "' AND ITACTUAL IS NULL " _
   '        & "AND ITQTY>0 AND ITPSNUMBER='' AND ITINV='')"
   
   sSql = "SELECT DISTINCT SOTYPE, SONUMBER" & vbCrLf _
          & "FROM SohdTable so" & vbCrLf _
          & "join SoitTable item on item.ITSO = so.SONUMBER" & vbCrLf _
          & "WHERE SOCUST='" & sCust & "'" & vbCrLf _
          & "AND ITACTUAL IS NULL AND ITQTY>0 AND ITPSNUMBER='' AND ITINVOICE=0 AND ITCANCELED=0" & vbCrLf _
          & "ORDER BY SONUMBER DESC"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         cmbSon = "" & Trim(!SOTYPE) & Format(!SoNumber, SO_NUM_FORMAT)
         Do Until .EOF
            AddComboStr cmbSon.hWnd, "" & Trim(!SOTYPE) & Format$(!SoNumber, SO_NUM_FORMAT)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
      GetSoShipTo False
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'2/11/02

Private Function GetSoShipTo(Optional DisplayErrorMessage As Boolean = True) As Boolean
   GetSoShipTo = False
   Dim RdoSto As ADODB.Recordset
   On Error GoTo DiaErr1
   If (Len(Trim(cmbSon)) = 6) Then
      lSoNum = Val(Right(Trim(cmbSon), (SO_NUM_SIZE - 1)))
   Else
      lSoNum = Val(Right(Trim(cmbSon), SO_NUM_SIZE))
   End If
   ' Get the ship VIA information
   Dim sCust As String
   sCust = Compress(cmbCst)
   sSql = "SELECT SONUMBER,SOSTNAME,SOSTADR, SOVIA,SOSTERMS, SOITAREAR FROM SohdTable " _
          & "WHERE SONUMBER=" & lSoNum & " and SOCUST ='" & sCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSto, ES_FORWARD)
   If bSqlRows Then
      With RdoSto
         lblStnme = "" & Trim(!SOSTNAME)
         lblStadr = "" & Trim(!SOSTADR)
         lblVia = "" & Trim(!SOVIA)
         lblTerms = "" & Trim(!SOSTERMS)
         isItarEarSo = !SOITAREAR
         ClearResultSet RdoSto
         GetSoShipTo = True
      End With
   Else
      If DisplayErrorMessage Then
         MsgBox lSoNum & " does not belong to customer " & sCust
      End If
      lSoNum = 0
      lblStnme = ""
      lblStadr = ""
      lblVia = ""
      lblTerms = ""
'      If (lSoNum = 0) Then
'         MsgBox ("Primary SO number is Zero")
'      End If
   End If
   Set RdoSto = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getsoship "
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
