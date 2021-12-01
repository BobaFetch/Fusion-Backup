VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel Manufacturing Order Completions"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkReverse 
      Caption         =   "Remove Completed MO's from Inventory"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   2040
      Width           =   3975
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1500
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Qualified Part Numbers (CO)"
      Top             =   300
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Contains Qualified Runs"
      Top             =   1020
      Width           =   1095
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5760
      TabIndex        =   5
      ToolTipText     =   "Cancel MO Completion"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2940
      FormDesignWidth =   6825
   End
   Begin VB.Label lblCantReverse 
      Caption         =   "(Cannot remove if one or more completed lots have activity.)"
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   2340
      Width           =   4395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date "
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   17
      Top             =   1515
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete"
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   16
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5760
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5760
      TabIndex        =   14
      Top             =   2460
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblLvl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   2460
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   5760
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Yield Quantity"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   10
      Top             =   345
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1260
      TabIndex        =   9
      Top             =   660
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3180
      TabIndex        =   7
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   0
      Left            =   2460
      TabIndex        =   6
      Top             =   1020
      Width           =   1095
   End
End
Attribute VB_Name = "ShopSHf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'7/22/05 Added clean up to set INPQTY and INADATE
Option Explicit
'Dim RdoQry As rdoQuery
Dim bOnLoad As Byte
Dim bGoodRun As Byte
Dim bGoodJrn As Byte
Dim lRunno As Long

Dim cYield As Currency
Dim cRunExp As Currency
Dim cRunHours As Currency
Dim cRunLabor As Currency
Dim cRunMatl As Currency
Dim cRunOvHd As Currency

'Dim sLotNumber As String
Dim sPartNumber As String

Dim sCreditAcct As String
Dim sDebitAcct As String
'WIP
Dim sWipLabAcct As String
Dim sWipMatAcct As String
Dim sWipExpAcct As String
Dim sWipOhdAcct As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetRunCosts()
   Dim RdoCst As ADODB.Recordset
   cRunExp = 0
   cRunHours = 0
   cRunLabor = 0
   cRunMatl = 0
   cRunOvHd = 0
   
   sProcName = "getruncosts"
   sSql = "SELECT RUNREF,RUNNO,RUNCOST,RUNOHCOST,RUNCMATL," _
          & "RUNCEXP,RUNCHRS,RUNCLAB FROM RunsTable WHERE " _
          & "RUNREF='" & sPartNumber & "' AND RUNNO=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         cRunExp = !RUNCEXP
         cRunHours = !RUNCHRS
         cRunLabor = !RUNCLAB
         cRunMatl = !RUNCMATL
         cRunOvHd = !RUNOHCOST
         ClearResultSet RdoCst
      End With
   End If
   Set RdoCst = Nothing
   
End Sub

Private Function CheckShippedMO(strPartRef As String, lRunno As Long) As String
   'return number of lots where there has been activity for this MO
   Dim RdoPS As ADODB.Recordset
   Dim strDate As String
   Dim strLotNum As String
   
   On Error GoTo DiaErr1
   If strPartRef <> "" And lRunno <> 0 Then
      
      sSql = "SELECT DISTINCT INPSNUMBER,INLOTNumber,INADATE FROM RunsTable,InvaTable,PshdTable" _
               & " WHERE RUNREF = '" & Compress(strPartRef) & "' and Runno = '" & CStr(lRunno) & "'" _
            & " AND runlotnumber = inlotnumber AND inpsnumber <> '' " _
            & " AND INLOTNumber <> '' AND INTYPE = 25 AND " _
            & " inpsnumber = psnumber AND PSSHIPPED = 1"
      
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPS, ES_FORWARD)
      If bSqlRows Then
         With RdoPS
            If Not IsNull(!INPSNUMBER) Then
               CheckShippedMO = !INPSNUMBER
               strLotNum = !INLOTNumber
               strDate = "" & Trim(!INADATE)
               
            Else
               CheckShippedMO = ""
            End If
            ClearResultSet RdoPS
         End With
      End If
   End If
   Set RdoPS = Nothing
   
   If (CheckShippedMO <> "") Then
      ' Check if there is any return MO  = 4
      If (CheckAnyReturnMO(strLotNum, strDate) = True) Then
         CheckShippedMO = ""
      End If
   End If
   
   Exit Function

DiaErr1:
   sProcName = "checklots"
   CheckShippedMO = ""
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

'Private Function CheckLots() As Integer
'   'return number of lots where there has been activity for this MO
'   Dim RdoLots As ADODB.recordset
'   On Error GoTo DiaErr1
'   'If sLotNumber <> "" Then
'      sSql = "SELECT COUNT(LOINUMBER) As LotCount FROM LoitTable " _
'             & "WHERE LOINUMBER='" & sLotNumber & "'"
'      bsqlrows = clsadocon.getdataset(ssql, RdoLots, ES_FORWARD)
'      If bSqlRows Then
'         With RdoLots
'            If Not IsNull(!LotCount) Then
'               CheckLots = !LotCount
'            Else
'               CheckLots = 0
'            End If
'            ClearResultSet RdoLots
'         End With
'      End If
'   'End If
'   Set RdoLots = Nothing
'   Exit Function
'
'DiaErr1:
'   sProcName = "checklots"
'   CheckLots = 0
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Function


Private Function CheckAnyReturnMO(strLotNum As String, strDate As String) As Boolean
   'return number of lots where there has been activity for this MO
   Dim RdoInv As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT Count(*) as Cnt FROM InvaTable WHERE inlotnumber = '" & strLotNum & "' " _
            & " AND INTYPE = 4  AND INADATE >= '" & Format(strDate, "mm/dd/yy") & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   If bSqlRows Then
      With RdoInv
         If (!Cnt) > 0 Then
            CheckAnyReturnMO = True
         Else
            CheckAnyReturnMO = False
         End If
         ClearResultSet RdoInv
      End With
   End If
   Set RdoInv = Nothing
   
   '
   Exit Function

DiaErr1:
   sProcName = "CheckAnyReturnMO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function


Private Sub cmbPrt_Click()
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   GetRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) Then
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
      GetRuns
   End If
   
End Sub


Private Sub cmbRun_Click()
   bGoodRun = GetCurrRun()
   End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   bGoodRun = GetCurrRun()
   End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   cmbPrt = ""
   
End Sub


Private Sub cmdDel_Click()
   On Error GoTo DiaErr1
   bGoodJrn = RecheckJournal()
   If bGoodJrn = 0 Then Exit Sub
   If bGoodRun = 0 Then
      MsgBox "Requires A Valid Run. See Help.", _
         vbExclamation, Caption
   Else
      Dim strPartRef As String
      Dim lRunno As Long
      Dim strPSNumber As String
      
      strPartRef = Compress(cmbPrt)
      lRunno = Val(cmbRun)
      
      strPSNumber = CheckShippedMO(strPartRef, lRunno)
      If strPSNumber <> "" Then
         MsgBox "This MO has PackSlip and it is Shipped." & vbCr _
            & "MO Completion Cannot Be Canceled.", _
            vbInformation, Caption
      Else
         GetWipAccounts
         GetRunCosts
         CancelRunComplete
      End If
   End If
   Exit Sub
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4151
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   'sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE " _
   '       & "RUNREF= ? AND RUNSTATUS='CO' "
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ShopSHf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub FillCombo()
   Dim b As Byte
   On Error GoTo DiaErr1
   sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yyyy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For This Period.", _
         vbExclamation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   End If
   sProcName = "fillcombo"
   cmbPrt.Clear
'   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF" & vbCrLf _
'      & "FROM RunsTable" & vbCrLf _
'      & "JOIN PartTable ON PARTREF = RUNREF" & vbCrLf _
'      & "WHERE RUNSTATUS = 'CO'" & vbCrLf _
'      & "ORDER BY PARTREF"

      'get CO MO's and any non-CO MO's with partial completions
      sSql = "SELECT DISTINCT PARTREF, PARTNUM" & vbCrLf _
      & "FROM RunsTable" & vbCrLf _
      & "JOIN PartTable ON PARTREF = RUNREF" & vbCrLf _
      & "WHERE RUNSTATUS = 'CO'" & vbCrLf _
      & "UNION " & vbCrLf _
      & "SELECT DISTINCT PARTREF, PARTNUM" & vbCrLf _
      & "FROM RunsTable" & vbCrLf _
      & "JOIN PartTable ON PARTREF = RUNREF" & vbCrLf _
      & "JOIN InvaTable ON RUNREF = INMOPART AND RUNNO = INMORUN " & vbCrLf _
      & "AND INREF1 = 'COMPLETED RUN' AND RUNSTATUS NOT IN ('CA','CL','CO')" & vbCrLf _
      & "ORDER BY PARTNUM"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc, True)
   End If
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetRuns()
   Dim rdo As ADODB.Recordset
   cmbRun.Clear
   sPartNumber = Compress(cmbPrt)
   'RdoQry(0) = sPartNumber
   'bSqlRows = clsAdoCon.GetQuerySet(RdoRns, RdoQry, ES_FORWARD)
   'If bSqlRows Then
   sSql = "SELECT DISTINCT RUNNO" & vbCrLf _
      & "FROM RunsTable" & vbCrLf _
      & "WHERE RUNSTATUS = 'CO' AND RUNREF = '" & sPartNumber & "'" & vbCrLf _
      & "UNION " & vbCrLf _
      & "SELECT DISTINCT RUNNO" & vbCrLf _
      & "FROM RunsTable" & vbCrLf _
      & "JOIN InvaTable ON RUNREF = INMOPART AND RUNNO = INMORUN " & vbCrLf _
      & "AND INREF1 = 'COMPLETED RUN' AND RUNSTATUS NOT IN ('CA','CL','CO')" & vbCrLf _
      & "AND RUNREF = '" & sPartNumber & "'" & vbCrLf _
      & "ORDER BY RUNNO"
   
   If clsADOCon.GetDataSet(sSql, rdo) Then
      With rdo
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format(!Runno, "####0")
            .MoveNext
         Loop
      End With
   End If
   
   Set rdo = Nothing
   
   If cmbRun.ListCount > 0 Then
      cmbRun = cmbRun.List(0)
      If GetPreferenceValue("AutoSelectLastRun") = "1" Then cmbRun = cmbRun.List(cmbRun.ListCount - 1)
      bGoodRun = GetCurrRun()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetCurrRun() As Byte
   Dim RdoRun As ADODB.Recordset
   ' There is global variable defined
   'Dim lRunno As Long
   
   lRunno = Val(cmbRun)
   sPartNumber = Compress(cmbPrt)
   On Error GoTo DiaErr1
'   sSql = "SELECT PARTREF,PALEVEL,PAPRODCODE,RUNREF,RUNNO," _
'          & "RUNSTATUS,RUNYIELD,RUNCOMPLETE,RUNLOTNUMBER FROM PartTable,RunsTable " _
'          & "WHERE PARTREF=RUNREF AND (RUNREF='" & sPartNumber & "' " _
'          & "AND RUNNO=" & lRunno & ")"
   sSql = "SELECT PARTREF,PALEVEL,PAPRODCODE,RUNREF,RUNNO," _
          & "RUNSTATUS,RUNYIELD,RUNCOMPLETE FROM PartTable,RunsTable " _
          & "WHERE PARTREF=RUNREF AND (RUNREF='" & sPartNumber & "' " _
          & "AND RUNNO=" & lRunno & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblStat = "" & Trim(!RUNSTATUS)
         lblQty = Format(!RUNYIELD, ES_QuantityDataFormat)
         lblCode = "" & Trim(!PAPRODCODE)
         lblLvl = Format$(!PALEVEL, "0")
         lblDate = Format(!RUNCOMPLETE, "mm/dd/yyyy")
         'sLotNumber = "" & Trim(!RUNLOTNUMBER)
         cYield = !RUNYIELD
         ClearResultSet RdoRun
      End With
   Else
      'sLotNumber = ""
      lblStat = "**"
      lblQty = "0.000"
   End If
'   If lblStat = "CO" Then
'      GetCurrRun = 1
'   Else
   If lblStat <> "CA" And lblStat <> "CL" Then
      GetCurrRun = 1
   Else
      GetCurrRun = 0
      lblQty = "0.000"
      lblStat = "**"
   End If
   
   'if there is lot activity, lots cannot be reversed
   sSql = "select count(*) from LohdTable" & vbCrLf _
      & "where LOTMOPARTREF='" & sPartNumber & "' and LOTMORUNNO=" & lRunno _
      & " and LOTORIGINALQTY<>LOTREMAININGQTY"
   If clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD) Then
      If RdoRun.Fields(0).Value = 0 Then
         Me.chkReverse.Enabled = True
         chkReverse.Value = vbChecked
         'Me.lblCantReverse.Visible = False
      Else
         chkReverse.Enabled = False
         'lblCantReverse.Visible = True
         chkReverse.Value = vbUnchecked
      End If
   Else
      chkReverse.Enabled = False
      'lblCantReverse.Visible = True
      chkReverse.Value = vbUnchecked
   End If
      
   Set RdoRun = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcurrrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub lblStat_Change()
   If lblStat = "**" Then
      lblStat.ForeColor = ES_RED
   Else
      lblStat.ForeColor = Es_TextForeColor
   End If
   
End Sub


'Lots 3/21/02

Private Sub CancelRunComplete()
   Dim bResponse As Byte
   
   Dim iLength As Integer
   
   Dim lCOUNTER As Long
   Dim lSysCount As Long
   
   Dim cComQty As Currency
   Dim cStdCost As Currency
   
   Dim sDate As String
   Dim sMsg As String
   
   sPartNumber = Compress(cmbPrt)
   lRunno = Val(cmbRun)
   sDate = txtDte
   On Error GoTo DiaErr1
   sMsg = "This Procedure Cancels The Run Completion." & vbCr _
          & "Do You Really Want To Cancel This Completion?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      cComQty = Abs(Val(lblQty))
      cStdCost = GetPartCost(sPartNumber, ES_STANDARDCOST)
      iLength = Len(Trim(str(cmbRun)))
      iLength = 5 - iLength
      If iLength < 0 Then iLength = 0
      bResponse = GetPartAccounts(sPartNumber, sCreditAcct, sDebitAcct)
      lCOUNTER = GetLastActivity() + 1
      lSysCount = lCOUNTER
      
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      Dim sStatus As String
      Dim mo As New ClassMO
      sStatus = mo.GetOpenMoStatus(sPartNumber, lRunno)
      
      sSql = "UPDATE RunsTable" & vbCrLf _
         & "SET RUNSTATUS='PC'" & vbCrLf _
         & "WHERE RUNREF='" & sPartNumber & "'" & vbCrLf _
         & "AND RUNNO=" & lRunno & " AND RUNSTATUS = 'CO'"
      clsADOCon.ExecuteSQL sSql
      
      'reverse inventory activity if requested
      If chkReverse.Value = vbChecked Then
      
         sSql = "UPDATE RunsTable" & vbCrLf _
                & "SET RUNCOMPLETE=NULL,RUNYIELD=0,RUNPARTIALQTY=0," _
                & "RUNREMAININGQTY=RUNQTY,RUNCOST=0,RUNOHCOST=0," & vbCrLf _
                & "RUNCMATL=0,RUNCEXP=0," _
                & "RUNCHRS=0,RUNCLAB=0" & vbCrLf _
                & "WHERE RUNREF='" _
                & sPartNumber & "' AND RUNNO=" & lRunno & " "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "INSERT INTO LoitTable" & vbCrLf _
            & "(LOINUMBER,LOIRECORD,LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," & vbCrLf _
            & "LOIMOPARTREF,LOIMORUNNO,LOIACTIVITY,LOICOMMENT)" & vbCrLf _
                
         sSql = sSql & "select LOINUMBER,(select isnull(max(b.LOIRECORD),0) + 1 from LoitTable b WHERE " & vbCrLf _
               & " LOINUMBER in (select LOTNUMBER from LohdTable where LOTORIGINALQTY=LOTREMAININGQTY" & vbCrLf _
            & " and LOTREMAININGQTY<>0" & vbCrLf _
            & " and LOTMOPARTREF='" & sPartNumber & "' and LOTMORUNNO = " & lRunno & "))" & vbCrLf _
            & "," & IATYPE_CanceledMoCompletion & "," _
            & "LOIPARTREF,'" & sDate & "',-LOIQUANTITY," & vbCrLf _
            & "LOIMOPARTREF,LOIMORUNNO,(SELECT MAX(b.INNUMBER)+ 1 FROM InvaTable b) INNUMBER,'Cancel MO Run Compl'" & vbCrLf _
            & "from LoitTable" & vbCrLf _
            & "join InvaTable on INLOTNUMBER = LOINUMBER and INTYPE=" & IATYPE_MoCompletion & vbCrLf _
            & "where LOITYPE=" & IATYPE_MoCompletion & vbCrLf _
            & "and LOINUMBER in (select LOTNUMBER from LohdTable where LOTORIGINALQTY=LOTREMAININGQTY" & vbCrLf _
            & "and LOTREMAININGQTY<>0" & vbCrLf _
            & "and LOTMOPARTREF='" & sPartNumber & "' and LOTMORUNNO = " & lRunno & ")"
         
         Debug.Print sSql
         
         clsADOCon.ExecuteSQL sSql
         
         sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," & vbCrLf _
            & "INPDATE,INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," & vbCrLf _
            & "INMOPART,INMORUN,INTOTMATL,INTOTLABOR,INTOTEXP," & vbCrLf _
            & "INTOTOH,INTOTHRS,INWIPLABACCT,INWIPMATACCT," & vbCrLf _
            & "INDRLABACCT,INDRMATACCT,INDREXPACCT,INDROHDACCT," & vbCrLf _
            & "INCRLABACCT,INCRMATACCT,INCREXPACCT,INCROHDACCT," & vbCrLf _
            & "INWIPOHDACCT,INWIPEXPACCT,INLOTNUMBER,INUSER) " & vbCrLf
                
         sSql = sSql & "select " & IATYPE_CanceledMoCompletion & ",INPART,'CANCELED RUN COMPL',INREF2," & vbCrLf _
            & "'" & sDate & "','" & sDate & "',-INPQTY,-INAQTY,INAMT,INCREDITACCT,INDEBITACCT," & vbCrLf _
            & "INMOPART,INMORUN,INTOTMATL,INTOTLABOR,INTOTEXP," & vbCrLf _
            & "INTOTOH,INTOTHRS,INWIPLABACCT,INWIPMATACCT," & vbCrLf _
            & "INDRLABACCT,INDRMATACCT,INDREXPACCT,INDROHDACCT," & vbCrLf _
            & "INCRLABACCT,INCRMATACCT,INCREXPACCT,INCROHDACCT," & vbCrLf _
            & "INWIPOHDACCT,INWIPEXPACCT,INLOTNUMBER,'" & sInitials & "'" & vbCrLf _
            & "from InvaTable where INMOPART='" & sPartNumber & "' and INMORUN=" & lRunno & " and INTYPE=" & IATYPE_MoCompletion & vbCrLf _
            & "and INLOTNUMBER in (select LOTNUMBER from LohdTable where LOTORIGINALQTY=LOTREMAININGQTY" & vbCrLf _
            & "and LOTREMAININGQTY<>0" & vbCrLf _
            & "and LOTMOPARTREF='" & sPartNumber & "' and LOTMORUNNO = " & lRunno & ")"
         clsADOCon.ExecuteSQL sSql
         
         Debug.Print sSql
         
         'create reversing lot items
         
         ' Get Lot Number to get he
         

         
         
         'update part quantities
         sSql = "UPDATE PartTable SET PAQOH=PAQOH-" _
            & "(select sum(LOTREMAININGQTY) from LohdTable where LOTMOPARTREF='" & sPartNumber & "' and LOTMORUNNO = " & lRunno & ")," & vbCrLf _
            & "PALOTQTYREMAINING=PALOTQTYREMAINING-" _
            & "(select sum(LOTREMAININGQTY) from LohdTable where LOTMOPARTREF='" & sPartNumber & "' and LOTMORUNNO = " & lRunno & ")" & vbCrLf _
            & "where PARTREF='" & sPartNumber & "'"
         clsADOCon.ExecuteSQL sSql
         
         'zero out the lot records themselves
         sSql = "UPDATE LohdTable" & vbCrLf _
               & "SET LOTORIGINALQTY=0," & vbCrLf _
                & "LOTREMAININGQTY=0,LOTUNITCOST=0,LOTDATECOSTED=NULL," & vbCrLf _
                & "LOTTOTMATL=0,LOTTOTLABOR=0,LOTTOTEXP=0,LOTTOTOH=0," & vbCrLf _
                & "LOTTOTHRS=0,LOTAVAILABLE=0" & vbCrLf _
                & "where LOTMOPARTREF='" & sPartNumber & "' and LOTMORUNNO = " & lRunno & vbCrLf _
                & "and LOTREMAININGQTY <> 0"
         clsADOCon.ExecuteSQL sSql

      End If
      clsADOCon.CommitTrans
      AverageCost sPartNumber
      UpdateWipColumns lSysCount
      
      sMsg = "The run completion was canceled, Set To PC" & vbCrLf _
             & "and the parts were removed from inventory."
      MsgBox sMsg, vbInformation, Caption
      FillCombo
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "cancelrunc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetWipAccounts()
   sProcName = "getlaboracct"
   sWipLabAcct = GetLaborAcct(sPartNumber, lblCode, Val(lblLvl))
   sProcName = "getexpenseacct"
   sWipExpAcct = GetExpenseAcct(sPartNumber, lblCode, Val(lblLvl))
   sProcName = "getmaterialacct"
   sWipMatAcct = GetMaterialAcct(sPartNumber, lblCode, Val(lblLvl))
   sProcName = "getoverheadacct"
   sWipOhdAcct = GetOverHeadAcct(sPartNumber, lblCode, Val(lblLvl))
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   If Format(txtDte, "yyyy/mm/dd") < Format(lblDate, "yyyy/mm/dd") Then
      Beep
      txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   End If
   
End Sub



Private Function RecheckJournal() As Byte
   Dim b As Byte
   sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For " & txtDte & ".", _
         vbExclamation, Caption
      RecheckJournal = 0
   Else
      RecheckJournal = 1
   End If
   
End Function
