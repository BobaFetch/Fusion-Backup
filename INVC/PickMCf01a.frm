VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PickMCf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Pick List"
   ClientHeight    =   2370
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PickMCf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   1680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2370
      FormDesignWidth =   6870
   End
   Begin VB.CommandButton cmdCpl 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   6
      ToolTipText     =   "Cancel This Pick List"
      Top             =   480
      Width           =   915
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   5880
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   960
      Width           =   900
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select MO Part Number"
      Top             =   960
      Width           =   3545
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblStu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "PickMCf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/7/04 See GetOpenLots
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodRuns As Byte
Dim bGoodPick As Byte

Dim sPartNumber As String
Dim sCreditAcct As String
Dim sDebitAcct As String

Dim vItems(250, 5) As Variant
Dim sLots(30, 6) As String 'See GetOpenLots
Dim sPartsGroup(250) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub cmbPrt_Click()
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCancel Then Exit Sub
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbRun_Click()
   GetStatus
   
End Sub


Private Sub cmbRun_LostFocus()
   GetStatus
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = True
   
End Sub


Private Sub cmdCpl_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Left(lblStu, 1) <> "P" Then
      MsgBox "Requires Status PL,PP or PC.", _
         vbInformation, Caption
      Exit Sub
   End If
   If Not bGoodRuns Then
      MsgBox "That Run Wasn't Listed.", vbInformation, Caption
      Exit Sub
   End If
   sMsg = "Cancels The Entire Pick/Pick List And Returns Status To SC. " & vbCr _
          & "Are You Sure That You Want To Continue?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      'On Error GoTo McnplCn1
      cmdCpl.Enabled = False
      bGoodPick = CancelPick()
      If bGoodPick Then
         MouseCursor 0
         MsgBox "Pick List Was Canceled.", vbInformation, Caption
         FillCombo
      Else
         MouseCursor 0
         MsgBox "Could Not Cancel The Pick..", vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
'McnplCn1:
'   Resume McnplCn2
'   CurrError.Description = Err.Description
'McnplCn2:
'   MouseCursor 0
'   On Error Resume Next
'   RdoCon.RollbackTrans
'   sMsg = CurrError.Description & vbCr _
'          & "Could Not Complete Pick List Cancel."
'   MsgBox sMsg, vbExclamation, Caption
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5250"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND RUNSTATUS LIKE 'P_' "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 30
   AdoParameter.Type = adChar
   AdoQry.Parameters.Append AdoParameter

   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set PickMCf01a = Nothing
   
End Sub



Private Sub FillCombo()
   Dim RdoPcl As ADODB.Recordset
   
   Dim b As Byte
   Dim sTempPart As String
   
   On Error GoTo DiaErr1
   sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
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
   
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PALEVEL,RUNREF,RUNSTATUS" & vbCrLf _
      & "FROM PartTable" & vbCrLf _
      & "JOIN RunsTable ON RUNREF = PARTREF" & vbCrLf _
      & "WHERE RUNSTATUS IN ('PL', 'PP', 'PC')" & vbCrLf _
      & "ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPcl)
   If bSqlRows Then
      With RdoPcl
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         Do Until .EOF
            If sTempPart <> Trim(!PartNum) Then
               AddComboStr cmbPrt.hWnd, "" & Trim(!PartNum)
               sTempPart = Trim(!PartNum)
            End If
            .MoveNext
         Loop
         ClearResultSet RdoPcl
      End With
      If cmbPrt.ListCount > 0 Then bGoodRuns = GetRuns()
   Else
      MsgBox "No Matching Runs Recorded.", _
         vbInformation, Caption
   End If
   On Error Resume Next
   Set RdoPcl = Nothing
   cmbPrt.SetFocus
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Function GetRuns()
   Dim RdoRns As ADODB.Recordset
   cmbRun.Clear
   FindPart Compress(cmbPrt)
   On Error GoTo DiaErr1
   AdoQry.Parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry)
   If bSqlRows Then
      With RdoRns
         cmbRun = Format(!Runno, "####0")
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      cmdCpl.Enabled = True
      GetRuns = True
   Else
      sPartNumber = ""
      cmdCpl.Enabled = False
      GetRuns = False
   End If
   GetStatus
   On Error Resume Next
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub


'Add lots 4/29/02

Private Function CancelPick() As Byte
   Dim RdoPck As ADODB.Recordset
   Dim bLots As Byte
   Dim b As Byte
   Dim bByte As Byte
   
   Dim A As Integer
   Dim iRow As Integer
   Dim iTrans As Integer
   Dim iRef As Integer
   
   Dim lCOUNTER As Long
   Dim lLOTRECORD As Long
   Dim lSysCount As Long
   
   Dim cCost As Currency
   Dim clineCost As Currency
   Dim cLotQty As Currency
   Dim cQuantity As Currency
   Dim sMoPart As String * 31
   Dim sMoRun As String * 9
   Dim sPkPartRef As String
   Dim sPartNumber As String
   Dim sPkDate As String
   
   'On Error GoTo DiaErr1
   On Error GoTo whoops
   MouseCursor 13
   sPartNumber = Compress(cmbPrt)
   sSql = "SELECT PKPARTREF,PKMOPART,PKMORUN,PKMORUN," _
          & "PKADATE,PKAQTY,PKAMT FROM MopkTable WHERE PKMOPART='" _
          & sPartNumber & "' AND PKMORUN=" & Val(cmbRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPck, ES_FORWARD)
   If bSqlRows Then
      With RdoPck
         sMoPart = cmbPrt
         sPkDate = Format(!PKADATE, "mm/dd/yy")
         iRow = Len(Trim(str(cmbRun)))
         iRow = 5 - iRow
         sMoRun = "RUN" & Space$(iRow) & cmbRun
         iRow = 0
         Do Until .EOF
            iRow = iRow + 1
            vItems(iRow, 0) = "" & Trim(!PKPARTREF)
            sPartsGroup(iRow) = "" & Trim(!PKPARTREF)
            vItems(iRow, 1) = !PKAQTY
            vItems(iRow, 2) = !PKAMT
            cCost = cCost + !PKAMT
            .MoveNext
         Loop
         ClearResultSet RdoPck
      End With
      Set RdoPck = Nothing
      If iRow > 0 Then
         A = iRow
         'On Error Resume Next
         clsADOCon.ADOErrNum = 0
         clsADOCon.BeginTrans
         
         lCOUNTER = GetLastActivity()
         lSysCount = lCOUNTER + 1
         For iRow = 1 To A
            sPkPartRef = sPartsGroup(iRow)
            bByte = GetPartAccounts(sPkPartRef, sDebitAcct, sCreditAcct)
            
            cQuantity = Format(Val(vItems(iRow, 1)), ES_QuantityDataFormat)
            clineCost = Format((Val(vItems(iRow, 1)) * vItems(iRow, 2)), ES_QuantityDataFormat)
            
            If cQuantity > 0 Then
               Dim strLoiNum As String
               Dim strPartRef As String
               Dim strLoiQty As String
               Dim strMOPartRef As String
               Dim strMORunNo As String
               Dim strLotUnitCost As String
               
               'lCOUNTER = lCOUNTER + 1
'               sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPQTY,INAQTY,INAMT," _
'                      & "INCREDITACCT,INDEBITACCT,INMOPART,INMORUN,INNUMBER,INUSER)  " _
'                      & "VALUES(11,'" & sPkPartRef & "','CANCELED PICK','" & sMoPart & sMoRun & "'," _
'                      & cQuantity & "," & cQuantity & "," & vItems(iRow, 2) & ",'" _
'                      & sCreditAcct & "','" & sDebitAcct & "','" _
'                      & sPartNumber & "'," & Val(cmbRun) & "," & lCOUNTER & ",'" & sInitials & "')"
'               RdoCon.Execute sSql, rdExecDirect
               
               bLots = GetOpenLots(sPkPartRef, sPartNumber, Val(cmbRun), sPkDate)
               cLotQty = 0
               'insert lot transaction here
               For b = 1 To bLots
                  strLoiNum = CStr(sLots(b, 0))
                  strPartRef = CStr(sLots(b, 1))
                  strLoiQty = CStr(sLots(b, 2))
                  strMOPartRef = CStr(sLots(b, 3))
                  strMORunNo = CStr(sLots(b, 4))
            
                  lLOTRECORD = GetNextLotRecord(sLots(b, 0))
                  cLotQty = Format(cLotQty + Val(sLots(b, 2)), ES_QuantityDataFormat)
                  
                  lCOUNTER = lCOUNTER + 1
                  strLotUnitCost = GetLotUnitCost(strLoiNum, strPartRef, strMOPartRef, strMORunNo)
                  
'                  sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPQTY,INAQTY,INAMT," _
'                         & "INCREDITACCT,INDEBITACCT,INLOTNUMBER, INMOPART,INMORUN,INNUMBER,INUSER)  " _
'                         & "VALUES(11,'" & strPartRef & "','CANCELED PICK','" & strMOPartRef & strMORunNo & "'," _
'                         & strLoiQty & "," & strLoiQty & ",'" & Trim(strLotUnitCost) & "','" _
'                         & sCreditAcct & "','" & sDebitAcct & "','" & sLots(b, 0) & "','" _
'                         & strMOPartRef & "'," & Val(cmbRun) & "," & lCOUNTER & ",'" & sInitials & "')"
                  
                  sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPQTY,INAQTY,INAMT," _
                         & "INCREDITACCT,INDEBITACCT,INLOTNUMBER, INMOPART,INMORUN,INNUMBER,INUSER)  " _
                         & "VALUES(11,'" & strPartRef & "','CANCELED ALL MO PICK','" & strMOPartRef & strMORunNo & "'," _
                         & strLoiQty & "," & strLoiQty & ",'" & Trim(strLotUnitCost) & "','" _
                         & sCreditAcct & "','" & sDebitAcct & "','" & sLots(b, 0) & "','" _
                         & strMOPartRef & "'," & Val(cmbRun) & "," & lCOUNTER & ",'" & sInitials & "')"
                  clsADOCon.ExecuteSQL sSql
                  
'                  sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
'                         & "LOITYPE,LOIPARTREF,LOIQUANTITY," _
'                         & "LOIMOPARTREF,LOIMORUNNO," _
'                         & "LOIACTIVITY,LOICOMMENT) " _
'                         & "VALUES('" & sLots(b, 0) & "'," _
'                         & lLOTRECORD & ",11,'" & sLots(b, 1) & "'," _
'                         & Val(sLots(b, 2)) & ",'" & sLots(b, 3) & "'," & Val(cmbRun) & "," _
'                         & lCOUNTER & ",'Canceled MO Pick')"

                  sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                         & "LOITYPE,LOIPARTREF,LOIQUANTITY," _
                         & "LOIMOPARTREF,LOIMORUNNO," _
                         & "LOIACTIVITY,LOICOMMENT) " _
                         & "VALUES('" & sLots(b, 0) & "'," _
                         & lLOTRECORD & ",11,'" & sLots(b, 1) & "'," _
                         & Val(sLots(b, 2)) & ",'" & sLots(b, 3) & "'," & Val(cmbRun) & "," _
                         & lCOUNTER & ",'Canceled All MO Picks')"
                  clsADOCon.ExecuteSQL sSql
                  
                  'Update the open lot in LoitTable talbe as MO canceled
                  
                  sSql = "UPDATE LoitTable SET LOIMOPKCANCEL=1 " _
                           & " WHERE LOINUMBER='" & sLots(b, 0) & "' AND " _
                           & " LOIPARTREF = '" & sLots(b, 1) & "' AND " _
                           & " LOIMOPARTREF = '" & sLots(b, 3) & "' AND " _
                           & " LOIMORUNNO = '" & sLots(b, 4) & "' AND " _
                           & " LOIACTIVITY = '" & sLots(b, 5) & "'"
                           
                  clsADOCon.ExecuteSQL sSql
                  
                  'Update Lot Header
                  sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
                         & "+" & Val(sLots(b, 2)) & " WHERE LOTNUMBER='" & sLots(b, 0) & "'"
                  clsADOCon.ExecuteSQL sSql
               Next
               sSql = "UPDATE PartTable SET PAQOH=PAQOH+" & Abs(cQuantity) _
                      & ",PALOTQTYREMAINING=PALOTQTYREMAINING+" & Abs(cQuantity) _
                      & " WHERE PARTREF='" & sPartsGroup(iRow) & "' "
               clsADOCon.ExecuteSQL sSql
            End If
            
            'Journal
            If iTrans > 0 And clineCost > 0 Then
               'Credit
               iRef = iRef + 1
               If Len(vItems(iRow, 0)) > 20 Then vItems(iRow, 0) = Left(vItems(iRow, 0), 20)
               sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT,DCACCTNO," _
                      & "DCDATE,DCDESCR,DCDESC,DCPARTNO,DCRUNNO) VALUES('" _
                      & sJournalID & "'," _
                      & iTrans & "," _
                      & iRef & "," _
                      & clineCost & ",'" _
                      & sCreditAcct & "','" _
                      & Format(ES_SYSDATE, "mm/dd/yy") & "','" _
                      & "CAPick" & "','" _
                      & sPartsGroup(iRow) & "','" _
                      & sPartNumber & "'," _
                      & Val(cmbRun) & ")"
               clsADOCon.ExecuteSQL sSql
               'Debit
               iRef = iRef + 1
               sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT,DCACCTNO," _
                      & "DCDATE,DCDESCR,DCDESC,DCPARTNO,DCRUNNO) VALUES('" _
                      & sJournalID & "'," _
                      & iTrans & "," _
                      & iRef & "," _
                      & clineCost & ",'" _
                      & sDebitAcct & "','" _
                      & Format(ES_SYSDATE, "mm/dd/yy") & "','" _
                      & "CAPick" & "','" _
                      & sPartsGroup(iRow) & "','" _
                      & sPartNumber & "'," _
                      & Val(cmbRun) & ")"
               clsADOCon.ExecuteSQL sSql
            End If
         Next
         
         ' Changed the Runs status to SC as the msg box says that
         ' the status is changed to SC
         'sSql = "UPDATE RunsTable SET RUNSTATUS='RL',"
         
         sSql = "UPDATE RunsTable SET RUNSTATUS='SC'," _
                & "RUNCMATL=RUNCMATL-" & cCost & "," _
                & "RUNCOST=RUNCOST-" & cCost & " " _
                & "WHERE RUNREF='" & sPartNumber & "' " _
                & "AND RUNNO=" & cmbRun & " "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "DELETE FROM MopkTable WHERE PKMOPART='" & sPartNumber & "' " _
                & "AND PKMORUN=" & cmbRun & ""
         clsADOCon.ExecuteSQL sSql
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            sSql = "UPDATE InvaTable SET INPDATE=INADATE WHERE INTYPE=11 AND " _
                   & "INPDATE IS NULL"
            clsADOCon.ExecuteSQL sSql
            CancelPick = 1
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            CancelPick = 0
         End If
      Else
         'On Error Resume Next
         ' Changed the Runs status to SC as the msg box says that
         ' the status is changed to SC
         'sSql = "UPDATE RunsTable SET RUNSTATUS='RL',"
         sSql = "UPDATE RunsTable SET RUNSTATUS='SC'," _
                & "RUNCMATL=RUNCMATL-" & cCost & "," _
                & "RUNCOST=RUNCOST-" & cCost & " " _
                & "WHERE RUNREF='" & sPartNumber & "' " _
                & "AND RUNNO=" & cmbRun & " "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "DELETE FROM MopkTable WHERE PKMOPART='" & sPartNumber & "' " _
                & "AND PKMORUN=" & cmbRun & ""
         clsADOCon.ExecuteSQL sSql
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            sSql = "UPDATE InvaTable SET INPDATE=INADATE WHERE " _
                   & "INTYPE=11 AND INPDATE IS NULL"
            clsADOCon.ExecuteSQL sSql
            UpdateWipColumns lSysCount
            CancelPick = 1
         Else
            clsADOCon.RollbackTrans
            CancelPick = 0
         End If
         MouseCursor 0
      End If
   Else
      ' Changed the Runs status to SC as the msg box says that
      ' the status is changed to SC
      'sSql = "UPDATE RunsTable SET RUNSTATUS='RL',"
      sSql = "UPDATE RunsTable SET RUNSTATUS='SC'," _
             & "RUNCMATL=RUNCMATL-" & cCost & "," _
             & "RUNCOST=RUNCOST-" & cCost & " " _
             & "WHERE RUNREF='" & sPartNumber & "' " _
             & "AND RUNNO=" & cmbRun & " "
      clsADOCon.ExecuteSQL sSql
      MouseCursor 0
      CancelPick = 1
   End If
   Erase vItems
   Exit Function
   
'DiaErr1:
'   sProcName = "cancelpick"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
whoops:
   ProcessError "CancelPick"
End Function


Private Sub GetStatus()
   Dim RdoStu As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = '" & Compress(cmbPrt) & "' " _
          & "AND RUNNO=" & CInt("0" & cmbRun)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStu, ES_FORWARD)
   If bSqlRows Then
      lblStu = "" & Trim(RdoStu!RUNSTATUS)
   Else
      lblStu = ""
   End If
   Set RdoStu = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getstatus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'4/26/02 Lots = find lots for the Pick Transaction
'10/7/04 Added MAX() Provision to get lots

Private Function GetOpenLots(sLotPart As String, sMoNum As String, lMoRun As Long, sADate As Variant) As Byte
   Dim RdoPlot As ADODB.Recordset
   Dim b As Byte
   Dim iRow As Integer
   Dim iTotalLots As Integer
   Dim sOldLots(50, 2) As String
   
   Erase sLots
   GetOpenLots = 0
   On Error GoTo DiaErr1
   sSql = "SELECT LOINUMBER, MAX(LOIRECORD) AS LOTRECORD," _
          & "LOIACTIVITY FROM LoitTable WHERE (LOIPARTREF='" _
          & sLotPart & "' AND LOIMOPARTREF='" & sMoNum & "' AND " _
          & "LOIMORUNNO=" & lMoRun & " AND LOITYPE=10 AND " _
          & "(LOIMOPKCANCEL IS NULL OR LOIMOPKCANCEL <> 1)) GROUP " _
          & "BY LOINUMBER,LOIACTIVITY"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPlot, ES_FORWARD)
   If bSqlRows Then
      With RdoPlot
         Do Until .EOF
            iTotalLots = iTotalLots + 1
            sOldLots(iTotalLots, 0) = "" & Trim(!LOINUMBER)
            sOldLots(iTotalLots, 1) = Trim$(str$(!LOTRECORD))
            .MoveNext
         Loop
         ClearResultSet RdoPlot
      End With
   End If
   Set RdoPlot = Nothing
   For iRow = 1 To iTotalLots
      sSql = "SELECT LOINUMBER,LOITYPE,LOIPARTREF,LOIADATE,LOIQUANTITY,LOIMOPARTREF," _
             & "LOIMORUNNO,LOIACTIVITY FROM LoitTable WHERE (LOITYPE=10 AND LOINUMBER='" _
             & sOldLots(iRow, 0) & "' AND LOIRECORD=" & sOldLots(iRow, 1) & " AND " _
             & " (LOIMOPKCANCEL IS NULL OR LOIMOPKCANCEL <> 1))"
             
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPlot, ES_FORWARD)
      If bSqlRows Then
         With RdoPlot
            Do Until .EOF
               GetOpenLots = GetOpenLots + 1
               sLots(GetOpenLots, 0) = "" & Trim(!LOINUMBER)
               sLots(GetOpenLots, 1) = "" & Trim(!LOIPARTREF)
               sLots(GetOpenLots, 2) = "" & Trim(str(Abs(!LOIQUANTITY)))
               sLots(GetOpenLots, 3) = "" & Trim(!LOIMOPARTREF)
               sLots(GetOpenLots, 4) = "" & Trim(str(!LOIMORUNNO))
               sLots(GetOpenLots, 5) = "" & Trim(str(!LoiActivity))
               .MoveNext
            Loop
            ClearResultSet RdoPlot
         End With
      End If
   Next
   Set RdoPlot = Nothing
   Exit Function
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetLotUnitCost(ByVal strLoiNum As String, ByVal strPartRef As String, _
                  ByVal strMOPartRef As String, _
                  strMORunNo As String) As Currency

   Dim RdoUnitCost As ADODB.Recordset
   On Error Resume Next
   
   sSql = "SELECT LOTUNITCOST FROM lohdTable WHERE " _
             & " LOTNUMBER = '" & strLoiNum & "'" _
            & " AND LOTPARTREF = '" & strPartRef & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoUnitCost, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoUnitCost!LotUnitCost) Then
         GetLotUnitCost = RdoUnitCost!LotUnitCost
      Else
         GetLotUnitCost = 0
      End If
   End If
   Set RdoUnitCost = Nothing

End Function



