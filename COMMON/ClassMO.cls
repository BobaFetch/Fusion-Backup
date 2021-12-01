VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sPartNo As String
Private SPartRef As String
Private nPartLevel As Byte
Private sProductCode As String
'get wip accounts
Private sWipLabAcct As String, sWipExpAcct As String, sWipMatAcct As String, sWipOhdAcct As String

Private nRunNo As Long
Private bPartHasBeenRead As Boolean
Private bRequireInvoices As Boolean
Private bIgnoreUnpicked As Boolean
Private bIgnoreExpendables As Boolean
Private dtCloseDate As Date
Private dtCompletionDate As Date
Private sStatus As String
Private iLogNumber As Integer
Private bCantClose As Boolean
Private bLogDetails As Boolean
Private bLoggingEnabled As Boolean
Private cRunCost As Currency     'total cost for MO
Private cRunOvHd As Currency     'overhead - from GetLaborCosts()
Private cRunLabor As Currency    'direct labor - from GetLaborCosts()
Private cRunMatl As Currency     'from GetMaterialCosts()
Private cRunExp As Currency      'from GetExpenseCosts()
Private cRunHours As Currency    'from GetLaborCosts()
Private cYield As Currency
Private bDiagnoseOnly As Boolean 'True if MO's are not to be updated
Private dtCompletion As Date
Private sLotNumber As String
Private vTimeFormat As Variant

' vRunOps subscripts
   '0 = Center
   '1 = Shop
   '2 = OPNO
   '3 = QHRS
   '4 = MHRS
   '5 = OPSUHRS
   '6 = OPUNITHRS
Private Const RUNOP_RunHours = 6
   '7 = Sched Complete
Private Const RUNOP_CompleteDate = 7
   '8 = Run Start Date
Private Const RUNOP_StartDate = 8
   '9 = SU Start Date
Private Const RUNOP_SetupStartDate = 9
   '10 = Move Date
Private Const RUNOP_MoveDate = 10
   '11 = Queue Date
Private Const RUNOP_QueueDate = 11
   
Public Property Let PartNumber(PartNo As String)
   If StrComp(Trim(PartNo), sPartNo, vbTextCompare) <> 0 Then
      sPartNo = Trim(PartNo)
      SPartRef = Compress(sPartNo)
      
      'get part parameters
      sSql = "select * from PartTable where PARTREF = '" & SPartRef & "'"
      Dim rdo As ADODB.Recordset
      If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
         nPartLevel = rdo!PALEVEL
         sProductCode = Trim(rdo!PAPRODCODE)
      End If
      Set rdo = Nothing
      
      'get WIP accounts
      sWipLabAcct = GetLaborAcct(SPartRef, sProductCode, nPartLevel)
      sWipExpAcct = GetExpenseAcct(SPartRef, sProductCode, nPartLevel)
      sWipMatAcct = GetMaterialAcct(SPartRef, sProductCode, nPartLevel)
      sWipOhdAcct = GetOverHeadAcct(SPartRef, sProductCode, nPartLevel)
      
   End If
End Property

Public Property Get PartNumber() As String
   PartNumber = sPartNo
End Property

Public Property Get PartRef() As String
   PartNumber = SPartRef
End Property

Public Property Get lotNumber() As String
   lotNumber = sLotNumber
End Property

Public Property Let RunNumber(Runno As Integer)
   If nRunNo <> Runno Then
      nRunNo = Runno
   End If
End Property

Public Property Get RunNumber() As Integer
   RunNumber = nRunNo
End Property

Public Property Let CloseRequiresInvoices(Value As Boolean)
   bRequireInvoices = Value
End Property

Public Property Get CloseRequiresInvoices() As Boolean
   CloseRequiresInvoices = bRequireInvoices
End Property

Public Property Let LoggingEnabled(Value As Boolean)
   bLoggingEnabled = Value
   If Value Then
      TruncateReportLog
   End If
End Property

Public Property Get LoggingEnabled() As Boolean
   LoggingEnabled = bLoggingEnabled
End Property

Public Property Let CloseIgnoresUnpicked(Value As Boolean)
   bIgnoreUnpicked = Value
End Property

Public Property Get CloseIgnoresUnpicked() As Boolean
   CloseIgnoresUnpicked = bIgnoreUnpicked
End Property

Public Property Let CloseIgnoresExpendables(Value As Boolean)
   bIgnoreExpendables = Value
End Property

Public Property Get CloseIgnoresExpendables() As Boolean
   CloseIgnoresExpendables = bIgnoreExpendables
End Property

Public Property Let DiagnoseOnly(Value As Boolean)
   bDiagnoseOnly = Value
End Property

Public Property Get DiagnoseOnly() As Boolean
   DiagnoseOnly = bDiagnoseOnly
End Property

Public Property Let CloseDate(Value As Date)
   dtCloseDate = Value
End Property

Public Property Get CloseDate() As Date
   CloseDate = dtCloseDate
End Property

Property Get RUNCOST() As Currency
   RUNCOST = cRunCost   'total cost for MO
End Property

Property Get RUNOVHD() As Currency
   RUNOVHD = cRunOvHd   'overhead - from GetLaborCosts()
End Property

Property Get RunLabor() As Currency
   RunLabor = cRunLabor       'direct labor - from GetLaborCosts()
End Property

Property Get RUNMATL() As Currency
    RUNMATL = cRunMatl 'from GetMaterialCosts()
End Property

Property Get RUNEXP() As Currency
   RUNEXP = cRunExp   'from GetExpenseCosts()
End Property

Property Get RunHours() As Currency
   RunHours = cRunHours 'from GetLaborCosts()
End Property

Property Get Yield() As Currency
   Yield = cYield
End Property


Private Sub Class_Initialize()
   sPartNo = ""
   SPartRef = ""
   nRunNo = 0
   bPartHasBeenRead = False
   bRequireInvoices = True
   bIgnoreUnpicked = False
   bIgnoreExpendables = False
   dtCloseDate = Date
   bLogDetails = False
   'TruncateReportLog
End Sub

'Public Function CanMoBeClosed() As Boolean
'
'   Dim rdo As ADODB.Recordset
'
'   sSql = "select count(*) from viewCompletedMOs " & vbCrLf _
'      & "where RUNREF='" & sPartRef & "' and RUNNO=" & nRunNo & vbCrLf _
'      & "and OpenPoItems=0 and UncostedLots=0" & vbCrLf
'
'   If RequireInvoices Then
'      sSql = sSql & "and RecPoItems = 0" & vbCrLf
'   End If
'
'   If Not IgnoreUnpicked Then
'      If IgnoreExpendables Then
'         sSql = sSql & "and OpenNonType5PickItems = 0" & vbCrLf
'      Else
'         sSql = sSql & "and OpenPickItems = 0" & vbCrLf
'      End If
'   End If
'
'   bsqlrows = clsadocon.getdataset(ssql, rdo, ES_FORWARD)
'   If bSqlRows Then
'      If CLng(rdo.Fields(0).Value) = 0 Then
'         CanMoBeClosed = True
'      Else
'         CanMoBeClosed = False
'      End If
'   End If
'
'End Function

Public Function CloseMO() As Boolean
   'return = True if successful
   bCantClose = False
   ReadProperties
   'iLogNumber = 0
   Log ""
   Log "Close MO " & SPartRef & " Run " & nRunNo & " completed " & dtCompletion
   If sStatus <> "CO" Then
      Log "* Can't close.  MO status = " & sStatus
      Exit Function
   End If
   
   If DateDiff("d", dtCompletionDate, dtCloseDate) < 0 Then
      Log "* Close date (" & dtCloseDate & ") must be >= the completion date (" & dtCompletionDate & ")"
      GoTo CantClose
   End If
   
   'show any completions and cancellations (if logging)
   ShowCompletions
   
   'determine whether it is OK to close
   GetUnInvoicedPoItems
   GetUnpickedItems
   GetUncostedLots
   If bDiagnoseOnly Then
      Log ""
      Exit Function
   End If
   
   If bCantClose Then
      Log ""
      GoTo CantClose
   Else
      GetExpenseCosts
      GetLaborCosts
      GetMaterialCosts
      If CloseMoFinalize Then
         'succeeded
         CloseMO = True
      End If
   End If
   
   Log ""
   Exit Function
   
CantClose:
   CloseMO = False
   'WriteReportLog
   

End Function

Public Function ViewCompletedMosWhereClause() As String
   
   Dim sql As String
   
   sql = "where OpenPoItems=0 and UncostedLots=0" & vbCrLf
   
   If bRequireInvoices Then
      sql = sql & "and RecPoItems = 0" & vbCrLf
   End If
   
   If Not bIgnoreUnpicked Then
      If bIgnoreExpendables Then
         sql = sql & "and OpenNonType5PickItems = 0" & vbCrLf
      Else
         sql = sql & "and OpenPickItems = 0" & vbCrLf
      End If
   End If
   
   ViewCompletedMosWhereClause = sql
End Function

Private Sub ReadProperties()
   Dim rdo As ADODB.Recordset
   sSql = "select * from RunsTable where RUNREF='" & SPartRef & "' and RUNNO=" & nRunNo
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      dtCompletionDate = rdo!RUNCOMPLETE
      sStatus = rdo!RUNSTATUS
      cYield = rdo!RUNYIELD
      dtCompletion = rdo!RUNCOMPLETE
   End If
   Set rdo = Nothing
End Sub

Private Sub TruncateReportLog()
   sSql = "TRUNCATE TABLE EsReportClosedRunsLog"
   clsADOCon.ExecuteSql sSql
   iLogNumber = 0
End Sub

Public Sub Log(sText As String)

   DoEvents
   
   iLogNumber = iLogNumber + 1
   sText = Replace(sText, "'", "''")
   If Len(sText) > 80 Then sText = Left$(sText, 80)
   'Debug.Print iLogNumber & ": " & sText
   If LoggingEnabled Then
      sSql = "INSERT INTO EsReportClosedRunsLog (LOG_NUMBER,LOG_TEXT, LOG_PARTNO, LOG_RUNNO, LOG_CLOSED) " _
             & "VALUES( " & iLogNumber & ", '" & sText & "', '" & sPartNo & "', " & nRunNo & ",0)"
      clsADOCon.ExecuteSql sSql
   End If
End Sub

'Private Sub WriteReportLog()
'   Dim iList As Integer
'   'sSql = "TRUNCATE TABLE EsReportClosedRunsLog"
'   'clsAdoCon.ExecuteSQL sSql
'   For iList = 1 To iLogNumber
'      sSql = "INSERT INTO EsReportClosedRunsLog (LOG_NUMBER,LOG_TEXT) " _
'             & "VALUES( (select isnull(max(LOG_NUMBER),0) + 1 from EsReportClosedRunsLog)," & vbCrLf _
'             & "'" & sLogNote(iList) & "')"
'      clsAdoCon.ExecuteSQL sSql
'   Next
'
'End Sub

Private Sub GetUnInvoicedPoItems()
   'log any unreceived/uninvoiced PO items that prevent closing
   
   Dim RdoInv As ADODB.Recordset
   
   'if invoices required, list unreceived items.
   'if invoices required, list uninvoiced items too.
   sSql = "select case when PITYPE=14 then 'Open' else 'Recd' end as Status," & vbCrLf _
      & "PINUMBER, PIITEM, PIREV, PIPART from PoitTable" & vbCrLf _
      & "where PIRUNPART='" & sPartNo & "' and PIRUNNO=" & nRunNo & vbCrLf
   If bRequireInvoices Then
      sSql = sSql & "and (PITYPE=" & IATYPE_PoOpenItem & " or PITYPE=" & IATYPE_PoReceipt & ")" & vbCrLf
   Else
      sSql = sSql & "and PITYPE=" & IATYPE_PoOpenItem & vbCrLf
   End If
   sSql = sSql & "order by PIPART"
         
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   If bSqlRows Then
      bCantClose = True
      If bRequireInvoices Then
         Log "Unreceived or uninvoiced purchase orders allocations:"
      Else
         Log "Unreceived purchase order allocations:"
      End If
   
      With RdoInv
         Do Until .EOF
            Log "PO " & Format(!PINUMBER, "00000") & " item " & !PIITEM & !PIREV _
                     & " " & !PIPART & " " & !Status
            .MoveNext
         Loop
         ClearResultSet RdoInv
      End With
   Else
      If bRequireInvoices Then
         Log "All items where invoiced"
      Else
         Log "All items were received or invoiced"
      End If
   End If
   Set RdoInv = Nothing
   
End Sub

Private Sub GetUnpickedItems()
   'log any unpicked items that prevent closing
   Dim RdoPck As ADODB.Recordset
   
   If bIgnoreUnpicked Then
      Log "Ignoring unpicked items"
      Exit Sub
   End If
   
   sSql = "select INPART, INPQTY " _
          & "from InvaTable" & vbCrLf _
          & "join PartTable on PARTREF=INPART" & vbCrLf _
          & "where (INTYPE=" & CStr(IATYPE_PickOpenItem) & " or INTYPE=" & CStr(IATYPE_PickSubstitution) & ")" & vbCrLf _
          & "and INMOPART='" & SPartRef & "' AND INMORUN=" & nRunNo & vbCrLf
   
   If bIgnoreExpendables Then
      sSql = sSql & "and PALEVEL<>5" & vbCrLf
   End If
   
   sSql = sSql & "order by PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPck, ES_FORWARD)
   If bSqlRows Then
      bCantClose = True
      If bIgnoreExpendables Then
         Log "Unpicked items (ignoring expandables):"
      Else
         Log "Unpicked items:"
      End If
      
      With RdoPck
         Do Until .EOF
            Log "* " & !INPART & " " & "Qty: " & !INPQTY & " not picked"
            .MoveNext
         Loop
         ClearResultSet RdoPck
      End With
   Else
      If bIgnoreExpendables Then
         Log "There were no unpicked items (ignoring expandables)"
      Else
         Log "There were no unpicked items"
      End If
   End If
   Set RdoPck = Nothing
End Sub


Public Sub GetExpenseCosts()
   'place expenses for MO in RunExp property
   
   Dim RdoExp As ADODB.Recordset
   Dim cMOEXPENSE As Currency
   Dim cFREIGHT As Currency
   Dim cTAXES As Currency
   Dim cNOPO As Currency
   sProcName = "GetExpenseCosts"
   
   cRunExp = 0
   
   'Purchased Expense Items
   'Log "Invoiced Purchase Order Expenses (Services):"
   
'   sSql = "select isnull(sum( PIAQTY * PIAMT ), 0) as ServiceCost" & vbCrLf _
'      & "from PoitTable join PartTable on PIPART = PARTREF" & vbCrLf _
'      & "where PIRUNPART='" & SPartRef & "' and PIRUNNO=" & nRunNo & vbCrLf _
'      & "and PITYPE=17 AND (PALEVEL=6 or PALEVEL=7)"
'   If GetDataSet(RdoExp, ES_FORWARD) Then
'      cMOEXPENSE = RdoExp!ServiceCost
'      If cMOEXPENSE > 0 Then
'         Debug.Print cMOEXPENSE
'      End If
'   End If
   
   sSql = "select isnull( sum" & vbCrLf _
      & "(" & vbCrLf _
      & "case PAUSEACTUALCOST " & vbCrLf _
      & " when 0 then" & vbCrLf _
      & "    case when OpSvcCost <> 0 then" & vbCrLf _
      & "       OpSvcCost" & vbCrLf _
      & "    else" & vbCrLf _
      & "       PoSvcCost" & vbCrLf _
      & "    end " & vbCrLf _
      & " else" & vbCrLf _
      & "    PoSvcCost" & vbCrLf _
      & "end" & vbCrLf _
      & ") , 0 ) as ServiceCost" & vbCrLf _
      & "from viewExpensesByMoDetails" & vbCrLf _
      & "where OPREF = '" & SPartRef & "' and OPRUN = " & nRunNo & vbCrLf
   
   'Debug.Print sSql
   If clsADOCon.GetDataSet(sSql, RdoExp, ES_FORWARD) Then
      cMOEXPENSE = CCur(RdoExp.Fields(0))
      If cMOEXPENSE > 0 Then
         Debug.Print cMOEXPENSE
      End If
   End If
   
   Log "Invoiced Purchase Order Expenses (Services): " & cMOEXPENSE
   
   'freight and tax are included in inventory costs
   cFREIGHT = 0
   cTAXES = 0
   
   'Invoices without PO's
   sSql = "select isnull(SUM( VITQTY * VITCOST ), 0) as NoPoCost" & vbCrLf _
      & "from ViitTable" & vbCrLf _
      & "where VITPO=0 AND VITPOITEM=0" & vbCrLf _
      & "and VITMO='" & SPartRef & "' and VITMORUN=" & nRunNo
   Debug.Print sSql
   If clsADOCon.GetDataSet(sSql, RdoExp, ES_FORWARD) Then
      cNOPO = RdoExp!NoPoCost
   End If
   Log "Invoiced expenses without a purchase order: " & cNOPO
   
   cRunExp = cMOEXPENSE + cFREIGHT + cTAXES + cNOPO
   Log "Total expenses: " & cRunExp
   Set RdoExp = Nothing
End Sub

Public Sub GetLaborCosts()
   'place labor for MO in RunLabor, RunOvHd, and RunHours  properties
   Dim RdoLab As ADODB.Recordset
   cRunHours = 0
   cRunOvHd = 0
   cRunLabor = 0
   
   sProcName = "GetLaborCosts"
   
   ' (TCHOURS * TCOHRATE * TCRATE) / 100 - Changed the calculation
   ' AS TCOHRATE is already calculated.
   'get total hours, OH (% + fixed), and labor
   
   
   Dim RdoOVH As ADODB.Recordset
   Dim bShareOVH As Integer
   
   bShareOVH = 0
   
   sSql = "select isnull(CONOTCSHAREOVH, 0) as CONOTCSHAREOVH " & vbCrLf _
      & " FROM ComnTable"
      
   Debug.Print sSql
   If clsADOCon.GetDataSet(sSql, RdoOVH, ES_FORWARD) Then
      bShareOVH = RdoOVH!CONOTCSHAREOVH
   End If
   Set RdoOVH = Nothing
   
   If (bShareOVH = 1) Then
      sSql = "select isnull(sum(TCHOURS),0) as TotalHours," & vbCrLf _
         & "isnull(sum(((DATEDIFF( minute, TCSTARTTIME, TCSTOPTIME) /60) * TCOHRATE ) + " & vbCrLf _
         & "((DATEDIFF( minute, TCSTARTTIME, TCSTOPTIME) /60) * TCOHFIXED)),0) as TotalOverhead," & vbCrLf _
         & "isnull(sum(TCRATE * TCHOURS),0) as TotalLabor" & vbCrLf _
         & "from TcitTable join TchdTable on TCCARD=TMCARD" & vbCrLf _
         & "where TCPARTREF='" & SPartRef & "' AND TCRUNNO=" & nRunNo
   Else
      sSql = "select isnull(sum(TCHOURS),0) as TotalHours," & vbCrLf _
         & "isnull(sum((TCHOURS * TCOHRATE ) + (TCHOURS * TCOHFIXED)),0) as TotalOverhead," & vbCrLf _
         & "isnull(sum(TCRATE * TCHOURS),0) as TotalLabor" & vbCrLf _
         & "from TcitTable join TchdTable on TCCARD=TMCARD" & vbCrLf _
         & "where TCPARTREF='" & SPartRef & "' AND TCRUNNO=" & nRunNo
   End If
   Set RdoOVH = Nothing
   
   
   
   Debug.Print sSql
   
   If clsADOCon.GetDataSet(sSql, RdoLab, ES_FORWARD) Then
      cRunHours = RdoLab!TotalHours
      cRunOvHd = RdoLab!TotalOverhead
      cRunLabor = RdoLab!TotalLabor
      Log "Hours: " & Format$(cRunHours, "#####0.00") _
         & " Labor: " & Format$(cRunLabor, "#####0.00") _
         & " Overhead: " & Format$(cRunOvHd, "#####0.00")
   Else
      Log "No Labor Costs Recorded"
   End If
   Set RdoLab = Nothing
End Sub

Public Sub GetMaterialCosts()
   'place material costs for MO in RunMatl property
   Dim RdoMat As ADODB.Recordset
   Dim cLotCosts As Currency
   Dim cNonLotCosts As Currency
   Dim cReturns As Currency
   
   sProcName = "GetMaterialCosts"
   
   cLotCosts = 0
   cNonLotCosts = 0
   cReturns = 0
   
   'get costs for lot-costed items
   sSql = "select TotalCost" & vbCrLf _
      & "from viewLotCostsByMoSummary" & vbCrLf _
      & "where MoPart='" & SPartRef & "' and MoRun=" & nRunNo
               
   Debug.Print sSql
   If clsADOCon.GetDataSet(sSql, RdoMat, ES_FORWARD) Then
      cLotCosts = RdoMat!totalCost
      Log "Total of lot-costed parts: " & cLotCosts
   Else
      Log "No lot-costed parts."
   End If
   
   'get standard costs for other parts
   sSql = "select TotalCost" & vbCrLf _
      & "from viewNonLotCostsByMoSummary" & vbCrLf _
      & "where MoPart='" & SPartRef & "' and MoRun=" & nRunNo
               
   Debug.Print sSql
   If clsADOCon.GetDataSet(sSql, RdoMat, ES_FORWARD) Then
      cNonLotCosts = RdoMat!totalCost
      Log "Total of non-lot-costed parts: " & cNonLotCosts
   Else
      Log "No non-lot-costed parts."
   End If
   
   
   'Pick return costs
   sSql = "select cast(isnull(sum(INAQTY*INAMT), 0) as decimal(15,4)) as TotalCost " & vbCrLf _
      & "from InvaTable" & vbCrLf _
      & "join PartTable on INPART = PARTREF" & vbCrLf _
      & "where INTYPE between 11 and 21" & vbCrLf _
      & "and INMOPART='" & SPartRef & "' and INMORUN=" & nRunNo & vbCrLf
   If bIgnoreExpendables Then
      sSql = sSql & "and PALEVEL<>5" & vbCrLf
   End If
            
   Debug.Print sSql
   If clsADOCon.GetDataSet(sSql, RdoMat, ES_FORWARD) Then
      cReturns = RdoMat!totalCost
   End If
   If bIgnoreExpendables Then
      Log "Picks returned to inventory (excluding expendables): " & cReturns
   Else
      Log "Picks returned to inventory (including expendables): " & cReturns
   End If
   
   cRunMatl = cLotCosts + cNonLotCosts - cReturns
   Log "Total Material: " & cRunMatl
   Set RdoMat = Nothing
End Sub


Private Sub GetUncostedLots()
   Dim RdoMat As ADODB.Recordset
   'select *, Part, LotUserID, Quantity, UnitCost

   sSql = "select Part, LotUserID, Quantity, UnitCost, LOTDATECOSTED" & vbCrLf _
      & "from viewLotCostsByMoDetails" & vbCrLf _
      & "where MoPart='" & SPartRef & "' and MoRun=" & nRunNo & vbCrLf _
      & "and TotalCost = 0"
'Debug.Print sSql

   Debug.Print sSql
   Dim uncostedLots As Long
   If clsADOCon.GetDataSet(sSql, RdoMat, ES_FORWARD) Then
      With RdoMat
         Do Until .EOF
            'allow zero cost if lot has been costed
            If IsNull(!LOTDATECOSTED) Then
               bCantClose = True
               Log "* " & !part & " uncosted lot " & Trim(!LotUserId) _
                  & " qty=" & !Quantity
               uncostedLots = uncostedLots + 1
            ElseIf DateDiff("d", !LOTDATECOSTED, dtCloseDate) < 0 Then
               bCantClose = True
               Log "* " & !part & " lot " & Trim(!LotUserId) _
                  & " costed " & Format(!LOTDATECOSTED, "MM/DD/YY") _
                  & " > close date"
               uncostedLots = uncostedLots + 1
            Else
               Log !part & " from zero-cost lot " & Trim(!LotUserId) _
                  & " qty=" & !Quantity & " allowed."
            End If
            .MoveNext
         Loop
         ClearResultSet RdoMat
      End With
   End If
   If uncostedLots = 0 Then
      Log "There were no uncosted lots"
   End If
   Set RdoMat = Nothing
End Sub

Public Function AreTherePicksForUnclosedMos() As Boolean
   'return = false if no unclosed subsidiary MO's
   '     = true otherwise (see log)
   Dim RdoMat As ADODB.Recordset
   Dim unclosedMos As Boolean

   sSql = "select rtrim(RUNREF) AS RUNREF, RUNNO, RUNSTATUS," & vbCrLf _
      & "rtrim(INMOPART) as MoPart, INMORUN as MoRun, lohd.LOTNUMBER" & vbCrLf _
      & "from InvaTable" & vbCrLf _
      & "join PartTable on PARTREF=INPART" & vbCrLf _
      & "join LoitTable loit on loit.LOIACTIVITY = INNUMBER" & vbCrLf _
      & "join LohdTable lohd on lohd.LotNumber = loit.LOINUMBER" & vbCrLf _
      & "join RunsTable on RUNREF = LOTMOPARTREF and RUNNO = LOTMORUNNO" & vbCrLf _
      & "Where INTYPE = 10 And PALEVEL < 5 AND RUNSTATUS <> 'CL'" & vbCrLf _
      & "and INMOPART = '" & SPartRef & "' and INMORUN = " & RunNumber
   
   If clsADOCon.GetDataSet(sSql, RdoMat, ES_FORWARD) Then
      With RdoMat
         Do Until .EOF
            Log ""
            Log "Close MO " & SPartRef & " Run " & nRunNo & " completed " & dtCompletion
            Log "* Pick for unclosed MO " & !RUNREF & " # " & !Runno & " status " & !RUNSTATUS
            AreTherePicksForUnclosedMos = True
            .MoveNext
         Loop
         ClearResultSet RdoMat
      End With
   End If
   If Not AreTherePicksForUnclosedMos Then
      'Log "There were no picks for unclosed MOs"
   End If
   Set RdoMat = Nothing
End Function






Private Function CloseMoFinalize() As Boolean
   'returns true if successful
   
   'Dim RdoInv As ADODB.Recordset
   cRunCost = cRunOvHd + cRunMatl + cRunExp + cRunLabor
   'If cYield = 0 Then cYield = 1      WHY?
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "update RunsTable" & vbCrLf _
      & "set RUNSTATUS='CL'," _
      & "RUNCOST=" & cRunCost & "," _
      & "RUNOHCOST=" & cRunOvHd & "," _
      & "RUNCMATL=" & cRunMatl & "," _
      & "RUNCEXP=" & cRunExp & "," _
      & "RUNCHRS=" & cRunHours & "," _
      & "RUNCLAB=" & cRunLabor & "," _
      & "RUNCLOSED='" & dtCloseDate & "'," _
      & "RUNREVBY='" & sInitials & "' " & vbCrLf _
      & "WHERE RUNREF='" & SPartRef & "' and RUNNO=" & nRunNo & ""
   clsADOCon.ExecuteSql sSql
      
   '@@@ don't commit if this fails.
   '@@@ track down timeout
   Dim returnMessage As String
   returnMessage = UpdateMOCosts(SPartRef, nRunNo, "CLOSED RUN", _
            cRunMatl, cRunLabor, cRunExp, cRunOvHd, cRunHours)
   Log returnMessage
   If InStr(1, returnMessage, "failed") = 0 Then
      clsADOCon.CommitTrans
      CloseMoFinalize = True
      Log "MO " & SPartRef & " Run " & nRunNo & " Was Closed.  Total cost: " & cRunCost
   Else
      clsADOCon.RollbackTrans
      CloseMoFinalize = False
      Log "MO " & SPartRef & " Run " & nRunNo & " Closure failed"
   End If
   
End Function


Public Function UpdateMOCosts(PartNumber As String, _
   RunNumber As Long, _
   Description As String, _
   ByVal ActualTotMaterialCost As Currency, _
   ByVal ActualTotLaborCost As Currency, _
   ByVal ActualTotExpenseCost As Currency, _
   ByVal ActualTotOverheadCost As Currency, _
   ByVal ActualTotHours As Currency, _
   Optional UpdateIaEvenIfJournalClosed As Boolean) As String
   
   'determine whether to use standard or actual costs and update corresponding
   'inventory record.
   'Description = "*" to retain current description
   'return =   "Used standard costs"
   '           or "Used actual costs"
   '           or "ClassMO.UpdateMOCosts failed"
   
   'note: there could be multiple invatable records here
   'if lot tracking was turned on after one was created, we might or might not
   'get the INLOTTRACK and INUSEACTUAL reflecting this, so Max() is used
   Dim Quantity As Currency
   Dim ActualUnitCost As Currency
   Dim iaCount As Long
   On Error GoTo whoops
   UpdateMOCosts = "Used actual costs"
   
   'get part parameters
   sSql = "SELECT PALOTTRACK, PAUSEACTUALCOST, PASTDCOST, PATOTMATL," & vbCrLf _
      & "PATOTLABOR, PATOTEXP, PATOTOH, PATOTHRS" & vbCrLf _
      & "from PartTable where PARTREF = '" & PartNumber & "' "
   
   Dim rdo As ADODB.Recordset
   Dim bLotTracked As Boolean, bUseActualCost As Boolean, usingStandardCost As Boolean
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
      
         Dim RdoInv As ADODB.Recordset
         sSql = "SELECT COALESCE(max(cast(INLOTTRACK as tinyint))," & !PALOTTRACK & ", 0) as LOTTRACK," & vbCrLf _
            & "COALESCE(max(cast(INUSEACTUALCOST as tinyint))," & !PAUSEACTUALCOST & ", 0) as USEACTUALCOST, " & vbCrLf _
            & "isnull(sum(INAQTY),0) as TOTALQTY, count(*) as CT" & vbCrLf _
            & "from InvaTable" & vbCrLf _
            & "WHERE (INTYPE=" & IATYPE_MoCompletion & vbCrLf _
            & " OR INTYPE=" & IATYPE_CanceledMoCompletion & ") " & vbCrLf _
            & "AND INPART='" & PartNumber & "' " _
            & "AND INMORUN=" & RunNumber
            
   Debug.Print sSql
         If clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD) Then
            bLotTracked = RdoInv!LOTTRACK
            bUseActualCost = RdoInv!useActualCost
            Quantity = RdoInv!TOTALQTY
            iaCount = RdoInv!ct
            
         
         'no completion quantity so no ia records
         Else
            bLotTracked = !PALOTTRACK
            bUseActualCost = !PAUSEACTUALCOST
         End If
         'RdoInv.Close
         Set RdoInv = Nothing
         
         Log "Total quantity: " & Quantity & " in " & iaCount & " completion(s)"
         
         'get quantity = sum of inaqty records for partial or full completions
         If Quantity = 0 Then
            ActualUnitCost = 0
         Else
            ActualUnitCost = (ActualTotMaterialCost + ActualTotLaborCost _
               + ActualTotExpenseCost + ActualTotOverheadCost) / Quantity
         End If
         
         'if using standard cost
         'If Not bLotTracked Or (bLotTracked And Not bUseActualCost) Then
         If (Not bLotTracked And Not bUseActualCost) Then
            usingStandardCost = True
            ActualUnitCost = !PASTDCOST
            ActualTotMaterialCost = !PATOTMATL * Quantity
            ActualTotLaborCost = !PATOTLABOR * Quantity
            ActualTotExpenseCost = !PATOTEXP * Quantity
            ActualTotOverheadCost = !PATOTOH * Quantity
            ActualTotHours = !PATOTHRS * Quantity
            UpdateMOCosts = "Used standard costs"
         Else
            usingStandardCost = False
         End If
      End With
   End If
   'rdo.Close
   Set rdo = Nothing
   
   'calculate unit component costs
   Dim ActualUnitMaterialCost As Currency, ActualUnitLaborCost As Currency, ActualUnitExpenseCost As Currency
   Dim ActualUnitOverheadCost As Currency, ActualUnitHours As Currency
   If Quantity > 0 Then
      ActualUnitMaterialCost = ActualTotMaterialCost / Quantity
      ActualUnitLaborCost = ActualTotLaborCost / Quantity
      ActualUnitExpenseCost = ActualTotExpenseCost / Quantity
      ActualUnitOverheadCost = ActualTotOverheadCost / Quantity
      ActualUnitHours = ActualTotHours / Quantity
   End If
   
'   'update inventory activity record(s)
'   sSql = "UPDATE InvaTable SET " & vbCrLf _
'      & "INREF1='" & Description & "',INAMT=" & ActualUnitCost & ", " & vbCrLf
      
   'update inventory activity record(s)
   sSql = "UPDATE InvaTable SET " & vbCrLf
   If Description <> "*" Then
      sSql = sSql & "INREF1='" & Description & "'," & vbCrLf
   End If
   sSql = sSql & "INAMT=" & ActualUnitCost & ", " & vbCrLf
      
   'if total quantity > 0 or using standard cost
   If Quantity > 0 Or usingStandardCost Then
      sSql = sSql & "INTOTMATL=" & ActualUnitMaterialCost & " * INAQTY," & vbCrLf _
         & "INTOTLABOR=" & ActualUnitLaborCost & " * INAQTY, " & vbCrLf _
         & "INTOTEXP=" & ActualUnitExpenseCost & " * INAQTY, " & vbCrLf _
         & "INTOTOH=" & ActualUnitOverheadCost & " * INAQTY," & vbCrLf _
         & "INTOTHRS=" & ActualUnitHours & " * INAQTY, " & vbCrLf
   
   'if total quantity = 0 and using actual cost, insert real actual costs even though quantity = 0
   Else
      sSql = sSql & "INTOTMATL=" & ActualTotMaterialCost & "," & vbCrLf _
         & "INTOTLABOR=" & ActualTotLaborCost & ", " & vbCrLf _
         & "INTOTEXP=" & ActualTotExpenseCost & ", " & vbCrLf _
         & "INTOTOH=" & ActualTotOverheadCost & "," & vbCrLf _
         & "INTOTHRS=" & ActualTotHours & ", " & vbCrLf
   End If
   sSql = sSql & "INLOTTRACK=" & bLotTracked * -1 & ", " _
      & "INUSEACTUALCOST=" & bUseActualCost * -1 & "," & vbCrLf _
      & "INCOSTEDBY='" & sInitials & "'," & vbCrLf _
      & "INMAINTCOSTED=1" & vbCrLf _
      & "WHERE INTYPE=" & IATYPE_MoCompletion & " AND INPART='" & PartNumber & "' " _
      & "AND INMORUN=" & RunNumber & vbCrLf
   If Not UpdateIaEvenIfJournalClosed Then
      sSql = sSql & "AND INGLPOSTED = 0"
   End If
      
   Debug.Print sSql
   clsADOCon.ExecuteSql sSql
   
   'update corresponding lot record(s)
   Dim vAdate As Variant
   vAdate = GetServerDateTime
   sSql = "Update LohdTable" & vbCrLf _
      & "set LOTDATECOSTED='" & vAdate & "'," _
      & "LOTUNITCOST=" & ActualUnitCost & "," & vbCrLf _
      & "LOTMAINTCOSTED=1," & vbCrLf
      
   'if total quantity > 0 or using standard cost
   If Quantity > 0 Or usingStandardCost Then
      sSql = sSql & "LOTTOTMATL=" & ActualUnitMaterialCost & " * LOTORIGINALQTY," & vbCrLf _
      & "LOTTOTLABOR=" & ActualUnitLaborCost & " * LOTORIGINALQTY," & vbCrLf _
      & "LOTTOTEXP=" & ActualUnitExpenseCost & " * LOTORIGINALQTY," & vbCrLf _
      & "LOTTOTOH=" & ActualUnitOverheadCost & " * LOTORIGINALQTY," & vbCrLf _
      & "LOTTOTHRS=" & ActualUnitHours & " * LOTORIGINALQTY" & vbCrLf
   
   'if total quantity = 0 and using actual cost, insert real actual costs even though quantity = 0
   Else
      sSql = sSql & "LOTTOTMATL=" & ActualTotMaterialCost & "," & vbCrLf _
      & "LOTTOTLABOR=" & ActualTotLaborCost & "," & vbCrLf _
      & "LOTTOTEXP=" & ActualTotExpenseCost & "," & vbCrLf _
      & "LOTTOTOH=" & ActualTotOverheadCost & "," & vbCrLf _
      & "LOTTOTHRS=" & ActualTotHours & vbCrLf
   End If
   
   sSql = sSql & "WHERE (LOTMOPARTREF='" & PartNumber & "' " _
      & "AND LOTMORUNNO=" & RunNumber & ")"
   Debug.Print sSql
   clsADOCon.ExecuteSql sSql
   
   'now recost any pick items that are being drawn fom the lots
   sSql = "update InvaTable set " & vbCrLf _
      & "INAMT = LOTUNITCOST, " & vbCrLf _
      & "INTOTMATL = - LOTTOTMATL * INAQTY / LOTORIGINALQTY," & vbCrLf _
      & "INTOTLABOR = - LOTTOTLABOR * INAQTY / LOTORIGINALQTY, " & vbCrLf _
      & "INTOTEXP = - LOTTOTEXP * INAQTY / LOTORIGINALQTY, " & vbCrLf _
      & "INTOTOH = - LOTTOTOH * INAQTY / LOTORIGINALQTY," & vbCrLf _
      & "INTOTHRS = - LOTTOTHRS * INAQTY / LOTORIGINALQTY, " & vbCrLf _
      & "INLOTTRACK = " & IIf(bLotTracked, 1, 0) & ", " & vbCrLf _
      & "INUSEACTUALCOST = " & IIf(bUseActualCost, 1, 0) & "," & vbCrLf _
      & "INCOSTEDBY = '" & sInitials & "'," & vbCrLf _
      & "INMAINTCOSTED = 1" & vbCrLf _
      & "from InvaTable ia" & vbCrLf _
      & "join LoitTable on LOIACTIVITY = INNUMBER" & vbCrLf _
      & "join lohdtable on LOINUMBER = LOTNUMBER" & vbCrLf _
      & "WHERE LOTMOPARTREF='" & PartNumber & "' AND LOTMORUNNO=" & RunNumber & vbCrLf _
      & "and ia.INAQTY < 0 and LOTORIGINALQTY > 0" & vbCrLf
   If Not UpdateIaEvenIfJournalClosed Then
      sSql = sSql & "AND INGLPOSTED = 0"
   End If
   Debug.Print sSql
   clsADOCon.ExecuteSql sSql
   
   'make sure run yield quantity is correct.
   sSql = "update RunsTable" & vbCrLf _
      & "set RUNYIELD=" & Quantity & "," & vbCrLf _
      & "RUNMAINTCOSTED=1" & vbCrLf _
      & "WHERE RUNREF='" & PartNumber & "' and RUNNO=" & RunNumber & ""
   Debug.Print sSql
   clsADOCon.ExecuteSql sSql
   
   Exit Function
   
whoops:
   sProcName = "UpdateMOCosts"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   UpdateMOCosts = "ClassMO.UpdateMOCosts failed"
   MsgBox UpdateMOCosts & " " & Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
End Function

Public Sub GetRunCosts()
   'calculate total run costs.  Make available in
   'RunHours
   'RunOvHd
   'RunLabor
   'RunMatl
   'RunExp
   'RunCost
   
   GetLaborCosts
   GetMaterialCosts
   GetExpenseCosts
   cRunCost = cRunOvHd + cRunMatl + cRunExp + cRunLabor
   Log "Total MO cost: " & cRunCost
End Sub

Private Sub ShowCompletions()

   If Not LoggingEnabled Then Exit Sub
   
   Dim rdo As ADODB.Recordset
   Dim ia As New ClassInventoryActivity
   
   sSql = "select * " _
          & "from InvaTable" & vbCrLf _
          & "where INTYPE in ( " & CStr(IATYPE_MoCompletion) & ", " & IATYPE_CanceledMoCompletion & ")" & vbCrLf _
          & "and INMOPART='" & SPartRef & "' AND INMORUN=" & nRunNo & vbCrLf
   
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         Do Until .EOF
            Log ia.IATypeName(!INTYPE) & " " & Format(!INADATE, "mm/dd/yy") & " quantity: " & !INAQTY & " lot " & !INLOTNUMBER
            .MoveNext
         Loop
      End With
   End If
   Set rdo = Nothing
End Sub

Public Function RecalculateCosts() As Boolean
   GetRunCosts
   Dim sReturnMsg As String
   If Not DiagnoseOnly Then
      sReturnMsg = UpdateMOCosts(PartNumber, RunNumber, "*", _
          RUNMATL, RunLabor, RUNEXP, RUNOVHD, RunHours)
   End If
   RecalculateCosts = True
End Function

Public Function RecalCostsIgnoreGLPost() As Boolean
   GetRunCosts
   Dim sReturnMsg As String
   If Not DiagnoseOnly Then
      sReturnMsg = UpdateMOCosts(PartNumber, RunNumber, "*", _
          RUNMATL, RunLabor, RUNEXP, RUNOVHD, RunHours, True)
   End If
   RecalCostsIgnoreGLPost = True
End Function

Public Sub FillComboBoxWithMoParts(cbo As ComboBox, whereClause As String)
   'Fill combo box with part numbers for runs satisfying where clause
   'Example FillComboBoxWithMoParts cbo, "WHERE RUNSTATUS IN ( 'CO', 'CL' )"
   
   On Error GoTo DiaErr1
   cbo.Clear
   sSql = "SELECT DISTINCT PARTNUM" & vbCrLf _
         & "FROM RunsTable" & vbCrLf _
         & "JOIN PartTable ON PARTREF=RUNREF" & vbCrLf _
         & whereClause
   LoadComboBox cbo, -1
   If cbo.ListCount > 0 Then
      cbo = cbo.List(0)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "FillComboBoxWithMoParts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
   
End Sub

Public Sub FillComboBoxWithMoRuns(cbo As ComboBox, whereClause As String)
   'Fill combo box with run numbers for runs satisfying where clause
   'Example FillComboBoxWithMoRuns cbo, "WHERE RUNSTATUS IN ( 'CO', 'CA' ) AND RUNREF = '12345678'"
   
   On Error GoTo DiaErr1
   cbo.Clear
   sSql = "SELECT DISTINCT RUNNO" & vbCrLf _
         & "FROM RunsTable" & vbCrLf _
         & whereClause
   LoadComboBox cbo, -1
   If cbo.ListCount > 0 Then
      cbo = cbo.List(0)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "FillComboBoxWithMoRuns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
End Sub

Public Sub UpdatePrePickMO(strMoPreRef As String, nMOPreRun As Integer)
      
   If (strMoPreRef <> "") Then
      sSql = "UPDATE RunsTable SET RUNPREPKMOREF='" & strMoPreRef & "'," _
         & " RUNPREPKNO=" & nMOPreRun & " " _
         & " WHERE RUNREF='" & SPartRef & "' AND " _
         & " RUNNO=" & nRunNo & " "
      clsADOCon.ExecuteSql sSql
   End If
End Sub

Public Sub CompleteMo(bPartialCompletion As Boolean, CompletionQty As Currency, _
   dateCompleted As Date, Optional bPrePick As Boolean = False)
   'THIS MUST HAPPEN INSIDE A TRANSACTION
   
   Dim bByte As Byte
   Dim bResponse As Byte
   Dim bLots As Byte
   Dim iOps As Integer
   Dim lCOUNTER As Long
   'Dim lSysCount As Long
   Dim sMsg As String
   Dim sStatus As String
   Dim vDate As Variant
   
   vDate = Format(ES_SYSDATE, "mm/dd/yy hh:mm")
   
   On Error GoTo DiaErr1
   
   GetRunCosts
   Dim sCreditAcct As String, sDebitAcct As String
   bResponse = GetPartAccounts(SPartRef, sCreditAcct, sDebitAcct)
   cRunCost = cRunOvHd + cRunMatl + cRunLabor + cRunExp
   Err = 0
   
   'Check lots
   bLots = CheckLotStatus()
   lCOUNTER = GetLastActivity() + 1
   sLotNumber = GetNextLotNumber()
   'lSysCount = lCOUNTER
   
   If CompletionQty > 0 Then
      sSql = "UPDATE PartTable SET PAQOH = ISNULL(PAQOH,0)+" & CompletionQty & "," _
         & "PALOTQTYREMAINING = ISNULL(PALOTQTYREMAINING,0)+" & CompletionQty & " " _
         & "WHERE PARTREF='" & SPartRef & "'"
      clsADOCon.ExecuteSql sSql
   End If
   
   'create ia & lot records if the quantity is > 0
   'or if qty = 0, this is a final completion, and there are no other completions
   Dim rdo As ADODB.Recordset
   Dim totalMoQty As Currency
   totalMoQty = 0
   If CompletionQty = 0 Then
      sSql = "select RUNYIELD from RunsTable" & vbCrLf _
         & "WHERE RUNREF='" & SPartRef & "' AND " _
         & "RUNNO=" & nRunNo & " "
      If clsADOCon.GetDataSet(sSql, rdo) Then
         totalMoQty = CCur(rdo.Fields(0).Value)
      End If
   End If
   
   If CompletionQty > 0 Or (CompletionQty = 0 And totalMoQty = 0) Then
      
      'format runno with leading spaces to make them line up
      Dim formattedRunNo As String
      formattedRunNo = CStr(nRunNo)
      If Len(formattedRunNo) < 5 Then
         formattedRunNo = Space(5 - Len(formattedRunNo)) & formattedRunNo
      End If
      
      'note: inpdate = selected date closed and inadate = system datetime is wrong,
      'but I'm afraid to change it.  Some maint scripts rely on times agreeing to match entries
      sSql = "INSERT INTO InvaTable" & vbCrLf _
         & "(INTYPE,INPART,INREF1,INREF2," _
         & "INPDATE,INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," & vbCrLf _
         & "INMOPART,INMORUN,INTOTMATL,INTOTLABOR,INTOTEXP," & "INTOTOH,INTOTHRS," & vbCrLf _
         & "INWIPLABACCT,INWIPMATACCT,INWIPOHDACCT,INWIPEXPACCT,INNUMBER,INLOTNUMBER,INUSER) " & vbCrLf _
         & "VALUES(" & IATYPE_MoCompletion & ",'" & SPartRef & "','COMPLETED RUN'," _
         & "'RUN " & formattedRunNo & "'," _
         & "'" & dateCompleted & "','" & vDate & "'," & CompletionQty & "," & CompletionQty & "," _
         & cRunCost & ",'" & sCreditAcct & "','" & sDebitAcct & "'," & vbCrLf _
         & "'" & SPartRef & "'," & nRunNo & "," & cRunMatl & "," & cRunLabor & "," _
         & cRunExp & "," & cRunOvHd & "," & cRunHours & ",'" & vbCrLf _
         & sWipLabAcct & "','" & sWipMatAcct & "','" _
         & sWipExpAcct & "','" & sWipOhdAcct & "'," & lCOUNTER & ",'" _
         & sLotNumber & "','" & sInitials & "')"
      clsADOCon.ExecuteSql sSql
   
      'find a lot userid not in use
      If bPartialCompletion Then
         sMsg = "MO PA-"
      ElseIf (bPrePick = True) Then
         sMsg = "MO PRE-"
      ElseIf GetComnTableBit("CoUseAbbreviatedLotNumbers") Then
         sMsg = ""
      Else
         sMsg = "MO CO-"
      End If
      
      sMsg = sMsg & sPartNo & " R" & nRunNo
      Dim suffix As String
      Dim lot As New ClassLot
      Do
         If Not lot.IsUserLotIdInUseForAnotherLot(sMsg & suffix, sLotNumber, False) Then
            Exit Do
         End If
         If suffix = "" Then
            suffix = "A"
         Else
            suffix = Chr(Asc(suffix) + 1)
         End If
      Loop While True
      
      sSql = "INSERT INTO LohdTable" & vbCrLf _
         & "(LOTNUMBER,LOTUSERLOTID," _
         & "LOTPARTREF,LOTPDATE,LOTADATE,LOTORIGINALQTY,LOTREMAININGQTY," & vbCrLf _
         & "LOTUNITCOST,LOTMOPARTREF,LOTMORUNNO," _
         & "LOTTOTMATL,LOTTOTLABOR,LOTTOTEXP,LOTTOTOH,LOTTOTHRS) " & vbCrLf _
         & "VALUES('" & sLotNumber & "','" & sMsg & suffix _
         & "','" & SPartRef & "','" & vDate & "','" & vDate & "'," & CompletionQty & "," & CompletionQty & "," & vbCrLf _
         & cRunCost & ",'" & SPartRef & "'," _
         & nRunNo & "," & cRunMatl & "," & cRunLabor & "," _
         & cRunExp & "," & cRunOvHd & "," & cRunHours & ")"
      clsADOCon.ExecuteSql sSql
      If bPartialCompletion Then
         sMsg = "MO Run Partial Comp"
      Else
         sMsg = "MO Run Completion"
      End If
      sSql = "INSERT INTO LoitTable" & vbCrLf _
         & "(LOINUMBER,LOIRECORD,LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
         & "LOIMOPARTREF,LOIMORUNNO,LOIACTIVITY,LOICOMMENT)" & vbCrLf _
         & "VALUES('" & sLotNumber & "',1,6,'" & SPartRef _
         & "','" & dateCompleted & "'," & Trim(str(CompletionQty)) _
         & ",'" & SPartRef & "'," & nRunNo & "," _
         & lCOUNTER & ",'MO Run Completion')"
      clsADOCon.ExecuteSql sSql
   End If
   
   'partial completion
   If bPartialCompletion Then
      Err.Clear
      sSql = "UPDATE RunsTable SET " _
         & "RUNYIELD=RUNYIELD+" & CompletionQty & "," _
         & "RUNPARTIALQTY=RUNPARTIALQTY+" & CompletionQty & "," _
         & "RUNPARTIALDATE='" & dateCompleted & "'," _
         & "RUNREMAININGQTY=RUNREMAININGQTY-" & CompletionQty & "," _
         & "RUNLOTNUMBER='" & sLotNumber & "' " _
         & "WHERE RUNREF='" & SPartRef & "' AND " _
         & "RUNNO=" & nRunNo & " "
      clsADOCon.ExecuteSql sSql
      
   'full completion
   Else
      sSql = "UPDATE RnopTable SET OPCOMPDATE='" & dateCompleted & "'," _
         & "OPCOMPLETE=1 WHERE OPREF='" & SPartRef & "' AND " _
         & "OPRUN=" & nRunNo & " AND OPCOMPLETE=0 "
      clsADOCon.ExecuteSql sSql
      sSql = "UPDATE RunsTable SET RUNCOMPLETE='" & dateCompleted & "'," _
         & "RUNYIELD=" & CompletionQty & ",RUNSTATUS='CO'," _
         & "RUNCOST=" & cRunCost & "," _
         & "RUNOHCOST=" & cRunOvHd & "," _
         & "RUNCMATL=" & cRunMatl & "," _
         & "RUNCEXP=" & cRunExp & "," _
         & "RUNCHRS=" & cRunHours & "," _
         & "RUNCLAB=" & cRunLabor & "," _
         & "RUNREMAININGQTY=0," _
         & "RUNLOTNUMBER='" & sLotNumber & "' " _
         & "WHERE RUNREF='" & SPartRef & "' AND " _
         & "RUNNO=" & nRunNo & " "
      clsADOCon.ExecuteSql sSql
   End If
   
   Dim sReturnMsg As String
   sReturnMsg = UpdateMOCosts(SPartRef, nRunNo, "COMPLETED RUN", _
      cRunMatl, cRunLabor, cRunExp, cRunOvHd, cRunHours)

   AverageCost (SPartRef)
   
   UpdateWipColumns lCOUNTER
'   SysMsg "Manufacturing Order Is Complete.", True, Me
   Set rdo = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "ClassMo.CompleteMO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
   
End Sub

Public Function GetPreviousCompletions() As Currency
   'update run quantities
   'return quantity in
   
   Dim RdoPrev As ADODB.Recordset
   Dim cQtyIn As Currency 'Type 6 = IATYPE_MoCompletion
   Dim cQtyOut As Currency 'Type 38 = IATYPE_CanceledMoCompletion
   Dim cQtyBal As Currency 'Total Completed
   
   'Complete
   On Error GoTo DiaErr1
   'MouseCursor 13
   sSql = "SELECT SUM(INAQTY) AS QtyComplete FROM InvaTable WHERE (INTYPE=" & IATYPE_MoCompletion & vbCrLf _
          & "AND INMOPART='" & SPartRef & "' AND INMORUN=" & nRunNo & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrev, ES_FORWARD)
   If bSqlRows Then
      With RdoPrev
         If Not IsNull(!QtyComplete) Then
            cQtyIn = !QtyComplete
         Else
            cQtyIn = 0
         End If
         ClearResultSet RdoPrev
      End With
   End If
   
   If cQtyIn > 0 Then
      sSql = "SELECT SUM(INAQTY) AS QtyComplete FROM InvaTable WHERE (INTYPE=" & IATYPE_CanceledMoCompletion & vbCrLf _
             & "AND INMOPART='" & SPartRef & "' AND INMORUN=" & nRunNo & ")"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrev, ES_FORWARD)
      If bSqlRows Then
         With RdoPrev
            If Not IsNull(!QtyComplete) Then
               cQtyOut = !QtyComplete
            Else
               cQtyOut = 0
            End If
            ClearResultSet RdoPrev
         End With
      End If
   End If
   cQtyBal = cQtyIn - Abs(cQtyOut)
   If cQtyBal < 0 Then cQtyBal = 0
   sSql = "UPDATE RunsTable SET RUNPARTIALQTY=" & cQtyBal & "," & vbCrLf _
      & "RUNREMAININGQTY=RUNQTY - " & cQtyBal & vbCrLf _
      & "WHERE (RUNREF='" & SPartRef & "' AND RUNNO=" & nRunNo & ")"
   clsADOCon.ExecuteSql sSql
   'lblPartial = Format(cQtyIn, ES_QuantityDataFormat)
   GetPreviousCompletions = cQtyIn
   'MouseCursor 0
   Set RdoPrev = Nothing
   Exit Function
   
DiaErr1:
'   Resume DiaErr2
'DiaErr2:
   'MouseCursor 0
   'Set RdoPrev = Nothing
   
End Function

Public Function GetOpenMoStatus(part As String, run As Long) As String
   'retrieve the status of an MO based upon pick operations and run parameters
   'returns only SC, RL, PL, PP, PC
   'other statuses are set by transactions (CO, CL, CA, SC)
   
   sSql = "select PKTYPE, count(*) as ct from MopkTable" & vbCrLf _
      & "where PKMOPART = '" & Compress(part) & "' AND PKMORUN = " & run & vbCrLf _
      & "group by PKTYPE order by PKTYPE"
   Dim rdo As ADODB.Recordset

   Dim openItems As Integer
   Dim pickedItems As Integer
   If clsADOCon.GetDataSet(sSql, rdo) Then
      While Not rdo.EOF
         If rdo!PKTYPE = IATYPE_PickOpenItem Then
            openItems = rdo!ct
         ElseIf rdo!PKTYPE = IATYPE_PickedItem Then
            pickedItems = rdo!ct
         End If
         rdo.MoveNext
      Wend
      If openItems <> 0 And pickedItems <> 0 Then
         GetOpenMoStatus = "PP"
      ElseIf openItems = 0 And pickedItems = 0 Then
         GetOpenMoStatus = "SC"
      ElseIf openItems <> 0 Then
         GetOpenMoStatus = "PL"
      Else
         GetOpenMoStatus = "PC"
      End If
   Else
      GetOpenMoStatus = "SC"
   End If
   
   'if no pick items, determine whether MO is RL or SC
   If GetOpenMoStatus = "SC" Then
      sSql = "select RUNRELEASED from RunsTable" & vbCrLf _
         & "where RUNREF = '" & Compress(part) & "' AND RUNNO = " & run
      If clsADOCon.GetDataSet(sSql, rdo) Then
         If rdo!RUNRELEASED = 1 Then
            GetOpenMoStatus = "RL"
         End If
      End If
   End If
   Set rdo = Nothing
End Function

Public Function ScheduleOperations(sPartNumber As String, lblRun As Long, cRunqty As Currency, _
   txtCmp As Date, Optional shMsg As Boolean)
   'sPartNumber = MO Part Reference
   'lblRun = Run number
   'cRunQty = MO Quantity
   'txtCmp = MO Completion date
   
   'from ShopShe01a.BackSchedule
   
   Dim bGoodWCCal As Boolean
   Dim bGoodCoCal As Byte
   Dim RdoOps As ADODB.Recordset
   Dim bDays As Byte
   Dim bFirstOp As Byte
   Dim bQMConvert As Byte 'Hour conversion for Q&M
   Dim bCounter As Byte 'Fall through for Q&M
   
   Dim A As Integer
   Dim iList As Integer
   Dim n As Integer
   Dim d As Long
   Dim iDay As Integer
   
   Dim iTotalOps As Integer
   Dim iSunday As Integer
   Dim iSaturday As Integer
   
   Dim cOpHrs As Currency
   Dim cMoveHrs As Currency
   Dim cQueueHrs As Currency
   'Dim cRunQty As Currency
   Dim cSetupHrs As Currency
   Dim cUnitHrs As Currency
   Dim cUnusedHrs As Currency
   Dim cTestHours As Currency
   
   Dim vSchedDate As Date
   Dim vPreSchedDate As Date
   Dim vStartDate As Date
   Dim vMoveDate As Date
   Dim vQueueDate As Date
   
   Dim sCenter As String
   Dim sShop As String
   Dim sOpNum As String
   Dim sOldCenter As String
   Dim sLogData As String
   Dim bAvalHrs As Boolean
   Dim vRunOps(300, 12) As Variant
   '0 = Center
   '1 = Shop
   '2 = OPNO
   '3 = QHRS
   '4 = MHRS
   '5 = OPSUHRS
   '6 = OPUNITHRS
   '7 = Sched Complete
   '8 = Run Start Date
   '9 = SU Start Date
   '10 = Move Date
   '11 = Queue Date
   
   
   'MouseCursor 13
   On Error GoTo whoops
   bQMConvert = GetQNMConversion()
   'Initialize variant data
   For iList = 1 To 298
      vRunOps(iList, 1) = ""
      vRunOps(iList, 2) = ""
      vRunOps(iList, 3) = 0
      vRunOps(iList, 4) = 0
      vRunOps(iList, 5) = 0
      vRunOps(iList, RUNOP_RunHours) = 0
      vRunOps(iList, RUNOP_CompleteDate) = Format(ES_SYSDATE, "mm/dd/yy")
      vRunOps(iList, RUNOP_StartDate) = Format(ES_SYSDATE, "mm/dd/yy")
      vRunOps(iList, RUNOP_SetupStartDate) = Format(ES_SYSDATE, "mm/dd/yy")
   Next
   vRunOps(iList, 1) = ""
   vRunOps(iList, 2) = ""
   vRunOps(iList, 3) = 0
   vRunOps(iList, 4) = 0
   vRunOps(iList, 5) = 0
   vRunOps(iList, RUNOP_RunHours) = 0
   vRunOps(iList, RUNOP_CompleteDate) = Format(ES_SYSDATE, "mm/dd/yy")
   vRunOps(iList, RUNOP_StartDate) = Format(ES_SYSDATE, "mm/dd/yy")
   vRunOps(iList, RUNOP_SetupStartDate) = Format(ES_SYSDATE, "mm/dd/yy")
   
   'cRunQty = Val(txtQty)
   'get distinct Shops,Work Centers
   'PrgBar.Value = 10
   'On Error GoTo DiaErr1
   
   'PrgBar.Value = 20
   'get the workcenters and ops in reverse order
   sSql = "SELECT OPREF,OPRUN,OPNO,OPSHOP,OPCENTER,OPQHRS,OPMHRS,OPSUHRS,OPUNITHRS" & vbCrLf _
      & " FROM RnopTable WHERE OPREF='" & sPartNumber & "' AND OPRUN=" & lblRun & vbCrLf _
      & " ORDER BY OPNO DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOps, ES_STATIC)
   If bSqlRows Then
      With RdoOps
         Do Until .EOF
            iTotalOps = iTotalOps + 1
            vRunOps(iTotalOps, 0) = "" & Trim(!OPSHOP)
            vRunOps(iTotalOps, 1) = "" & Trim(!OPCENTER)
            vRunOps(iTotalOps, 2) = !opNo
            vRunOps(iTotalOps, 3) = !OPQHRS
            vRunOps(iTotalOps, 4) = !OPMHRS
            vRunOps(iTotalOps, 5) = !OPSUHRS
            vRunOps(iTotalOps, RUNOP_RunHours) = !OPUNITHRS * cRunqty
            .MoveNext
         Loop
         ClearResultSet RdoOps
      End With
      vRunOps(1, RUNOP_CompleteDate) = Format(txtCmp, "mm/dd/yy")
   End If
   Set RdoOps = Nothing
   
   'Get Workcenter Calendars and time
   vRunOps(0, 4) = vRunOps(1, 4)
   'Release the memory
   
   'PrgBar.Value = 50
   
   ' Hard coded the 16:00, shuoul dbe read from the database later
   ' 8 hr shift
   'FormatScheduleTime
   vTimeFormat = "mm/dd/yy hh:mm"
   vSchedDate = Format(txtCmp, vTimeFormat)
   Dim curSchedDate As Date
   'Schedule 'em
   'We'll Allow a Calendar or Work Center Hours.
   'If neither then they are bumming
   'On Error Resume Next
   cTestHours = 0
   sOldCenter = ""
   cUnusedHrs = 0
   bFirstOp = 1
   sLogData = "MO Schedule for PartNumber = " + sPartNumber + _
                " Run " + CStr(lblRun) + " Qty = " + CStr(cRunqty) + "." + vbCrLf
    sLogData = sLogData + "==============================" + vbCrLf
    
    
   For iList = 1 To iTotalOps 'one op at a time
      ' Allow move 6/6/00
      If cUnusedHrs < 0 Then cUnusedHrs = 0
      
      
      'If PrgBar.Value < 100 Then PrgBar.Value = PrgBar.Value + 10
      
      Dim cExactQueuehrs As Currency, cExactMoveHrs As Currency, cExactSetupHrs As Currency, cExactUnitHrs As Currency
      cExactQueuehrs = Val(vRunOps(iList, 3))
      cExactMoveHrs = Val(vRunOps(iList, 4))
      cExactSetupHrs = Val(vRunOps(iList, 5))
      cExactUnitHrs = Val(vRunOps(iList, RUNOP_RunHours))
      
      cQueueHrs = RoundUpToInteger(cExactQueuehrs)
      cMoveHrs = RoundUpToInteger(cExactMoveHrs)
      cSetupHrs = RoundUpToInteger(cExactSetupHrs)
      cUnitHrs = RoundUpToInteger(cExactUnitHrs)
      
      sOpNum = vRunOps(iList, 2)
      sShop = vRunOps(iList, 0)
      sCenter = vRunOps(iList, 1)
      vPreSchedDate = vSchedDate
      
'      If sOpNum = "30" Or sOpNum = "40" Then
'         MsgBox sOpNum
'         Dim xyz As Integer
'         xyz = 0     'breakpoint
'      End If
'
      bGoodWCCal = GetThisCalendar(vSchedDate, sShop, sCenter)
      If bGoodWCCal Then
         bAvalHrs = GetValidWCCalSchDt(vSchedDate, sShop, sCenter)
      Else
         bAvalHrs = GetValidWCSchDt(vSchedDate, sCenter)
      End If
    
      If Not bAvalHrs Then
       GoTo ErrorCalender
      End If
      
      If sCenter <> sOldCenter Then
         ' Get the last shift end date
         GetLastShiftTime vSchedDate, sCenter, Format(vSchedDate, "w")

         If (sOldCenter <> "") Then
            If (CDate(vSchedDate) > CDate(vPreSchedDate)) Then
               vSchedDate = vPreSchedDate
            End If
            
            ' added 11/9/2017: Get the unused hours remaining for this operation
            ' previously cUnusedHrs was for the previous workcenter
            cUnusedHrs = WcHoursBeforeThisTime(vSchedDate, sShop, sCenter)
         End If
      
      End If

      sOldCenter = sCenter
      
      'end move
      vRunOps(iList, RUNOP_CompleteDate) = vSchedDate
      sOpNum = vRunOps(iList, 2)
      sShop = vRunOps(iList, 0)
      sCenter = vRunOps(iList, 1)
      
      ' Init the Current date
      curSchedDate = CDate(vSchedDate)
      ' Cascade the Log file
      sLogData = sLogData + "OPNO = " + sOpNum + " : SHOP = " + sShop + " : CENTER = " + sCenter + vbCrLf
      sLogData = sLogData & "Q=" & CStr(cExactQueuehrs) & " MOVE=" & CStr(cExactMoveHrs) _
         & " SU=" & CStr(cExactSetupHrs) & " RUN=" & CStr(cExactUnitHrs) & vbCrLf
         
      
      '11/5/05
      'Move Hrs
      If cMoveHrs > 0 Then
         'If bQMConvert = 12 Then cMoveHrs = cMoveHrs * 2
         bGoodWCCal = GetThisCalendar(vSchedDate, sShop, sCenter)
         If bGoodWCCal Then
            If bFirstOp = 1 Then
               ' Get the last shift end date
               curSchedDate = CDate(vSchedDate)
               GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
               If cUnusedHrs <= 0 Then cUnusedHrs = GetCenterCalHours(curSchedDate, sShop, sCenter, Format(curSchedDate, "d"))
               bFirstOp = 0
            End If
            Do Until cMoveHrs <= 0
              If cUnusedHrs >= cMoveHrs Then
                  cUnusedHrs = cUnusedHrs - cMoveHrs
                  
                  vSchedDate = Format(DateAdd("h", -cMoveHrs, vSchedDate), "mm/dd/yy hh:mm")
                  curSchedDate = CDate(vSchedDate)
                  ' If we have a UnusedHrs we need to change the logic
                  bAvalHrs = GetValidWCCalSchDt(vSchedDate, sShop, sCenter)
                    If Not bAvalHrs Then
                     GoTo ErrorCalender
                    End If
                   
                  cMoveHrs = 0
                  Exit Do
               Else
                  cMoveHrs = cMoveHrs - cUnusedHrs
                  If cMoveHrs < 0 Then cMoveHrs = 0
                  ' Get the last shift end date
                  curSchedDate = curSchedDate - 1
                  'vSchedDate = Format(vSchedDate - 1, vTimeFormat)
                  bAvalHrs = GetValidWCCalSchDt(Format(curSchedDate), sShop, sCenter)
                    If Not bAvalHrs Then
                     GoTo ErrorCalender
                    End If
                  GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
                  cUnusedHrs = GetCenterCalHours(vSchedDate, sShop, sCenter, Format(curSchedDate, "d"))
                  ' I am not sure we need this hard coded value
                  If cUnusedHrs = 0 Then cUnusedHrs = 8
                  
               End If
            Loop
         Else
            Do Until cMoveHrs <= 0
                bAvalHrs = GetValidWCSchDt(vSchedDate, sCenter)
                If Not bAvalHrs Then
                 GoTo ErrorCalender
                End If
                cUnusedHrs = GetCenterHours(sCenter, Format(vSchedDate, "w"))
                  cMoveHrs = cMoveHrs - cUnusedHrs
                  If cMoveHrs > 0 Then
                     If cMoveHrs > cUnusedHrs Then
                        ' Get the last shift end date
                        curSchedDate = curSchedDate - 1
                        GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
                        'vSchedDate = Format(vSchedDate - 1, vTimeFormat)
                     Else
                        vSchedDate = Format(DateAdd("h", -cMoveHrs, vSchedDate), "mm/dd/yy hh:mm")
                        curSchedDate = CDate(vSchedDate)
                        bAvalHrs = GetValidWCSchDt(vSchedDate, sCenter)
                        If Not bAvalHrs Then
                         GoTo ErrorCalender
                        End If
                        Exit Do
                     End If
                  End If
            Loop
         End If
      End If
      
      vMoveDate = vSchedDate
      vRunOps(iList, RUNOP_MoveDate) = vMoveDate
      ' Cascade the Log file
      'sLogData = sLogData + "Move StartDate = " + CStr(vRunOps(iList, RUNOP_MoveDate)) + vbCrLf
      sLogData = sLogData + "Move StartDate = " + Format(CDate(vRunOps(iList, RUNOP_MoveDate)), "mm/dd/yyyy hh:mm AM/PM") + vbCrLf
      'Format$(Now, "hh:mm AM/PM")
      'run date
      If cUnitHrs > 0 Then
      
         'this is redundant with logic above - refactor
         bGoodWCCal = GetThisCalendar(vSchedDate, sShop, sCenter)
         If bGoodWCCal Then
            If bFirstOp = 1 Then
               ' Get the last shift end date
               curSchedDate = CDate(vSchedDate)
               GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
               If cUnusedHrs <= 0 Then cUnusedHrs = GetCenterCalHours(curSchedDate, sShop, sCenter, Format(curSchedDate, "d"))
               bFirstOp = 0
            End If
            Do Until cUnitHrs <= 0
               If cUnusedHrs >= cUnitHrs Then
                  cUnusedHrs = cUnusedHrs - cUnitHrs
                  vSchedDate = Format(DateAdd("h", -cUnitHrs, vSchedDate), "mm/dd/yy hh:mm")
                  curSchedDate = CDate(vSchedDate)
                  ' If we have a UnusedHrs we need to change the logic
                  bAvalHrs = GetValidWCCalSchDt(vSchedDate, sShop, sCenter)
                   If Not bAvalHrs Then
                    GoTo ErrorCalender
                   End If
                   
                  cUnitHrs = 0
                  Exit Do
               Else
                  cUnitHrs = cUnitHrs - cUnusedHrs
                  If cUnitHrs < 0 Then cUnitHrs = 0
                  ' Get the last shift end date
                  curSchedDate = curSchedDate - 1
                  
                  'MM vSchedDate = Format(vSchedDate - 1, vTimeFormat)
                  bAvalHrs = GetValidWCCalSchDt(curSchedDate, sShop, sCenter)
                   If Not bAvalHrs Then
                    GoTo ErrorCalender
                   End If
                  
                  GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
                  
                  cUnusedHrs = GetCenterCalHours(curSchedDate, sShop, sCenter, Format(curSchedDate, "d"))
                  ' I am not sure we need this hard coded value
                  If cUnusedHrs = 0 Then cUnusedHrs = 8
                  
               End If
            Loop
         Else
            'work center
            Do Until cUnitHrs <= 0
            
                bAvalHrs = GetValidWCSchDt(vSchedDate, sCenter)
                If Not bAvalHrs Then
                 GoTo ErrorCalender
                End If
                cUnusedHrs = GetCenterHours(sCenter, Format(vSchedDate, "w"))
                  cUnitHrs = cUnitHrs - cUnusedHrs
                  If cUnitHrs > 0 Then
                     If cUnitHrs > cUnusedHrs Then
                        ' Get the last shift end date
                        curSchedDate = curSchedDate - 1
                        GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
                        'vSchedDate = Format(vSchedDate - 1, vTimeFormat)
                     Else
                        vSchedDate = Format(DateAdd("h", -cUnitHrs, vSchedDate), "mm/dd/yy hh:mm")
                        curSchedDate = CDate(vSchedDate)
                        bAvalHrs = GetValidWCSchDt(vSchedDate, sCenter)
                        If Not bAvalHrs Then
                         GoTo ErrorCalender
                        End If
                        Exit Do
                     End If
                  End If
            Loop
         End If
      End If
    
      ' Add the run op start date
      vRunOps(iList, RUNOP_StartDate) = Format(vSchedDate, "mm/dd/yy hh:mm")
      ' Cascade the Log file
      'sLogData = sLogData + "RUNOP StartDate= " + CStr(vRunOps(iList, RUNOP_StartDate)) + vbCrLf
      sLogData = sLogData + "RUNOP StartDate = " + Format(CDate(vRunOps(iList, RUNOP_StartDate)), "mm/dd/yyyy hh:mm AM/PM") + vbCrLf
      
      
      If cUnusedHrs < 0 Then cUnusedHrs = 0
      'Setup date
      If cSetupHrs > 0 Then
         If bGoodWCCal Then
            If bFirstOp = 1 Then
               ' Get the last shift end date
               curSchedDate = CDate(vSchedDate)
               GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
               If cUnusedHrs <= 0 Then cUnusedHrs = GetCenterCalHours(curSchedDate, sShop, sCenter, Format(curSchedDate, "d"))
               bFirstOp = 0
            End If
            
            Do Until cSetupHrs <= 0
               If cUnusedHrs >= cSetupHrs Then
                  cUnusedHrs = cUnusedHrs - cSetupHrs
                  vSchedDate = Format(DateAdd("h", -cSetupHrs, vSchedDate), "mm/dd/yy hh:mm")
                  curSchedDate = CDate(vSchedDate)
                  ' call WC Calender
                  bAvalHrs = GetValidWCCalSchDt(vSchedDate, sShop, sCenter)
                    If Not bAvalHrs Then
                     GoTo ErrorCalender
                    End If
                  cSetupHrs = 0
                  Exit Do
               Else
                  cSetupHrs = cSetupHrs - cUnusedHrs
                  ' Get the last shift end date
                  curSchedDate = curSchedDate - 1
                  bAvalHrs = GetValidWCCalSchDt(curSchedDate, sShop, sCenter)
                  If Not bAvalHrs Then
                   GoTo ErrorCalender
                  End If
                  GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
                  'vSchedDate = Format(vSchedDate - 1, vTimeFormat)
                  
                  cUnusedHrs = GetCenterCalHours(curSchedDate, sShop, sCenter, Format(curSchedDate, "d"))
                  ' I am not sure we need this hard coded value
                  If cUnusedHrs = 0 Then cUnusedHrs = 8
               End If
            Loop
         Else
            'work center
            Do Until cSetupHrs <= 0
                bAvalHrs = GetValidWCSchDt(vSchedDate, sCenter)
                If Not bAvalHrs Then
                 GoTo ErrorCalender
                End If
                cUnusedHrs = GetCenterHours(sCenter, Format(vSchedDate, "w"))
                  
                  cSetupHrs = cSetupHrs - cUnusedHrs
                  If cSetupHrs > 0 Then
                     If cSetupHrs > cUnusedHrs Then
                        ' Get the last shift end date
                        curSchedDate = curSchedDate - 1
                        GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
                        'vSchedDate = Format(vSchedDate - 1, vTimeFormat)
                     Else
                        vSchedDate = Format(DateAdd("h", -cSetupHrs, vSchedDate), "mm/dd/yy hh:mm")
                        curSchedDate = CDate(vSchedDate)
                        bAvalHrs = GetValidWCSchDt(vSchedDate, sCenter)
                        If Not bAvalHrs Then
                         GoTo ErrorCalender
                        End If

                        Exit Do
                     End If
                  End If
             Loop
         End If
      End If
      
      
      vRunOps(iList, RUNOP_SetupStartDate) = Format(vSchedDate, "mm/dd/yy hh:mm")
      ' Cascade the Log file
      'sLogData = sLogData + "Setup StartDate = " + CStr(vRunOps(iList, RUNOP_SetupStartDate)) + vbCrLf
      sLogData = sLogData + "Setup StartDate = " + Format(CDate(vRunOps(iList, RUNOP_SetupStartDate)), "mm/dd/yyyy hh:mm AM/PM") + vbCrLf
      
      Dim bLastShifDone As Integer
      Dim strDateTime As String
      '       Queue Hours
      If cQueueHrs > 0 Then
         bGoodWCCal = GetThisCalendar(vSchedDate, sShop, sCenter)
         If bGoodWCCal Then
            bLastShifDone = 0
            If bFirstOp = 1 Then
               ' Get the last shift end date
               curSchedDate = CDate(vSchedDate)
               GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
               
               strDateTime = Format(vSchedDate, "mm/dd/yy hh:mm")
               If cUnusedHrs <= 0 Then cUnusedHrs = GetCenterCalHours(curSchedDate, sShop, sCenter, Format(curSchedDate, "d"))
               bFirstOp = 0
               bLastShifDone = 1
            End If
            
            Do Until cQueueHrs <= 0
              If cUnusedHrs >= cQueueHrs Then
                  cUnusedHrs = cUnusedHrs - cQueueHrs
                  vSchedDate = Format(DateAdd("h", -cQueueHrs, vSchedDate), "mm/dd/yy hh:mm")
                  curSchedDate = CDate(vSchedDate)
                  ' If we have a UnusedHrs we need to change the logic
                  bAvalHrs = GetValidWCCalSchDt(vSchedDate, sShop, sCenter)
                   
                    If Not bAvalHrs Then
                     GoTo ErrorCalender
                    End If
                  Exit Do
               Else
                  cQueueHrs = cQueueHrs - cUnusedHrs
                  If cQueueHrs < 0 Then cQueueHrs = 0
                  ' Get the last shift end date
                  curSchedDate = curSchedDate - 1
                  'vSchedDate = Format(vSchedDate - 1, vTimeFormat)
                  
                  'vSchedDate = vSchedDate - 1
                  'vSchedDate = Format(vSchedDate - 1, vTimeFormat)
                  bAvalHrs = GetValidWCCalSchDt(curSchedDate, sShop, sCenter)
                  If Not bAvalHrs Then
                     GoTo ErrorCalender
                  End If
                  GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
                  
                  cUnusedHrs = GetCenterCalHours(curSchedDate, sShop, sCenter, Format(curSchedDate, "d"))
                  ' I am not sure we need this hard coded value
                  If cUnusedHrs = 0 Then cUnusedHrs = 8
                  
               End If
            Loop
         Else
            Do Until cQueueHrs <= 0
               
                bAvalHrs = GetValidWCSchDt(vSchedDate, sCenter)
                If Not bAvalHrs Then
                 GoTo ErrorCalender
                End If
                cUnusedHrs = GetCenterHours(sCenter, Format(vSchedDate, "w"))
                
                  ' I am not sure we need this hard coded value
                  ' Just for now 5/14/2009
                  If cUnusedHrs = 0 Then cUnusedHrs = 8
                  cQueueHrs = cQueueHrs - cUnusedHrs
                  If cQueueHrs > 0 Then
                     If cQueueHrs > cUnusedHrs Then
                        ' Get the last shift end date
                        curSchedDate = curSchedDate - 1
                        GetLastShiftRevTime curSchedDate, vSchedDate, sCenter, Format(curSchedDate, "w")
                        'vSchedDate = Format(vSchedDate - 1, vTimeFormat)
                     Else
                        vSchedDate = Format(DateAdd("h", -cQueueHrs, vSchedDate), "mm/dd/yy hh:mm")
                        curSchedDate = CDate(vSchedDate)
                        bAvalHrs = GetValidWCSchDt(vSchedDate, sCenter)
                        If Not bAvalHrs Then
                         GoTo ErrorCalender
                        End If
                        Exit Do
                     End If
                  End If
            Loop
         End If
      End If
      vRunOps(iList, RUNOP_QueueDate) = vSchedDate
      ' Cascade the Log file
      'sLogData = sLogData + "Queue StartDate = " + CStr(vRunOps(iList, RUNOP_QueueDate)) + vbCrLf
      sLogData = sLogData + "Queue StartDate = " + Format(CDate(vRunOps(iList, RUNOP_QueueDate)), "mm/dd/yyyy hh:mm AM/PM") + vbCrLf
      sLogData = sLogData + "==============================" + vbCrLf
      
      'end queue
   Next
   
   vStartDate = vSchedDate
   sLogData = sLogData + "Schedule StartDate = " + CStr(vSchedDate) + vbCrLf
   
   'Debug.Print sLogData
   
   'If shMsg = True Then
   ' Ticket number : 37682
   ' Removed the pop up information.
   ' MsgBox sLogData, vbInformation
   'End If
   'testing
   'On Error GoTo DiaErr1
   For iList = 1 To iTotalOps
      sSql = "UPDATE RnopTable SET OPSCHEDDATE='" & vRunOps(iList, RUNOP_CompleteDate) & "'," & vbCrLf _
             & "OPRUNDATE='" & vRunOps(iList, RUNOP_StartDate) & "'," & vbCrLf _
             & "OPSUDATE='" & vRunOps(iList, RUNOP_SetupStartDate) & "'," & vbCrLf _
             & "OPMDATE='" & vRunOps(iList, RUNOP_MoveDate) & "', " & vbCrLf _
             & "OPQDATE='" & vRunOps(iList, RUNOP_QueueDate) & "', " & vbCrLf _
             & "OPRUNHRS=" & vRunOps(iList, RUNOP_RunHours) & vbCrLf _
             & "WHERE OPREF='" & sPartNumber & "' AND OPRUN=" & lblRun _
             & " AND OPNO=" & vRunOps(iList, 2) & " "
      clsADOCon.ExecuteSql sSql
   Next
   
   sSql = "UPDATE RunsTable SET RUNSCHED='" & Format(txtCmp, "mm/dd/yy 16:00") & "'," _
          & "RUNPKSTART='" & Format(vStartDate - 1, "mm/dd/yy hh:nn") & "'," _
          & "RUNSTART='" & Format(vStartDate, "mm/dd/yy 07:00") & "'," _
          & "RUNQTY=" & cRunqty & ", RUNREMAININGQTY=" & cRunqty & "  WHERE RUNREF='" & sPartNumber & "' " _
          & "AND RUNNO=" & lblRun & " "
   
   clsADOCon.ExecuteSql sSql
   
   'On Error Resume Next
   'Update picks
   sSql = "UPDATE MopkTable SET PKPDATE='" & vStartDate & "' " _
          & "WHERE PKMOPART='" & sPartNumber & "' AND PKMORUN=" & lblRun & " " _
          & "AND PKTYPE<>12 AND PKAQTY=0 "
   clsADOCon.ExecuteSql sSql
   
   Set RdoOps = Nothing
   '   New stuff to Here
   'optDmy.Value = vbChecked
   'On Error GoTo 0
   ScheduleOperations = True
   Exit Function
   
'DiaErr1:
'   Resume DiaErr2
'DiaErr2:
'   MouseCursor 0
'   'PrgBar.Visible = False
'   If optAllocate.Value = vbUnchecked Then
'      ShopSHe02a.lblFrom = cmbPrt
'      ShopSHe02a.optFrom.Value = vbChecked
'      ShopSHe02a.txtPri = txtPri
'      ShopSHe02a.txtQty = txtQty
'      ShopSHe02a.lblSch = txtCmp
'      ShopSHe02a.cmbPrt = cmbPrt
'      ShopSHe02a.cmbRun = txtRun
'   End If
'   MsgBox "Couldn't Schedule MO.", vbExclamation, Caption
'   FinishMo
'   On Error GoTo 0
   
ErrorCalender:
   MsgBox "ClassMO.ScheduleOperations: Could not find hours for " & sPartNumber & " run " & CStr(lblRun) & " in WC " & sShop & "/" & sCenter & " calenders within 10 days"
   Exit Function
whoops:
   sProcName = "ScheduleOperations"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   MsgBox "ClassMO.ScheduleOperations failed: " & Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm

End Function

Public Function GetWeekEnd(TestDate As Variant) As Byte
   Dim RdoWe As ADODB.Recordset
   GetWeekEnd = 0
   If Left(Format(TestDate, "ddd"), 1) <> "S" Then Exit Function
   
   sSql = "SELECT SUM(COCSHT1+COCSHT2+COCSHT3+COCSHT4) AS AvailHours FROM " _
          & "CoclTable WHERE COCREF='" & Left(TestDate, 3) & "-" _
          & Right(TestDate, 4) & " ' AND COCDAY=" & Val(Mid(TestDate, 5, 2)) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWe, ES_FORWARD)
   If Not IsNull(RdoWe!AvailHours) Then GetWeekEnd = RdoWe!AvailHours
   
   If GetWeekEnd = 0 Then
      If Format(TestDate, "ddd") = "Sun" Then GetWeekEnd = 2 _
                Else GetWeekEnd = 1
   End If
   Set RdoWe = Nothing
   
End Function


Public Function GetQNMConversion() As Byte
   Dim RdoGet As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT QueueMoveConversion FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         If Not IsNull(!QueueMoveConversion) Then
            GetQNMConversion = !QueueMoveConversion
         Else
            GetQNMConversion = 24
         End If
         .Cancel
      End With
      ClearResultSet RdoGet
   Else
      GetQNMConversion = 24
   End If
   Set RdoGet = Nothing
   Exit Function
   
modErr1:
   GetQNMConversion = 24
   
End Function


Public Function TestWeekEnd(CalMonth As String, CalDay As String, CalShop As String, _
                            Calcenter As String) As Integer
   Dim RdoWend As ADODB.Recordset
   sSql = "SELECT SUM(WCCSHH1+WCCSHH2+WCCSHH3+WCCSHH4) As TotHours " _
          & "FROM WcclTable WHERE (WCCREF='" & CalMonth & "' AND " _
          & "DATENAME(dw,WCCDATE)Like '" & CalDay & "%' AND " _
          & "WCCSHOP='" & CalShop & "' AND WCCCENTER='" & Calcenter & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWend, ES_FORWARD)
   If bSqlRows Then
      With RdoWend
         If Not IsNull(!TotHours) Then
            TestWeekEnd = !TotHours
         Else
            TestWeekEnd = 0
         End If
      End With
   Else
      TestWeekEnd = 0
   End If
   If TestWeekEnd < 2 Then TestWeekEnd = 0 _
                                         Else TestWeekEnd = 1
   
   Set RdoWend = Nothing
End Function



Public Function GetThisCoCalendar(dTdate As Date) As Boolean
   Dim RdoCal As ADODB.Recordset
   Dim sTMonth As String
   Dim sTYear As String
   Dim sTDay As Integer
   sTMonth = Format(dTdate, "mmm")
   sTYear = sTMonth & "-" & Format(dTdate, "yyyy")
   sTDay = Format(dTdate, "d")
   sSql = "SELECT COCREF,COCDAY FROM CoclTable WHERE " _
          & "COCREF='" & sTYear & "' AND COCDAY=" & sTDay & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCal, ES_FORWARD)
   If bSqlRows Then
      GetThisCoCalendar = True
   Else
      GetThisCoCalendar = False
   End If
   Set RdoCal = Nothing
   
End Function


Public Function GetQMCalHours(dTdate As Date) As Currency
   Dim RdoTme As ADODB.Recordset
   Dim sTMonth As String
   Dim iTDay As Integer
   sTMonth = Format(dTdate, "mmm") & "-" & Format(dTdate, "yyyy")
   iTDay = Format(dTdate, "d")
'ALTER proc [dbo].[Qry_GetCompanyCalendarTime]
'(@Month char(8),@Day tinyint) AS
'   SELECT sum(COCSHT1+COCSHT2+COCSHT3+COCSHT4)
'   From CoclTable
'   WHERE (COCREF=@Month AND COCDAY=@Day)
   sSql = "Qry_GetCompanyCalendarTime '" & sTMonth & "'," & iTDay & " "
   Set RdoTme = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   'Set RdoTme = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not RdoTme.BOF And Not RdoTme.EOF Then
      With RdoTme
         If Not IsNull(.Fields(0)) Then
            GetQMCalHours = .Fields(0)
         Else
            GetQMCalHours = 1
         End If
         ClearResultSet RdoTme
      End With
   Else
      GetQMCalHours = 1
   End If
   Set RdoTme = Nothing
End Function


Function FormatScheduleTime(Optional cHours As Currency) As Variant
   If cHours = 0 Then cHours = 8
   Select Case cHours
      Case Is < 8.5
         FormatScheduleTime = "mm/dd/yy 14:30"
      Case 8.5 To 16
         MsgBox "FormatScheduleTime cHours = " & cHours & " shouldn't happen"
         FormatScheduleTime = "mm/dd/yy 21:30"
      Case Is > 16
         MsgBox "FormatScheduleTime cHours = " & cHours & " shouldn't happen"
         FormatScheduleTime = "mm/dd/yy 23:59"
   End Select
   vTimeFormat = FormatScheduleTime
   
End Function


Public Function GetCenterCalHours(dSchedDt As Date, sShop As String, sCenter As String, iDay As Integer) As Currency
   Dim RdoTme As ADODB.Recordset
   Dim cResources As Currency
   Dim sMonth As String
   sMonth = Format(dSchedDt, "mmm") & "-" & Format(dSchedDt, "yyyy")
'ALTER proc [dbo].[Qry_GetWorkCenterTimes]
'(@Month char(8),@Shop char(12),@Center char(12),@Day tinyint) AS
'   SELECT sum(WCCSHH1+WCCSHH2+WCCSHH3+WCCSHH4)
'   From WcclTable
'   WHERE WCCREF=@Month AND (WCCSHOP=@Shop AND WCCCENTER=@Center AND WCCDAY=@Day)
   sSql = "Qry_GetWorkCenterTimes '" & sMonth & "','" & sShop & "','" & sCenter & "'," _
          & iDay & " "
   Set RdoTme = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   'Set RdoTme = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not RdoTme.BOF And Not RdoTme.EOF Then
      With RdoTme
         If Not IsNull(.Fields(0)) Then
            GetCenterCalHours = .Fields(0)
         Else
            GetCenterCalHours = 0
         End If
         ClearResultSet RdoTme
      End With
   Else
      GetCenterCalHours = 0
   End If
   Set RdoTme = Nothing
   
   
End Function

'no calendar...try the workcenters
'6/2/00
'bGoodTime = GetCenterHours("CENTER", 2)
'Local errors

Public Function GetCenterHours(sCenter As String, iDayNum As Integer) As Currency
   Dim RdoTme As ADODB.Recordset
   Dim sWkDay As String
   Select Case iDayNum
      Case 2
         sWkDay = "MON"
      Case 3
         sWkDay = "TUE"
      Case 4
         sWkDay = "WED"
      Case 5
         sWkDay = "THU"
      Case 6
         sWkDay = "FRI"
      Case 7
         sWkDay = "SAT"
      Case Else
         sWkDay = "SUN"
   End Select
   
   sWkDay = "WCN" & sWkDay & "HR1+WCN" & sWkDay & "HR2+WCN" _
            & sWkDay & "HR3+WCN" & sWkDay & "HR4"
   sSql = "SELECT SUM(" & sWkDay & ") FROM WcntTable WHERE " _
          & "WCNREF='" & sCenter & "'"
   Set RdoTme = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   'Set RdoTme = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not RdoTme.BOF And Not RdoTme.EOF Then
      With RdoTme
         If Not IsNull(.Fields(0)) Then
            GetCenterHours = .Fields(0)
         Else
            GetCenterHours = 0
         End If
         ClearResultSet RdoTme
      End With
   Else
      GetCenterHours = 0
   End If
   Set RdoTme = Nothing
   
End Function


Public Function GetLastShiftTime(dtSchedule As Date, sCenter As String, iDayNum As Integer)
   Dim RdoTme As ADODB.Recordset
   Dim sWkDay As String
   Dim dt As Date
   Dim I As Integer
   Dim sStartHrs As String
   Dim iShiftHrs As Integer
   Select Case iDayNum
      Case 2
         sWkDay = "MON"
      Case 3
         sWkDay = "TUE"
      Case 4
         sWkDay = "WED"
      Case 5
         sWkDay = "THU"
      Case 6
         sWkDay = "FRI"
      Case 7
         sWkDay = "SAT"
      Case Else
         sWkDay = "SUN"
   End Select
   
   sWkDay = "WCN" & sWkDay & "SH1, WCN" & sWkDay & "HR1, WCN" _
            & sWkDay & "SH2, WCN" & sWkDay & "HR2, WCN" _
            & sWkDay & "SH3, WCN" & sWkDay & "HR3, WCN" _
            & sWkDay & "SH4, WCN" & sWkDay & "HR4"
            
   sSql = "SELECT " & sWkDay & " FROM WcntTable WHERE " _
          & "WCNREF='" & sCenter & "'"
    Set RdoTme = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
'   Set RdoTme = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not RdoTme.BOF And Not RdoTme.EOF Then
      With RdoTme
         For I = 3 To 0 Step -1 ' 3 shifts
            If Not IsNull(.Fields(I * 2)) Then
                sStartHrs = Trim(.Fields((I * 2)))
                iShiftHrs = Trim(.Fields((I * 2) + 1))
                If sStartHrs <> "" Then
                    Exit For
                End If
            End If
         Next
         
         If (iShiftHrs <> 0) Then
            dt = Format(dtSchedule, "mm/dd/yy " + sStartHrs)
            dtSchedule = DateAdd("h", iShiftHrs, dt)
         End If
         
         ClearResultSet RdoTme
      End With
   End If
   Set RdoTme = Nothing
   
End Function

Public Function GetLastShiftRevTime(dtCurSchedule As Date, dtSchedule As Date, sCenter As String, iDayNum As Integer)
   Dim RdoTme As ADODB.Recordset
   Dim sWkDay As String
   Dim dt As Date
   Dim I As Integer
   Dim sStartHrs As String
   Dim iShiftHrs As Integer
   Select Case iDayNum
      Case 2
         sWkDay = "MON"
      Case 3
         sWkDay = "TUE"
      Case 4
         sWkDay = "WED"
      Case 5
         sWkDay = "THU"
      Case 6
         sWkDay = "FRI"
      Case 7
         sWkDay = "SAT"
      Case Else
         sWkDay = "SUN"
   End Select
   
   iShiftHrs = 0
   sWkDay = "WCN" & sWkDay & "SH1, WCN" & sWkDay & "HR1, WCN" _
            & sWkDay & "SH2, WCN" & sWkDay & "HR2, WCN" _
            & sWkDay & "SH3, WCN" & sWkDay & "HR3, WCN" _
            & sWkDay & "SH4, WCN" & sWkDay & "HR4"
            
   sSql = "SELECT " & sWkDay & " FROM WcntTable WHERE " _
          & "WCNREF='" & sCenter & "'"
    Set RdoTme = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
'   Set RdoTme = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not RdoTme.BOF And Not RdoTme.EOF Then
      With RdoTme
         For I = 3 To 0 Step -1 ' 3 shifts
            If Not IsNull(.Fields(I * 2)) Then
                sStartHrs = Trim(.Fields((I * 2)))
                iShiftHrs = Trim(.Fields((I * 2) + 1))
                If sStartHrs <> "" Then
                    Exit For
                End If
            End If
         Next
         
         If (iShiftHrs <> 0) Then
            dt = Format(dtCurSchedule, "mm/dd/yy " + sStartHrs)
            dtSchedule = DateAdd("h", iShiftHrs, dt)
         Else
            dtSchedule = dtCurSchedule
         End If
         
         ClearResultSet RdoTme
      End With
   End If
   Set RdoTme = Nothing
   
End Function



Public Function GetThisCalendar(dSchedDt As Date, sShop As String, sCenter As String) As Boolean
   On Error Resume Next
   Dim sMonth As String
   sMonth = Format(dSchedDt, "mmm") & "-" & Format(dSchedDt, "yyyy")
   sSql = "SELECT WCCREF,WCCSHOP,WCCCENTER FROM WcclTable WHERE " _
          & "WCCREF='" & sMonth & "' AND (WCCSHOP='" & sShop & "' " _
          & "AND WCCCENTER='" & sCenter & "')"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.RowsAffected = 0 Then
      GetThisCalendar = False
   Else
      GetThisCalendar = True
   End If

End Function

'3/24/2009 Allows for Sat and Sun in schedules
Public Function GetValidWCCalSchDt(SchedDate As Date, sShop As String, sCenter As String) As Boolean
    Dim dNewDate As Date
    Dim iDay As Integer
    Dim I As Integer
    Dim cAvailableHrs As Integer
    On Error Resume Next
    dNewDate = Format(SchedDate, "mm/dd/yy hh:mm")

    ' Hardcoded for NOW, will be changed later
    For I = 1 To 10
         iDay = Format(dNewDate, "d")
        cAvailableHrs = GetCenterCalHours(dNewDate, sShop, sCenter, iDay)
        If (cAvailableHrs > 0) Then
            SchedDate = dNewDate
            Exit For
        Else
            dNewDate = Format(dNewDate - 1, vTimeFormat)
        End If
    Next
    ' For NOW
    If I = 10 Then
        GetValidWCCalSchDt = False
    Else
        GetValidWCCalSchDt = True
    End If
End Function


Public Function GetCalculatedMoCost(ByRef cUnitCost As Currency) As Boolean
   GetRunCosts
   Dim sReturnMsg As String
   sReturnMsg = CalculatedMoCost(PartNumber, RunNumber, cUnitCost, _
       RUNMATL, RunLabor, RUNEXP, RUNOVHD, RunHours)

   GetCalculatedMoCost = True
End Function

Private Function CalculatedMoCost(PartNumber As String, _
                     RunNumber As Long, _
                     ByRef ActualUnitCost As Currency, _
                     ByRef ActualTotMaterialCost As Currency, _
                     ByRef ActualTotLaborCost As Currency, _
                     ByRef ActualTotExpenseCost As Currency, _
                     ByRef ActualTotOverheadCost As Currency, _
                     ByRef ActualTotHours As Currency) As String
   Dim Quantity As Currency
   Dim iaCount As Long
   On Error GoTo whoops
   
   'get part parameters
   sSql = "SELECT PALOTTRACK, PAUSEACTUALCOST, PASTDCOST, PATOTMATL," & vbCrLf _
      & "PATOTLABOR, PATOTEXP, PATOTOH, PATOTHRS" & vbCrLf _
      & "from PartTable where PARTREF = '" & PartNumber & "' "
   
   Dim rdo As ADODB.Recordset
   Dim bLotTracked As Boolean, bUseActualCost As Boolean, usingStandardCost As Boolean
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
      
         ' 4/17/2011 - Included MO cancelled Qty (-) if there are any cancels.
         Dim RdoInv As ADODB.Recordset
         sSql = "SELECT COALESCE(max(cast(INLOTTRACK as tinyint))," & !PALOTTRACK & ", 0) as LOTTRACK," & vbCrLf _
            & "COALESCE(max(cast(INUSEACTUALCOST as tinyint))," & !PAUSEACTUALCOST & ", 0) as USEACTUALCOST, " & vbCrLf _
            & "isnull(sum(INAQTY),0) as TOTALQTY, count(*) as CT" & vbCrLf _
            & "from InvaTable" & vbCrLf _
            & "WHERE (INTYPE=" & IATYPE_MoCompletion & vbCrLf _
            & " OR INTYPE=" & IATYPE_CanceledMoCompletion & ") " & vbCrLf _
            & "AND INPART='" & PartNumber & "' " _
            & "AND INMORUN=" & RunNumber
            
   'Debug.Print sSql
         If clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD) Then
            bLotTracked = RdoInv!LOTTRACK
            bUseActualCost = RdoInv!useActualCost
            Quantity = RdoInv!TOTALQTY
            iaCount = RdoInv!ct
         
         'no completion quantity so no ia records
         Else
            bLotTracked = !PALOTTRACK
            bUseActualCost = !PAUSEACTUALCOST
         End If
         'RdoInv.Close
         Set RdoInv = Nothing
         
         'get quantity = sum of inaqty records for partial or full completions
         If Quantity = 0 Then
            ActualUnitCost = 0
         Else
            ActualUnitCost = (ActualTotMaterialCost + ActualTotLaborCost _
               + ActualTotExpenseCost + ActualTotOverheadCost) / Quantity
         End If
         
         'if using standard cost
         If Not bLotTracked Or (bLotTracked And Not bUseActualCost) Then
            usingStandardCost = True
            ActualUnitCost = !PASTDCOST
            ActualTotMaterialCost = !PATOTMATL * Quantity
            ActualTotLaborCost = !PATOTLABOR * Quantity
            ActualTotExpenseCost = !PATOTEXP * Quantity
            ActualTotOverheadCost = !PATOTOH * Quantity
            ActualTotHours = !PATOTHRS * Quantity
         Else
            usingStandardCost = False
         End If
      End With
   End If
   'rdo.Close
   Set rdo = Nothing
   Exit Function
   
whoops:
   sProcName = "CalculatedMoCost"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   CalculatedMoCost = "ClassMO.CalculatedMoCost failed"
   MsgBox CalculatedMoCost & " " & Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
End Function

'Private Function GetCancelledMOQty(strINPart As String, lRunMONum As Long) As Boolean
'   On Error Resume Next
'   Dim Qty As Currency
'
'   sSql = "SELECT WCCREF,WCCSHOP,WCCCENTER FROM WcclTable WHERE " _
'          & "WCCREF='" & sMonth & "' AND (WCCSHOP='" & sShop & "' " _
'          & "AND WCCCENTER='" & sCenter & "')"
'   clsAdoCon.ExecuteSQL sSql
'   If clsadocon.rowsaffected = 0 Then
'      GetThisCalendar = False
'   Else
'      GetThisCalendar = True
'   End If
'
'End Function


'3/24/2009 Allows for Sat and Sun in schedules
Public Function GetValidWCSchDt(SchedDate As Date, sCenter As String) As Boolean
    Dim dNewDate As Date
    Dim iWeekDay As Integer
    Dim I As Integer
    Dim cAvailableHrs As Integer
    On Error Resume Next
    dNewDate = Format(SchedDate, "mm/dd/yy hh:mm")

    ' Hardcoded for NOW, as there are 8 day in a week
    For I = 1 To 8
        iWeekDay = Format(dNewDate, "w")
        cAvailableHrs = GetCenterHours(sCenter, iWeekDay)
        If (cAvailableHrs > 0) Then
            SchedDate = dNewDate
            Exit For
        Else
            dNewDate = Format(dNewDate - 1, vTimeFormat)
        End If
    Next
    ' For NOW
    If I = 8 Then
        GetValidWCSchDt = False
    Else
        GetValidWCSchDt = True
    End If
End Function

Private Function RoundUpToInteger(cNumber As Currency) As Currency
   If cNumber >= 0 Then
      RoundUpToInteger = IIf(Fix(cNumber) = cNumber, cNumber, Fix(cNumber) + 1)
   Else
      RoundUpToInteger = IIf(Fix(cNumber) = cNumber, cNumber, Fix(cNumber) - 1)
   End If
End Function


Private Function WcHoursBeforeThisTime(vSchedDate As Date, sShop As String, sCenter As String) As Currency
   Dim rs As ADODB.Recordset
   Dim hours As Currency
   sSql = "select dbo.WCHoursBeforeTime( '" & sShop & "','" & sCenter & "','" & Format(vSchedDate, "mm/dd/yyyy hh:mm AM/PM") & "')"
   If clsADOCon.GetDataSet(sSql, rs, ES_FORWARD) Then
      'If Not rs.BOF And Not rs.EOF Then
      hours = rs.Fields(0)
      'Else
      '   hours = 0
      'End If
   End If
   Set rs = Nothing
   WcHoursBeforeThisTime = hours
End Function




