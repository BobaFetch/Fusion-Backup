VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInventoryActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sPartNo As String
'Private SPartRef As String
'Private nRunNo As Integer
Private bPartHasBeenRead As Boolean
Private bLoggingEnabled As Boolean
Private iLogNumber As Integer
Private bDiagnoseOnly As Boolean 'True if items are not to be updated

Public Enum eIATypes
   IATYPE_BeginningBalance = 1
   IATYPE_ShippedSoItem = 3
   IATYPE_ReturnedSoItem = 4
   IATYPE_CanceledSoItem = 5
   IATYPE_MoCompletion = 6
   IATYPE_MoClosed = 7
   IATYPE_PickOpenItem = 9    'Pick Only - not used in InvaTable
   IATYPE_PickedItem = 10
   IATYPE_PickCancelled = 12
   IATYPE_PoOpenItem = 14
   IATYPE_PoReceipt = 15
   IATYPE_PoCanceledItem = 16
   IATYPE_PoInvoiced = 17
   IATYPE_ManualAdjustment = 19
   IATYPE_PickRestock = 21
   IATYPE_PickScrap = 22
   IATYPE_PickSubstitution = 23
   IATYPE_PackingSlip = 25
   IATYPE_ABCAdjustment = 30
   IATYPE_InvTransfer = 32
   IATYPE_CancPackSlip = 33
   IATYPE_CanceledMoCompletion = 38
   IATYPE_PoRejectItem = 39
End Enum

Public Property Let PartNumber(PartNo As String)
   If StrComp(PartNo, sPartNo, vbTextCompare) <> 0 Then
      sPartNo = PartNo
      'SPartRef = Compress(sPartNo)
      bPartHasBeenRead = False
   End If
End Property

Public Property Get PartNumber() As String
   PartNumber = sPartNo
End Property

Public Property Let DiagnoseOnly(Value As Boolean)
   bDiagnoseOnly = Value
End Property

Public Property Get DiagnoseOnly() As Boolean
   DiagnoseOnly = bDiagnoseOnly
End Property

'Public Property Let RunNumber(RunNo As Integer)
'   If nRunNo <> RunNo Then
'      nRunNo = RunNo
'      bPartHasBeenRead = False
'   End If
'End Property
'
'Public Property Get RunNumber() As Integer
'   RunNumber = nRunNo
'End Property

Public Property Let LoggingEnabled(Value As Boolean)
   bLoggingEnabled = Value
   If Value Then
      TruncateReportLog
   End If
End Property

Public Property Get LoggingEnabled() As Boolean
   LoggingEnabled = bLoggingEnabled
End Property

Private Sub Class_Initialize()
   sPartNo = ""
   'SPartRef = ""
   'nRunNo = 0
   bPartHasBeenRead = False
End Sub

'Public Function UpdateReceiptCosts(PartNumber As String, _
'   PoNo As Long, _
'   ItemNo As Integer, _
'   RevisionLetter As String, _
'   Quantity As Currency, _
'   ByVal ActualUnitCost As Currency, _
'   ShowCostedInfo As Boolean, _
'   Optional UpdateIaEvenIfJournalClosed As Boolean) As String
'
'   'determine whether to use standard or actual costs and update corresponding
'   'receipt inventory record.  Costs for lot header and allocated pick
'   'inventory activity records (type 10) are also updated.
'   'This will be part of the enclosing transaction.
'   '
'   'ShowCostedInfo = True to include LOTDATECOSTED
'   '
'   'return =   "Used standard cost in inventory activity record"
'   '           or "Used actual cost in inventory activity record"
'   '           or "UpdateReceiptCosts failed"
'
'   Dim iaCount As Long
'   On Error GoTo whoops
'   Dim poReference As String
'   poReference = RefStringPoItem(PoNo, ItemNo, RevisionLetter)
'
'   UpdateReceiptCosts = "UpdateReceiptCosts failed"
'
'   Dim totalMaterialCost As Currency, totalExpenseCost As Currency
'   Dim partType As Integer
'   Dim rdo As ADODB.Recordset
'   Dim bLotTracked As Boolean, bUseActualCost As Boolean
'   Dim sLotNumber As String
'   Dim unitMaterialCost As Currency, unitExpenseCost As Currency
'
'   Log "PO " & PoNo & " item " & ItemNo & RevisionLetter & " part " & PartNumber
'
'   'default costs (should be revised below)
'   unitMaterialCost = ActualUnitCost
'   totalMaterialCost = Quantity * ActualUnitCost
'   unitExpenseCost = 0
'   totalExpenseCost = 0
'
'   sSql = "SELECT COALESCE(INLOTTRACK, PALOTTRACK, 0) as LOTTRACK," & vbCrLf _
'      & "COALESCE(INUSEACTUALCOST, PAUSEACTUALCOST, 0) as USEACTUALCOST, " & vbCrLf _
'      & "PASTDCOST, PATOTMATL, PATOTLABOR, PATOTEXP, PATOTOH, PATOTHRS, PALEVEL," & vbCrLf _
'      & "INLOTNUMBER" & vbCrLf _
'      & "from InvaTable join PartTable on INPART = PARTREF" & vbCrLf _
'      & "WHERE INTYPE=" & IATYPE_PoReceipt & vbCrLf _
'      & "AND INREF2='" & poReference & "' "
'   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
'      With rdo
'         bLotTracked = !LOTTRACK
'         bUseActualCost = !useActualCost
'         partType = !PALEVEL
'         sLotNumber = !INLOTNUMBER
'         ' 4/1/2009 The logic when Actual Cost is selected then use Actual Cost
''         If Not bLotTracked Or (bLotTracked And Not bUseActualCost) Then
''            ActualUnitCost = !PASTDCOST
''            UpdateReceiptCosts = "Used standard costs"
''         Else
''            UpdateReceiptCosts = "Used actual costs"
''         End If
'
'         If Not bUseActualCost Then
'            ActualUnitCost = !PASTDCOST
'            UpdateReceiptCosts = "Used standard costs"
'         Else
'            UpdateReceiptCosts = "Used actual costs"
'         End If
'
'
'         If partType <= 4 Then
'            unitMaterialCost = ActualUnitCost
'            totalMaterialCost = Quantity * ActualUnitCost
'            unitExpenseCost = 0
'            totalExpenseCost = 0
'         Else
'            unitMaterialCost = 0
'            totalMaterialCost = 0
'            unitExpenseCost = ActualUnitCost
'            totalExpenseCost = Quantity * ActualUnitCost
'         End If
'
'      End With
'   End If
'   'rdo.Close
'
'   Set rdo = Nothing
'
'   If bLoggingEnabled Then
'      Log "    Unit cost=" & ActualUnitCost & " matl=" & totalMaterialCost _
'         & " exp=" & totalExpenseCost & " lots=" & bLotTracked & " actual=" & bUseActualCost
'   End If
'
'   If Not bDiagnoseOnly Then
'
'      'update costs in inventory activity record
'      sSql = "UPDATE InvaTable SET INAMT=" & ActualUnitCost & ", " _
'         & "INTOTMATL=" & totalMaterialCost & "," & vbCrLf _
'         & "INTOTEXP=" & totalExpenseCost & "," & vbCrLf _
'         & "INCOSTEDBY='" & sInitials & "'," & vbCrLf _
'         & "INMAINTCOSTED=1" & vbCrLf _
'         & "WHERE INTYPE=" & IATYPE_PoReceipt & " AND INREF2='" & poReference & "'" & vbCrLf
'      If Not UpdateIaEvenIfJournalClosed Then
'         sSql = sSql & "AND INGLPOSTED = 0"
'      End If
'      clsADOCon.ExecuteSQL sSql
'
'      'update corresponding lot record
'      If sLotNumber <> "" Then
'         Dim vAdate As Variant
'         vAdate = GetServerDateTime
'
'         sSql = "Update LohdTable" & vbCrLf _
'            & "set "
'         If ShowCostedInfo Then
'            sSql = sSql & "LOTDATECOSTED='" & vAdate & "'," & vbCrLf
'         End If
'         sSql = sSql & "LOTUNITCOST=" & ActualUnitCost & "," & vbCrLf _
'            & "LOTTOTMATL=" & totalMaterialCost & "," & vbCrLf _
'            & "LOTTOTEXP=" & totalExpenseCost & "," & vbCrLf _
'            & "LOTMAINTCOSTED=1" & vbCrLf _
'            & "WHERE LOTNUMBER='" & sLotNumber & "' "
'
'         clsADOCon.ExecuteSQL sSql
'
''         'update allocated pick inventory activity records for this lot
''        CAN'T USE INLOTNUMBER
''         sSql = "UPDATE InvaTable SET " & vbCrLf _
''                & "INAMT=" & ActualUnitCost & ", " _
''                & "INTOTMATL=-" & unitMaterialCost & " * INAQTY," & vbCrLf _
''                & "INTOTEXP=-" & unitExpenseCost & " * INAQTY, " _
''                & "INLOTTRACK=" & bLotTracked * -1 & ", " _
''                & "INUSEACTUALCOST=" & bUseActualCost * -1 & "," & vbCrLf _
''                & "INCOSTEDBY='" & sInitials & "'," & vbCrLf _
''                & "INMAINTCOSTED=1" & vbCrLf _
''                & "WHERE INTYPE=" & IATYPE_PickedItem & " AND INLOTNUMBER='" & sLotNumber & "' "
''         clsAdoCon.ExecuteSQL sSql
'
'         'update allocated pick inventory activity records for this lot
'         'revised 12/3/08 to use INNUMBER rather than INLOTNUMBER (there can be
'         'multiple lots for a single pick ia, so INLOTNUMBER is not always correct.)
'         sSql = "UPDATE InvaTable SET " & vbCrLf _
'            & "INAMT=" & ActualUnitCost & ", " _
'            & "INTOTMATL=-" & unitMaterialCost & " * INAQTY," & vbCrLf _
'            & "INTOTEXP=-" & unitExpenseCost & " * INAQTY, " _
'            & "INLOTTRACK=" & bLotTracked * -1 & ", " _
'            & "INUSEACTUALCOST=" & bUseActualCost * -1 & "," & vbCrLf _
'            & "INCOSTEDBY='" & sInitials & "'," & vbCrLf _
'            & "INMAINTCOSTED=1" & vbCrLf _
'            & "FROM InvaTable ia" & vbCrLf _
'            & "join LoitTable li on li.LOIACTIVITY = ia.INNUMBER" & vbCrLf _
'            & "AND ia.INTYPE=" & IATYPE_PickedItem & " AND li.LOINUMBER='" & sLotNumber & "' "
'         If Not UpdateIaEvenIfJournalClosed Then
'            sSql = sSql & "AND INGLPOSTED = 0"
'         End If
'         clsADOCon.ExecuteSQL sSql
'
'      End If
'
'   End If
'
'   Exit Function
'
'whoops:
'   sProcName = "UpdateReceiptCosts"
'   CurrError.Number = Err
'   CurrError.Description = Err.Description
'   UpdateReceiptCosts = "ClassInventoryActivity.UpdateReceiptCosts failed: "
'   MsgBox UpdateReceiptCosts & Err.Description
'   Dim frm As New ClassErrorForm
'   DoModuleErrors frm
'End Function
'

Public Function UpdateReceiptCosts(PartNumber As String, _
   PoNo As Long, _
   ItemNo As Integer, _
   RevisionLetter As String, _
   Quantity As Currency, _
   ByVal ActualUnitCost As Currency, _
   ShowCostedInfo As Boolean, _
   Optional UpdateIaEvenIfJournalClosed As Boolean) As String
   
   'determine whether to use standard or actual costs and update corresponding
   'receipt inventory record.  Costs for lot header and allocated pick
   'inventory activity records (type 10) are also updated.
   'This will be part of the enclosing transaction.
   '
   'ShowCostedInfo = True to include LOTDATECOSTED
   '
   'return =   "Used standard cost in inventory activity record"
   '           or "Used actual cost in inventory activity record"
   '           or "UpdateReceiptCosts failed"
   '
   'Revised Oct 2016 by Terry to support purchase unit conversion for sheet inventory
   
   Dim iaCount As Long
   On Error GoTo whoops
   Dim poReference As String
   poReference = RefStringPoItem(PoNo, ItemNo, RevisionLetter)
   
   UpdateReceiptCosts = "UpdateReceiptCosts failed"
   
   Dim totalMaterialCost As Currency, totalExpenseCost As Currency
   Dim partType As Integer
   Dim rdo As ADODB.Recordset
   Dim bLotTracked As Boolean, bUseActualCost As Boolean
   Dim sLotNumber As String
   Dim unitMaterialCost As Currency, unitExpenseCost As Currency
   
   Log "PO " & PoNo & " item " & ItemNo & RevisionLetter & " part " & PartNumber
   
   'default costs (should be revised below)
   unitMaterialCost = ActualUnitCost
   totalMaterialCost = Quantity * ActualUnitCost
   unitExpenseCost = 0
   totalExpenseCost = 0
   
   sSql = "SELECT COALESCE(INLOTTRACK, PALOTTRACK, 0) as LOTTRACK," & vbCrLf _
      & "COALESCE(INUSEACTUALCOST, PAUSEACTUALCOST, 0) as USEACTUALCOST, " & vbCrLf _
      & "PASTDCOST, PATOTMATL, PATOTLABOR, PATOTEXP, PATOTOH, PATOTHRS, PALEVEL," & vbCrLf _
      & "INLOTNUMBER, PARTREF, INNO, INAQTY" & vbCrLf _
      & "from InvaTable join PartTable on INPART = PARTREF" & vbCrLf _
      & "WHERE INTYPE=" & IATYPE_PoReceipt & vbCrLf _
      & "AND INREF2='" & poReference & "' "
   Dim success As Boolean
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         ' these parameters do not change regardless of lot
         Dim SPartRef As String
         Dim cAdjUnitCost As Currency
         SPartRef = !PartRef
         bLotTracked = !LOTTRACK
         partType = !PALEVEL
         cAdjUnitCost = ActualUnitCost
         bUseActualCost = !useActualCost
         If Not bUseActualCost Then
            ActualUnitCost = !PASTDCOST
            UpdateReceiptCosts = "Used standard costs"
         Else
            UpdateReceiptCosts = "Used actual costs"
         End If
         
         Dim ActNo() As Long
         Dim lotNumber() As String
         Dim lotQty() As Currency
         Dim count As Integer
         Dim ActivityNo As Long
         Dim cAdjActQty As Currency
         count = 0
         
         Do Until .EOF
            'these parameters are lot-specific
            count = count + 1
            ReDim Preserve ActNo(count)
            ReDim Preserve lotNumber(count)
            ReDim Preserve lotQty(count)
'            Dim cAdjActQty As Currency
'            cAdjActQty = !INAQTY
            lotQty(count) = !INAQTY
'            sLotNumber = !INLOTNUMBER
'            ActivityNo = !INNO
            lotNumber(count) = !INLOTNUMBER
            ActNo(count) = !INNO
            rdo.MoveNext
         Loop
         
         Dim i As Integer
         For i = 1 To UBound(ActNo)
            cAdjActQty = lotQty(i)
            sLotNumber = lotNumber(i)
            ActivityNo = ActNo(i)
            
            ' if purchase units and inventory units are different:
            If partType <= 4 Then
               unitMaterialCost = cAdjUnitCost
               totalMaterialCost = cAdjActQty * cAdjUnitCost
               unitExpenseCost = 0
               totalExpenseCost = 0
            Else
               unitMaterialCost = 0
               totalMaterialCost = 0
               unitExpenseCost = cAdjUnitCost
               totalExpenseCost = cAdjActQty * cAdjUnitCost
            End If
            
   
            If bLoggingEnabled Then
               Log "    Unit cost=" & cAdjUnitCost & " matl=" & totalMaterialCost _
                  & " exp=" & totalExpenseCost & " lots=" & bLotTracked & " actual=" & bUseActualCost
            End If
            
            If Not bDiagnoseOnly Then
               
               'update costs in inventory activity record
               sSql = "UPDATE InvaTable SET INAMT=" & cAdjUnitCost & ", " _
                  & "INTOTMATL=" & totalMaterialCost & "," & vbCrLf _
                  & "INTOTEXP=" & totalExpenseCost & "," & vbCrLf _
                  & "INCOSTEDBY='" & sInitials & "'," & vbCrLf _
                  & "INMAINTCOSTED=1" & vbCrLf _
                  & "WHERE INTYPE=" & IATYPE_PoReceipt & " AND INREF2='" & poReference & "'" & vbCrLf _
                  & "AND INNO=" & CStr(ActivityNo) & vbCrLf
               If Not UpdateIaEvenIfJournalClosed Then
                  sSql = sSql & "AND INGLPOSTED = 0"
               End If
               clsADOCon.ExecuteSql sSql
               
               'update corresponding lot record
               If sLotNumber <> "" Then
                  Dim vAdate As Variant
                  vAdate = GetServerDateTime
                  
                  sSql = "Update LohdTable" & vbCrLf _
                     & "set "
                  If ShowCostedInfo Then
                     sSql = sSql & "LOTDATECOSTED='" & vAdate & "'," & vbCrLf
                  End If
                  sSql = sSql & "LOTUNITCOST=" & cAdjUnitCost & "," & vbCrLf _
                     & "LOTTOTMATL=" & totalMaterialCost & "," & vbCrLf _
                     & "LOTTOTEXP=" & totalExpenseCost & "," & vbCrLf _
                     & "LOTMAINTCOSTED=1" & vbCrLf _
                     & "WHERE LOTNUMBER='" & sLotNumber & "' "
                  
                  clsADOCon.ExecuteSql sSql
                  
                  'update allocated pick inventory activity records for this lot
                  'revised 12/3/08 to use INNUMBER rather than INLOTNUMBER (there can be
                  'multiple lots for a single pick ia, so INLOTNUMBER is not always correct.)
                  sSql = "UPDATE InvaTable SET " & vbCrLf _
                     & "INAMT=" & cAdjUnitCost & ", " _
                     & "INTOTMATL=-" & unitMaterialCost & " * INAQTY," & vbCrLf _
                     & "INTOTEXP=-" & unitExpenseCost & " * INAQTY, " _
                     & "INLOTTRACK=" & bLotTracked * -1 & ", " _
                     & "INUSEACTUALCOST=" & bUseActualCost * -1 & "," & vbCrLf _
                     & "INCOSTEDBY='" & sInitials & "'," & vbCrLf _
                     & "INMAINTCOSTED=1" & vbCrLf _
                     & "FROM InvaTable ia" & vbCrLf _
                     & "join LoitTable li on li.LOIACTIVITY = ia.INNUMBER" & vbCrLf _
                     & "AND ia.INTYPE=" & IATYPE_PickedItem & " AND li.LOINUMBER='" & sLotNumber & "' "
                  If Not UpdateIaEvenIfJournalClosed Then
                     sSql = sSql & "AND INGLPOSTED = 0"
                  End If
                  clsADOCon.ExecuteSql sSql
               
               End If 'diagnostic
         End If
      Next
      End With
      Set rdo = Nothing
   End If
   
   Exit Function
   
whoops:
   sProcName = "UpdateReceiptCosts"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   UpdateReceiptCosts = "ClassInventoryActivity.UpdateReceiptCosts failed: "
   MsgBox UpdateReceiptCosts & Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
End Function




Public Sub UpdatePackingSlipCosts(PsNumber As String, Optional UpdateIaEvenIfJournalClosed As Boolean)

   'update IA's for a packing slip to use the current costs of the lots
   'from which they are drawn
   Err.Clear
   On Error GoTo whoops

   Log "Update costs for PS " & PsNumber
   
'   sSql = "update InvaTable" & vbCrLf _
'      & "set INAMT = LOTUNITCOST," & vbCrLf _
'      & "INTOTMATL = cast ( abs( INAQTY ) * LOTTOTMATL / LOTORIGINALQTY as decimal(12,4))," & vbCrLf _
'      & "INTOTLABOR = cast ( abs( INAQTY ) * LOTTOTLABOR / LOTORIGINALQTY as decimal(12,4))," & vbCrLf _
'      & "INTOTEXP = cast ( abs( INAQTY ) * LOTTOTEXP / LOTORIGINALQTY as decimal(12,4)), " & vbCrLf _
'      & "INTOTOH = cast ( abs( INAQTY ) * LOTTOTOH / LOTORIGINALQTY as decimal(12,4))," & vbCrLf _
'      & "INTOTHRS = cast ( abs( INAQTY ) * LOTTOTHRS / LOTORIGINALQTY as decimal(12,4))" & vbCrLf _
'      & "from LoitTable " & vbCrLf _
'      & "join LohdTable ON LOINUMBER = LOTNUMBER" & vbCrLf _
'      & "join InvaTable ia2 ON INNUMBER = LOIACTIVITY" & vbCrLf _
'      & "where ia2.INPSNUMBER = '" & PsNumber & "'" & vbCrLf _
'      & "and LOTORIGINALQTY <> 0" & vbCrLf
''      & "and LOITYPE = " & IATYPE_PackingSlip & vbCrLf _

   sSql = "UpdatePackingSlipCosts '" & PsNumber & "', " & IIf(UpdateIaEvenIfJournalClosed, 1, 0)
   clsADOCon.ExecuteSql sSql
   Exit Sub

whoops:
   sProcName = "UpdatePackingSlipCosts"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   MsgBox "UpdatePackingSlipCosts Error: " & Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm

End Sub

Public Function RefStringPoItem(PoNo As Long, ItemNo As Integer, RevisionLetter As String) As String
   
   'Seatch for PO items with this key in INREF2
   'Note:  remove the leading zeros on the PO number when all places
   'generating it or using it have been found.
   
    RefStringPoItem = "PO " & Format(PoNo, "000000") & "-" _
                     & "ITEM " & ItemNo & RevisionLetter
   
End Function

Public Function IATypeName(IaType As eIATypes) As String
   Select Case IaType
   Case 1
      IATypeName = "Begining Balance"
   Case 3
      IATypeName = "Shipped SO Item"
   Case 4
      IATypeName = "Returned SOItem"
   Case 5
      IATypeName = "Canceled SO Item"
   Case IATYPE_MoCompletion:
      IATypeName = "MO Completion"
   Case IATYPE_MoClosed:
      IATypeName = "Closed MO"
   Case IATYPE_PickOpenItem:
      IATypeName = "Open Pick Item"
   Case IATYPE_PickedItem:
      IATypeName = "Picked Item"
   Case IATYPE_PickCancelled:
      IATypeName = "Canceled Pick Item"
   Case IATYPE_PoOpenItem:
      IATypeName = "Open PO Item"
   Case IATYPE_PoReceipt:
      IATypeName = "Received PO Item"
   Case IATYPE_PoCanceledItem:
      IATypeName = "Canceled PO Item"
   Case IATYPE_PoInvoiced:
      IATypeName = "Invoiced PO Item"
   Case IATYPE_ManualAdjustment:
      IATypeName = "Manual Adjustment"
   Case IATYPE_PickRestock:
      IATypeName = "Pick Restock"
   Case IATYPE_PickScrap:
      IATypeName = "Pick Scrap"
   Case IATYPE_PickSubstitution:
      IATypeName = "Pick Substitution"
   Case IATYPE_PackingSlip
      IATypeName = "Packing Slip"
   Case IATYPE_ABCAdjustment
      IATypeName = "ABC Cycle Count Adjustment"
   Case IATYPE_InvTransfer
      IATypeName = "Inventory Transfer"
   Case IATYPE_CancPackSlip
      IATypeName = "Canceled Packing Slip"
   Case IATYPE_CanceledMoCompletion:
      IATypeName = "Canceled MO Completion"
   Case Else
      IATypeName = "IA Type " & CInt(IaType)
   End Select
   
'                  Select Case !LOITYPE
'                     Case 1
'                        sActivity = "Begining Balance"
'                     Case 3
'                        sActivity = "Shipped Item"
'                     Case 4
'                        sActivity = "Returned Item"
'                     Case 5
'                        sActivity = "Canceled SO Item"
'                     Case 6
'                        sActivity = "Completed MO"
'                     Case 7
'                        sActivity = "Closed MO"
'                     Case 9
'                        sActivity = "Pick Request"
'                     Case 10
'                        sActivity = "Actual Pick"
'                     Case 11
'                        sActivity = "Pick On Dock"
'                     Case 12
'                        sActivity = "Canceled Pick Req"
'                     Case 13
'                        sActivity = "Pick Surplus"
'                     Case 14
'                        sActivity = "Open PO Item"
'                     Case 15
'                        sActivity = "PO Receipt"
'                     Case 16
'                        sActivity = "Canceled PO Item"
'                     Case 17
'                        sActivity = "Invoiced PO Item"
'                     Case 18
'                        sActivity = "On Dock"
'                     Case 19
'                        sActivity = "Manual Adjustment"
'                     Case 21
'                        sActivity = "Restocked Pick"
'                     Case 22
'                        sActivity = "Scrapped Pick Item"
'                     Case 23
'                        sActivity = "Pick Substitute"
'                     Case 21
'                        sActivity = "Pick From Freight"
'                     Case 38
'                        sActivity = "Canceled MO Completion"
'                     Case Else
'                        sActivity = "Undocumented Activity"
'                  End Select

'   cmbAct.AddItem "19 - Manual Adjustment"
'   cmbAct.AddItem "32 - Inventory Transfer"
'   cmbAct.AddItem "10 - Actual Pick"
'   cmbAct.AddItem "13 - Pick Surplus"
'   cmbAct.AddItem "21 - Restocked Item"
'   cmbAct.AddItem "22 - Scrapped Pick Item"
'   cmbAct.AddItem "23 - Pick Substitute"
'   cmbAct.AddItem "15 - PO Receipt"
'   cmbAct.AddItem "16 - Canceled PO Receipt"
'   cmbAct.AddItem "17 - Invoiced PO Item"
'   cmbAct.AddItem "03 - Shipped Item (No Packing Slip)"
'   cmbAct.AddItem "04 - Returned Item"
'   cmbAct.AddItem "25 - Packing Slip (Out)"
'   cmbAct.AddItem "33 - Canceled Packing Slip Item (In)"
'   cmbAct.AddItem "06 - Completed MO"
'   cmbAct.AddItem "07 - Closed MO"
'   cmbAct.AddItem "38 - Canceled MO Completion"
  
   
   
End Function

Public Sub Log(sText As String)

   DoEvents
   
   iLogNumber = iLogNumber + 1
   sText = Replace(sText, "'", "''")
   If Len(sText) > 80 Then sText = Left$(sText, 80)
   'Debug.Print iLogNumber & ": " & sText
   If LoggingEnabled Then
      sSql = "INSERT INTO EsReportClosedRunsLog (LOG_NUMBER,LOG_TEXT,LOG_PARTNO, LOG_RUNNO, LOG_CLOSED) " _
             & "VALUES( " & iLogNumber & ", '" & sText & "','" & sPartNo & "',0,0)"
      clsADOCon.ExecuteSql sSql
   End If
End Sub

Private Sub TruncateReportLog()
   sSql = "TRUNCATE TABLE EsReportClosedRunsLog"
   clsADOCon.ExecuteSql sSql
   iLogNumber = 0
End Sub


