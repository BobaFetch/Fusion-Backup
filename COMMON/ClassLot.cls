VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ConsolidateLots(PartNumber As String, NewQuantity As Currency, _
                                userLotID As String)
   'Clear all existing lots and put new quantity, if any, into one new lot.
   'return = true if new lot created
   'return = false if new lot not created
   
   Dim RdoLot As ADODB.Recordset
   
   Dim iList As Integer
   Dim iLots As Integer
   Dim lCOUNTER As Long
   Dim lLOTRECORD As Long
   
   'Dim cActQty     As Currency
   Dim cInvQty As Currency
   Dim cLotQty As Currency
   Dim cItmQty As Currency
   
   Dim sLotNum As String
   Dim sPartNumber As String
   Dim vAdate As Variant
   
   Dim sLots(20000) As String
   
   ConsolidateLots = False
   On Error GoTo DiaErr1
   'Collect the lots
   'MouseCursor 11
   Erase sLots
   'cmdOrg.Enabled = False
   iLots = 0
   'cActQty = NewQuantity
   sPartNumber = Compress(PartNumber)
   vAdate = Format(GetServerDateTime(), "mm/dd/yy hh:mm")
   
   'get standard cost for part
   Dim useActualCost As Boolean, unitCost As Currency
   sSql = "SELECT PAUSEACTUALCOST, PASTDCOST from PartTable " _
          & "WHERE PARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         useActualCost = CBool(!PAUSEACTUALCOST)
         unitCost = !PASTDCOST
         ClearResultSet RdoLot
      End With
   End If
   Set RdoLot = Nothing
   
   'if using actual cost, calculate average from lots
   If useActualCost Then
      sSql = "select CAST( ISNULL( SUM( LOTUNITCOST *  LOTORIGINALQTY ), 0 )" & vbCrLf _
             & "/ case ISNULL( SUM( LOTORIGINALQTY ), 1 )" & vbCrLf _
             & "    when 0 then 1 else ISNULL( SUM( LOTORIGINALQTY ), 1 ) end" & vbCrLf _
             & "    as DECIMAL(15,4) ) as AVERAGECOST" & vbCrLf _
             & "from LohdTable where LOTPARTREF = '" & sPartNumber & "'" & vbCrLf
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
      If bSqlRows Then
         With RdoLot
            unitCost = !AverageCost
            ClearResultSet RdoLot
         End With
      End If
   End If
   Set RdoLot = Nothing
   
   'Get the Lots for this part
   sSql = "SELECT LOTNUMBER,LOTPARTREF FROM LohdTable " _
          & "WHERE LOTPARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         Do Until .EOF
            iLots = iLots + 1
            sLots(iLots) = "" & Trim(!lotNumber)
            .MoveNext
         Loop
         ClearResultSet RdoLot
      End With
   End If
   Set RdoLot = Nothing
   
   'If we have some then process them
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   On Error Resume Next
   For iList = 1 To iLots
   
      sSql = "SELECT SUM(LOIQUANTITY) FROM LoitTable " _
             & "WHERE LOINUMBER='" & sLots(iList) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_STATIC) ' Added Static cursor
      If bSqlRows Then
         With RdoLot
            Do Until .EOF
               If Not IsNull(.Fields(0)) Then
                  cLotQty = .Fields(0)
               Else
                  cLotQty = 0
               End If
               If cLotQty <> 0 Then
                  'Debug.Print "lot " & sLots(iList) & " qty " & cLotQty
                  sLotNum = sLots(iList)
                  lCOUNTER = GetLastActivity() + 1
                  lLOTRECORD = GetNextLotRecord(sLotNum)
                  
                  sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                         & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
                         & "LOIACTIVITY,LOICOMMENT) " _
                         & "VALUES('" _
                         & sLotNum & "'," & lLOTRECORD & ",19,'" & sPartNumber _
                         & "','" & vAdate & "'," & -cLotQty _
                         & "," & lCOUNTER & ",'" _
                         & "Manual Inventory Adjustment" & "')"
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "UPDATE LohdTable SET LOTREMAININGQTY=0," _
                      & "LOTAVAILABLE=0 WHERE LOTNUMBER='" & sLots(iList) & "'"
               clsADOCon.ExecuteSQL sSql
               'MsgBox iList
               DoEvents
               .MoveNext
            Loop
            ClearResultSet RdoLot
         End With
      End If
   Next
   Set RdoLot = Nothing
   
   'Processed the lot items.  Now the Part and associated rows
   sSql = "UPDATE PartTable SET PAQOH=" & NewQuantity & "," _
          & "PALOTQTYREMAINING=" & NewQuantity & " WHERE " _
          & "PARTREF='" & sPartNumber & "'"
   clsADOCon.ExecuteSQL sSql
   
   sLotNum = GetNextLotNumber()
   sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
          & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
          & "LOTUNITCOST,LOTDATECOSTED,LOTCOMMENTS) " _
          & "VALUES('" _
          & sLotNum & "','" & userLotID & "','" & sPartNumber _
          & "','" & vAdate & "'," & NewQuantity & "," & NewQuantity _
          & "," & unitCost & ",'" & vAdate & "','Manual Lot Re-Org')"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
          & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
          & "LOIACTIVITY,LOICOMMENT) " _
          & "VALUES('" _
          & sLotNum & "',1,19,'" & sPartNumber _
          & "','" & vAdate & "'," & NewQuantity _
          & "," & lCOUNTER & ",'" _
          & "Manual Manual Re-org" & "')"
   clsADOCon.ExecuteSQL sSql
   
   'Now to square Activity away
   sSql = "SELECT ISNULL(SUM(INAQTY), 0 ) FROM InvaTable WHERE INPART='" _
          & sPartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         'If Not IsNull(.Fields(0)) Then
         cInvQty = .Fields(0)
         'Else
         '    cInvQty = 0
         'End If
         lCOUNTER = GetLastActivity() + 1
         
         'less than zero
         '            If cInvQty < 0 Then
         '                cInvQty = Abs(cInvQty) + NewQuantity
         '                sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
         '                    & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
         '                & "VALUES(19,'" & sPartNumber & "','Manual Adjust Reorg','" & UserLotID & "'," _
         '                    & "'" & vAdate & "','" & vAdate & "'," & cInvQty _
         '                    & "," & cInvQty & ",0,'',''," & lCOUNTER & ",'" & sInitials & "')"
         '                clsAdoCon.ExecuteSQL sSql
         '
         '            ElseIf cInvQty = 0 Then
         '                cInvQty = NewQuantity
         '                sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
         '                    & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
         '                & "VALUES(19,'" & sPartNumber & "','Manual Adj Reorg','" & UserLotID & "'," _
         '                    & "'" & vAdate & "','" & vAdate & "'," & cInvQty _
         '                    & "," & cInvQty & ",0,'',''," & lCOUNTER & ",'" & sInitials & "')"
         '                clsAdoCon.ExecuteSQL sSql
         '
         '            Else
         '                'Greater than zero
         '                cInvQty = cInvQty - NewQuantity
         '                cInvQty = cInvQty - (2 * cInvQty)
         '                cInvQty = NewQuantity - cInvQty
         '                sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
         '                    & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
         '                & "VALUES(19,'" & sPartNumber & "','Manual Adj Reorg','" & UserLotID & "'," _
         '                    & "'" & vAdate & "','" & vAdate & "'," & cInvQty _
         '                    & "," & cInvQty & ",0,'',''," & lCOUNTER & ",'" & sInitials & "')"
         '                clsAdoCon.ExecuteSQL sSql
         '            End If
         
         'now create inventory activity record
         Dim adjustQty As Currency
         adjustQty = NewQuantity - cInvQty
         If adjustQty <> 0 Then
            sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
                   & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
                   & "VALUES(19,'" & sPartNumber & "','Manual Adjust Reorg','" & userLotID & "'," _
                   & "'" & vAdate & "','" & vAdate & "'," & cInvQty _
                   & "," & cInvQty & "," & unitCost & ",'',''," & lCOUNTER & ",'" & sInitials & "')"
            clsADOCon.ExecuteSQL sSql
         End If
      End With
   End If
   'MouseCursor 0
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      UpdateWipColumns lCOUNTER
      'SysMsg "Transaction Was Completed", True
      ConsolidateLots = True
   Else
      'MsgBox Err.Description
      'Debug.Print Err.Description
      'MsgBox "Could Not Complete The Transaction.", _
      vbExclamation, Caption
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
   End If
   'cmdOrg.Enabled = True
   'Sleep 1000
   Set RdoLot = Nothing
   'bGoodPart = GetPart()
   Exit Function
   
DiaErr1:
   sProcName = "ClassLots.ConsolidateLots"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
   
End Function

Public Function GetLotUnitCost(PartNumber As String, strLotNumber As String) As Currency
    On Error GoTo whoops
    
    Dim PartRef As String
    PartRef = Compress(PartNumber)
     
    sSql = "SELECT LOTUNITCOST FROM LohdTable WHERE " & _
                " LOTPARTREF = '" & PartRef & "' AND " & _
            " LOTNUMBER = '" & strLotNumber & "'"
    
    Dim rdo As ADODB.Recordset
    bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
    
    If bSqlRows Then
        With rdo
            GetLotUnitCost = !LotUnitCost
            ClearResultSet rdo
        End With
    Else
        GetLotUnitCost = "0.00"
    End If
    Set rdo = Nothing
whoops:
    
    
End Function
Public Function IsUserLotIdInUseForAnotherLot(UserLot As String, SystemLot As String, DisplayMsg As Boolean) As Boolean
   Dim rdo As ADODB.Recordset
   sSql = "SELECT COUNT(*) FROM LohdTable" & vbCrLf _
      & "WHERE LOTUSERLOTID='" & UserLot & "' AND LOTNUMBER <> '" & SystemLot & "'"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      If rdo.Fields(0) > 0 Then
         IsUserLotIdInUseForAnotherLot = True
         If DisplayMsg Then
            MsgBox "That User Lot ID Is In Use.", vbInformation, "Revise User Lot Number"
         End If
      End If
   End If
   Set rdo = Nothing
End Function

Public Function GetPickRestockLots(MOPart As String, MORUN As Long, PickPart As String) As ADODB.Recordset
   Dim rdo As ADODB.Recordset
   Dim sSql1  As String
   Dim sSql2 As String
   '4/25/2009 - Restock issue
'    sSql = "SELECT LOTUSERLOTID, LOTLOCATION, LOINUMBER as LOTNUMBER, " & vbCrLf _
'        & "- sum(LOIQUANTITY) as ReturnQtyAvail, LOIPDATE " & vbCrLf _
'        & "FROM Loittable, LohdTable " & vbCrLf _
'            & "WHERE Loittable.LOINUMBER = LohdTable.LOTNUMBER  AND " & vbCrLf _
'            & "Loittable.LOIMOPARTREF='" & Compress(MOPart) & "' AND " & vbCrLf _
'            & "Loittable.LOIMORUNNO='" & MORUN & "' " & vbCrLf _
'          & "AND Loittable.LOIPARTREF='" & Compress(PickPart) & "' " & vbCrLf _
'          & "AND Loittable.LOITYPE in (10) " & vbCrLf _
'          & " GROUP BY LOTUSERLOTID, LOTLOCATION, " & vbCrLf _
'                & " LOIPDATE, LOINUMBER having sum(LOIQUANTITY) < 0 " & vbCrLf _

   sSql1 = "select DISTINCT LOTUSERLOTID, LOTLOCATION, CONVERT(varchar(12),LOIPDATE, 101) LOIPDATE, x.*" & vbCrLf _
      & "from LoitTable loit " & vbCrLf _
      & "join LohdTable lohd on LOINUMBER = LOTNUMBER and LOITYPE = 10" & vbCrLf _
      & "join " & vbCrLf _
      & "(select LOINUMBER as LOTNUMBER, - sum(LOIQUANTITY) as ReturnQtyAvail" & vbCrLf _
      & "from Loittable" & vbCrLf
      
  sSql2 = " WHERE LOIMOPARTREF='" & Compress(MOPart) & "' and LOIMORUNNO=" & MORUN & vbCrLf _
      & "and LOIPARTREF='" & Compress(PickPart) & "'" & vbCrLf _
      & "and LOITYPE in (10,21)" & vbCrLf _
      & "group by LOINUMBER having sum(LOIQUANTITY) < 0 )" & vbCrLf _
      & "as x" & vbCrLf _
      & "on x.LOTNUMBER = lohd.LOTNUMBER " & vbCrLf _
      & "WHERE (loit.LOIMOPARTREF = '" & Compress(MOPart) & "' AND  LOIMORUNNO=" & MORUN & ")" & vbCrLf
      
   sSql = sSql1 & sSql2
      '& "WHERE (loit.LOIMOPARTREF = 'RDOEMTANDEMBOX')" & vbCrLf
      
      
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   Set GetPickRestockLots = rdo
   
End Function

Public Sub AllocatePickRestocks(MOPart As String, MORUN As Long, PickPart As String, RestockQty As Currency)
   'generates lot restock information for picked parts to lots from whence they came
   'returns Es_TotalLots lots in the Lots global array used by lot selection screens
   'use this for non-lot-tracked parts
   
   Dim rdo As ADODB.Recordset
   Set rdo = GetPickRestockLots(MOPart, MORUN, PickPart)

   Dim qtyToRestock As Currency
   qtyToRestock = RestockQty
   Erase lots
   Es_TotalLots = 0
   Dim dateString As String
   dateString = Format(Now, "mm/dd/yy")
   'For irow = 1 To iTotalItems
   With rdo
      Do Until (.EOF Or qtyToRestock <= 0)
'            cReStock = !ReturnQtyAvail
'            lblLots(bByte) = "" & Trim(!LotNumber)
'            lblLotu(bByte) = "" & Trim(!LOTUSERLOTID)
'            lblLotLoc(bByte) = "" & Trim(!LOTLOCATION)
'            lblQty(bByte) = Format(cReStock, ES_QuantityDataFormat)
'            lblDate(bByte) = Format(!LOIPDATE, "mm/dd/yy")
         Es_TotalLots = Es_TotalLots + 1
         ReDim Preserve lots(Es_TotalLots)
         lots(Es_TotalLots).LotSysId = "" & Trim(!lotNumber)
         lots(Es_TotalLots).LotUserId = "" & Trim(!LOTUSERLOTID)
         lots(Es_TotalLots).LotPartRef = Compress(PickPart)
         lots(Es_TotalLots).LotADate = dateString
         
         If !ReturnQtyAvail < qtyToRestock Then
            lots(Es_TotalLots).LotSelQty = !ReturnQtyAvail
            qtyToRestock = qtyToRestock - !ReturnQtyAvail
         Else
            lots(Es_TotalLots).LotSelQty = qtyToRestock
            qtyToRestock = 0
         End If
      Loop
   End With
   Es_LotsSelected = 1
   Set rdo = Nothing
End Sub

Public Function AutoAllocateLots(part As String, QtyRequired As Currency) As Boolean
   'uses fifo or lifo as required and allocates lots
   'results returned in lots array (type LotsAvailable)
   'lots are in lots(1) to to lots(Es_TotalLots)
   'return = False if insufficient quantity in lots
      
   Dim rdo As ADODB.Recordset
   Dim qty As Currency
   Erase lots
   Es_TotalLots = 0
   qty = QtyRequired
   
   If UsingFifo Then
     sSql = "SELECT LOTNUMBER, LOTREMAININGQTY" & vbCrLf _
      & "FROM LohdTable WHERE LOTPARTREF='" & part & "'" & vbCrLf _
      & "AND LOTREMAININGQTY > 0 AND LOTAVAILABLE = 1 ORDER BY LOTNUMBER ASC" & vbCrLf
   Else
     sSql = "SELECT LOTNUMBER, LOTREMAININGQTY" & vbCrLf _
      & "FROM LohdTable WHERE LOTPARTREF='" & part & "'" & vbCrLf _
      & "AND LOTREMAININGQTY > 0 AND LOTAVAILABLE = 1 ORDER BY LOTNUMBER DESC" & vbCrLf
   End If
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         Do Until .EOF
            If qty <= 0 Then Exit Do
            Es_TotalLots = Es_TotalLots + 1
            ReDim Preserve lots(Es_TotalLots)
            lots(Es_TotalLots).LotSysId = "" & Trim(!lotNumber)
            If !LOTREMAININGQTY >= qty Then
               lots(Es_TotalLots).LotSelQty = qty
               qty = 0
               Exit Do
            Else
               lots(Es_TotalLots).LotSelQty = !LOTREMAININGQTY
               qty = qty = !LOTREMAININGQTY '!LOTREMAINGQTY
            End If
            .MoveNext
         Loop
         ClearResultSet rdo
      End With
   End If
   If qty <= 0 Then
      AutoAllocateLots = True
   Else
      Es_TotalLots = 0
      Erase lots
   End If
   Set rdo = Nothing
End Function
