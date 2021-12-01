VERSION 1.0 CLASS
BEGIN
MultiUse = -1 'True
Persistable = 0 'NotPersistable
DataBindingBehavior = 0 'vbNone
DataSourceBehavior = 0 'vbNone
MTSTransactionMode = 0 'NotAnMTSObject
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
   
   Dim RdoLot As rdoResultset
   
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
   bSqlRows = GetDataSet(RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         useActualCost = CBool(!PAUSEACTUALCOST)
         unitCost = !PASTDCOST
         ClearResultSet RdoLot
      End With
   End If
   
   'if using actual cost, calculate average from lots
   If useActualCost Then
      sSql = "select CAST( ISNULL( SUM( LOTUNITCOST *  LOTORIGINALQTY ), 0 )" & vbCrLf _
             & "/ case ISNULL( SUM( LOTORIGINALQTY ), 1 )" & vbCrLf _
             & "    when 0 then 1 else ISNULL( SUM( LOTORIGINALQTY ), 1 ) end" & vbCrLf _
             & "    as DECIMAL(15,4) ) as AVERAGECOST" & vbCrLf _
             & "from LohdTable where LOTPARTREF = '" & sPartNumber & "'" & vbCrLf
      bSqlRows = GetDataSet(RdoLot, ES_FORWARD)
      If bSqlRows Then
         With RdoLot
            unitCost = !AverageCost
            ClearResultSet RdoLot
         End With
      End If
   End If
   
   'Get the Lots for this part
   sSql = "SELECT LOTNUMBER,LOTPARTREF FROM LohdTable " _
          & "WHERE LOTPARTREF='" & sPartNumber & "' "
   bSqlRows = GetDataSet(RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         Do Until .EOF
            iLots = iLots + 1
            sLots(iLots) = "" & Trim(!LOTNUMBER)
            .MoveNext
         Loop
         ClearResultSet RdoLot
      End With
   End If
   
   'If we have some then process them
   RdoCon.BeginTrans
   On Error Resume Next
   For iList = 1 To iLots
      sSql = "SELECT SUM(LOIQUANTITY) FROM LoitTable " _
             & "WHERE LOINUMBER='" & sLots(iList) & "'"
      bSqlRows = GetDataSet(RdoLot, ES_FORWARD)
      If bSqlRows Then
         With RdoLot
            Do Until .EOF
               If Not IsNull(.rdoColumns(0)) Then
                  cLotQty = .rdoColumns(0)
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
                         & "','" & vAdate & "'," & - cLotQty _
                         & "," & lCOUNTER & ",'" _
                         & "Manual Inventory Adjustment" & "')"
                  RdoCon.Execute sSql, rdExecDirect
               End If
               sSql = "UPDATE LohdTable SET LOTREMAININGQTY=0," _
                      & "LOTAVAILABLE=0 WHERE LOTNUMBER='" & sLots(iList) & "'"
               RdoCon.Execute sSql, rdExecDirect
               'MsgBox iList
               DoEvents
               .MoveNext
            Loop
            ClearResultSet RdoLot
         End With
      End If
   Next
   
   'Processed the lot items.  Now the Part and associated rows
   sSql = "UPDATE PartTable SET PAQOH=" & NewQuantity & "," _
          & "PALOTQTYREMAINING=" & NewQuantity & " WHERE " _
          & "PARTREF='" & sPartNumber & "'"
   RdoCon.Execute sSql, rdExecDirect
   
   sLotNum = GetNextLotNumber()
   sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
          & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
          & "LOTUNITCOST,LOTDATECOSTED,LOTCOMMENTS) " _
          & "VALUES('" _
          & sLotNum & "','" & userLotID & "','" & sPartNumber _
          & "','" & vAdate & "'," & NewQuantity & "," & NewQuantity _
          & "," & unitCost & ",'" & vAdate & "','Manual Lot Re-Org')"
   RdoCon.Execute sSql, rdExecDirect
   
   sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
          & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
          & "LOIACTIVITY,LOICOMMENT) " _
          & "VALUES('" _
          & sLotNum & "',1,19,'" & sPartNumber _
          & "','" & vAdate & "'," & NewQuantity _
          & "," & lCOUNTER & ",'" _
          & "Manual Manual Re-org" & "')"
   RdoCon.Execute sSql, rdExecDirect
   
   'Now to square Activity away
   sSql = "SELECT ISNULL(SUM(INAQTY), 0 ) FROM InvaTable WHERE INPART='" _
          & sPartNumber & "'"
   bSqlRows = GetDataSet(RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         'If Not IsNull(.rdoColumns(0)) Then
         cInvQty = .rdoColumns(0)
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
         '                RdoCon.Execute sSql, rdExecDirect
         '
         '            ElseIf cInvQty = 0 Then
         '                cInvQty = NewQuantity
         '                sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
         '                    & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
         '                & "VALUES(19,'" & sPartNumber & "','Manual Adj Reorg','" & UserLotID & "'," _
         '                    & "'" & vAdate & "','" & vAdate & "'," & cInvQty _
         '                    & "," & cInvQty & ",0,'',''," & lCOUNTER & ",'" & sInitials & "')"
         '                RdoCon.Execute sSql, rdExecDirect
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
         '                RdoCon.Execute sSql, rdExecDirect
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
            RdoCon.Execute sSql, rdExecDirect
         End If
      End With
   End If
   'MouseCursor 0
   If Err = 0 Then
      RdoCon.CommitTrans
      UpdateWipColumns lCOUNTER
      'SysMsg "Transaction Was Completed", True
      ConsolidateLots = True
   Else
      'MsgBox Err.Description
      Debug.Print Err.Description
      'MsgBox "Could Not Complete The Transaction.", _
      vbExclamation, Caption
      RdoCon.RollbackTrans
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
   DoModuleErrors Me
   
End Function
