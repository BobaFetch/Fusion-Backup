Attribute VB_Name = "ESICOST"
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'************************************************************************************************
' ESICOST - ES/2002 Standard Costing Logic
'
' Notes:
'
' Created: 03/27/02 (nth)
' Revisons:
'   04/17/02 (nth) Changed code to multiply (part type 4) material by BOM qty
'   12/02/03 (nth) Removed the RRQ multiplier from part expense calcs.
'   06/14/04 (nth) Added the service unit cost to expenses.
'
'************************************************************************************************

Option Explicit

' Part costing data structure

Type PartCost
   nMaterial As Single
   nLabor As Single
   nExpense As Single
   nHours As Single
   nOverhead As Single
   bHasBom As Byte
   bHasRouting As Byte
End Type

' All 9 possible lower level cost.
' This global varible structure is populated by CostPartBOM
' and is access directly from the caller after IniStdCost and
' CostPartBOM is called and returned successful
Public BomCost(9) As PartCost

Public Function IniStdCost()
   Dim i As Integer
   
   For i = 0 To 9
      BomCost(i).bHasBom = 0
      BomCost(i).bHasRouting = 0
      BomCost(i).nExpense = 0
      BomCost(i).nHours = 0
      BomCost(i).nLabor = 0
      BomCost(i).nMaterial = 0
      BomCost(i).nOverhead = 0
   Next
End Function

Public Function CostPart(SPartRef As String, Optional bUpdateCost As Byte) As PartCost
   Dim RdoRtn As ADODB.Recordset
   Dim RdoPrt As ADODB.Recordset
   Dim RdoBom As ADODB.Recordset
   Dim nRRQ As Single
   Dim nRate As Single
   Dim nUnit As Single
   Dim nSetup As Single
   Dim nConversion As Single
   Dim nTemp As Single
   
   On Error GoTo modErr1
   sSql = "SELECT PARTREF, PARTNUM, PADESC, PALEVEL, PAREVDATE," _
          & "PAEXTDESC, PAMAKEBUY, PALEVLABOR, PALEVEXP, PALEVMATL, PALEVOH," _
          & "PALEVHRS, PASTDCOST, PABOMLABOR, PABOMEXP, PABOMMATL, PABOMOH," _
          & "PABOMHRS, PABOMREV, PAPREVLABOR, PAPREVEXP, PAPREVMATL, PAPREVOH," _
          & "PAPREVHRS, PATOTHRS, PATOTEXP, PATOTLABOR, PATOTMATL, PATOTOH,PAROUTING," _
          & "PARRQ,PAEOQ FROM PartTable WHERE PARTREF = '" & Trim(SPartRef) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   
   If bSqlRows Then
      With RdoPrt
         ' If RRQ is 0 then default it to 1
         If !PARRQ = 0 Then nRRQ = 1 Else nRRQ = !PARRQ
         
         ' Check is part has a BOM associated with it.
         ' If so cost material for this level.
         sSql = "SELECT DISTINCT BMASSYPART,BMCONVERSION,BMQTYREQD,BMSETUP," _
                & "BMADDER,PALEVEL,PASTDCOST FROM BmplTable,PartTable WHERE " _
                & "BMPARTNUM = PARTREF AND BMASSYPART = '" & SPartRef & "' " _
                & "AND BMREV = '" & "" & Trim(!PABOMREV) & "'"
         CostPart.bHasBom = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
         
         If CostPart.bHasBom Then
            CostPart.nMaterial = 0
            While Not RdoBom.EOF
               If RdoBom!PALEVEL = 4 Then
                  If RdoBom!BMCONVERSION < 1 Then
                     nConversion = 1
                  Else
                     nConversion = RdoBom!BMCONVERSION
                  End If
                  
                  ' ((Required Qty + Wasted Qty + (Setup Qty / Recomended Run Qty)) / Inventory Conversion Units) * UnitCost
                  
                  CostPart.nMaterial = CostPart.nMaterial + _
                                       ((RdoBom!BMQTYREQD + RdoBom!BMADDER) _
                                       + (RdoBom!BMSETUP / nRRQ) / nConversion) * RdoBom!PASTDCOST
               End If
               RdoBom.MoveNext
            Wend
         Else
            ' No BOM parts list found just use PALEVMATL
            CostPart.nMaterial = !PALEVMATL
         End If
         Set RdoBom = Nothing
         
         ' Check if the part has a routing. If so cost this level labor, overhead, and expense.
         If Len(Trim(!PAROUTING)) Then
            ' Operations, workcenters, and shops associated with part
            sSql = "SELECT OPREF,OPNO,OPSETUP,OPUNIT,WCNSTDRATE,WCNOHFIXED,SHPRATE," _
                   & "SHPOHTOTAL,WCNOHPCT,OPSERVPART,PASTDCOST,PAEOQ,OPSVCUNIT FROM RtopTable " _
                   & "INNER JOIN WcntTable ON RtopTable.OPCENTER = WcntTable.WCNREF " _
                   & "INNER JOIN ShopTable ON RtopTable.OPSHOP = ShopTable.SHPREF " _
                   & "LEFT JOIN PartTable On RtopTable.OPSERVPART = PartTable.PARTREF " _
                   & "WHERE RtopTable.OPREF = '" & Trim(!PAROUTING) & "'"
            CostPart.bHasRouting = clsADOCon.GetDataSet(sSql, RdoRtn)
            
            ' Loop through all op's
            If CostPart.bHasRouting Then
               nSetup = 0
               nUnit = 0
               While Not RdoRtn.EOF
                  
                  ' Labor Cost
                  If RdoRtn!WCNSTDRate > 0 Then
                     nRate = RdoRtn!WCNSTDRate
                  Else
                     ' If no work center assigned the use the shop rate
                     nRate = RdoRtn!SHPRATE
                  End If
                  'nSetup = (RdoRtn!opsetup * nRate)
                  'nUnit = ((RdoRtn!opunit * nRRQ) * nRate)
                  'CostPart.nLabor = CostPart.nLabor + ((nSetup + nUnit) / nRRQ)
                  
                  nTemp = ((RdoRtn!opsetup + (RdoRtn!opunit * nRRQ)) * nRate) / nRRQ
                  CostPart.nLabor = CostPart.nLabor + nTemp
                  
                  Debug.Print CostPart.nLabor
                  
                  
                  ' Overhead Cost
                  
                  If RdoRtn!WCNOHPCT = 0 Then
                     ' Overhead = (Overhead$ Hr) * (Setup + EOQ * Unit) / EOQ)
                     'nTemp =  * (RdoRtn!OPSETUP _
                     '    + (nRRQ * RdoRtn!OPUNIT) / nRRQ)
                     
                     nTemp = ((RdoRtn!opsetup + (RdoRtn!opunit * nRRQ)) * RdoRtn!WCNOHFIXED) / nRRQ
                     
                     CostPart.nOverhead = CostPart.nOverhead + nTemp
                  Else
                     ' Overhead = (Overhead %) * (((Setup * Rate) + (EOQ * Unit * Rate)) / EOQ)
                     'nTemp = (RdoRtn!WCNOHPCT / 100) * (((RdoRtn!opsetup _
                     '    * nRate) + (nRRQ * RdoRtn!opunit * nRate)) / nRRQ)
                     
                     nTemp = ((RdoRtn!opsetup + (RdoRtn!opunit * nRRQ)) * RdoRtn!WCNOHPCT) / nRRQ
                     
                     CostPart.nOverhead = CostPart.nOverhead + nTemp
                  End If
                  
                  
                  
                  
                  ' Expense Cost
                  'If IsNull(RdoRtn!PASTDCOST) Then
                  If IsNull(RdoRtn!OPSVCUNIT) Then
                     nTemp = 0
                  Else
                     nTemp = RdoRtn!OPSVCUNIT
                  End If
                  If Trim(RdoRtn!OPSERVPART) <> "" And !PARRQ <> 0 Then
                     CostPart.nExpense = CostPart.nExpense + (nTemp / !PARRQ)
                  End If
                  RdoRtn.MoveNext
               Wend
               
            End If
            Set RdoRtn = Nothing
         Else
            CostPart.nLabor = !PALEVLABOR
            CostPart.nOverhead = !PALEVOH
            CostPart.nExpense = !PALEVEXP
            CostPart.nHours = !PALEVHRS
         End If
         CostPart.nHours = !PALEVHRS
      End With
   End If
   Set RdoPrt = Nothing
   Exit Function
   
modErr1:
   sProcName = "CostPart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function


Public Sub CostPartBOM(sUsedOnPart As String, sRev As String, bLevel As Byte, Optional bUpdateCost As Byte)
   Dim RdoBom As ADODB.Recordset
   'Dim RdoPrt As ADODB.Recordset
   'Dim RdoRtn As ADODB.Recordset
   Dim rdoLowerRev As ADODB.Recordset
   Dim PartBOM As PartCost
   Dim sParent As String
   Dim sLowerRev As String
   Dim nTemp As Single
   
   On Error GoTo modErr1
   sSql = "SELECT * FROM BmplTable WHERE BMASSYPART = '" & sUsedOnPart _
          & "' AND BMREV = '" & sRev & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   
   If bSqlRows Then
      With RdoBom
         sParent = !BMASSYPART
         
         Do While Not .EOF
            If Err > 0 Or bLevel > 10 Then Exit Do
            
            ' Get Default BOM for current part in assembly
            sSql = "SELECT PABOMREV FROM PartTable WHERE PartRef = '" _
                   & Trim(!BMPARTREF) & "'"
            bSqlRows = clsADOCon.GetDataSet(sSql, rdoLowerRev)
            
            If bSqlRows Then
               sLowerRev = "" & Trim(rdoLowerRev!PABOMREV)
               Set rdoLowerRev = Nothing
               CostPartBOM Trim(!BMPARTREF), sLowerRev, bLevel + 1
            End If
            
            ' Cost current part
            PartBOM = CostPart(Trim(!BMPARTREF))
            
            BomCost(bLevel).nMaterial = BomCost(bLevel).nMaterial + PartBOM.nMaterial
            BomCost(bLevel).nExpense = BomCost(bLevel).nExpense + PartBOM.nExpense
            BomCost(bLevel).nLabor = BomCost(bLevel).nLabor + PartBOM.nLabor
            BomCost(bLevel).nHours = BomCost(bLevel).nHours + PartBOM.nHours
            BomCost(bLevel).nOverhead = BomCost(bLevel).nOverhead + PartBOM.nOverhead
            
            ' If part has a BOM then multiply cost by BOM qty required
            If PartBOM.bHasBom Then
               Debug.Print bLevel & " " & !BMPARTREF & PartBOM.nMaterial & Chr(9) & BomCost(bLevel).nMaterial
               
               
               nTemp = BomCost(bLevel).nMaterial + BomCost(bLevel).nExpense + BomCost(bLevel).nLabor + BomCost(bLevel).nOverhead
               BomCost(bLevel - 1).nMaterial = (nTemp * !BMQTYREQD)
               
               'BomCost.nMaterial = BomCost.nMaterial + PartBOM.nMaterial
               'BomCost.nExpense = BomCost.nExpense + PartBOM.nExpense
               'BomCost.nLabor = BomCost.nLabor + PartBOM.nLabor
               'BomCost.nHours = BomCost.nHours + PartBOM.nHours
               'BomCost.nOverhead = BomCost.nOverhead + PartBOM.nOverhead
               
               'PartBOM.nMaterial = PartBOM.nMaterial * !BMQTYREQD
               'PartBOM.nExpense = PartBOM.nExpense * !BMQTYREQD
               'PartBOM.nLabor = PartBOM.nLabor * !BMQTYREQD
               'PartBOM.nOverhead = PartBOM.nOverhead * !BMQTYREQD
               'PartBOM.nHours = PartBOM.nHours * !BMQTYREQD
               
            End If
            
            
            .MoveNext
         Loop
         Set rdoLowerRev = Nothing
      End With
   End If
   
   Set RdoBom = Nothing
   Exit Sub
   
modErr1:
   sProcName = "costpartbom"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Sub
