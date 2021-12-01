VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moPartNo As String
Private moPartRef As String
Private moRunNo As Long
Private nPartLevel As Byte
Private sProductCode As String
Private sJournalID As String
Private iTrans As Integer
Private iRef As Integer
Private sWipLabAcct As String, sWipExpAcct As String, sWipMatAcct As String, sWipOhdAcct As String
Private RUN_nnnn As String
Private iSelIndex As Integer

Public Property Let MoPartNumber(PartNo As String)
   If StrComp(Trim(PartNo), moPartNo, vbTextCompare) <> 0 Then
      moPartNo = Trim(PartNo)
      moPartRef = Compress(PartNo)
      
      'get part parameters
      sSql = "select * from PartTable where PARTREF = '" & moPartRef & "'"
      Dim Ado As ADODB.Recordset
      If clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD) Then
         nPartLevel = Ado!PALEVEL
         sProductCode = Trim(Ado!PAPRODCODE)
      End If
      Set Ado = Nothing
      
      'get WIP accounts
      sWipLabAcct = GetLaborAcct(moPartRef, sProductCode, nPartLevel)
      sWipExpAcct = GetExpenseAcct(moPartRef, sProductCode, nPartLevel)
      sWipMatAcct = GetMaterialAcct(moPartRef, sProductCode, nPartLevel)
      sWipOhdAcct = GetOverHeadAcct(moPartRef, sProductCode, nPartLevel)
      
   End If
End Property

Public Property Get MoPartNumber() As String
   MoPartNumber = moPartNo
End Property

Public Property Let PickPartIndex(iIndex As Integer)
   iSelIndex = iIndex
End Property

Public Property Get PickPartIndex() As Integer
   PickPartIndex = iSelIndex
End Property


Public Property Let MoRunNumber(RunNumber As Long)
   If RunNumber <> moRunNo Then
      moRunNo = RunNumber
      Dim I As Integer
      I = 5 - Len(CStr(moRunNo))
      RUN_nnnn = "RUN" & Space$(I) & moRunNo
   End If
End Property

Public Property Get MoRunNumber() As Long
   MoRunNumber = moRunNo
End Property

Public Function PickPart(PartNumber As String, PickQty As Currency, RequiredQty As Currency, _
   UnitOfMeasure As String, unitCost As Currency, PickRecordNo As Long, wipLocation As String, _
   pickComplete As Boolean, Optional bSelIndex As Boolean = False) As Boolean

   'pick a part for an MO (Properties MoPartNumber and MoRunNumber)
   '12/10/08 new version with 1:1 relationship between InvaTable and LoitTable
   'THIS METHOD MUST BE CALLED FROM WITHIN A TRANSACTION
   'PartNumber = the part being picked
   'requires entries in TempPickLots table for any lot-tracked parts (see ClearLotSelections & AddLotSelection)
   'UnitOfMeasure = "?" to use part UOM (PAUNITS)
   'UnitCost = -1 to use PASTDCOST
   'PickRecordNo = -1 to create a new picked pick list item
   'return  = True if successful
   '        = False if failed
      
   On Error GoTo whoops
   
   Dim Ado As ADODB.Recordset
   Dim sDebitAcct As String, sCreditAcct As String
   Dim lCOUNTER As Long
   Dim vAdate As Variant
   Dim pickPartRef As String
   Dim I As Long
   Dim bResponse As Byte
   
   pickPartRef = Compress(PartNumber)
   vAdate = Format(GetServerDateTime(), "mm/dd/yy hh:mm")
   bResponse = GetPartAccounts(pickPartRef, sDebitAcct, sCreditAcct)
   'lCOUNTER = GetLastActivity() + 1
   
   'UnitOfMeasure = "?" to use part UOM (PAUNITS)
   'UnitCost = -1 to use PASTDCOST
   If UnitOfMeasure = "?" Or unitCost = -1 Then
      sSql = "select PAUNITS, PASTDCOST from PartTable where PARTREF = '" & pickPartRef & "'"
      If clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD) Then
         If UnitOfMeasure = "?" Then
            UnitOfMeasure = Ado!PAUNITS
         End If
         If unitCost = -1 Then
            unitCost = Ado!PASTDCOST
         End If
      End If
      Set Ado = Nothing
   End If
   
   'PickRecordNo = -1 to create a new picked pick list item
   If PickRecordNo = -1 Then
      PickRecordNo = AddUnpickedItem(PartNumber, PickQty, UnitOfMeasure)
   End If
   
   'get standard costs
   Dim part As New ClassPart
   Dim cMaterial As Currency, cLabor As Currency, cExpense As Currency, cOverhead As Currency, cHours As Currency
   Dim cExtendedCost As Currency
   part.PartNumber = PartNumber
   'Allocate selected lots
   Dim lLOTRECORD As Long
   Dim bUseActCost As Boolean
   
   lCOUNTER = GetLastActivity() + 1
'   sSql = "select * from TempPickLots" & vbCrLf _
'      & "where MoPartRef = '" & moPartRef & "' and " & vbCrLf _
'      & " MoRunNo = " & moRunNo & " and PickPartRef = '" & pickPartRef & "'" & vbCrLf _
'      & " AND LotQty = " & PickQty

   If (bSelIndex = True) Then
      sSql = "select * from TempPickLots" & vbCrLf _
         & "where MoPartRef = '" & moPartRef & "' and " & vbCrLf _
         & " MoRunNo = " & moRunNo & " and PickPartRef = '" & pickPartRef & "'" & vbCrLf _
         & " AND selIndex = " & iSelIndex
      ' MM need to find a way to fix for duplicate pick items.
   Else
      sSql = "select * from TempPickLots" & vbCrLf _
         & "where MoPartRef = '" & moPartRef & "' and " & vbCrLf _
         & " MoRunNo = " & moRunNo & " and PickPartRef = '" & pickPartRef & "'"
   End If

      If clsADOCon.GetDataSet(sSql, Ado, ES_STATIC) Then
      With Ado
         Do Until .EOF
         
'''''''' 12/10/08 do for each lot item
            
            part.GetExtendedStdCosts !lotQty, cMaterial, cLabor, cExpense, cOverhead, cHours
            
            ' Find if we need to use Actual Cost
            bUseActCost = False
            bUseActCost = part.GetUseActualCost(pickPartRef)
            If (bUseActCost) Then
                Dim clsLot As New ClassLot
                unitCost = clsLot.GetLotUnitCost(pickPartRef, !LotID)
                Set clsLot = Nothing
            End If
            ' Else use the standard cost
            cExtendedCost = !lotQty * unitCost
            
            sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE,INADATE,INPQTY,INAQTY,INAMT," & vbCrLf _
                   & "INCREDITACCT,INWIPMATACCT,INMOPART,INMORUN,INNUMBER,INUSER," & vbCrLf _
                   & "INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS,INLOTNUMBER)" & vbCrLf _
                   & "VALUES(" & IATYPE_PickedItem & ",'" & pickPartRef & "','PICK','" & moPartRef & " " & RUN_nnnn & "'," _
                   & "'" & vAdate & " ','" & vAdate & "'," & vbCrLf _
                   & "-" & !lotQty & "," & "-" & !lotQty & "," & unitCost & ",'" & sCreditAcct & "','" _
                   & sDebitAcct & "','" & moPartRef & "'," & moRunNo & "," & lCOUNTER _
                   & ",'" & sInitials & "'," _
                   & cMaterial & "," & cLabor & "," & cExpense & "," & cOverhead & "," & cHours & "," & vbCrLf _
                   & "'" & !LotID & "')"
            clsADOCon.ExecuteSql sSql
            'AverageCost pickPartRef
         
'''''''' 12/10/08 END do for each lot item
            
            lLOTRECORD = GetNextLotRecord(!LotID)
            sSql = "INSERT INTO LoitTable (LOINUMBER, LOIRECORD," _
               & "LOITYPE, LOIPARTREF, LOIQUANTITY," _
               & "LOIMOPARTREF, LOIMORUNNO, " _
               & "LOIACTIVITY, LOICOMMENT, LOIUNITS)" & vbCrLf _
               & "VALUES('" & !LotID & "'," _
               & lLOTRECORD & ", " & IATYPE_PickedItem & ", '" & pickPartRef & "', " _
               & "-" & !lotQty & ", '" & moPartRef & "', " & moRunNo & ", " _
               & lCOUNTER & ", 'MO Picked Item', '" & UnitOfMeasure & "')"
            clsADOCon.ExecuteSql sSql
            
            sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
                   & "-" & !lotQty & " WHERE LOTNUMBER='" & !LotID & "'"
            clsADOCon.ExecuteSql sSql
      
            lCOUNTER = lCOUNTER + 1
            .MoveNext
         Loop
         AverageCost pickPartRef
      End With
   End If
   
   'update pick record
   sSql = "UPDATE MopkTable" & vbCrLf _
      & "SET PKAQTY = " & PickQty & ", " & "PKTYPE = " & IATYPE_PickedItem & "," & vbCrLf _
      & "PKAMT = " & unitCost & ", PKADATE = '" & vAdate & "', " _
      & "PKWIP = '" & wipLocation & "'" & vbCrLf _
      & "WHERE PKPARTREF='" & pickPartRef & "' " _
      & "AND PKMOPART='" & moPartRef & "' AND PKMORUN=" & moRunNo & " " _
      & "AND PKRECORD=" & PickRecordNo & " "
   clsADOCon.ExecuteSql sSql
   
   'update QOH
   sSql = "UPDATE PartTable SET PAQOH=PAQOH - " & PickQty _
          & ",PALOTQTYREMAINING=PALOTQTYREMAINING - " & PickQty & vbCrLf _
          & "WHERE PARTREF='" & pickPartRef & "' "
   clsADOCon.ExecuteSql sSql
   
   'Journal entries
   If iTrans > 0 And cExtendedCost > 0 Then
      Dim sNewPart As String
      'Credit
      If Len(pickPartRef) > 20 Then
         sNewPart = Left(pickPartRef, 20)
      Else
         sNewPart = pickPartRef
      End If
   
      iRef = iRef + 1
      
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT,DCACCTNO," _
         & "DCDATE,DCDESCR,DCDESC,DCPARTNO,DCRUNNO)" & vbCrLf _
         & "VALUES('" _
         & sJournalID & "'," _
         & iTrans & "," _
         & iRef & "," _
         & cExtendedCost & ",'" _
         & sCreditAcct & "','" _
         & vAdate & "','" _
         & "Pick" & "','" _
         & sNewPart & "','" _
         & moPartRef & "'," _
         & moRunNo & ")"
      clsADOCon.ExecuteSql sSql
      
      'Debit
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT,DCACCTNO," _
         & "DCDATE,DCDESCR,DCDESC,DCPARTNO,DCRUNNO)" & vbCrLf _
         & "VALUES('" _
         & sJournalID & "'," _
         & iTrans & "," _
         & iRef & "," _
         & cExtendedCost & ",'" _
         & sDebitAcct & "','" _
         & vAdate & "','" _
         & "Pick" & "','" _
         & sNewPart & "','" _
         & moPartRef & "'," _
         & moRunNo & ")"
      clsADOCon.ExecuteSql sSql
   End If
      
   'if partial pick, add a pick item for the difference
   If PickQty < RequiredQty Then
   
      If Not pickComplete Then
         
         'get next PKRECORD number
         Dim iPkRecord As Long
         iPkRecord = GetNextPkRecord()
         
         'create new pick record
         Dim bomQty As Currency
         Dim opNo As Integer
         Dim sNewDate As String, sMsg As String, sComment As String
         sSql = "SELECT * FROM MopkTable WHERE PKRECORD=" & PickRecordNo & vbCrLf _
                & "AND PKMOPART='" & moPartRef & "' AND PKMORUN=" & moRunNo
         bSqlRows = clsADOCon.GetDataSet(sSql, Ado, ES_DYNAMIC)
         If bSqlRows Then
            bomQty = Ado!PKBOMQTY
            opNo = Ado!PKMORUNOP
            ' 11/26/2009 MM
            ' Adjust the Original Qty, BOM and PKPQty to the picked qty.
            Ado!PKPQTY = PickQty
            Ado!PKORIGQTY = PickQty
            Ado!PKBOMQTY = PickQty
            Ado.Update
            sMsg = "" & Ado!PKREFERENCE
            sComment = "" & Ado!PKCOMT
            If Not IsNull(Ado!PKPDATE) Then
               sNewDate = Format(Ado!PKPDATE, "mm/dd/yy")
            Else
               sNewDate = vAdate
            End If
            Ado.AddNew
            Ado!PKPARTREF = pickPartRef
            Ado!PKMOPART = moPartRef
            Ado!PKMORUN = moRunNo
            Ado!PKMORUNOP = opNo
            Ado!PKTYPE = IATYPE_PickOpenItem
            'ado!PKREV = sNewRev       not used consistently
            'ado!PKADATE = vAdate      should be null
            Ado!PKPDATE = sNewDate
            Ado!PKPQTY = RequiredQty - PickQty
            Ado!PKORIGQTY = RequiredQty - PickQty
            Ado!PKBOMQTY = bomQty - PickQty
            Ado!PKREFERENCE = sMsg
            Ado!PKCOMT = sComment
            Ado!PKRECORD = iPkRecord
            Ado.Update
         Else
            'this should never happen
            MsgBox "ERROR: No pick item found for " & moPartRef & " run " & moRunNo & " record " & PickRecordNo
         End If
         Set Ado = Nothing
         
      End If
   End If

   'update runstatus to partial pick (PP) or pick complete (PC)
   Dim sMOStatus As String
   sSql = "select count(*) from MopkTable" & vbCrLf _
      & "where PKAQTY=0 AND PKTYPE=" & IATYPE_PickOpenItem & vbCrLf _
      & "AND PKMOPART='" & moPartRef & "' AND PKMORUN=" & moRunNo
   
   sMOStatus = "PC"
   If clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD) Then
      If CLng(Ado.Fields(0)) > 0 Then
         sMOStatus = "PP"
      End If
   End If
   Set Ado = Nothing
   
   sSql = "UPDATE RunsTable SET RUNSTATUS='" & sMOStatus & "'," _
      & "RUNCMATL=RUNCMATL+" & cMaterial & "," _
      & "RUNCOST=RUNCOST+" & cMaterial & " " & vbCrLf _
      & "WHERE RUNREF='" & moPartRef & "' AND RUNNO=" & moRunNo & " "
   clsADOCon.ExecuteSql sSql
   
   UpdateWipColumns lCOUNTER
   PickPart = True
   Set Ado = Nothing
   Exit Function
      
whoops:
   sProcName = "ClassPick.PickPart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
   PickPart = False
End Function



Private Sub Class_Initialize()
   sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then sJournalID = ""
   If sJournalID <> "" Then iTrans = GetNextTransaction(sJournalID)
   iRef = 0
End Sub

Public Sub ClearLotSelections()
   sSql = "delete from TempPickLots" & vbCrLf _
      & "where ( MoPartRef = '" & moPartRef & "' and MoRunNo = " & moRunNo & " )" & vbCrLf _
      & "or DateDiff( hour, WhenCreated, getdate() ) > 24"
   clsADOCon.ExecuteSql sSql
End Sub

Public Sub AddLotSelection(PickedPartNumber As String, LotID As String, PickedQty As Currency)
   sSql = "INSERT INTO TempPickLots ( MoPartRef, MoRunNo, PickPartRef, LotID, LotQty )" & vbCrLf _
      & "Values ( '" & moPartRef & "', " & moRunNo & ", '" & Compress(PickedPartNumber) & "'," & vbCrLf _
      & "'" & LotID & "', " & PickedQty & " )"
   clsADOCon.ExecuteSql sSql
End Sub

Public Function GetNextPkRecord() As Long
   Dim Ado As ADODB.Recordset
   
   sSql = "select ISNULL(max(PKRECORD), 0) from MopkTable" & vbCrLf _
      & "where PKMOPART='" & moPartRef & "' AND PKMORUN=" & moRunNo & " "
   If clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD) Then
      GetNextPkRecord = CLng(Ado.Fields(0)) + 1
   Else
      GetNextPkRecord = 1
   End If
   Set Ado = Nothing
End Function

Public Function AddUnpickedItem(PartNumber As String, Quantity As Currency, UnitOfMeasure As String) As Long
   'returns PKRECORD
   
   On Error GoTo whoops
   
   'get next PKRECORD
   Dim iPkRecord As Long
   iPkRecord = GetNextPkRecord()
   
   'Dim bomQty As Currency
   'Dim opNo As Integer
   'Dim sNewDate As String, sMsg As String, sComment As String
   Dim Ado As ADODB.Recordset
   sSql = "SELECT * FROM MopkTable WHERE PKRECORD = -1" & vbCrLf _
          & "AND PKMOPART='" & moPartRef & "' AND PKMORUN=" & moRunNo
   bSqlRows = clsADOCon.GetDataSet(sSql, Ado, ES_DYNAMIC)
   Ado.AddNew
   Ado!PKPARTREF = Compress(PartNumber)
   Ado!PKMOPART = moPartRef
   Ado!PKMORUN = moRunNo
   Ado!PKMORUNOP = 0
   Ado!PKTYPE = IATYPE_PickOpenItem
   Ado!PKPDATE = Format(GetServerDateTime(), "mm/dd/yy")
   Ado!PKPQTY = Quantity
   Ado!PKORIGQTY = Quantity
   Ado!PKBOMQTY = Quantity
   Ado!PKREFERENCE = "Pre-Pick"
   Ado!PKCOMT = " "
   Ado!PKRECORD = iPkRecord
   Ado!PKUNITS = UnitOfMeasure
   Ado.Update
   Set Ado = Nothing
   AddUnpickedItem = iPkRecord
   Exit Function
   
whoops:
   sProcName = "ClassPick.AddUnpickedItem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
End Function

Public Function AddPickList(partRef As String, runNo As Integer, dtRunSched As Date, cRunqty As Currency) As Boolean
   Dim RdoLst As ADODB.Recordset
   
   Dim bGoodHeader As Byte
   Dim bGoodPl As Byte
   Dim bResponse As Byte
   Dim bOrphanedParts As Byte
   
   Dim iRow As Integer
   Dim iTotalItems As Integer
   Dim n As Integer
   
   Dim cQuantity As Currency
   Dim cConversion As Currency
   Dim cSetup As Currency
   'Dim cRunqty As Currency
   Dim sMsg As String
   Dim sBomRev As String
   Dim iPkRecord As Integer
   Dim Caption As String
   Caption = "Pick List Generation"
   Dim runSched As String
   runSched = Format(dtRunSched, "mm/dd/yy")
   
   On Error GoTo DiaErr2
   
   iPkRecord = 0
   sBomRev = GetRevisions(partRef)
      
   'determine whether any part list for this part and rev
   sSql = "SELECT BMASSYPART FROM BmplTable " & vbCrLf _
          & "WHERE BMASSYPART = '" & partRef & "'" & vbCrLf _
          & "AND BMREV = '" & sBomRev & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      bGoodPl = True
      ClearResultSet RdoLst
      Set RdoLst = Nothing
   Else
      MouseCursor 0
      bGoodPl = False
      MsgBox "part " & partRef & " does not have a parts list rev " & sBomRev, vbInformation, Caption
      AddPickList = False
      Exit Function
   End If

   sSql = "SELECT BMHREF,BMHREV,BMHOBSOLETE,BMHRELEASED,BMHEFFECTIVE " & vbCrLf _
          & "FROM BmhdTable" & vbCrLf _
          & "WHERE BMHREF='" & partRef & "' AND BMHREV='" & sBomRev & "' " & vbCrLf _
          & "AND (BMHOBSOLETE IS NULL OR BMHOBSOLETE >='" & runSched & "') AND BMHRELEASED=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      bGoodHeader = True
   Else
      bGoodHeader = False
   End If
   
   If Not bGoodHeader Then
      'oops the header is gone, date invalid or not released?
      MouseCursor 0
      MsgBox "The Parts List Is Not Valid, Released, " & vbCr _
         & "Or Outdated For This Part.", vbInformation, Caption
      AddPickList = False
      Exit Function
   End If
      
   sSql = "SELECT PARTREF,PAUNITS, * FROM BmplTable" & vbCrLf _
      & "LEFT OUTER JOIN PartTable ON PARTREF=BMPARTREF " & vbCrLf _
      & "WHERE BMASSYPART='" & partRef & "'" & vbCrLf _
      & "AND BMREV='" & sBomRev & "'" & vbCrLf _
      & "ORDER BY BMSEQUENCE"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_DYNAMIC)
   bOrphanedParts = 0
   If bSqlRows Then
      With RdoLst
         Do Until .EOF
            If Not IsNull(!BMSETUP) Then
               cSetup = !BMSETUP
            Else
               cSetup = 0
            End If
            
            If (SetupQtyEnabled = True) Then
               cQuantity = Format(((cRunqty + cSetup) * (!BMQTYREQD + !BMADDER)), "######0.000")
            Else
               cQuantity = Format(((cRunqty * (!BMQTYREQD + !BMADDER)) + cSetup), "######0.000")
            End If
            If !BMCONVERSION <> 0 Then
               cQuantity = cQuantity / !BMCONVERSION
            End If
   
'            'if phantom item, then explode it
'            If !BMPHANTOM = 1 Then
'               InsertPhantom Trim(!BMPARTREF), Trim(!BMPARTREV), cQuantity
'            Else
               iPkRecord = iPkRecord + 1
               sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
                      & "PKTYPE,PKPDATE,PKPQTY,PKBOMQTY,PKRECORD,PKUNITS," _
                      & "PKCOMT) VALUES('" & Trim(!BMPARTREF) & "','" _
                      & Compress(partRef) & "'," & CStr(runNo) & ",9,'" & runSched _
                      & "'," & cQuantity & "," & cQuantity & "," & iPkRecord & "," _
                      & "'" & Trim(!PAUNITS) & "','" & Trim(!BMCOMT) & "') "
               If Len(Trim(!partRef)) = 0 Then
                   bOrphanedParts = 1
               Else
                   clsADOCon.ExecuteSql sSql
               End If
               
'            End If
               
            .MoveNext
         Loop
         ClearResultSet RdoLst
         Set RdoLst = Nothing
         AddPickList = True
      End With
         
      If bOrphanedParts Then
        MsgBox "Pick List Added Successfully. However, your BOM Parts List has Orphaned Parts." & vbCrLf & "Please Contact Fusion Support"
      End If
   Else
      MouseCursor 0
      MsgBox "Couldn't Find Items For Revision " & sBomRev & ".", vbInformation, Caption
      AddPickList = False
   End If
   Set RdoLst = Nothing
   Exit Function
   
DiaErr2:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function GetRevisions(strPartRef As String) As String
   
   Dim RdoLst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREV FROM BmhdTable WHERE BMHREF='" _
          & Compress(strPartRef) & "' ORDER BY BMHREVDATE DESC"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      With RdoLst
         GetRevisions = "" & Trim(!BMHREV)
      End With
      
      ClearResultSet RdoLst
      Set RdoLst = Nothing
   Else
      GetRevisions = ""
   End If
   
   Exit Function
   
DiaErr1:
   sProcName = "getrevisions"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



