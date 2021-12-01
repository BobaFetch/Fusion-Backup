VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private PartNum As String
Private PartRef As String
Private nPartLevel As Byte
Private sProductCode As String
Private sUnitPrice As String

'std costs
Private stdMaterial As Currency
Private stdLabor As Currency
Private stdExpense As Currency
Private stdOverHead  As Currency
Private stdHours  As Currency

Private Location As String

Public Enum ePartType
   PARTTYPE_1_TopAssembly = 1
   PARTTYPE_2_IntermediateAssembly = 2
   PARTTYPE_3_BasicAssembly = 3
   PARTTYPE_4_RawMaterial = 4
   PARTTYPE_5_MiscInventory = 5
   PARTTYPE_6_Expense = 6
   PARTTYPE_7_OutsideServices = 7
   PARTTYPE_8_TopAssembly = 8
End Enum

'Public Enum eMakeBuyEither
'   MAKEBYTEITHER_Make = "M"
'   MAKEBYTEITHER_Buy = "B"
'   MAKEBYTEITHER_Either = "E"
'End Enum

Public Property Let PartNumber(PartNo As String)
   If StrComp(Trim(PartNo), PartNum, vbTextCompare) <> 0 Then
      PartNum = Trim(PartNo)
      PartRef = Compress(PartNo)
      
      'get part parameters
      sSql = "select * from PartTable where PARTREF = '" & PartRef & "'"
      Dim rdo As ADODB.Recordset
      If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
         With rdo
            nPartLevel = rdo!PALEVEL
            sProductCode = Trim(rdo!PAPRODCODE)
            stdMaterial = !PATOTMATL
            stdLabor = !PATOTLABOR
            stdExpense = !PATOTEXP
            stdOverHead = !PATOTOH
            stdHours = !PATOTHRS
            Location = !PALOCATION
                        sUnitPrice = !PAPRICE
         End With
      End If
      'rdo.Close
      
    End If
    Set rdo = Nothing
End Property

Public Property Get PartNumber() As String
   PartNumber = PartNum
End Property

Public Property Get defaultLocation() As String
   defaultLocation = Location
End Property

Public Function GetPartUnitPrice(Optional bBookPrice As Boolean = False, Optional strBook As String = "") As String
   Dim strUnitPrice As String
   
   If (bBookPrice = True) Then
      GetBookPrice PartRef, strBook, strUnitPrice
   Else
      strUnitPrice = sUnitPrice
   End If
   
   GetPartUnitPrice = sUnitPrice
End Function

Public Sub PopulatePartCombo(cbo As ComboBox)
   'populate a part combo box (normally called from dropdown event)
   sSql = "select PARTNUM from PartTable where PARTREF like '" & Compress(cbo.Text) & "%'" & vbCrLf _
      & "order by PARTNUM"
   LoadComboBox cbo, -1
End Sub

Public Sub PopulatePartComboTest(cbo As ComboBox, ShowAll As Boolean)
   'populate a part combo box (normally called from dropdown event)
   'ShowAll = True to show an <ALL> entry.
   
   MouseCursor ccHourglass
   cbo.Enabled = False
   
   Dim startCharacters As String
   Dim leadingCharacters As String
   
   startCharacters = cbo
   leadingCharacters = cbo
   If leadingCharacters = "<ALL>" Then
      leadingCharacters = ""
   End If
   
   If ShowAll Then
      sSql = "SELECT '<ALL>' AS PARTNUM" & vbCrLf _
         & "UNION" & vbCrLf
   Else
      sSql = ""
   End If
   
   sSql = sSql & "SELECT RTRIM(PARTNUM) FROM PartTable" & vbCrLf
   If leadingCharacters <> "" Then
      sSql = sSql & "WHERE PARTREF like '" & leadingCharacters & "%'" & vbCrLf
   End If
   sSql = sSql & "ORDER BY PARTNUM"
   LoadComboBoxAndSelect cbo, leadingCharacters

   'if there are no matches, just show all parts regardless of wait time involved
   If (ShowAll And cbo.ListCount = 1) Or (Not ShowAll And cbo.ListCount = 0) Then
      sSql = "SELECT '<ALL>' AS PARTNUM" & vbCrLf _
         & "UNION" & vbCrLf
      sSql = sSql & "SELECT PARTNUM FROM PartTable" & vbCrLf _
         & "ORDER BY PARTNUM"
      LoadComboBoxAndSelect cbo, leadingCharacters
   End If
   
'   'position at first entry >= starting entry
'   Dim i As Long
'   If Len(leadingCharacters) = 0 Then
'      Exit Sub
'   End If
'   For i = 0 To cbo.ListCount - 1
'      If cbo.List(i) >= leadingCharacters Then
'         cbo.ListIndex = i
'         MouseCursor ccArrow
'         cbo.Enabled = True
'         Exit Sub
'      End If
'   Next
   MouseCursor ccArrow
   cbo.Enabled = True
   
End Sub

Public Function GetStdCosts(Material As Currency, Labor As Currency, _
   Expense As Currency, OverHead As Currency, Hours As Currency) As Byte
   
   Material = stdMaterial
   Labor = stdLabor
   Expense = stdExpense
   OverHead = stdOverHead
   Hours = stdHours
End Function

Public Function GetExtendedStdCosts(Quantity As Currency, Material As Currency, Labor As Currency, _
   Expense As Currency, OverHead As Currency, Hours As Currency) As Byte
   
   Material = Quantity * stdMaterial
   Labor = Quantity * stdLabor
   Expense = Quantity * stdExpense
   OverHead = Quantity * stdOverHead
   Hours = Quantity * stdHours
End Function

Public Sub CreateInitialLot(PartNo As String, UOM As String)
   'THIS MUST BE CALLED AS part of a transaction to create the initial lot for a part
   'after the part has been created.
   'a creation trigger creates the initial InvaTable entry (sadly without the lot)
   'and the initial BOM entry
   
   
   'bResponse = GetNewPartAccounts(sDebitAcct, sCreditAcct)
   Dim sDebitAcct As String, sCreditAcct As String
   Dim sLotNumber As String
   Dim innumber As Long
   Dim compressedPartNo As String
   Dim cCost As Currency                 'not defined at this time
   compressedPartNo = Compress(PartNo)
   
   Dim activityDate As String
   activityDate = Format$(Now, "mm/dd/yy")
   
   Dim part As New ClassPart
   part.GetPartAccounts PartNo, True, sDebitAcct, sCreditAcct
   
   '  Colin added a trigger for these inserts 5/25/99
   '  trigger was removed in R77.  Call CreateNewPart method in this object instead

   '        sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2) " _
   '            & "VALUES(1,'" & compressedPartNo & "','BEGINNING BALANCE','INITIAL BALANCE') "

   '        sSql = "INSERT INTO BmhdTable (BMHREF,BMHPARTNO,BMHPART) " _
   '            & "VALUES('" & compressedPartNo & "','" & cmbPrt & "','" & compressedPartNo & "') "

   innumber = GetLastActivity()    'get the number already created by the trigger

   'Create beginning lot
   sLotNumber = GetNextLotNumber()
   sSql = "INSERT INTO LohdTable (LOTNUMBER,LOTUSERLOTID," _
          & "LOTPARTREF,LOTPDATE,LOTORIGINALQTY,LOTREMAININGQTY," _
          & "LOTUNITCOST,LOTDATECOSTED,LOTCOMMENTS) " _
          & "VALUES('" _
          & sLotNumber & "','Beginning Balance-" & sLotNumber & "','" & compressedPartNo _
          & "','" & activityDate & "',0,0" _
          & "," & cCost & ",'" & activityDate & "','New Part Number')"
   ' Not need to ad empty lot record.
   ' MM clsADOCon.ExecuteSQL sSql

   sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
          & "LOITYPE,LOIPARTREF,LOIPDATE,LOIQUANTITY," _
          & "LOIACTIVITY,LOICOMMENT,LOIUNITS) " _
          & "VALUES('" _
          & sLotNumber & "',1,1,'" & compressedPartNo _
          & "','" & activityDate & "',0" _
          & "," & innumber & ",'" _
          & "Beginning Balance" & "','" & UOM & "')"
   ' MM clsADOCon.ExecuteSQL sSql

   sSql = "UPDATE InvaTable" & vbCrLf _
      & "SET INCREDITACCT='" & sDebitAcct & "'," _
      & "INDEBITACCT='" & sCreditAcct & "'," _
      & "INLOTNUMBER='" & sLotNumber & "'" & vbCrLf _
      & "WHERE INTYPE=1 AND INPART='" & compressedPartNo & "' "
   ' MM clsADOCon.ExecuteSQL sSql
End Sub

Public Sub GetPartAccounts(PartNumber As String, InventoryIn As Boolean, _
   DebitAccount As String, CreditAccount As String)
   'If InventoryIn = True then DebitAccount = Part's inventory account
   '                       and CreditAccount = Part's CGS (Material or Expense) account
   'If InventoryIn = False then DebitAccount = Part's CGS (Material or Expense) account
   '                       and CreditAccount =Part's inventory account
   '
   
   Dim rdoAct As ADODB.Recordset
   Dim compressedPart As String
   
   compressedPart = Compress(PartNumber)
   On Error GoTo modErr1
   
   DebitAccount = ""
   CreditAccount = ""
   
   sSql = "select dbo.fnGetPartCgsAccount ( '" & compressedPart & "' )"
   If clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD) Then
      If InventoryIn Then
         CreditAccount = "" & rdoAct.Fields(0)
      Else
         DebitAccount = "" & rdoAct.Fields(0)
      End If
   End If
   
   sSql = "select dbo.fnGetPartInvAccount ( '" & compressedPart & "' )"
   If clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD) Then
      If InventoryIn Then
         DebitAccount = "" & rdoAct.Fields(0)
      Else
         CreditAccount = "" & rdoAct.Fields(0)
      End If
   End If
   
   Set rdoAct = Nothing
   Exit Sub
   
modErr1:
sProcName = "GetPartAccounts"
CurrError.Number = Err.Number
CurrError.Description = Err.Description
DoModuleErrors MdiSect.ActiveForm

End Sub

Public Sub UpdateAllLotLocations()
   
   sSql = "select count(*)" & vbCrLf _
      & "from parttable part join lohdtable lot on part.partref = lot.lotpartref" & vbCrLf _
      & "Where lotlocation <> palocation"
   Dim differentLocations As Long
   Dim rdo As ADODB.Recordset
   If clsADOCon.GetDataSet(sSql, rdo) Then
      differentLocations = rdo.Fields(0)
   End If
   
   sSql = "select count(*) from lohdtable"
   Dim lots As Long
   If clsADOCon.GetDataSet(sSql, rdo) Then
      lots = rdo.Fields(0)
   End If
   
   If differentLocations = 0 Then
      MsgBox "The locations of all " & lots & " lots are current." & vbCrLf _
         & "No update required."
      Exit Sub
   End If
   
   If MsgBox("Update " & differentLocations & " of " & lots & " lot locations for all lots with locations <> locations " _
      & "defined in the part records?" & vbCrLf _
      & "If you use multiple locations for any part, this information will be lost.", vbQuestion + vbYesNo) <> vbYes Then
      Set rdo = Nothing
      Exit Sub
   End If
   
   MouseCursor ccHourglass
   
   sSql = "Update lohdtable" & vbCrLf _
      & "Set lotlocation = part.palocation" & vbCrLf _
      & "from parttable part join lohdtable lot on part.partref = lot.lotpartref" & vbCrLf _
      & "Where lotlocation <> palocation"
   
   clsADOCon.ExecuteSQL sSql
   
   MouseCursor ccDefault
   
   MsgBox clsADOCon.RowsAffected & " of " & lots & " lot locations were revised."
   Set rdo = Nothing
End Sub

Public Sub UpdatePartLotLocations(PartNumber As String, Location As String, _
   DisplayIfUpdateNotRequired As Boolean)
   
   sSql = "select count(*) from lohdtable" & vbCrLf _
      & "where lotpartref = '" & Compress(PartNumber) & "'" & vbCrLf _
      & "and lotlocation <> '" & Location & "'"
   Dim differentLocations As Long
   Dim rdo As ADODB.Recordset

   If clsADOCon.GetDataSet(sSql, rdo) Then
      differentLocations = rdo.Fields(0)
   End If
   
   sSql = "select count(*) from lohdtable" & vbCrLf _
      & "where lotpartref = '" & Compress(PartNumber) & "'"
   Dim lots As Long
   If clsADOCon.GetDataSet(sSql, rdo) Then
      lots = rdo.Fields(0)
   End If
   Set rdo = Nothing
   If differentLocations = 0 Then
      If DisplayIfUpdateNotRequired Then
         MsgBox "All " & lots & " lot locations for this part are current."
      End If
      Exit Sub
   Else
   
      If MsgBox("Update lot locations for " & differentLocations & " of " & lots _
         & " lots for part " & PartNumber & " to " & Location & "?", _
         vbQuestion + vbYesNo) <> vbYes Then
         Exit Sub
      End If
   End If
     
   sSql = "Update lohdtable" & vbCrLf _
      & "Set lotlocation = '" & Location & "'" & vbCrLf _
      & "Where lotlocation <> '" & Location & "'" & vbCrLf _
      & "and lotpartref = '" & Compress(PartNumber) & "'"
   
   clsADOCon.ExecuteSQL sSql
   MsgBox clsADOCon.RowsAffected & " of " & lots & " lot locations were updated."
   
End Sub

Public Function IsPartTypeChangeOK(part As String, NewType As Integer) As Integer
   'returns -1 if change to new part type is OK
   'displays error message and returns prior part type otherwise
   
   'get current type.  if same, just return
   Dim PartRef As String
   PartRef = Compress(part)
   Dim rdo As ADODB.Recordset
   Dim oldType As Integer
   sSql = "select PALEVEL from PartTable where PARTREF = '" & PartRef & "'"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      oldType = rdo!PALEVEL
   Else
      MsgBox "Part " & part & " not found."
      Set rdo = Nothing
      Exit Function
   End If
   
'   If OldType = NewType Then
'      IsPartTypeChangeOK = True
'   End If
'
   'part type must be less than or equal to any assembly on which it is used
   sSql = "select BMASSYPART, PALEVEL, PARTNUM from BmplTable" & vbCrLf _
      & "join PartTable on BMASSYPART = PARTREF" & vbCrLf _
      & "where BMPARTREF = '" & PartRef & "'" & vbCrLf _
      & "and PALEVEL > " & NewType
   Dim msg As String
   Dim ct As Integer
   If clsADOCon.GetDataSet(sSql, rdo) Then
      Do While Not rdo.EOF
         ct = ct + 1
         If ct <= 10 Then
            msg = msg & "Part type " & NewType & " cannot be on part list for type " & rdo!PALEVEL _
               & ": " & rdo!PartNum & vbCrLf
         ElseIf ct = 11 Then
            msg = msg & "..."
         End If
         rdo.MoveNext
      Loop
      MsgBox msg
      IsPartTypeChangeOK = oldType
      Set rdo = Nothing
      Exit Function
   End If
   
   'part type must be greater than or equal to any of its components
   sSql = "select BMASSYPART, PALEVEL, PARTNUM from BmplTable" & vbCrLf _
      & "join PartTable on BMPARTREF = PARTREF" & vbCrLf _
      & "where BMASSYPART = '" & PartRef & "'" & vbCrLf _
      & "and PALEVEL < " & NewType
   ct = 0
   If clsADOCon.GetDataSet(sSql, rdo) Then
      Do While Not rdo.EOF
         ct = ct + 1
         If ct <= 10 Then
            msg = msg & "Part type " & NewType & " cannot have component of type " & rdo!PALEVEL _
               & ": " & rdo!PartNum & vbCrLf
         ElseIf ct = 11 Then
            msg = msg & "..."
         End If
         rdo.MoveNext
      Loop
      MsgBox msg
      IsPartTypeChangeOK = oldType
      Set rdo = Nothing
      Exit Function
   End If
   Set rdo = Nothing
   IsPartTypeChangeOK = -1
   
End Function

Private Function DefaultPartCosting() As Byte
   Dim rdo As ADODB.Recordset
   
   On Error Resume Next
   DefaultPartCosting = "0"
   sSql = "SELECT DEFCOSTINGMETHOD FROM ComnTable WHERE COREF=1 "
   If clsADOCon.GetDataSet(sSql, rdo, ES_STATIC) Then DefaultPartCosting = "" & rdo!DEFCOSTINGMETHOD
   Set rdo = Nothing
End Function

Public Function CreateNewPart(PartNumber As String, partType As ePartType, _
   PartDescription As String, MakeBuyEither As String) As Boolean
   'create a new part record.
   'Call this from within a transaction
   'return = True if successful
   
   On Error GoTo whoops
   
   Dim PartRef As String
   PartRef = Compress(PartNumber)

   sSql = "INSERT INTO PartTable (PARTREF,PARTNUM,PALEVEL,PADESC,PAMAKEBUY,PADOCLISTREV, " _
      & "PAUSEACTUALCOST)" & vbCrLf _
      & "VALUES('" & PartRef & "','" & PartNumber & "'," & partType & "," & vbCrLf _
      & "'" & Replace(PartDescription, "'", "''") & "','" & MakeBuyEither & "','NONE'," & DefaultPartCosting & ")"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "INSERT INTO BmhdTable (BMHREF,BMHPARTNO,BMHPART)" & vbCrLf _
     & "VALUES('" & PartRef & "','" & PartNumber & "','" & PartRef & "')"
   clsADOCon.ExecuteSQL sSql
   
   CreateNewPart = True
   
whoops:
End Function


Public Function GetUseActualCost(PartNumber As String) As Boolean
    On Error GoTo whoops
    
    Dim PartRef As String
    PartRef = Compress(PartNumber)
     
    sSql = "SELECT PAUSEACTUALCOST FROM PartTable WHERE PARTREF = '" & PartRef & "'"
    Dim rdo As ADODB.Recordset
    bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
    
    If bSqlRows Then
        With rdo
            GetUseActualCost = CBool(!PAUSEACTUALCOST)
            ClearResultSet rdo
        End With
    Else
        GetUseActualCost = False
    End If
    Set rdo = Nothing
whoops:
    
        
End Function
