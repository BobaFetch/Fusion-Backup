Attribute VB_Name = "EsiInvc"
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***

Option Explicit
'Activity Codes
'MO Items

Public Y As Byte
Public bFoundPart As Byte
Public iAutoIncr As Integer

Public sCurrForm As String
Public sPassedRout As String
Public sPassedMo As String
Public sPassedPart As String
Public sSelected As String

Public sFavorites(13) As String
Public sRecent(10) As String
Public sSession(50) As String

Public Function GetActivityQuantity(PartNumber As String) As Currency
   Dim RdoSum As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT SUM(INAQTY) FROM InvaTable WHERE INPART='" _
          & PartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSum, ES_FORWARD)
   If bSqlRows Then
      With RdoSum
         If Not IsNull(.Fields(0)) Then
            GetActivityQuantity = .Fields(0)
         Else
            GetActivityQuantity = 0
         End If
         ClearResultSet RdoSum
      End With
   End If
   Set RdoSum = Nothing
   
End Function

'1/13/05 Adjusts the QOH of InvaTable to the Part Qoh
'Adjustment is PAQOH - SUM(INAQTY) for the Part

Public Sub RepairInventory(PartNumber As String, Adjustment As Currency, Optional InCounter As Long)
   Dim vAdate As Variant
   If Adjustment <> 0 Then
      If InCounter = 0 Then InCounter = (GetLastActivity() + 1)
      vAdate = Format(ES_SYSDATE, "mm/dd/yy hh:mm")
      sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
             & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
             & "VALUES(19,'" & PartNumber & "','Inventory Adjustment','RepairInventory'," _
             & "'" & vAdate & "','" & vAdate & "'," & Adjustment & "," & Adjustment & "," _
             & "0,'',''," & InCounter & ",'" & sInitials & "')"
      clsADOCon.ExecuteSql sSql
      UpdateWipColumns InCounter
   End If
   
End Sub

'10/7/04

Public Function GetPODataFormat() As String
   Dim RdoFormat As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT PurchasedDataFormat FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFormat, ES_FORWARD)
   If bSqlRows Then
      With RdoFormat
         If Not IsNull(!PurchasedDataFormat) Then
            GetPODataFormat = "" & Trim(!PurchasedDataFormat)
         Else
            GetPODataFormat = ES_QuantityDataFormat
         End If
         ClearResultSet RdoFormat
      End With
   End If
   If GetPODataFormat = "" Then GetPODataFormat = ES_QuantityDataFormat
   Set RdoFormat = Nothing
   Exit Function
   
modErr1:
   GetPODataFormat = ES_QuantityDataFormat
   
End Function

'3/5/02
'Procedure to go into mom
'Scrapped for Rel 2.0 3/25/02 See Procedure in MOM

Public Sub AddLotTracking()
   Exit Sub
  
End Sub

Public Sub FillBuyers()
   Dim RdoByr As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_GetBuyerList"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoByr, ES_FORWARD)
   If bSqlRows Then
      With RdoByr
         Do Until .EOF
            AddComboStr MdiSect.ActiveForm.cmbByr.hWnd, "" & Trim(!BYNUMBER)
            .MoveNext
         Loop
         ClearResultSet RdoByr
      End With
   End If
   Set RdoByr = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillbuyers"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Public Sub GetCurrentBuyer(sBuyer As String)
   Dim RdoByr As ADODB.Recordset
   sBuyer = UCase$(Compress(sBuyer))
   sSql = "SELECT BYNUMBER,BYLSTNAME,BYFSTNAME,BYMIDINIT FROM " _
          & "BuyrTable WHERE BYREF='" & sBuyer & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoByr, ES_FORWARD)
   If bSqlRows Then
      With RdoByr
         MdiSect.ActiveForm.cmbByr = "" & Trim(!BYNUMBER)
         MdiSect.ActiveForm.lblByr = "" & Trim(!BYFSTNAME) _
                                     & " " & Trim(!BYMIDINIT) & " " & Trim(!BYLSTNAME)
         ClearResultSet RdoByr
      End With
   Else
      If Len(Trim(sBuyer)) > 0 Then
         MdiSect.ActiveForm.lblByr = "*** Buyer Wasn't Found ***"
      Else
         MdiSect.ActiveForm.lblByr = ""
      End If
   End If
   Set RdoByr = Nothing
   Exit Sub
   
modErr1:
   sProcName = "getcurrentbuyer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Forms(1)
   
   
End Sub



'8/9/05 Close open sets

Public Sub FormUnload(Optional bDontShowForm As Byte)
   Dim iList As Integer
   Dim iResultSets As Integer
   On Error Resume Next
   MdiSect.lblBotPanel.Caption = MdiSect.Caption
'   If Forms.Count < 3 Then
'      iResultSets = RdoCon.rdoResultsets.Count
'      For iList = iResultSets - 1 To 0 Step -1
'         RdoCon.rdoResultsets(iList).Close
'      Next
'   End If
   If bDontShowForm = 0 Then
      Select Case cUR.CurrentGroup
         Case "Invc"
            zGr1Invc.Show
         Case "Matl"
            zGr2Pick.Show
         Case "Lots"
            zGr5Lots.Show
         Case "Recv"
            zGr3Recv.Show
         Case "Invm"
            zGr4Matl.Show
      End Select
      Erase bActiveTab
      cUR.CurrentGroup = ""
   End If
   
End Sub

'Find a favorite from the list
'1/5/04 removed Raw/FG and WIP

Public Sub OpenFavorite(sSelected As String)
   CloseForms
   If LTrim$(sSelected) = "" Then Exit Sub
   MouseCursor 13
   On Error GoTo OpenFavErr1
   cUR.CurrentGroup = ""
   Select Case sSelected
      Case "Parts"
         cUR.CurrentGroup = "Invc"
         InvcINe01a.Show
      Case "Parts By Part Number"
         cUR.CurrentGroup = "Invc"
         InvcINp01a.Show
      Case "Parts By Part Description"
         cUR.CurrentGroup = "Invc"
         InvcINp02a.Show
      Case "Part Unit of Measure and Weight"
         cUR.CurrentGroup = "Invc"
         InvcINe02a.Show
      Case "Pick Items"
         cUR.CurrentGroup = "Matl"
         PickMCe01a.z1(2).Visible = True
         PickMCe01a.optInd.Visible = True
         PickMCe01a.optExp.Visible = True
         PickMCe01a.Show
      Case "Pick Substitution"
         cUR.CurrentGroup = "Matl"
         PickMCe02a.Show
      Case "Revise A Pick List"
         cUR.CurrentGroup = "Matl"
         PickMCe01a.Caption = "Revise A Pick List"
         PickMCe01a.Show
      Case "Scrap/Restock Information"
         cUR.CurrentGroup = "Matl"
         PickMCe03a.Show
      Case "Charge Material To A Project"
         cUR.CurrentGroup = "Matl"
         PickMCe04a.Show
      Case "Cancel A Pick List"
         cUR.CurrentGroup = "Matl"
         PickMCf01a.Show
      Case "Individual Pick List"
         cUR.CurrentGroup = "Matl"
         PickMCp01a.Show
      Case "Vendors"
         cUR.CurrentGroup = "Recv"
         PurcPRe03a.Show
      Case "Purchase Order Receipt"
         cUR.CurrentGroup = "Recv"
         RecvRVe01a.Show
      Case "Cancel A Purchase Order Receipt"
         cUR.CurrentGroup = "Recv"
         RecvRVf01a.Show
      Case "Receiving Log By Date"
         cUR.CurrentGroup = "Recv"
         RecvRVp01a.Show
      Case "Receiving Log By Vendor"
         cUR.CurrentGroup = "Recv"
         RecvRVp02a.Show
      Case "Negative Inventory Report"
         cUR.CurrentGroup = "Invc"
         InvcINp04a.Show
         'Case "Raw Material And Finished Goods Inventory"
         '    diaPin05.Show
      Case "Part Activity With Cost"
         cUR.CurrentGroup = "Invm"
         MatlMMp01a.Show
      Case "Part Quantity Activity"
         cUR.CurrentGroup = "Invm"
         MatlMMp02a.Show
      Case "Revised Parts"
         cUR.CurrentGroup = "Invc"
         InvcINp03a.Show
      Case "Copy A Part Number"
         cUR.CurrentGroup = "Invc"
         InvcINf02a.Show
         'Case "Work In Progress Report"
         '    diaPin06.Show
      Case "Standard Cost"
         cUR.CurrentGroup = "Invm"
         MatlMMe02a.Show
      Case "Pick Expediting"
         cUR.CurrentGroup = "Matl"
         PickMCp02a.Show
      Case "Change A Part Number"
         cUR.CurrentGroup = "Invc"
         InvcINf03a.Show
      Case "Adjust Part Quantity"
         cUR.CurrentGroup = "Invm"
         MatlMMf01a.Show
         'Case "Part Markup Matrix"
         '    diaPmatx.Show
      Case "Delete A Part Number"
         cUR.CurrentGroup = "Invc"
         InvcINf01a.Show
      Case "Part Aliases"
         cUR.CurrentGroup = "Invc"
         InvcINe03a.Show
      Case "Parts With Aliased Numbers"
         cUR.CurrentGroup = "Invc"
         InvcINp05a.Show
      Case "E-Commerce Parts"
         cUR.CurrentGroup = "Invc"
         InvcINp06a.Show
      Case "Revise Lots"
         cUR.CurrentGroup = "Lots"
         LotsLTe01a.Show
      Case "Lots By Part Number"
         cUR.CurrentGroup = "Lots"
         LotsLTp01a.Show
      Case "On Dock (Delivered)"
         cUR.CurrentGroup = "Recv"
         DockODe02a.Show
      Case "Parts On Dock (Delivered)"
         cUR.CurrentGroup = "Recv"
         DockODp02a.Show
      Case "Set Lot Tracking Requirements"
         cUR.CurrentGroup = "Lots"
         LotsLTe02a.Show
      Case "Lots Available By Part Number"
         cUR.CurrentGroup = "Lots"
         LotsLTp02a.Show
      Case "Uncosted Lots"
         cUR.CurrentGroup = "Lots"
         LotsLTp03a.Show
      Case "Lot Organization-Single Part Number"
         cUR.CurrentGroup = "Lots"
         LotsLTf01a.Show
      Case "Lot Organization-Groups Of Part Numbers"
         cUR.CurrentGroup = "Lots"
         ' diaLotGrp.Show
      Case "Lot Organization"
         cUR.CurrentGroup = "Lots"
         LotsLTf01a.Show
      Case "Cancel A Pick List Item"
         cUR.CurrentGroup = "Matl"
         PickMCf02a.Show
      Case "Part Locations"
         cUR.CurrentGroup = "Invc"
         InvcINe04a.Show
      Case "Assign ABC Classes"
         cUR.CurrentGroup = "Invm"
         MatlMMe01a.Show
      Case "Part Numbers Without An ABC Class"
         cUR.CurrentGroup = "Invm"
         MatlMMp03a.Show
      Case "Create ABC Classes"
         cUR.CurrentGroup = "Invm"
         CyclCYe01a.Show
      Case "ABC Classes By Part Number"
         cUR.CurrentGroup = "Invm"
         MatlMMp04a.Show
      Case "Update ABC Class Codes By Standard Cost"
         cUR.CurrentGroup = "Invm"
         CyclCYf01a.Show
      Case "Update Inspection Dates For ABC Classes"
         cUR.CurrentGroup = "Invm"
         CyclCYf02a.Show
      Case "Part Numbers Without An Inventory Location"
         cUR.CurrentGroup = "Invc"
         InvcINp07a.Show
      Case "Inactive Inventory Report"
         cUR.CurrentGroup = "Invc"
         InvcINp08a.Show
      Case "Excess Inventory Report"
         cUR.CurrentGroup = "Invc"
         InvcINp09a.Show
      Case "Inventory Due By ABC Class And Date"
         cUR.CurrentGroup = "Invm"
         MatlMMp05a.Show
      Case "Add A Pick List Item"
         cUR.CurrentGroup = "Matl"
         PickMCe05a.Show
'      Case "Revise A Cycle Count"
'         cUR.CurrentGroup = "Invm"
'         CyclCYe03a.Show
      Case "Assign Parts to a Cycle Count"
         cUR.CurrentGroup = "Invm"
         CyclCYe04.Show
      Case "Initialize a Cycle Count"
         cUR.CurrentGroup = "Invm"
         CyclCYe02a.Show
      Case "Delete A Cycle Count"
         cUR.CurrentGroup = "Invm"
         CyclCYf04a.Show
      Case "Unlock A Cycle Count"
         cUR.CurrentGroup = "Invm"
         CyclCYf03a.Show
      Case "Preview A Cycle Count"
         cUR.CurrentGroup = "Invm"
         CyclCYp01a.Show
      Case "Print Cycle Count Sheets"
         cUR.CurrentGroup = "Invm"
         CyclCYp02a.Show
      Case "ABC Inventory Reconciliation"
         cUR.CurrentGroup = "Invm"
         CyclCYf07.Show
      Case "Mark Cycle Count Audited"
         cUR.CurrentGroup = "Invm"
         CyclCYf06a.Show
      Case "Update Inventory Activity Standard Costs"
         cUR.CurrentGroup = "Invm"
         MatlMMf02a.Show
      Case "Split A Lot"
         cUR.CurrentGroup = "Lots"
         LotsLTe03a.Show
      Case "Split Lots"
         cUR.CurrentGroup = "Lots"
         LotsLTp04a.Show
      Case "Inventory Variance Reconciliation"
         cUR.CurrentGroup = "Invm"
         CyclCYp03a.Show
      Case "Cycle Count Status"
         cUR.CurrentGroup = "Invm"
         CyclCYp04a.Show
      Case "Mismatched Lots"
         cUR.CurrentGroup = "Lots"
         LotsLTp06a.Show
      Case "Lot Quantity Reconciliation"
         cUR.CurrentGroup = "Lots"
         LotsLTf02a.Show
      Case "Inventory Transfer"
         cUR.CurrentGroup = "Lots"
         LotsLTe04a.Show
      Case "Inventory Transfers (Report)"
         cUR.CurrentGroup = "Lots"
         LotsLTp05a.Show
      Case "Set Part QOH to Lot QOH"
         cUR.CurrentGroup = "Invm"
         MatlMMf03a.Show
      
      Case Else
         MouseCursor 0
   End Select
   On Error GoTo 0
   Exit Sub
   
OpenFavErr1:
   Resume OpenFavErr2
OpenFavErr2:
   MouseCursor 0
   MsgBox "ActiveX Error. Can't Load Form..", 48, "System    "
   On Error GoTo 0
   
End Sub

Sub Main()
   Dim sApptitle As String
   If App.PrevInstance Then
      On Error Resume Next
      sApptitle = App.Title
      App.Title = "E1eInv"
      SysMsgBox.Width = 3800
      SysMsgBox.msg.Width = 3200
      SysMsgBox.tmr1.Enabled = True
      SysMsgBox.msg = sApptitle & " Is Already Open."
      SysMsgBox.Show
      Sleep 5000
      AppActivate sApptitle
   End
   Exit Sub
End If
' Set the Module name before loading the form
sProgName = "Inventory"
MainLoad "invc"
GetFavorites "EsiInvc"
' MM 9/10/2009
'sProgName = "Inventory"
MdiSect.Show

End Sub

'Pick up permissions for this use


Public Sub FindPart(sGetPart As String, Optional NoMessage As Byte)
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo modErr1
   If Len(Trim(sGetPart)) > 0 Then
      sSql = "Qry_GetINVCfindPart '" & Compress(sGetPart) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
      If bSqlRows Then
         With RdoPrt
            On Error Resume Next
            MdiSect.ActiveForm.cmbPrt = "" & Trim(!PartNum)
            MdiSect.ActiveForm.lblDsc = "" & !PADESC
            MdiSect.ActiveForm.lblTyp = Format(0 + !PALEVEL, "0")
            MdiSect.ActiveForm.lblUom = "" & Trim(!PAUNITS)
         End With
         ClearResultSet RdoPrt
         bFoundPart = 1
      Else
         On Error Resume Next
         If Not NoMessage Then
            MsgBox "Part Wasn't Found.", 48, MdiSect.ActiveForm.Caption
            MdiSect.ActiveForm.cmbPrt = ""
         End If
         MdiSect.ActiveForm.lblDsc = "*** Part Number Wasn't Found ***"
         MdiSect.ActiveForm.lblTyp = ""
         bFoundPart = 0
      End If
      Set RdoPrt = Nothing
   Else
      On Error Resume Next
      MdiSect.ActiveForm.cmbPrt = "NONE"
      MdiSect.ActiveForm.lblDsc = "*** Part Number Wasn't Found ***"
      bFoundPart = 0
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
modErr1:
   sProcName = "findpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   bFoundPart = 0
   DoModuleErrors MdiSect.ActiveForm
   
End Sub



Public Function FindVendor(Optional AllowAll As Boolean = False) As Byte
   Dim RdoVed As ADODB.Recordset
   If (MdiSect.ActiveForm.cmbVnd = "ALL" And AllowAll = True) Or Len(MdiSect.ActiveForm.cmbVnd) = 0 Then
    If AllowAll Then
        MdiSect.ActiveForm.cmbVnd = "ALL"
        MdiSect.ActiveForm.txtNme = "* All Vendors *"
        FindVendor = 1
    End If
    Exit Function
   End If
   On Error GoTo modErr1
   sSql = "Qry_GetVendorBasics '" & Compress(MdiSect.ActiveForm.cmbVnd) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      On Error Resume Next
      With RdoVed
         MdiSect.ActiveForm.cmbVnd = "" & Trim(!VENICKNAME)
         MdiSect.ActiveForm.txtNme = "" & Trim(!VEBNAME)
         FindVendor = 1
         ClearResultSet RdoVed
      End With
   Else
      On Error Resume Next
      MdiSect.ActiveForm.cmbVnd = ""
      MdiSect.ActiveForm.txtNme = "No Valid Vendor Selected."
      FindVendor = 0
   End If
   Set RdoVed = Nothing
   Exit Function
   
modErr1:
   sProcName = "findvendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   FindVendor = False
   DoModuleErrors MdiSect.ActiveForm
   
End Function


Public Function GetThisVendor(sNick As String, sName As String)
   Dim RdoPrt As ADODB.Recordset
   sSql = "Qry_GetVendorBasics '" & Compress(sNick) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         GetThisVendor = "" & Trim(!VENICKNAME)
         sName = "" & Trim(!VEBNAME)
         ClearResultSet RdoPrt
      End With
   End If
   Set RdoPrt = Nothing
   
End Function


'Use To Add Columns to a table where necessary.
'Will only update if the Column doesn't exist or if
'SQL Server isn't open. The later won't make any difference
'anyway.
'Create tables, indexes,columns, etc here

Public Sub UpDateTables()
   If MdiSect.bUnloading = 1 Then Exit Sub
   'Dim RdoTest As ADODB.Recordset
   
   SaveSetting "Esi2000", "AppTitle", "invc", "ESI Inventory"
   SysOpen.Show
   SysOpen.prg1.Visible = True
   SysOpen.pnl = "Configuration Settings."
   SysOpen.pnl.Refresh
   
   '2/15/01 number for activity table
   On Error Resume Next
   SysOpen.prg1.Value = 20
   Sleep 500
   Err.Clear
   ' BuildKeys
   SysOpen.prg1.Value = 60
   Sleep 500
   
   GoTo modErr2
   Exit Sub
   
modErr1:
   Resume modErr2
modErr2:
   On Error GoTo 0
   SysOpen.Timer1.Enabled = True
   SysOpen.prg1.Value = 100
   SysOpen.Refresh
   Sleep 500
   
End Sub

Public Function GetNextPickRecord(sMoPartRef As String, lRunno As Long) As Integer
   Dim rdoRec As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT MAX(PKRECORD) FROM MopkTable WHERE " _
          & "PKMOPART='" & sMoPartRef & "' AND " _
          & "PKMORUN=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoRec, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(rdoRec.Fields(0)) Then
         GetNextPickRecord = rdoRec.Fields(0) + 1
      Else
         GetNextPickRecord = 1
      End If
      ClearResultSet rdoRec
   Else
      GetNextPickRecord = 1
   End If
   Set rdoRec = Nothing
   Exit Function
   
modErr1:
Resume modErr2:
modErr2:
   GetNextPickRecord = 1
   On Error GoTo 0
   
End Function


'12/20/04 For improper Setup
'03/13/07 Commented out

Public Sub SetUpAbcCodes()
   
End Sub



'8/22/05

Public Function GetUserLotID(UserLot As String) As Byte
   Dim RdoLot As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT DISTINCT LOTUSERLOTID FROM LohdTable WHERE " _
          & "LOTUSERLOTID='" & UserLot & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then GetUserLotID = 1 _
                                   Else GetUserLotID = 0
   ClearResultSet RdoLot
   If GetUserLotID = 1 Then MsgBox "That User Lot ID Is In Use.", _
                     vbInformation, "Revise A User Lot Number"
   Set RdoLot = Nothing
   Exit Function
   
modErr1:
   
End Function

'6/6/06
'11-21-06 Changed if Err

Private Sub BuildKeys()
   On Error Resume Next
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   sSql = "DROP INDEX CabcTable.AbcRef"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then
      sSql = "ALTER TABLE CabcTable ALTER COLUMN COABCROW TINYINT NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE CabcTable ADD Constraint PK_CabcTable_COABCROW PRIMARY KEY CLUSTERED (COABCROW) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX CchdTable.CycleRef"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE CchdTable ALTER COLUMN CCREF CHAR(20) NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE CchdTable ADD Constraint PK_CchdTable_CCREF PRIMARY KEY CLUSTERED (CCREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      sSql = "DELETE FROM CcitTable " & vbCr _
             & "FROM CcitTable LEFT JOIN CchdTable ON CcitTable.CIREF = CchdTable.CCREF " & vbCr _
             & "WHERE (CchdTable.CCREF Is Null)"
      clsADOCon.ExecuteSql sSql
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX CcitTable.CycleRef"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE CcitTable ALTER COLUMN CIREF CHAR(20) NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE CcitTable ALTER COLUMN CIPARTREF CHAR(30) NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE CcitTable ADD Constraint PK_CcitTable_CIREF PRIMARY KEY CLUSTERED (CIREF,CIPARTREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      sSql = "ALTER TABLE CcitTable ADD CONSTRAINT FK_CcitTable_CchdTable FOREIGN KEY (CIREF) References CchdTable ON UPDATE CASCADE ON DELETE CASCADE"
      clsADOCon.ExecuteSql sSql
      
      Err.Clear
      sSql = "ALTER TABLE CcitTable ADD CONSTRAINT FK_CcitTable_PartTable FOREIGN KEY (CIPARTREF) References PartTable "
      clsADOCon.ExecuteSql sSql
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX CcltTable.CycleRef"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE CcltTable ALTER COLUMN CLREF CHAR(20) NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE CcltTable ALTER COLUMN CLPARTREF CHAR(30) NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE CcltTable ALTER COLUMN CLLOTNUMBER CHAR(15) NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE CcltTable ADD Constraint PK_CcitTable_CLREF PRIMARY KEY CLUSTERED (CLREF,CLPARTREF,CLLOTNUMBER) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      sSql = "ALTER TABLE CcltTable ADD CONSTRAINT FK_CcltTable_CchdTable FOREIGN KEY (CLREF) References CchdTable ON UPDATE CASCADE ON DELETE CASCADE"
      clsADOCon.ExecuteSql sSql
      
      Err.Clear
      sSql = "ALTER TABLE CcltTable ADD CONSTRAINT FK_CcltTable_PartTable FOREIGN KEY (CLPARTREF) References PartTable"
      clsADOCon.ExecuteSql sSql
      
      
      sSql = "DROP INDEX InvaTable.InNum"
      clsADOCon.ExecuteSql sSql
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX InvaTable.InvPart"
      clsADOCon.ExecuteSql sSql
      
      sSql = "DROP INDEX InvaTable.InvType"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE InvaTable ALTER COLUMN INPART CHAR(30) NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE InvaTable ALTER COLUMN INTYPE INT NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE InvaTable ALTER COLUMN INNUMBER INT NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "CREATE CLUSTERED INDEX INVCLUSTER ON InvaTable(INPART,INTYPE,INNUMBER) WITH FILLFACTOR = 80 "
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX LohdTable.LotRef"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE LohdTable ALTER COLUMN LOTNUMBER CHAR(15) NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE LohdTable ADD Constraint PK_LohdTable_LOTNUMBER PRIMARY KEY CLUSTERED (LOTNUMBER) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      sSql = "DELETE FROM LoitTable " & vbCr _
             & "FROM LoitTable LEFT JOIN LohdTable ON LoitTable.LOINUMBER = LohdTable.LOTNUMBER " & vbCr _
             & "WHERE (LohdTable.LOTNUMBER Is Null)"
      clsADOCon.ExecuteSql sSql
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX LoitTable.LotRef"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE LoitTable ALTER COLUMN LOINUMBER CHAR(15) NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE LoitTable ALTER COLUMN LOIRECORD INT NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE LoitTable ADD Constraint PK_LoitTable_LOINUMBER PRIMARY KEY CLUSTERED (LOINUMBER,LOIRECORD) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      clsADOCon.ADOErrNum = 0
      sSql = "ALTER TABLE LoitTable ADD CONSTRAINT FK_LoitTable_LohdTable FOREIGN KEY (LOINUMBER) References LohdTable"
      clsADOCon.ExecuteSql sSql
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX PaalTable.AlRef"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE PaalTable ALTER COLUMN ALPARTREF CHAR(30) NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE PaalTable ALTER COLUMN ALALIASREF CHAR(30) NOT NULL"
      clsADOCon.ExecuteSql sSql
      
      sSql = "ALTER TABLE PaalTable ADD Constraint PK_PaalTable_ALPARTREF PRIMARY KEY CLUSTERED (ALPARTREF,ALALIASREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      sSql = "ALTER TABLE PaalTable ADD CONSTRAINT FK_PaalTable_PartTable FOREIGN KEY (ALPARTREF) References PartTable ON DELETE CASCADE ON UPDATE CASCADE"
      clsADOCon.ExecuteSql sSql
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
   End If
   Exit Sub
   
KeysErr1:
   On Error Resume Next
   clsADOCon.RollbackTrans
   clsADOCon.ADOErrNum = 0
   
End Sub

Public Function GetMRPCreateDates(DateCreated As String, DateThrough As String)
   Dim RdoDate As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT MRP_ROW,MRP_CREATEDATE,MRP_THROUGHDATE FROM " _
          & "MrpdTable WHERE MRP_ROW=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDate, ES_FORWARD)
   If bSqlRows Then
      With RdoDate
         If Not IsNull(!MRP_CREATEDATE) Then
            DateCreated = Format$(!MRP_CREATEDATE, "mm/dd/yy")
         Else
            DateCreated = Format$(ES_SYSDATE, "mm/dd/yy")
         End If
         If Not IsNull(!MRP_CREATEDATE) Then
            DateThrough = Format$(!MRP_THROUGHDATE, "mm/dd/yy")
         Else
            DateThrough = Format$(ES_SYSDATE, "mm/dd/yy")
         End If
         .Cancel
      End With
   Else
      DateCreated = "  "
      DateThrough = "  "
   End If
   Set RdoDate = Nothing
   Exit Function
modErr1:
   Err.Clear
   DateCreated = "  "
   DateThrough = "  "
   
End Function



Public Function GetInventoryComment(ByVal ActivityNo As String)
    Dim RdoInv As ADODB.Recordset
    GetInventoryComment = ""
    On Error Resume Next
    sSql = "SELECT INREF2 FROM InvaTable WHERE INNUMBER = " & ActivityNo
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
    ' Syntax was wrong
    'If bSqlRows And Not RdoInv.EOF Then GetInventoryComment = Trim("" & !INREF2)
    If bSqlRows And Not RdoInv.EOF Then GetInventoryComment = Trim("" & RdoInv.Fields(0))
    Set RdoInv = Nothing
End Function
