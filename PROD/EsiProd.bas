Attribute VB_Name = "EsiProd"
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Customer permissions 6/29/03
'10/7/04 added GetPoDataFormat
'1/20/05 added New Services PO
'3/31/05 Removed Jet excepting AWI
'8/8/05 Checked KeySet Clearing
'10/31/05 Added Cur.CurrentGroup to OpenFavorite. Opens appropriate tab
'         when called from Recent/Favorites and closed.
'12/22/05 Removed Vendor RFQ and references
'12/26/05 Removed unused procedures
'1/12/06 Completed renaming dialogs to be consistent with Fina
'5/3/06  Converted MRP columns to Dec(12,4)
'5/16/06 Delete Triggers on RunsTable/PohdTable
'6/5/06 BuildKeys
'6/26/06 Removed Threed32.ocx
'8/8/06 SSTab32.OCX Free
'9/7/06 AWI Custom MO - Removed the last JET references
'1/10/07 Added GetThisVendor for reports 7.2.1
Option Explicit
Public bGoodSoMo As Byte
Public bPOCaption As Byte
Public bFoundPart As Byte
Public iAutoIncr As Integer
Public gbInSqlRows As Boolean

Public sCurrDate As String
Public sCurrEmployee As String
Public sCurrForm As String
Public sPassedRout As String
Public sPassedMo As String
Public sPassedPart As String
Public sSelected As String

Public sFavorites(13) As String
Public sRecent(10) As String
Public sSession(50) As String

Public vTimeFormat As Variant

'Column updates
Private RdoCol As ADODB.Recordset
'Private ER As rdoError
Private ADOError As ADODB.Error
Public gblnSqlRows As Boolean

Public Function GetUserLotID(UserLot As String) As Byte

   ' deprecated - use ClassLot.IsUserLotIdInUseForAnotherLot
   
   Dim RdoLot As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT DISTINCT LOTUSERLOTID FROM LohdTable WHERE " _
          & "LOTUSERLOTID='" & UserLot & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then GetUserLotID = 1 _
                                   Else GetUserLotID = 0
   ClearResultSet RdoLot
   If GetUserLotID = 1 Then MsgBox "That User Lot ID Is In Use.", _
                     vbInformation, "Revise A User Lot Number"
   Set RdoLot = Nothing

End Function

Public Function FindToolList(ToolNumber As String, ToolDesc As String, Optional _
                             DontShow As Byte) As String
   Dim RdoTlst As ADODB.Recordset
   
   On Error GoTo modErr1
   sSql = "Qry_GetToolList '" & Compress(ToolNumber) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTlst, ES_FORWARD)
   If bSqlRows Then
      With RdoTlst
         FindToolList = "" & Trim(!TOOLLIST_NUM)
         If DontShow = 0 Then MDISect.ActiveForm.lblLst = "" & Trim(!TOOLLIST_DESC)
         ClearResultSet RdoTlst
      End With
   Else
      On Error Resume Next
      FindToolList = ""
      If DontShow = 0 Then MDISect.ActiveForm.lblLst = "*** Tool List Wasn't Found ***"
   End If
   Set RdoTlst = Nothing
   Exit Function
   
modErr1:
   sProcName = "findtoollist"
   FindToolList = ""
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function


Public Function GetRoutCenter(CenterRef As String) As String
   Dim RdoShp As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_GetRoutCenter '" & CenterRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         GetRoutCenter = "" & Trim(!WCNNUM)
         ClearResultSet RdoShp
      End With
   Else
      GetRoutCenter = ""
   End If
   Set RdoShp = Nothing
   Exit Function
   
modErr1:
   sProcName = "getroutcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
   
End Function

'4/7/04

Public Function GetRoutShop(ShopRef As String) As String
   Dim RdoShp As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_GetShop '" & ShopRef & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         GetRoutShop = "" & Trim(!SHPNUM)
         ClearResultSet RdoShp
      End With
   Else
      GetRoutShop = ""
   End If
   Set RdoShp = Nothing
   Exit Function
   
modErr1:
   sProcName = "getroutshop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function


'11/14/02 to retrieve the next PKRECORD (index piece)

Public Function GetNextPickRecord(sMoPartRef As String, lRunno As Long) As Integer
   Dim RdoRec As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT MAX(PKRECORD) FROM MopkTable WHERE " _
          & "PKMOPART='" & sMoPartRef & "' AND " _
          & "PKMORUN=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRec, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoRec.Fields(0)) Then
         GetNextPickRecord = RdoRec.Fields(0) + 1
      Else
         GetNextPickRecord = 1
      End If
   Else
      GetNextPickRecord = 1
   End If
   Exit Function
   
modErr1:
Resume modErr2:
modErr2:
   GetNextPickRecord = 1
   On Error GoTo 0
   
End Function

'Time for Time Cards and Labor for distribution

Public Function GetDefTimeAccounts(AccountType As String) As String
   Dim RdoAcc As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT CODEFTIMEACCT,CODEFLABORACCT FROM ComnTable " _
          & "WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAcc, ES_FORWARD)
   If bSqlRows Then
      With RdoAcc
         If AccountType = "Time" Then
            GetDefTimeAccounts = "" & Trim(.Fields(0))
         Else
            GetDefTimeAccounts = "" & Trim(.Fields(1))
         End If
         ClearResultSet RdoAcc
      End With
   End If
   Set RdoAcc = Nothing
   Exit Function
   
modErr1:
   GetDefTimeAccounts = ""
   
End Function


Public Sub GetLastMrp()
   Dim RdoShp As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT DISTINCT MRP_ROW,MRP_CREATEDATE,MRP_CREATEDBY " _
          & "FROM MrpdTable WHERE MRP_ROW=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         MDISect.ActiveForm.lblMrp = Format(!MRP_CREATEDATE, "mm/dd/yy hh:mm AM/PM")
         MDISect.ActiveForm.lblUsr = "" & Trim(!MRP_CREATEDBY)
         MDISect.ActiveForm.lblMrp.ForeColor = Es_TextForeColor
         ClearResultSet RdoShp
      End With
   Else
      MDISect.ActiveForm.lblMrp = "No Current Mrp"
      MDISect.ActiveForm.lblMrp.ForeColor = ES_RED
      MDISect.ActiveForm.lblUsr = ""
   End If
   Set RdoShp = Nothing
   Exit Sub

modErr1:
   On Error GoTo 0

End Sub

'Function FormatScheduleTime(Optional cHours As Currency) As Variant
'   If cHours = 0 Then cHours = 8
'   Select Case cHours
'      Case Is < 8.5
'         FormatScheduleTime = "mm/dd/yy 14:30"
'      Case 8.5 To 16
'         FormatScheduleTime = "mm/dd/yy 21:30"
'      Case Is > 16
'         FormatScheduleTime = "mm/dd/yy 23:59"
'   End Select
'   vTimeFormat = FormatScheduleTime
'
'End Function
'

'11/21/06 Added .ListCount

Public Sub FillRoutings()
   Dim RdoRtg As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_FillRoutings "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRtg, ES_FORWARD)
   If bSqlRows Then
      With RdoRtg
         Do Until .EOF
            AddComboStr MDISect.ActiveForm.cmbRte.hwnd, "" & Trim(!RTNUM)
            .MoveNext
         Loop
         ClearResultSet RdoRtg
      End With
   End If
   If MDISect.ActiveForm.cmbRte.ListCount > 0 Then _
      MDISect.ActiveForm.cmbRte.Text = MDISect.ActiveForm.cmbRte.List(0)
   Set RdoRtg = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillroutings"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub

'8/9/05 Close open sets

Public Sub FormUnload(Optional bDontShowForm As Byte)
   Dim iList As Integer
   Dim iResultSets As Integer
   On Error Resume Next
   MDISect.lblBotPanel = MDISect.Caption
'   If Forms.Count < 3 Then
'      iResultSets = RdoCon.ADODB.Recordsets.Count
'      For iList = iResultSets - 1 To 0 Step -1
'         RdoCon.ADODB.Recordsets(iList).Close
'      Next
'   End If
   If bDontShowForm = 0 Then
      Select Case cUR.CurrentGroup
         Case "Capa"
            zGr2Capa.Show
         Case "Time"
            zGr4PAC.Show
         Case "Mrpl"
            zGr5Mrpl.Show
         Case "Purc"
            zGr3Purc.Show
         Case "Shop"
            zGr1Shop.Show
            ' Case "RFQS"
            '     tabRfq.Show
      End Select
      Erase bActiveTab
      cUR.CurrentGroup = ""
   End If
   
End Sub

'Find a favorite from the list

Public Sub OpenFavorite(sSelected As String)
   CloseForms
   If LTrim$(sSelected) = "" Then Exit Sub
   MouseCursor 13
   On Error GoTo OpenFavErr1
   Select Case sSelected
      Case "Shops"
         cUR.CurrentGroup = "Capa"
         CapaCPe02a.Show
      Case "Work Centers"
         cUR.CurrentGroup = "Capa"
         CapaCPe01a.Show
      Case "Negative Inventory Report"
         cUR.CurrentGroup = "Mrpl"
         InvcINp04a.Show
      Case "Work Centers Report"
         cUR.CurrentGroup = "Capa"
         CapaCPp01a.Show
      Case "Shop Information"
         cUR.CurrentGroup = "Capa"
         CapaCPp02a.Show
      Case "Company Calendar"
         cUR.CurrentGroup = "Capa"
         CapaCPe04a.Show
      Case "Company Calendar Template"
         cUR.CurrentGroup = "Capa"
         CapaCPe05a.Show
      Case "Work Center Calendars"
         cUR.CurrentGroup = "Capa"
         CapaCPe03a.Show
      Case "List Of Work Center Calendars"
         cUR.CurrentGroup = "Capa"
         CapaCPp05a.Show
      Case "Revise a Manufacturing Order"
         cUR.CurrentGroup = "Shop"
         ShopSHe02a.Show
      Case "New Manufacturing Order"
         cUR.CurrentGroup = "Shop"
         ShopSHe01a.Show
      Case "Manufacturing Orders"
         cUR.CurrentGroup = "Shop"
         If ES_CUSTOM = "WATERJET" Then
            awiShopSHp01a.Show
'         ElseIf ES_CUSTOM = "JEVCO" Then
'            jevShopSHp01a.Show
         Else
            ShopSHp01a.Show
         End If
      Case "Manufacturing Order History"
         cUR.CurrentGroup = "Shop"
         ShopSHp02a.Show
      Case "Operation Completions/Assignments"
         cUR.CurrentGroup = "Shop"
         ShopSHe03a.Show
         '11/7/03 Removed
         'Case "Work Center Schedule"
         '    diaCswcn.Show
         'Case "Work Center Chart"
         '    diaCcwcn.Show
      Case "Manufacturing Order Completions"
         cUR.CurrentGroup = "Shop"
         ShopSHe04a.Show
         '        Case "Daily Time Charges"
         '            cur.CurrentGroup = "Time"
         '            HumnHUe01a.Show
         '        Case "Revise Daily Time Charges"
         '            cur.CurrentGroup = "Time"
         '            HumnHUe02a.Show
         'Case "Employees By Name"
         '    cUR.CurrentGroup = "Time"
         '    HumnHUp01a.Show
         'Case "Employees By Number"
         '    cUR.CurrentGroup = "Time"
         '    HumnHUp02a.Show
         '        Case "Daily Employee Time Charges"
         '            cur.CurrentGroup = "Time"
         '            HumnHUp03a.Show
         '        Case "Weekly Time Charges (Report)"
         '            cur.CurrentGroup = "Time"
         '            HumnHUp05a.Show
         '        Case "Delete A Daily Time Charge"
         '            cur.CurrentGroup = "Time"
         '            HumnHUf01a.Show
         '        Case "Time Type Codes (Report)"
         '            cur.CurrentGroup = "Time"
         '            HumnHUp15a.Show
      Case "Vendors"
         cUR.CurrentGroup = "Purc"
         'BBS Changed lines below for Ticket #25640
         VendorEdit01.Tag = "1"
         VendorEdit01.Show
         'PurcPRe03a.Show
      Case "Vendor List"
         cUR.CurrentGroup = "Purc"
         PurcPRp02a.Show
      Case "Vendor Directory"
         cUR.CurrentGroup = "Purc"
         PurcPRp03a.Show
      Case "Revise A Purchase Order"
         cUR.CurrentGroup = "Purc"
         PurcPRe02a.Show
      Case "New Purchase Order"
         cUR.CurrentGroup = "Purc"
         PurcPRe02a.optNew.Value = vbChecked
         PurcPRe02a.Show
         PurcPRe01a.Show
      Case "Purchase Orders (Report)"
         cUR.CurrentGroup = "Purc"
         PurcPRp01a.Show
      Case "Purchase Order Log"
         cUR.CurrentGroup = "Purc"
         PurcPRp04a.Show
      Case "Work Center Load"
         cUR.CurrentGroup = "Capa"
         CapaCPp03a.Show
      Case "Late Manufacturing Orders"
         cUR.CurrentGroup = "Capa"
         CapaCPp04a.Show
      Case "Manufacturing Orders By Date"
         cUR.CurrentGroup = "Shop"
         ShopSHp03a.Show
      Case "Manufacturing Orders By Part"
         cUR.CurrentGroup = "Shop"
         ShopSHp04a.Show
      Case "Purchase Order Log By Date"
         cUR.CurrentGroup = "Purc"
         PurcPRp05a.Show
      Case "Cancel A Purchase Order"
         cUR.CurrentGroup = "Purc"
         PurcPRf01a.Show
      Case "Purchase Expediting Report"
         cUR.CurrentGroup = "Purc"
         PurcPRp06a.Show
      Case "Purchase Expediting Report By Buyer"
         cUR.CurrentGroup = "Purc"
         PurcPRp07a.Show
      Case "Purchase Order History By Vendor"
         cUR.CurrentGroup = "Purc"
         PurcPRp08a.Show
      Case "Part Purchasing Information"
         cUR.CurrentGroup = "Purc"
         PurcPRe04a.Show
      Case "Reassign Shops And Work Centers"
         cUR.CurrentGroup = "Capa"
         CapaCPf01a.Show
      Case "Delete Shops"
         cUR.CurrentGroup = "Capa"
         CapaCPf02a.Show
      Case "Cancel A Manufacturing Order"
         cUR.CurrentGroup = "Shop"
         ShopSHf01a.Show
      Case "Cancel Manufacturing Order Completions"
         cUR.CurrentGroup = "Shop"
         ShopSHf02a.Show
'      Case "Close A Manufacturing Order"
'         cUR.CurrentGroup = "Shop"
'         ShopSHf04a.Show
'      Case "Open A Closed Manufacturing Order"
'         cUR.CurrentGroup = "Shop"
'         ShopSHf05a.Show
      Case "Update MO Routings"
         cUR.CurrentGroup = "Shop"
         ShopSHf03a.Show
         'Case "Work In Progress Report"
         '    diaPin06.Show
      Case "Purchasing History By Manufacturing Order"
         cUR.CurrentGroup = "Purc"
         PurcPRp09a.Show
         'Case "Standard Cost"
         '    diaIsstd.Show
      Case "Shop Load By Work Center"
         cUR.CurrentGroup = "Capa"
         CapaCPp06a.Show
      Case "Manufacturing Order Status"
         cUR.CurrentGroup = "Shop"
         ShopSHp05a.Show
      Case "Shop Queue"
         cUR.CurrentGroup = "Shop"
         ShopSHp07a.Show
      Case "Work Center Queue"
         cUR.CurrentGroup = "Shop"
         ShopSHp08a.Show
      Case "Late MO's By Operation"
         cUR.CurrentGroup = "Shop"
         ShopSHp09a.Show
      Case "Production Report"
         cUR.CurrentGroup = "Shop"
         ShopSHp10a.Show
      Case "MO Status By Customer"
         cUR.CurrentGroup = "Shop"
         ShopSHp06a.Show
      Case "Delete A Vendor"
         cUR.CurrentGroup = "Purc"
         PurcPRf02a.Show
      Case "Manufacturing Order Sales Order Allocations"
         cUR.CurrentGroup = "Shop"
         ShopSHp12a.Show
      Case "Individual Pick List"
         cUR.CurrentGroup = "Shop"
         PickMCp01a.Show
      Case "Change A Vendor Nickname"
         cUR.CurrentGroup = "Purc"
         PurcPRf03a.Show
      Case "Sales Order Allocations By Customer"
         cUR.CurrentGroup = "Shop"
         ShopSHp11a.Show
      Case "Purchasing History By Part"
         cUR.CurrentGroup = "Purc"
         PurcPRp10a.Show
      Case "MO Priority/Work Center Schedules"
         cUR.CurrentGroup = "Shop"
         CustSh01.Show
      Case "Generate A New MRP"
         cUR.CurrentGroup = "Mrpl"
         MrplMRf01a.Show
      Case "Part Manufacturing Parameters"
         cUR.CurrentGroup = "Mrpl"
         ShopSHe05a.Show
      Case "MRP Requirements By Part(s)"
         cUR.CurrentGroup = "Mrpl"
         MrplMRp01a.Show
      Case "MRP Exceptions By Part(s)"
         cUR.CurrentGroup = "Mrpl"
         MrplMRp02a.Show
      Case "Buyers"
         cUR.CurrentGroup = "Purc"
         PurcPRe05a.Show
      Case "Change A Buyer ID"
         cUR.CurrentGroup = "Purc"
         PurcPRf04a.Show
      Case "Delete A Buyer ID"
         cUR.CurrentGroup = "Purc"
         PurcPRf05a.Show
      Case "Release Manufacturing Orders"
         cUR.CurrentGroup = "Shop"
         ShopSHe06a.Show
      Case "Work Centers Without Calendars"
         cUR.CurrentGroup = "Capa"
         CapaCPp07a.Show
      Case "Outside Services Requirements"
         cUR.CurrentGroup = "Purc"
         PurcPRp11a.Show
      Case "Capacity And Load"
         cUR.CurrentGroup = "Capa"
         CapaCPp08a.Show
      Case "Add A Pick List Item"
         cUR.CurrentGroup = "Shop"
         PickMCe05a.Show
      Case "Delete Work Centers"
         cUR.CurrentGroup = "Capa"
         CapaCPf03a.Show
      Case "MRP Part Number Bill of Material"
         cUR.CurrentGroup = "Mrpl"
         MrplMRp03a.Show
      Case "Revise A PO Line Item Price (Invoiced Item)"
         cUR.CurrentGroup = "Purc"
         PurcPRf06a.Show
      Case "Split A Purchase Order Item"
         cUR.CurrentGroup = "Purc"
         PurcPRf07a.Show
      Case "Close Manufacturing Orders"
         cUR.CurrentGroup = "Shop"
         ShopScrns.Show
      Case "Part Status (Report)"
         cUR.CurrentGroup = "Shop"
         ShopSHp13a.Show
      Case "New Services Purchase Order"
         cUR.CurrentGroup = "Purc"
         PurcPRe02a.optNew.Value = vbChecked
         PurcPRe02a.Show
         bPOCaption = 1
         PurcPRe01a.optSrv.Value = vbChecked
         PurcPRe01a.Caption = "New Services Purchase Order"
         PurcPRe01a.Show
      Case "Split A Manufacturing Order"
         cUR.CurrentGroup = "Shop"
         ShopSHf06a.Show
      Case "Manufacturing Orders Splits"
         cUR.CurrentGroup = "Shop"
         ShopSHp14a.Show
      Case "Work Center Used On"
         cUR.CurrentGroup = "Capa"
         CapaCPp09a.Show
      Case "Assign Parts To Buyers"
         cUR.CurrentGroup = "Purc"
         PurcPRe06a.Show
      Case "Part Numbers By Buyer"
         cUR.CurrentGroup = "Purc"
         PurcPRp12a.Show
      Case "Approved Supplier by Part Number"
         cUR.CurrentGroup = "Purc"
         PurcPRe07a.Show
      Case "Approved Supplier by Part Number List"
         cUR.CurrentGroup = "Purc"
         PurcPRp13a.Show
      Case "Open A Canceled Purchase Order"
         PurcPRf08a.Show
      Case "Change Purchase Order Requested By"
         PurcPRf09a.Show
      Case "Purchase Order Requested By"
         PurcPRp14a.Show
      Case "Pictures By Routing Operation"
         RoutRTp07a.Show
      Case "Work Center Load Analysis"
         CapaCPp10a.Show
      Case "Work Center Queue List"
         ShopSHp15a.Show
      Case "Manufacturers"
         PurcPRe08a.Show
      Case "Change A Manufacturer's Nickname"
         PurcPRf10a.Show
      Case "Delete A Manufacturer"
         PurcPRf11a.Show
      Case "List Of Manufacturers"
         PurcPRp15a.Show
      Case "Manufacturers By Part Number"
         PurcPRp16a.Show
      Case "Sales Backlog with MO Status"
         ShopSHp16a.Show
      Case "Re-Schedule Purchase Order Item"
         PurcPRe09a.Show
      Case "MRP Open Orders"
         MrplMRp08.Show
      Case "Revise MO Dates, Quantity and Status"
         MrplMRe03.Show
      Case "Create PO's from MRP Exceptions By Part(s)"
         MrplMRe02.Show
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

'11/21/06 Changed Name from GetDefaults

Public Sub GetRoutingIncrementDefault()
   Dim RdoDef As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT RTEINCREMENT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDef)
   If bSqlRows Then
      iAutoIncr = RdoDef!RTEINCREMENT
   Else
      iAutoIncr = 10
   End If
   If iAutoIncr = 0 Then iAutoIncr = 10
   RdoDef.Close
   Exit Sub
   
modErr1:
   sProcName = "getdefaults"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Forms(1)
   
End Sub






Sub Main()
   Dim sApptitle As String
   If App.PrevInstance Then
      On Error Resume Next
      sApptitle = App.Title
      App.Title = "E1ePr"
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
On Error Resume Next
' Set the Module name before loading the form
sProgName = "Production"
MainLoad "prod"
GetFavorites "EsiProd"
' save the setting in registry for the module
SetRegistryAppTitle ("EsiProd")
' MM sProgName = "Production"
   On Error Resume Next
   ' Set the Module name before loading the form
   sProgName = "Production"
   MainLoad "prod"
   GetFavorites "EsiProd"
   ' save the setting in registry for the module
   SetRegistryAppTitle ("EsiProd")
   ' MM sProgName = "Production"


   Dim arv() As String
   Dim strInDate As String
   Dim strTolerance As String
   
   If (Command <> "") Then
      arv = Split(Command, ",")
      Dim iLen As Integer
      iLen = UBound(arv) - LBound(arv) + 1
      
      If (iLen > 1) Then
         strInDate = Format(Trim(arv(1)), "mm/dd/yy 23:59")
      Else
         strInDate = GetSetting("Esi2000", "EsiProd", "MRPCutoffDate", Format(DateAdd("d", 1095, Now), "mm/dd/yy"))
         strInDate = Format(strInDate, "mm/dd/yy 23:59")
      End If
      
      If (arv(0) = "AUTORUN_MRP") Then
      
         strTolerance = GetSetting("Esi2000", "EsiProd", "MRPTolerance", "0")
         MrplMRf01a.txtTolerance = strTolerance
         MrplMRf01a.cmbCutoff = strInDate
         MrplMRf01a.vEndDate = strInDate
         MrplMRf01a.AutoGenerateMrp (strInDate)
         End
      Else
         MDISect.Show
      End If
   Else
      MDISect.Show
   End If
' MM
'MDISect.Show
End Sub

''Pick up permissions for this user
'Public Sub GetSectionPermissions()
'    Dim RdoUsr As ADODB.Recordset
'    '11/21/06
'    User.Group1 = 1
'    User.Group2 = 1
'    User.Group3 = 1
'    User.Group4 = 1
'    User.Group5 = 1
'    User.Group6 = 1
'
'    On Error GoTo ModErr1
'    sSql = "SELECT USERREF,USERADDUSER,USERLEVEL," _
'        & "USERPRODGR1,USERPRODGR2,USERPRODGR3," _
'        & "USERPRODGR4,USERPRODGR5,USERPRODGR6 " _
'        & "FROM UsscTable WHERE USERREF='" _
'        & UCase$(cur.CurrentUser) & "'"
'    bSqlRows = clsAdoCon.GetDataSet(ssql, RdoUsr)
'        If bSqlRows Then
'            With RdoUsr
'                User.Adduser = !UserAddUser
'                User.Level = !UserLevel
'                User.Group1 = !USERPRODGR1
'                User.Group2 = !USERPRODGR2
'                User.Group3 = !USERPRODGR3
'                User.Group4 = !USERPRODGR4
'                User.Group5 = !USERPRODGR5
'                User.Group6 = !USERPRODGR6
'                ClearResultSet RdoUsr
'            End With
'        End If
'    Set RdoUsr = Nothing
'    Exit Sub
'
'ModErr1:
'    Resume modErr2:
'modErr2:
'    On Error GoTo 0
'
'End Sub












Public Sub FindShop()
   'Use local errors
   Dim RdoShp As ADODB.Recordset
   If Len(Trim(MDISect.ActiveForm.cmbShp)) > 0 Then
      sSql = "Qry_GetShop '" & Compress(MDISect.ActiveForm.cmbShp) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
      If bSqlRows Then
         MDISect.ActiveForm.cmbShp = "" & Trim(RdoShp!SHPNUM)
      Else
         MsgBox "Shop Wasn't Found.", vbInformation, _
            MDISect.ActiveForm.Caption
         MDISect.ActiveForm.cmbShp = ""
      End If
      ClearResultSet RdoShp
   End If
   Set RdoShp = Nothing
   
End Sub


'Replaced (Mostly) by GetCurrentPart

Public Sub FindPart(sGetPart As String, Optional NoMessage As Byte)
   Dim RdoPrt As ADODB.Recordset
   sGetPart = Compress(sGetPart)
   On Error GoTo modErr1
   If Len(sGetPart) > 0 Then
      sSql = "Qry_GetINVCfindPart '" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
      On Error Resume Next
      If bSqlRows Then
         With RdoPrt
            MDISect.ActiveForm.cmbPrt = "" & Trim(!PartNum)
            MDISect.ActiveForm.lblDsc = "" & !PADESC
            MDISect.ActiveForm.lblTyp = Format(0 + !PALEVEL, "0")
            MDISect.ActiveForm.lblUom = "" & Trim(!PAUNITS)
         End With
         bFoundPart = 1
      Else
         If NoMessage = 0 Then
            MsgBox "Part Wasn't Found.", 48, MDISect.ActiveForm.Caption
            MDISect.ActiveForm.cmbPrt = ""
         End If
         MDISect.ActiveForm.lblDsc = "*** Part Number Wasn't Found ***"
         MDISect.ActiveForm.lblTyp = ""
         bFoundPart = 0
      End If
      Set RdoPrt = Nothing
   Else
      On Error Resume Next
      If NoMessage = 0 Then
         MDISect.ActiveForm.cmbPrt = "NONE"
         MDISect.ActiveForm.lblDsc = "*** Part Number Wasn't Found ***"
      End If
      bFoundPart = 0
   End If
   Exit Sub
   
modErr1:
   sProcName = "findpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   bFoundPart = 0
   DoModuleErrors MDISect.ActiveForm
   
End Sub



'Changed to controls 11/6/04


'BBS Added this new function to see if the vendor is approved or not
'I'm using the database to return the msg for me so I can hopefully
'centralize this logic

Public Function IsVendorApproved(strVendNickName As String, DisplayMsgToUser As Byte, ByRef Reason As String) As Byte
Dim RdoTest As ADODB.Recordset


    IsVendorApproved = 1
    Reason = ""
    
   Err.Clear
   On Error Resume Next
   
    'BBS Changed this query on 4/2/2010 because of a bug in the original query
    
    sSql = "SELECT CASE WHEN (VEAPPROVREQ=1) THEN " _
       & " CASE WHEN VEAPPDATE IS NULL THEN 'This Vendor is not approved' " _
       & " WHEN VEREVIEWDT IS NOT NULL AND NOT (getdate() between VEAPPDATE AND VEREVIEWDT) THEN 'Vendor Approval has expired'" _
       & " WHEN GETDATE() < VEAPPDATE THEN 'This Vendor is not approved yet' ELSE '' END " _
       & " WHEN (VESURVEY=1) THEN CASE WHEN VESURVSENT IS NULL THEN 'A survey is required for this vendor, but has not been sent'" _
       & " WHEN VESURVREC IS NULL THEN 'A survey has been sent, but not received' ELSE '' END ELSE '' END AS APPMSG " _
    & " FROM VndrTable WHERE VENICKNAME='" & strVendNickName & "'"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTest, ES_KEYSET)
   If bSqlRows Then
        Reason = "" & RdoTest!APPMSG
        If (Reason = "") Then
            IsVendorApproved = 1
        Else
            IsVendorApproved = 0
            If (DisplayMsgToUser <> 0) Then MsgBox "" & RdoTest!APPMSG, vbOKOnly, "Vendor Approval"
        End If
    
   End If
    ClearResultSet RdoTest
    Set RdoTest = Nothing
    
    
End Function


Public Function FindVendor(ContrlCombo As Control, ControlLabel As Control) As Byte
   Dim RdoVed As ADODB.Recordset
   Dim sVendRef As String
      
   sVendRef = Compress(ContrlCombo)
   If Len(sVendRef) = 0 Then Exit Function
   On Error GoTo modErr1
   sSql = "Qry_GetVendorBasics '" & sVendRef & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      On Error Resume Next
      With RdoVed
         ContrlCombo = "" & Trim(!VENICKNAME)
         ControlLabel = "" & Trim(!VEBNAME)
         FindVendor = True
         
         ClearResultSet RdoVed
      End With
   Else
      On Error Resume Next
      ContrlCombo = ""
      ControlLabel = "No Valid Vendor Selected."
      FindVendor = False
   End If
   Set RdoVed = Nothing
   Exit Function
   
modErr1:
   sProcName = "findvendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   FindVendor = False
   DoModuleErrors MDISect.ActiveForm
   
End Function

Public Sub FillRuns(frm As Form, sSearchString As String, Optional sComboName As String)
   Dim RdoFrn As ADODB.Recordset
   If sComboName = "" Then sComboName = "cmbPrt"
   On Error GoTo modErr1
   If sSearchString = "<> 'CA'" Then
      sSql = "Qry_RunsNotCanceled"
   Else
      sSql = "Qry_RunsNotLikeC"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFrn, ES_FORWARD)
   If bSqlRows Then
      With RdoFrn
         If sComboName = "cmbPrt" Then
            If Trim(frm.cmbPrt) = "" Then frm.cmbPrt = "" & Trim(!PartNum)
            Do Until .EOF
               AddComboStr frm.cmbPrt.hwnd, "" & Trim(!PartNum)
               .MoveNext
            Loop
         Else
            Do Until .EOF
               AddComboStr frm.cmbMon.hwnd, "" & Trim(!PartNum)
               .MoveNext
            Loop
         End If
         ClearResultSet RdoFrn
      End With
   End If
   On Error Resume Next
   Set RdoFrn = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

'Use To Add Columns to a table where necessary.
'Will only update if the Column doesn't exist or if
'SQL Server isn't open. The later won't make any difference
'anyway.
'Create tables, indexes,columns, etc here

Public Sub UpDateTables()
   If MDISect.bUnloading = 1 Then Exit Sub
   Dim RdoTest As ADODB.Recordset
   
   MouseCursor 13
   SaveSetting "Esi2000", "AppTitle", "prod", "ESI Production"
   SysOpen.Show
   SysOpen.prg1.Visible = True
   SysOpen.pnl = "Configuration Settings."
   SysOpen.pnl.Refresh
   
   On Error Resume Next
   SysOpen.prg1.Value = 20
   'moved to OldUpdate  1/3/03, 2/5/03, 5/9/05, 10/6/06
   
   '5/15/03 Patch for On Dock
   '*Leave in
   Err.Clear
   sSql = "SELECT VEREF,VENICKNAME,VEBNAME FROM VndrTable WHERE VEREF='NONE'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTest, ES_KEYSET)
   If Not bSqlRows Then
      With RdoTest
         .AddNew
         !VEREF = "NONE"
         !VENICKNAME = "NONE"
         !VEBNAME = "No Vendor Selected"
         .Update
      End With
   End If
   CheckTriggers
   Sleep 500
   '6/5/06
   Err.Clear
   BuildKeys
   
   '1/8/07 Email to Vendor AR 7.2.0
   Err.Clear
   clsADOCon.ADOErrNum = 0
   
   sSql = "SELECT VEAREMAIL FROM VndrTable WHERE VEAREMAIL='fubar'"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum > 0 Then
      Err.Clear
      clsADOCon.ADOErrNum = 0
      
      sSql = "ALTER TABLE VndrTable ADD VEAREMAIL VARCHAR(60) NULL DEFAULT('')"
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "UPDATE VndrTable SET VEAREMAIL='' WHERE VEAREMAIL IS NULL"
         clsADOCon.ExecuteSql sSql
      End If
      
   End If
   Err.Clear
   clsADOCon.ADOErrNum = 0
   '3/20/07 7.3.0 Add Actual Routing information
   sSql = "SELECT RUNRTNUM FROM RunsTable WHERE RUNRTNUM='FOOBAR'"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum > 0 Then
      Err.Clear
      clsADOCon.ADOErrNum = 0
      sSql = "ALTER TABLE RunsTable ADD " _
             & "RUNRTNUM CHAR(30) NULL DEFAULT('')," _
             & "RUNRTDESC CHAR(30) NULL DEFAULT('')," _
             & "RUNRTBY CHAR(20) NULL DEFAULT('')," _
             & "RUNRTAPPBY CHAR(20) NULL DEFAULT('')," _
             & "RUNRTAPPDATE CHAR(8) NULL DEFAULT('')"
      clsADOCon.ExecuteSql sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "UPDATE RunsTable SET RUNRTNUM=''," _
                & "RUNRTDESC='',RUNRTBY=''," _
                & "RUNRTAPPBY='',RUNRTAPPDATE='' " _
                & "WHERE RUNRTNUM IS NULL"
         clsADOCon.ExecuteSql sSql
      End If
   End If
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   
   SysOpen.prg1.Value = 80
   Sleep 500
   GoTo modErr2
   Exit Sub
   
modErr1:
   Resume modErr2
modErr2:
   Set RdoTest = Nothing
   Err.Clear
   On Error GoTo 0
   SysOpen.Timer1.Enabled = True
   SysOpen.prg1.Value = 100
   SysOpen.Refresh
   Sleep 500
   
End Sub


Public Sub FindMoPart()
   Dim RdoPrt As ADODB.Recordset
   Dim sGetPart As String
   
   sGetPart = Compress(MDISect.ActiveForm.cmbMon)
   On Error GoTo modErr1
   If Len(sGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable WHERE PARTREF='" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
      If bSqlRows Then
         On Error Resume Next
         With RdoPrt
            MDISect.ActiveForm.cmbMon = "" & Trim(!PartNum)
            MDISect.ActiveForm.lblMon = "" & !PADESC
         End With
      Else
         MDISect.ActiveForm.cmbMon = "NONE"
         MDISect.ActiveForm.lblMon = "*** Part Number Wasn't Found ***"
      End If
      Set RdoPrt = Nothing
   End If
   Exit Sub
   
modErr1:
   sProcName = "findmopart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   bFoundPart = 0
   DoModuleErrors MDISect.ActiveForm
   
End Sub

Public Sub Oldupdates()
End Sub


'See if there are calendars
'Syntax is bGoodCal = GetCenterCalendar(Me, Format$(SomeDate,"mm/dd/yy")
'6/1/00

Public Function GetCenterCalendar(frm As Form, Optional sMonth As String) As Boolean
   On Error Resume Next
   If sMonth = "" Then
      sMonth = Format(ES_SYSDATE, "mmm") & "-" & Format(ES_SYSDATE, "yyyy")
   Else
      sMonth = Format(sMonth, "mmm") & "-" & Format(sMonth, "yyyy")
   End If
   sSql = "SELECT WCCREF FROM WcclTable WHERE WCCREF='" & sMonth & "'"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.RowsAffected = 0 Then
      GetCenterCalendar = False
      MsgBox "There Are No Work Center Calendars " & vbCr _
         & "Open For This Period " & sMonth & ".", vbInformation, frm.Caption
   Else
      GetCenterCalendar = True
   End If
   
End Function

'Currency because it rounds better and is faster
'Use local errors


Public Sub FillAllRuns(Contrl As Control)
   Dim RdoRns As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "PartTable,RunsTable WHERE PARTREF=RUNREF ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr Contrl.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      If Contrl.ListCount > 0 Then Contrl = Contrl.List(0)
   End If
   Set RdoRns = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillallruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub

'Scripting for table
'Note non-normalized portions for speed and efficiency
'3/16/01

Public Sub MRPScript()
   'Types:
   '   Incoming (+)
   '   1 = Beginning balance
   '      2 = PO Items
   '      3 = MO Completions
   '   Out Going (-)
   '      4 = SO Items
   '      5 = Picks
   '      6 = Bills (Used On and no PL yet)
   '
   
End Sub

Public Sub FillBuyers()
   Dim RdoByr As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_GetBuyerList"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoByr, ES_FORWARD)
   If bSqlRows Then
      With RdoByr
         Do Until .EOF
            AddComboStr MDISect.ActiveForm.cmbByr.hwnd, "" & Trim(!BYNUMBER)
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
   DoModuleErrors MDISect.ActiveForm
   
End Sub

Public Sub GetCurrentBuyer(sBuyer As String, Optional HideLabel As Byte)
   Dim RdoByr As ADODB.Recordset
   On Error GoTo modErr1
   sBuyer = UCase$(Compress(sBuyer))
   sSql = "SELECT BYNUMBER,BYLSTNAME,BYFSTNAME,BYMIDINIT FROM " _
          & "BuyrTable WHERE BYREF='" & sBuyer & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoByr, ES_FORWARD)
   If bSqlRows Then
      With RdoByr
         If MDISect.ActiveForm.Caption = "Revise A Purchase Order" Then
         
            MDISect.ActiveForm.cmbByr = "" & Trim(!BYNUMBER)
            If HideLabel = 0 Then
               MDISect.ActiveForm.lblByr = "" & Trim(!BYFSTNAME) _
                                           & " " & Trim(!BYMIDINIT) & " " & Trim(!BYLSTNAME)
            End If
         End If
         ClearResultSet RdoByr
      End With
   Else
      If Len(Trim(sBuyer)) > 0 Then
         MDISect.ActiveForm.lblByr = "*** Buyer Wasn't Found ***"
      Else
         MDISect.ActiveForm.lblByr = ""
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

Public Function GetMoOperation(MONUMBER As String, Runno As Long, iOpno As Integer) As Byte
   Dim RdoOpr As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,OPREF,OPRUN,OPNO " _
          & "FROM RunsTable,RnopTable WHERE (RUNREF=OPREF AND " _
          & "RUNNO=OPRUN) AND (RUNREF='" & MONUMBER & "' AND RUNNO=" & Runno _
          & " AND OPNO=" & iOpno & " AND RUNSTATUS<>'CA' AND " _
          & "RUNSTATUS<>'CL')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOpr, ES_FORWARD)
   If bSqlRows Then GetMoOperation = 1 Else GetMoOperation = 0
   Set RdoOpr = Nothing
   Exit Function
   
modErr1:
   GetMoOperation = 0
   
End Function

'1/20/04 calculate the end of the month

Public Function GetMonthEnd(Optional vMonthEnd As Variant) As Variant
   Dim bMonth As Byte
   Dim bEnd As Byte
   Dim bYear As Integer
   Dim vTest As Variant
   
   On Error Resume Next
   'Trap to test empty vMonth
   vTest = Left(vMonthEnd, 1)
   If Err > 0 Then
      bMonth = Format(ES_SYSDATE, "m")
      bYear = Format(ES_SYSDATE, "yyyy")
   Else
      bMonth = Format(vMonthEnd, "m")
      bYear = Format(vMonthEnd, "yyyy")
   End If
   Select Case bMonth
      Case 1, 3, 5, 7, 8, 10, 12
         bEnd = 31
      Case 2
         bEnd = 28
      Case Else
         bEnd = 30
   End Select
   
   If bEnd = 28 Then
      If bYear = 2004 Or bYear = 2008 Or bYear = 2012 _
                 Or bYear = 2016 Or bYear = 2020 Or bYear = 2024 Then bEnd = 29
   End If
   vMonthEnd = Format(bMonth, "00") & "/" & bEnd & "/" & Right$(str$(bYear), 2)
   GetMonthEnd = vMonthEnd
   
End Function



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

'04/01/05

Public Function GetCompanyCalendar() As Byte
   Dim RdoCal As ADODB.Recordset
   Dim sCalYear As String
   Dim sCalMonth As String
   
   On Error Resume Next
   sCalYear = Format$(Now, "yyyy")
   sCalMonth = Format$(Now, "mmm")
   sSql = "Qry_GetCompanyCalendar '" & sCalMonth & "-" & sCalYear & " '"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCal, ES_FORWARD)
   ClearResultSet RdoCal
   GetCompanyCalendar = bSqlRows
   Set RdoCal = Nothing
   
End Function


Public Sub GetMRPCreateDates(DateCreated As String, DateThrough As String)
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
   Exit Sub
modErr1:
   Err.Clear
   DateCreated = "  "
   DateThrough = "  "
   
End Sub

'6/5/06

Private Sub BuildKeys()
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX RunsTable.RunRef"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum > 0 Then GoTo KeysErr1
   
   sSql = "DROP INDEX RunsTable.RunPart"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNREF CHAR(30) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNNO INT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RunsTable ADD Constraint PK_RunsTable_RUNREF PRIMARY KEY CLUSTERED (RUNREF,RUNNO) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX RnopTable.OpRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP INDEX RnopTable.OpPart"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnopTable ALTER COLUMN OPREF CHAR(30) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnopTable ALTER COLUMN OPRUN INT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnopTable ALTER COLUMN OPNO SMALLINT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnopTable ADD Constraint PK_RnopTable_OPREF PRIMARY KEY CLUSTERED (OPREF,OPRUN,OPNO) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   'No cascading
   sSql = "ALTER TABLE RnopTable ADD CONSTRAINT FK_RnopTable_RunsTable FOREIGN KEY (OPREF,OPRUN) References RunsTable"
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX ShopTable.ShpRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPREF CHAR(12) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE ShopTable ADD Constraint PK_ShopTable_OPREF PRIMARY KEY CLUSTERED (SHPREF) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX WcntTable.WcnRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP INDEX WcntTable.WcnShop"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNREF CHAR(12) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSHOP CHAR(12) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE WcntTable ADD Constraint PK_WcntTable_WCNREF PRIMARY KEY CLUSTERED (WCNREF,WCNSHOP) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   sSql = "ALTER TABLE WcntTable ADD CONSTRAINT FK_WcntTable_ShopTable FOREIGN KEY (WCNSHOP) References ShopTable ON UPDATE CASCADE"
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   sSql = "DROP INDEX CoclTable.CocRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE CoclTable ALTER COLUMN COCREF CHAR(8) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE CoclTable ALTER COLUMN COCDAY SMALLINT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE CoclTable ADD Constraint PK_CoclTable_COCREF PRIMARY KEY CLUSTERED (COCREF,COCDAY) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX CoclTable.CocRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE CoclTable ALTER COLUMN COCREF CHAR(8) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE CoclTable ALTER COLUMN COCDAY SMALLINT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE CoclTable ADD Constraint PK_CoclTable_COCREF PRIMARY KEY CLUSTERED (COCREF,COCDAY) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "DROP INDEX CctmTable.CalRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE CctmTable ALTER COLUMN CALREF CHAR(8) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE CctmTable ADD Constraint PK_CctmTable_CALREF PRIMARY KEY CLUSTERED (CALREF) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "DROP INDEX WcclTable.WcnRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCREF CHAR(8) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHOP CHAR(12) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCCENTER CHAR(12) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCDAY SMALLINT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE WcclTable ADD Constraint PK_WcclTable_COCREF PRIMARY KEY CLUSTERED (WCCREF,WCCSHOP,WCCCENTER,WCCDAY) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "DROP INDEX RnalTable.AllRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP INDEX RnalTable.RaRun"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP INDEX RnalTable.RaSo"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnalTable ALTER COLUMN RAREF CHAR(30) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnalTable ALTER COLUMN RARUN INTEGER NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnalTable ALTER COLUMN RASO INTEGER NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnalTable ALTER COLUMN RASOITEM SMALLINT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnalTable ALTER COLUMN RASOREV CHAR(2) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnalTable ADD Constraint PK_RnalTable_ALLOCATIONREF PRIMARY KEY CLUSTERED (RAREF,RARUN,RASO,RASOITEM,RASOREV) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnalTable ADD Constraint PK_RnalTable_RunsTable FOREIGN KEY (RAREF,RARUN) References RunsTable ON UPDATE CASCADE"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnalTable ADD Constraint PK_RnalTable_PartTable FOREIGN KEY (RAREF) References PartTable"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP INDEX RndlTable.DlsRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RndlTable ALTER COLUMN RUNDLSNUM SMALLINT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RndlTable ALTER COLUMN RUNDLSRUNREF CHAR(30) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RndlTable ALTER COLUMN RUNDLSRUNNO INTEGER NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RndlTable ADD Constraint PK_RndlTable_ALLOCATIONREF PRIMARY KEY CLUSTERED (RUNDLSNUM,RUNDLSRUNREF,RUNDLSRUNNO) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE RndlTable ADD CONSTRAINT FK_RndlTable_PartTable FOREIGN KEY (RUNDLSRUNREF) References PartTable ON UPDATE CASCADE ON DELETE CASCADE"
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX PohdTable.PohdRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE PohdTable ALTER COLUMN PONUMBER INT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE PohdTable ADD Constraint PK_PohdTable_PONUMBER PRIMARY KEY CLUSTERED (PONUMBER) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   sSql = "DELETE FROM PoitTable " & vbCr _
          & "FROM PoitTable LEFT JOIN PohdTable ON PoitTable.PINUMBER = PohdTable.PONUMBER " & vbCr _
          & "WHERE (PohdTable.PONUMBER Is Null)"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP INDEX PoitTable.PoitRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE PoitTable ALTER COLUMN PINUMBER INT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE PoitTable ALTER COLUMN PIITEM SMALLINT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE PoitTable ALTER COLUMN PIREV CHAR(2) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE PoitTable ADD Constraint PK_PoitTable_PINUMBER PRIMARY KEY CLUSTERED (PINUMBER,PIITEM,PIREV) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE PoitTable ADD CONSTRAINT FK_PoitTable_PohdTable FOREIGN KEY (PINUMBER) References PohdTable ON UPDATE CASCADE ON DELETE CASCADE"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP INDEX MopkTable.PickRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP INDEX MopkTable.PkRecord"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MopkTable ALTER COLUMN PKMOPART CHAR(30) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MopkTable ALTER COLUMN PKMORUN INT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MopkTable ALTER COLUMN PKRECORD SMALLINT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MopkTable ADD Constraint PK_MopkTable_MOPICK PRIMARY KEY CLUSTERED (PKMOPART,PKMORUN,PKRECORD) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE MopkTable ADD CONSTRAINT FK_MopkTable_PartTable FOREIGN KEY (PKPARTREF) References PartTable"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MopkTable ADD CONSTRAINT FK_MopkTable_RunsTable FOREIGN KEY (PKMOPART,PKMORUN) References RunsTable"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RnopTable ADD CONSTRAINT FK_RnopTable_WcntTable FOREIGN KEY (OPCENTER,OPSHOP) References WcntTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RtopTable ADD CONSTRAINT FK_RtopTable_WcntTable FOREIGN KEY (OPCENTER,OPSHOP) References WcntTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   sSql = "DELETE FROM BuyrTable WHERE BYREF IS NULL"
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX BuyrTable.BuyerRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuyrTable ALTER COLUMN BYREF CHAR(20) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuyrTable ADD Constraint PK_BuyrTable_BUYERID PRIMARY KEY CLUSTERED (BYREF) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   
   Err.Clear
   sSql = "DELETE FROM BuycTable " & vbCr _
          & "FROM BuycTable LEFT JOIN BuyrTable ON BuycTable.BYREF = BuyrTable.BYREF " & vbCr _
          & "WHERE (BuyrTable.BYREF Is Null)"
   clsADOCon.ExecuteSql sSql
   
   sSql = "DROP INDEX BuycTable.BuyercRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuycTable ALTER COLUMN BYREF CHAR(20) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuycTable ALTER COLUMN BYPRODCODE CHAR(6) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuycTable ADD Constraint PK_BuycTable_BUYERID PRIMARY KEY CLUSTERED (BYREF,BYPRODCODE) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "DELETE FROM BuypTable " & vbCr _
          & "FROM BuypTable LEFT JOIN BuyrTable ON BuypTable.BYREF = BuyrTable.BYREF " & vbCr _
          & "WHERE (BuyrTable.BYREF Is Null)"
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX BuypTable.BuyerpRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuypTable ALTER COLUMN BYREF CHAR(20) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuypTable ALTER COLUMN BYPARTNUMBER CHAR(30) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuypTable ADD Constraint PK_BuypTable_BUYERID PRIMARY KEY CLUSTERED (BYREF,BYPARTNUMBER) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "DELETE FROM BuyvTable " & vbCr _
          & "FROM BuyvTable LEFT JOIN BuyrTable ON BuyvTable.BYREF = BuyrTable.BYREF " & vbCr _
          & "WHERE (BuyrTable.BYREF Is Null)"
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "DROP INDEX BuyvTable.BuyervRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuyvTable ALTER COLUMN BYREF CHAR(20) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuyvTable ALTER COLUMN BYVENDOR CHAR(10) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuyvTable ADD Constraint PK_BuyvTable_BUYERID PRIMARY KEY CLUSTERED (BYREF,BYVENDOR) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE BuycTable ADD CONSTRAINT FK_BuycTable_BuyrTable FOREIGN KEY (BYREF) References BuyrTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuypTable ADD CONSTRAINT FK_BuypTable_BuyrTable FOREIGN KEY (BYREF) References BuyrTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuyvTable ADD CONSTRAINT FK_BuyvTable_BuyrTable FOREIGN KEY (BYREF) References BuyrTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuycTable ADD CONSTRAINT FK_BuycTable_PcodTable FOREIGN KEY (BYPRODCODE) References PcodTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuypTable ADD CONSTRAINT FK_BuypTable_PartTable FOREIGN KEY (BYPARTNUMBER) References PartTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE BuyvTable ADD CONSTRAINT FK_BuyvTable_VndrTable FOREIGN KEY (BYVENDOR) References VndrTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   sSql = "DELETE FROM RfvdTable WHERE RFVENDOR='' OR RFVENDOR IS NULL"
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   sSql = "DROP INDEX RfvdTable.RfvdTable_Unique"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RfvdTable ALTER COLUMN RFNO CHAR(12) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RfvdTable ALTER COLUMN RFITNO INT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RfvdTable ALTER COLUMN RFVENDOR CHAR(10) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE RfvdTable ADD Constraint PK_RfvdTable_RFREV PRIMARY KEY CLUSTERED (RFNO,RFITNO,RFVENDOR) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE RfvdTable ADD CONSTRAINT FK_RfvdTable_VndrTable FOREIGN KEY (RFVENDOR) References VndrTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   sSql = "DROP INDEX MrplTable.MrpRow"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MrplTable ALTER COLUMN MRP_ROW INT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MrplTable ADD Constraint PK_MrplTable_MRPREF PRIMARY KEY CLUSTERED (MRP_ROW) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE MrplTable ADD CONSTRAINT FK_MrplTable_PartTable FOREIGN KEY (MRP_PARTREF) References PartTable"
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "DROP INDEX MrppTable.MRPPartRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MrppTable ALTER COLUMN MRP_PARTREF CHAR(30) NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MrppTable ADD Constraint PK_MrppTable_MRPPARTREF PRIMARY KEY CLUSTERED (MRP_PARTREF) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "DROP INDEX MrpbTable.BillRef"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MrpbTable ALTER COLUMN MRPBILL_ORDER SMALLINT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MrpbTable ADD Constraint PK_MrpbTable_MRPPORDER PRIMARY KEY CLUSTERED (MRPBILL_ORDER) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE MrpdTable ADD CONSTRAINT FK_MrpbTable_PartTable FOREIGN KEY (MRPBILL_PARTREF) References PartTable"
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "DROP INDEX MrpdTable.MrpDateIdx"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MrpdTable ALTER COLUMN MRP_ROW INT NOT NULL"
   clsADOCon.ExecuteSql sSql
   
   sSql = "ALTER TABLE MrpdTable ADD Constraint PK_MrpdTable_MRPDATE PRIMARY KEY CLUSTERED (MRP_ROW) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   Exit Sub
   
KeysErr1:
   On Error Resume Next
   clsADOCon.RollbackTrans
   
   
End Sub

'ShopTable
'WcntTable
'RunsTable
'See ConvertRunOps

Private Sub ConvertProductionColumns()
   Dim bBadCol As Byte
   Dim sconstraint As String
   'Start ShopTable //Test the first one and bail if not Real (see Else)
   On Error Resume Next
   'SHPRATE
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPRATE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then 'See Else
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPRATE dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPRATE DEC(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPRATE'"
               clsADOCon.ExecuteSql sSql
            Else
               GoTo EndProc
            End If
         End If
      End If
   End With
   'SHPOH
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPOH"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOH dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  If Err > 0 Then
                     For Each ADOError In RdoCol.ActiveConnection.Errors
                        sconstraint = GetConstraint(ADOError.Description)
                        If sconstraint <> "" Then Exit For
                     Next ADOError
                  End If
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOH DEC(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPOH'"
               clsADOCon.ExecuteSql sSql
            End If
            
         End If
      End If
   End With
   'SHPOHTOTAL
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPOHTOTAL"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOHTOTAL dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOHTOTAL DEC(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPOHTOTAL'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SHPOHRATE
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPOHRATE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOHRATE dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOHRATE dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPOHRATE'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SHPSETUP
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPSETUP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSETUP dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSETUP dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPSETUP'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SHPUNIT
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPUNIT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPUNIT dec(5,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPUNIT dec(5,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPUNIT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SHPSECONDS
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPSECONDS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSECONDS dec(5,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSECONDS dec(5,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPSECONDS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SHPSUHRS
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPSUHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSUHRS dec(5,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSUHRS dec(5,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPSUHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SHPUNITHRS
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPUNITHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPUNITHRS dec(5,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPUNITHRS dec(5,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPUNITHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SHPQHRS
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPQHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPQHRS dec(5,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPQHRS dec(5,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPQHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SHPMHRS
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPMHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPMHRS dec(5,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPMHRS dec(5,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPMHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SHPESTRATE
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPESTRATE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPESTRATE dec(5,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPESTRATE dec(5,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPESTRATE'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'End ShopTable
   Err.Clear
   'Start WcntTable
   'WCNOHFIXED
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNOHFIXED"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNOHFIXED dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNOHFIXED dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNOHFIXED'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNOHPCT
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNOHPCT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNOHPCT dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNOHPCT dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNOHPCT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSTDRATE
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSTDRATE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSTDRATE dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSTDRATE dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSTDRATE'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSUHRS
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUHRS dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUHRS DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNUNITHRS
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNUNITHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNUNITHRS dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNUNITHRS DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNUNITHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNESTRATE
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNESTRATE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNESTRATE dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNESTRATE dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNESTRATE'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNQHRS
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNQHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNQHRS dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNQHRS DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNQHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNMHRS
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMHRS dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMHRS DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSUNHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSUNHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSUNHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSUNHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNMONHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNMONHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNMONHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNMONHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTUEHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTUEHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTUEHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTUEHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNWEDHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNWEDHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNWEDHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNWEDHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTHUHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTHUHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTHUHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTHUHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNFRIHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNFRIHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNFRIHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNFRIHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSATHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSATHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSATHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSATHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSUNMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNMU1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSUNMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNMU2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSUNMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNMU3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSUNMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNMU4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNMONMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONMU1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNMONMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONMU2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNMONMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONMU3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNMONMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONMU4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTUEMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEMU1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTUEMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEMU2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTUEMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEMU3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTUEMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEMU4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNWEDMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDMU1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNWEDMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDMU2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNWEDMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDMU3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNWEDMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDMU4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTHUMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUMU1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTHUMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUMU2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTHUMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUMU3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNTHUMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUMU4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNFRIMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIMU1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNFRIMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIMU2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNFRIMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIMU3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNFRIMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIMU4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSATMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU1 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU1 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATMU1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSATMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU2 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU2 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATMU2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSATMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU3 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU3 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATMU3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCNSATMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU4 dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU4 DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATMU4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'End WcntTable
   'Start RunsTable
   'RUNMATL
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNMATL"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNMATL dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNMATL dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNMATL'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNLABOR
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNLABOR"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNLABOR dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNLABOR dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNLABOR'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNSTDCOST
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNSTDCOST"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNSTDCOST dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNSTDCOST dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNSTDCOST'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
    ConvertProductionColumns2
    
EndProc:
   On Error Resume Next
   'Update Preferences
   sSql = "UPDATE Preferences SET ProdtoDecimalConvDate='" & Format(Now, "mm/dd/yy") & "' " _
          & "WHERE (ProdtoDecimalConvDate IS NULL AND PreRecord=1)"
   clsADOCon.ExecuteSql sSql
   RdoCol.Close
   
End Sub



Private Sub ConvertProductionColumns2()
   Dim bBadCol As Byte
   Dim sconstraint As String
   'Start ShopTable //Test the first one and bail if not Real (see Else)
   On Error Resume Next


   'RUNQTY
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNPKQTY
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNPKQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNPKQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNPKQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNPKQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNEXP
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNEXP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNEXP dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNEXP dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNEXP'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNYIELD
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNYIELD"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNYIELD dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNYIELD DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNYIELD'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNBUDLAB
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNBUDLAB"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDLAB dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDLAB DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNBUDLAB'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNBUDMAT
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNBUDMAT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDMAT dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDMAT DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNBUDMAT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNBUDEXP
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNBUDEXP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDEXP dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDEXP DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNBUDEXP'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNBUDOH
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNBUDOH"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDOH dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDOH DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNBUDOH'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNBUDHRS
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNBUDHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDHRS dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDHRS dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNBUDHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNCHARGED
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCHARGED"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCHARGED dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCHARGED DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCHARGED'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With

   'RUNCOST
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCOST"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCOST dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCOST DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCOST'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNOHCOST
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNOHCOST"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNOHCOST dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNOHCOST DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNOHCOST'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNCMATL
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCMATL"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCMATL dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCMATL DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCMATL'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNCEXP
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCEXP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCEXP dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCEXP DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCEXP'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNCHRS
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCHRS dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCHRS dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNCLAB
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCLAB"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCLAB dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCLAB DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCLAB'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNPARTIALQTY
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNPARTIALQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNPARTIALQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNPARTIALQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNPARTIALQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNSCRAP
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNSCRAP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNSCRAP dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNSCRAP dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNSCRAP'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNREWORK
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNREWORK"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNREWORK dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNREWORK dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNREWORK'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'RUNREMAININGQTY
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNREMAININGQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNREMAININGQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNREMAININGQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNREMAININGQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'End RunsTable
   ConvertRunOPs
   ConvertCalendarColumns

End Sub



Private Function GetConstraint(sDescription As String) As String
   Dim bByte As Byte
   GetConstraint = ""
   bByte = InStr(1, sDescription, "DF_")
   If bByte > 0 Then
      GetConstraint = Mid$(sDescription, bByte, Len(sDescription))
      bByte = InStr(1, GetConstraint, "'")
      If bByte > 0 Then GetConstraint = Left$(GetConstraint, bByte - 1)
   Else
      bByte = InStr(1, sDescription, "DEFZERO")
      If bByte > 0 Then
         sSql = "sp_unbindefault '" & RdoCol.Fields(2) & "." & RdoCol.Fields(3) & "'"
         clsADOCon.ExecuteSql sSql
      End If
   End If
   
   
End Function

Private Function CheckConvErrors() As Byte
   Dim iColCounter As Integer
   CheckConvErrors = 0
   For Each ADOError In RdoCol.ActiveConnection.Errors
      If Left(ADOError.Description, 5) = "22003" Then
         iColCounter = iColCounter + 1
         CheckConvErrors = 1
      End If
   Next ADOError
   
End Function


'RnopTable
'RnalTable
'MopkTable

Public Sub ConvertRunOPs()
   Dim bBadCol As Byte
   Dim sconstraint As String
   On Error Resume Next
   'Start RnopTable
   'OPSETUP
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPSETUP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSETUP dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSETUP DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPSETUP'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPUNIT
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPUNIT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPUNIT dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPUNIT DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPUNIT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPQHRS
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPQHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPQHRS dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPQHRS DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPQHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPMHRS
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPMHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPMHRS dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPMHRS DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPMHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPSVCUNIT
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPSVCUNIT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSVCUNIT dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSVCUNIT dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPSVCUNIT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPYIELD
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPYIELD"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPYIELD dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPYIELD DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPYIELD'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPSUHRS
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPSUHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSUHRS dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSUHRS DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPSUHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPUNITHRS
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPUNITHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPUNITHRS dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPUNITHRS DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPUNITHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPRUNHRS
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPRUNHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPRUNHRS dec(9,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPRUNHRS DEC(9,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPRUNHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPCHARGED
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPCHARGED"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCHARGED dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCHARGED DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPCHARGED'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPSHMUL
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPSHMUL"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSHMUL dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSHMUL DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPSHMUL'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPCOST
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPCOST"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCOST dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCOST DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPCOST'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPOHCOST
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPOHCOST"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPOHCOST dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPOHCOST DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPOHCOST'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPCONCUR
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPCONCUR"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCONCUR dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCONCUR DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPCONCUR'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPACCEPT
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPACCEPT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPACCEPT dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPACCEPT DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPACCEPT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPREJECT
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPREJECT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPREJECT dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPREJECT DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPREJECT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPSCRAP
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPSCRAP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSCRAP dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSCRAP dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPSCRAP'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'OPREWORK
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPREWORK"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then 'See Else
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPREWORK dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPREWORK dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPREWORK'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'End RnopTable
   Err.Clear
   'Start RnspTable
   'SPLIT_SPLQTY
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SPLIT_SPLORIGQTY
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLORIGQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLORIGQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLORIGQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLORIGQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SPLIT_SPLLABOR
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLLABOR"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLLABOR dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLLABOR DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLLABOR'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SPLIT_SPLOH
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLOH"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLOH dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLOH DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLOH'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SPLIT_SPLHRS
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLHRS dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLHRS dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'SPLIT_SPLEXP
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLEXP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLEXP dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLEXP dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLEXP'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'End RnspTable
   Err.Clear
   'Start RnalTable
   'RAQTY
   sSql = "sp_columns @table_name=RnalTable,@column_name=RAQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnalTable ALTER COLUMN RAQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnalTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE RnalTable ALTER COLUMN RAQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnalTable.RAQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'End RnalTable
   Err.Clear
   'Start MopkTable
   'PKPQTY
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKPQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKPQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKPQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKPQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PKAQTY
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKAQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKAQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKAQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKAQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PKAMT
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKAMT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKAMT dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKAMT DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKAMT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PKOHPCT
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKOHPCT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKOHPCT dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKOHPCT dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKOHPCT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PKORIGQTY
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKORIGQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKORIGQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKORIGQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKORIGQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PKINADDERS
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKINADDERS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKINADDERS dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKINADDERS DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKINADDERS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PKBOMQTY
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKBOMQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKBOMQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKBOMQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKBOMQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PKTOTMATL
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKTOTMATL"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTMATL dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTMATL DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKTOTMATL'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PKTOTLABOR
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKTOTLABOR"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTLABOR dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTLABOR DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKTOTLABOR'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PKTOTEXP
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKTOTEXP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTEXP dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTEXP DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKTOTEXP'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PKTOTOH
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKTOTOH"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTOH dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTOH DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKTOTOH'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PKTOTHRS
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKTOTHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTHRS dec(6,3)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTHRS dec(6,3) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKTOTHRS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'End MopkTable
   
End Sub

'PohdTable
'PoitTable
'VndrTable

Private Sub ConvertPurchasingColumns()
   Dim bBadCol As Byte
   Dim sconstraint As String
   'Start PohdTable //Test the first one and bail if not Real (see Else)
   On Error Resume Next
   'PODISCOUNT
   sSql = "sp_columns @table_name=PohdTable,@column_name=PODISCOUNT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then 'See Else
               sSql = "ALTER TABLE PohdTable ALTER COLUMN PODISCOUNT dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PohdTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PohdTable ALTER COLUMN PODISCOUNT dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PohdTable.PODISCOUNT'"
               clsADOCon.ExecuteSql sSql
            Else
               GoTo EndProc
            End If
         End If
      End If
   End With
   'End PohdTable
   Err.Clear
   'Start PoitTable
   'PIPQTY
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIPQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIPQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIPQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIPQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PIAQTY
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIAQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIAQTY dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIAQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIAQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PIAMT
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIAMT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIAMT dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIAMT DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIAMT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PIESTUNIT
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIESTUNIT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIESTUNIT dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIESTUNIT DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIESTUNIT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PIADDERS
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIADDERS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIADDERS dec(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIADDERS DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIADDERS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PILOT
   sSql = "sp_columns @table_name=PoitTable,@column_name=PILOT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PILOT DEC(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PILOT DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PILOT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PIFRTADDERS
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIFRTADDERS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIFRTADDERS dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIFRTADDERS dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIFRTADDERS'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PIREJECTED
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIREJECTED"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIREJECTED DEC(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIREJECTED DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIREJECTED'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PIWASTE
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIWASTE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIWASTE dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIWASTE dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIWASTE'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PIORIGSCHEDQTY
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIORIGSCHEDQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIORIGSCHEDQTY DEC(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIORIGSCHEDQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIORIGSCHEDQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PIONDOCKQTYACC
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIONDOCKQTYACC"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIONDOCKQTYACC DEC(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIONDOCKQTYACC DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIONDOCKQTYACC'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PIONDOCKQTYREJ
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIONDOCKQTYREJ"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIONDOCKQTYREJ DEC(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIONDOCKQTYREJ DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIONDOCKQTYREJ'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'PIODDELQTY
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIODDELQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIODDELQTY DEC(12,4)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIODDELQTY DEC(12,4) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIODDELQTY'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'End PoitTable
   Err.Clear
   'Start VndrTable
   'VEDISCOUNT
   sSql = "sp_columns @table_name=VndrTable,@column_name=VEDISCOUNT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE VndrTable ALTER COLUMN VEDISCOUNT DEC(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE VndrTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE VndrTable ALTER COLUMN VEDISCOUNT DEC(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'VndrTable.VEDISCOUNT'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   
   
EndProc:
   On Error Resume Next
   'Update Preferences
   sSql = "UPDATE Preferences SET PurctoDecimalConvDate='" & Format(Now, "mm/dd/yy") & "' " _
          & "WHERE (PurctoDecimalConvDate IS NULL AND PreRecord=1)"
   clsADOCon.ExecuteSql sSql
   RdoCol.Close
   
End Sub

Private Sub ConvertCalendarColumns()
   Dim bBadCol As Byte
   Dim sconstraint As String
   'Start WcclTable
   'WCCSHH1
   On Error Resume Next
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHH1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then 'see Else
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH1 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH1 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHH1'"
               clsADOCon.ExecuteSql sSql
            Else
               Exit Sub
            End If
         End If
      End If
   End With
   'WCCSHH2
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHH2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH2 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH2 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHH2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCCSHH3
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHH3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH3 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH3 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHH3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCCSHH4
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHH4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH4 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH4 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHH4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCCSHR1
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR1 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR1 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCCSHR2
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR2 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR2 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCCSHR3
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR3 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR3 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'WCCSHR4
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR4 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR4 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'End WcclTable
   Err.Clear
   'Start CctmTable
   'CALSUNHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSUNHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR1 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR1 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSUNHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALSUNHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSUNHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR2 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR2 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSUNHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALSUNHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSUNHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR3 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR3 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSUNHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALSUNHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSUNHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR4 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR4 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSUNHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALMONHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALMONHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR1 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR1 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALMONHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALMONHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALMONHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR2 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR2 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALMONHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALMONHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALMONHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR3 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR3 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALMONHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALMONHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALMONHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR4 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR4 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALMONHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALTUEHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTUEHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR1 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR1 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTUEHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALTUEHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTUEHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR2 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR2 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTUEHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALTUEHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTUEHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR3 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR3 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTUEHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALTUEHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTUEHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR4 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR4 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTUEHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALWEDHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALWEDHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR1 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR1 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALWEDHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALWEDHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALWEDHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR2 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR2 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALWEDHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALWEDHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALWEDHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR3 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR3 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALWEDHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALWEDHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALWEDHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR4 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR4 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALWEDHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALTHUHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTHUHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR1 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR1 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTHUHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALTHUHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTHUHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR2 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR2 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTHUHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALTHUHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTHUHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR3 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR3 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTHUHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALTHUHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTHUHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR4 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR4 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTHUHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALFRIHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALFRIHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR1 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR1 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALFRIHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALFRIHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALFRIHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR2 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR2 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALFRIHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALFRIHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALFRIHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR3 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR3 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALFRIHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALFRIHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALFRIHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR4 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR4 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALFRIHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALSATHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSATHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR1 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR1 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSATHR1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALSATHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSATHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR2 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR2 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSATHR2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALSATHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSATHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR3 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR3 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSATHR3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'CALSATHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSATHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR4 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR4 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSATHR4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'End CctmTable
   Err.Clear
   'Start CoclTable
   'COCSHT1
   sSql = "sp_columns @table_name=CoclTable,@column_name=COCSHT1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT1 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CoclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT1 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CoclTable.COCSHT1'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'COCSHT2
   sSql = "sp_columns @table_name=CoclTable,@column_name=COCSHT2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT2 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CoclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT2 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CoclTable.COCSHT2'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'COCSHT3
   sSql = "sp_columns @table_name=CoclTable,@column_name=COCSHT3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT3 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CoclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT3 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CoclTable.COCSHT3'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'COCSHT4
   sSql = "sp_columns @table_name=CoclTable,@column_name=COCSHT4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT4 dec(7,2)"
               clsADOCon.ExecuteSql sSql
               If Err > 0 Then
                  For Each ADOError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(ADOError.Description)
                     If sconstraint <> "" Then Exit For
                  Next ADOError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CoclTable DROP " & sconstraint
                  clsADOCon.ExecuteSql sSql
               End If
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT4 dec(7,2) "
               clsADOCon.ExecuteSql sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CoclTable.COCSHT4'"
               clsADOCon.ExecuteSql sSql
            End If
         End If
      End If
   End With
   'End CoclTable
   
   
End Sub

'Moved here 10/6/06 Leave in

Private Sub CheckTriggers()
   Err.Clear
   '5/16/06 Delete Trigger (Runs)
   sSql = "CREATE TRIGGER DT_RunsTable ON RunsTable" & vbCr _
          & "FOR  DELETE " & vbCr _
          & "  AS " & vbCr _
          & "  SAVE TRANSACTION SaveRows " & vbCr _
          & "  Rollback TRANSACTION"
   clsADOCon.ExecuteSql sSql
   
   Err.Clear
   '5/16/06 Delete Trigger (PO's)
   sSql = "CREATE TRIGGER DT_PohdTable ON PohdTable" & vbCr _
          & "FOR  DELETE " & vbCr _
          & "  AS " & vbCr _
          & "  SAVE TRANSACTION SavePoRows " & vbCr _
          & "  Rollback TRANSACTION"
   clsADOCon.ExecuteSql sSql
   
End Sub

'1/10/07 Added for reports

Public Sub GetThisVendor()
   Dim RdoRpt As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT VEBNAME FROM VndrTable WHERE VEREF='" _
          & Compress(MDISect.ActiveForm.cmbVnd) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      MDISect.ActiveForm.lblVEName = "" & Trim(RdoRpt!VEBNAME)
      ClearResultSet RdoRpt
   Else
      MDISect.ActiveForm.lblVEName = "*** A Range Of Vendors Selected ***"
   End If
   Set RdoRpt = Nothing
   Exit Sub
modErr1:
   On Error GoTo 0
   
End Sub


Public Function IsValidMORun(sMoNumber As String, iRunNo As Integer, bAllowCancelledAndCompleteRuns As Boolean, bShowMessage As Boolean) As Boolean
   Dim RdoRns As ADODB.Recordset
   Dim sRunStat As String
   
   On Error GoTo IVRErr1
       
   IsValidMORun = False
   If Len(Trim(sMoNumber)) = 0 Then
    IsValidMORun = True
    
    Exit Function
    End If
   MouseCursor 13
   sSql = "SELECT RUNNO, RUNSTATUS FROM RunsTable WHERE RUNREF='" & Compress(sMoNumber) & "' AND RUNNO=" & Trim(str(iRunNo))
   'If Not bAllowedCancelledRuns Then sSql = sSql & " AND RUNSTATUS NOT LIKE 'CA%'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns, ES_FORWARD)
   If bSqlRows Then sRunStat = "" & RdoRns!RUNSTATUS Else sRunStat = ""
   
   If Not bSqlRows Then
     If bShowMessage Then MsgBox "Invalid MO/Run Selected", vbOKOnly
     IsValidMORun = False
   ElseIf bSqlRows And Not bAllowCancelledAndCompleteRuns And Left(sRunStat, 1) = "C" Then
     If bShowMessage Then MsgBox "The run you selected has been cancelled", vbOKOnly
     If Left(sRunStat, 1) = "C" Then IsValidMORun = False Else IsValidMORun = True
   ElseIf bSqlRows And bAllowCancelledAndCompleteRuns Then
     IsValidMORun = True
     If bShowMessage And Left(sRunStat, 1) = "C" Then If MsgBox("This run has a status of " & sRunStat & ". Continue?", vbYesNo) = vbNo Then IsValidMORun = False
   Else
     IsValidMORun = False
   End If
   Set RdoRns = Nothing
   MouseCursor 0
   Exit Function
      
IVRErr1:
   MouseCursor 0
   On Error Resume Next
End Function

Public Function FormCount(ByVal frmName As String) As Long
    Dim frm As Form
    For Each frm In Forms
        If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
            FormCount = FormCount + 1
        End If
    Next
End Function
