Attribute VB_Name = "Esiadmn"
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/29/05 Removed all references to JetDb
'7/7/05 Dumped System Form Settings
'8/8/05 Corrected canceled KeySets
'10/31/05 Added Cur.CurrentGroup to OpenFavorite. Opens appropriate tab
'         when called from Recent/Favorites and closed.
'1/12/06  Completed renaming dialogs to be consistent with Fina
'5/25/06 Fixed ABC Report calls (Inv Man)
'6/1/06 BuildKeys
'6/23/06 Removed Threed32.ocx
'7/31/06 Tabs replaced
'8/11/06 All tabs replaced with TabStrip
Option Explicit
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public iUserIdx As Integer
Public iCurrentRec As Integer
Public iFreeIdx As Integer
Public iFreeDbf As Integer

Public y As Byte
Public bCannotOpen As Byte
Public bFoundPart As Byte

Public sCurrDate As String
Public sCurrEmployee As String
Public sCurrEstimator As String
Public sCurrForm As String
Public sPassedPart As String
Public sProgDir As String
Public sSelected As String

Public sFavorites(13) As String
Public sRecent(10) As String
Public sSession(50) As String
Public Const TENILETEST = "Tensile Test"
Public Const CHEMANALYSIS = "Chemical Analysis"
Public Const HEATTREATMENT = "Heat Treatment"
Public Const MELTERSLOG = "Melters Log"
Public Const MATCERT = "Material Cert"

'Type tUser
'    Adduser As Integer
'    Level   As Integer
'    Group1  As Integer
'    Group2  As Integer
'    Group3  As Integer
'    Group4  As Integer
'    Group5  As Integer
'    Group6  As Integer
'    Group7  As Integer
'    Group8  As Integer
'End Type
'Public User As tUser

'6/1/06 some in sales

Private Sub BuildKeys()
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   sSql = "DROP INDEX CdivTable.DivRef"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then
      sSql = "ALTER TABLE CdivTable ALTER COLUMN DIVREF CHAR(6) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE CdivTable ADD Constraint PK_CdivTable_DIVREF PRIMARY KEY CLUSTERED (DIVREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX CregTable.RegRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE CregTable ALTER COLUMN REGREF CHAR(4) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE CregTable ADD Constraint PK_CregTable_REGREF PRIMARY KEY CLUSTERED (REGREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX CntrTable.CouRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE CntrTable ALTER COLUMN COUREF CHAR(15) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE CntrTable ADD Constraint PK_CntrTable_COUREF PRIMARY KEY CLUSTERED (COUREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX PclsTable.ClassRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE PclsTable ALTER COLUMN CCREF CHAR(6) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE PclsTable ADD Constraint PK_PclsTable_CCREF PRIMARY KEY CLUSTERED (CCREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX PcodTable.CodeRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE PcodTable ALTER COLUMN PCREF CHAR(6) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE PcodTable ADD Constraint PK_PcodTable_PCREF PRIMARY KEY CLUSTERED (PCREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX SprsTable.SpRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE SprsTable ALTER COLUMN SPNUMBER CHAR(4) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE SprsTable ADD Constraint PK_SprsTable_SPNUMBER PRIMARY KEY CLUSTERED (SPNUMBER) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX StrmTable.TrmRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE StrmTable ALTER COLUMN TRMREF CHAR(2) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE StrmTable ADD Constraint PK_StrmTable_TRMREF PRIMARY KEY CLUSTERED (TRMREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX StcdTable.ComtRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE StcdTable ALTER COLUMN COMMENT_REF CHAR(20) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE StcdTable ALTER COLUMN COMMENT_ID CHAR(20) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE StcdTable ADD Constraint PK_StcdTable_COMMENTREF PRIMARY KEY CLUSTERED (COMMENT_REF,COMMENT_ID) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX StchTable.ComtRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE StchTable ALTER COLUMN COMMENT_CLASS CHAR(20) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE StchTable ADD Constraint PK_StchTable_COMMENTCLASS PRIMARY KEY CLUSTERED (COMMENT_CLASS) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      sSql = "ALTER TABLE StcdTable ADD CONSTRAINT FK_StcdTable_StchTable FOREIGN KEY (COMMENT_CLASS) References StchTable ON DELETE CASCADE ON UPDATE CASCADE"
      clsADOCon.ExecuteSQL sSql
      
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DROP INDEX CshpTable.ShipRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE CshpTable ALTER COLUMN SHIPREF CHAR(4) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE CshpTable ADD Constraint PK_CshpTable_SHIPREF PRIMARY KEY CLUSTERED (SHIPREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      sSql = "DROP INDEX CsteTable.StateRef"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE CsteTable ALTER COLUMN STATECODE CHAR(2) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE CsteTable ADD Constraint PK_CsteTable_STATECODE PRIMARY KEY CLUSTERED (STATECODE) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
      
      Err.Clear
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      sSql = "DROP INDEX TxcdTable.Index1"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE TxcdTable ALTER COLUMN TAXREF CHAR(12) NOT NULL"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE TxcdTable ADD Constraint PK_TxcdTable_TAXREF PRIMARY KEY CLUSTERED (TAXREF) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
   End If
   Exit Sub
   
KeysErr1:
   On Error Resume Next
   clsADOCon.RollbackTrans
   
End Sub

Public Function GetWorkStation() As String
   Dim intZeroPos As Integer
   Dim gintMAX_SIZE As Long
   Dim strBuf As String
   gintMAX_SIZE = 255 'Maximum buffer size
   
   strBuf = Space$(gintMAX_SIZE)
   On Error GoTo modErr1
   
   'Get the workstation and then trim the buffer to the exact length
   'returned and add a dir sep (backslash) if the API didn't return one
   '
   If GetComputerName(strBuf, gintMAX_SIZE) > 0 Then
      intZeroPos = InStr(strBuf, Chr$(0))
      If intZeroPos > 0 Then strBuf = Left$(strBuf, intZeroPos - 1)
      GetWorkStation = strBuf
   Else
      GetWorkStation = ""
   End If
   Exit Function
   
modErr1:
   On Error GoTo 0
   GetWorkStation = ""
   
End Function

'Changed to controls 11/6/04

Public Function FindVendor(ContrlCombo As Control, ControlLabel As Control) As Byte
   Dim RdoVed As ADODB.Recordset
   Dim sVendRef As String
   sVendRef = Compress(ContrlCombo)
   If Len(sVendRef) = 0 Then Exit Function
   On Error GoTo modErr1
   sSql = "SELECT VEREF,VENICKNAME,VEBNAME FROM VndrTable " _
          & "WHERE VEREF='" & sVendRef & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      On Error Resume Next
      With RdoVed
         ContrlCombo = "" & Trim(!VENICKNAME)
         ControlLabel = "" & Trim(!VEBNAME)
         ClearResultSet RdoVed
         FindVendor = 1
      End With
   Else
      On Error Resume Next
      ContrlCombo = ""
      ControlLabel = "No Valid Vendor Selected."
      FindVendor = 0
   End If
   Set RdoVed = Nothing
   Exit Function
   
modErr1:
   sProcName = "findvendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   FindVendor = 0
   DoModuleErrors MDISect.ActiveForm
   
End Function


Public Sub FindShop()
   'Use local errors
   Dim sGetShop As String
   Dim RdoShp As ADODB.Recordset
   sGetShop = Compress(MDISect.ActiveForm.cmbShp)
   If Len(sGetShop) > 0 Then
      sSql = "SELECT SHPREF,SHPNUM FROM ShopTable WHERE SHPREF='" & sGetShop & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp)
      If bSqlRows Then
         MDISect.ActiveForm.cmbShp = "" & Trim(RdoShp!SHPNUM)
      Else
         MsgBox "Shop Wasn't Found.", 48, MDISect.ActiveForm.Caption
         MDISect.ActiveForm.cmbShp = ""
      End If
      On Error Resume Next
      ClearResultSet RdoShp
      Set RdoShp = Nothing
   End If
   
End Sub

'10/7/03

Public Function GetMoOperation(MONUMBER As String, Runno As Long, OpNo As Integer) As Byte
   Dim RdoOpr As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,OPREF,OPRUN,OPNO " _
          & "FROM RunsTable,RnopTable WHERE (RUNREF=OPREF AND " _
          & "RUNNO=OPRUN) AND (RUNREF='" & MONUMBER & "' AND RUNNO=" & Runno _
          & " AND OPNO=" & OpNo & " AND RUNSTATUS<>'CA' AND " _
          & "RUNSTATUS<>'CL')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOpr, ES_FORWARD)
   If bSqlRows Then GetMoOperation = 1 Else GetMoOperation = 0
   ClearResultSet RdoOpr
   Set RdoOpr = Nothing
   Exit Function
   
modErr1:
   GetMoOperation = 0
   
End Function


'8/9/05 Close open sets

Public Sub FormUnload(Optional bDontShowForm As Byte)
   Dim iList As Integer
   'Dim iResultSets As Integer
   On Error Resume Next
   MDISect.lblBotPanel.Caption = MDISect.Caption
'   If Forms.Count < 3 Then
'      iResultSets = RdoCon.rdoResultsets.Count
'      For iList = iResultSets - 1 To 0 Step -1
'         RdoCon.rdoResultsets(iList).Close
'      Next
'   End If
   If bDontShowForm = 0 Then
      Select Case cUR.CurrentGroup
         Case "Admn"
            zGr1Admn.Show
         Case "Prod"
            zGr3Padm.Show
         Case "Sale"
            zGr2Sale.Show
         Case "Time"
            zGr4Humn.Show
         Case "Invc"
            zGr5Matl.Show
      End Select
      Erase bActiveTab
      cUR.CurrentGroup = ""
   End If
   
End Sub

Public Sub OpenFavorite(sSelected As String)
   CloseForms
   '    If sSelected = "User Profiles" Then _
   '        If User.Adduser = 0 Then sSelected = ""
   If LTrim$(sSelected) = "" Then Exit Sub
   MouseCursor 13
   On Error GoTo OpenFavErr1
   Select Case sSelected
      Case "User Profiles"
         cUR.CurrentGroup = "Admn"
         AdmnUuser2.Show
      Case "System Settings"
         cUR.CurrentGroup = "Admn"
         AdmnADe01a.Show
      Case "System Administration"
         cUR.CurrentGroup = "Admn"
         AdmnADe01a.Show
      Case "Divisions"
         cUR.CurrentGroup = "Sale"
         SadmSLe01a.Show
      Case "Regions"
         cUR.CurrentGroup = "Sale"
         SadmSLe02a.Show
      Case "Salespersons"
         cUR.CurrentGroup = "Sale"
         CommCOe01a.Show
      Case "Product Codes"
         cUR.CurrentGroup = "Prod"
         PadmPRe01a.Show
      Case "Product Classes"
         cUR.CurrentGroup = "Prod"
         PadmPRe02a.Show
      Case "Shipping Terms"
         cUR.CurrentGroup = "Sale"
         SadmSLe04a.Show
         '        Case "Employees"
         '            cur.CurrentGroup = "Time"
         '            HumnHUe03a.Show
         'Case "Employees By Name"
         '    cUR.CurrentGroup = "Time"
         '    HumnHUp01a.Show
         'Case "Employees By Number"
         '    cUR.CurrentGroup = "Time"
         '    HumnHUp02a.Show
      Case "Divisions Report"
         cUR.CurrentGroup = "Sale"
         SadmSLp01a.Show
      Case "Regions Report"
         SadmSLp02a.Show
      Case "Sales Persons Report"
         cUR.CurrentGroup = "Sale"
         SadmSLp03a.Show
         '        Case "Time Type Codes"
         '            cur.CurrentGroup = "Time"
         '            HumnHUe04a.Show
         '        Case "Time Type Codes Report"
         '            cur.CurrentGroup = "Time"
         '            HumnHUp15a.Show
      Case "Part Activity With Cost"
         cUR.CurrentGroup = "Invc"
         MatlMMp01a.Show
      Case "Part Quantity Activity"
         cUR.CurrentGroup = "Invc"
         MatlMMp02a.Show
         '        Case "Daily Time Charges"
         '            cur.CurrentGroup = "Time"
         '            HumnHUe01a.Show
         '        Case "Daily Employee Time Charges"
         '            cur.CurrentGroup = "Time"
         '            HumnHUp03a.Show
         '        Case "Delete A Daily Time Charge"
         '            cur.CurrentGroup = "Time"
         '            HumnHUf01a.Show
         '        Case "Weekly Time Charges (Report)"
         '            cur.CurrentGroup = "Time"
         '            HumnHUp05a.Show
      Case "State Codes"
         cUR.CurrentGroup = "Sale"
         SadmSLe05a.Show
         '10/25/04
         ' Case "Part Price Markup Matrix"
         '     diaPmatx.Show
      Case "Shipping Terms Report"
         cUR.CurrentGroup = "Sale"
         SadmSLp04a.Show
      Case "Product Codes Report"
         cUR.CurrentGroup = "Prod"
         PadmPRp01a.Show
      Case "Product Class Report"
         cUR.CurrentGroup = "Prod"
         PadmPRp02a.Show
      Case "Parts"
         InvcINe01a.Show
         cUR.CurrentGroup = "Invc"
      Case "Current Logons"
         cUR.CurrentGroup = "Admn"
         AdmnADf02a.Show
      Case "Delete A Product Code"
         cUR.CurrentGroup = "Prod"
         PadmPRf01a.Show
      Case "Delete A Product Class"
         cUR.CurrentGroup = "Prod"
         PadmPRf04a.Show
      Case "Change Product Codes"
         cUR.CurrentGroup = "Prod"
         PadmPRf02a.Show
      Case "Change Product Classes"
         cUR.CurrentGroup = "Prod"
         PadmPRf05a.Show
      Case "Update Product Codes By Part Type"
         cUR.CurrentGroup = "Prod"
         PadmPRf03a.Show
      Case "Update Product Classes By Part Type"
         cUR.CurrentGroup = "Prod"
         PadmPRf06a.Show
      Case "Parts By Part Number"
         cUR.CurrentGroup = "Invc"
         InvcINp01a.Show
         '        Case "Revise Daily Time Charges"
         '            cur.CurrentGroup = "Time"
         '            HumnHUe02a.Show
      Case "Country Codes"
         cUR.CurrentGroup = "Sale"
         SadmSLe06a.Show
      Case "Country Codes Report"
         cUR.CurrentGroup = "Sale"
         SadmSLp05a.Show
      Case "Standard Comments"
         cUR.CurrentGroup = "Admn"
         AdmnADe03a.Show
      Case "Standard Comments (Report)"
         cUR.CurrentGroup = "Admn"
         AdmnADp06a.Show
      Case "Custom Reports"
         cUR.CurrentGroup = "Admn"
         AdmnADf01a.Show
      Case "Custom Report Settings"
         cUR.CurrentGroup = "Admn"
         AdmnADp07a.Show
      Case "User Manager"
         If SecPw.UserAdmn = 1 Then
            AdmnUuser2.Show
         Else
            MouseCursor 0
         End If
         '        Case "User Profiles (Old)"
         '            AdmnUuser.Show
      Case "Part Locations"
         cUR.CurrentGroup = "Invc"
         InvcINe04a.Show
      Case "Inventory ABC Classes"
         cUR.CurrentGroup = "Invc"
         MatlMMe01a.Show
      Case "Part Numbers Without An ABC Class"
         cUR.CurrentGroup = "Invc"
         MatlMMp03a.Show
      Case "Inventory ABC Class Setup"
         cUR.CurrentGroup = "Invc"
         CyclCYe01a.Show
      Case "ABC Classes By Part Number"
         cUR.CurrentGroup = "Invc"
         MatlMMp04a.Show
         ' Case "WIP Test"
         '      cUR.CurrentGroup = "Invc"
         '      WipTest.Show
      Case "Workstation Lot Locations"
         cUR.CurrentGroup = "Invc"
         MatlMMe06a.Show
      Case "Ship To Locations"
         cUR.CurrentGroup = "Sale"
         SadmSLe07a.Show
      Case "Broadcast Message"
         cUR.CurrentGroup = "Admn"
         SysMessage.Show
      Case "Update Inventory Activity Standard Costs"
         cUR.CurrentGroup = "Invc"
         MatlMMf02a.Show
      Case "Standard Cost"
         cUR.CurrentGroup = "Invc"
         MatlMMe02a.Show
      Case Else
         MouseCursor 0
   End Select
   Exit Sub
   
OpenFavErr1:
   Resume OpenFavErr2
OpenFavErr2:
   MsgBox "ActiveX Error... Can't Load Form..", 48, "System    "
   On Error GoTo 0
   
End Sub

Sub GetCompanyName()
   Dim RdoGet As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT CONAME FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet)
   If bSqlRows Then
      With RdoGet
         sFacility = Trim(!CONAME)
         ClearResultSet RdoGet
      End With
   End If
   Set RdoGet = Nothing
   Exit Sub
   
modErr1:
   sProcName = "getcompna"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   sFacility = ""
   
End Sub


Sub Main()
   Dim sApptitle$
   If App.PrevInstance Then
      On Error Resume Next
      sApptitle = App.Title
      App.Title = "EsiAdmn"
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
sProgName = "Administration"
MainLoad "admn"
GetFavorites "EsiAdmn"
' MM 9/10/2009
'MM sProgName = "Administration"
MDISect.Show

End Sub









''Pick up permissions for this user
'Public Sub GetSectionPermissions()
'    Dim RdoUsr As rdoResultset
'    '11/21/06
'    User.Group1 = 1
'    User.Group2 = 1
'    User.Group3 = 1
'    User.Group4 = 1
'    User.Group5 = 1
'    User.Group6 = 1
'    cur.CurrentUser = UCase$(cur.CurrentUser)
'    On Error GoTo ModErr1
'    sSql = "SELECT USERREF,USERADDUSER,USERLEVEL," _
'        & "USERADMNGR1,USERADMNGR2,USERADMNGR3," _
'        & "USERADMNGR4,USERADMNGR5,USERADMNGR6 " _
'        & "FROM UsscTable WHERE USERREF='" _
'        & cur.CurrentUser & "'"
'    bSqlRows = GetDataSet(RdoUsr)
'        If bSqlRows Then
'            With RdoUsr
'                User.Adduser = !UserAddUser
'                User.Level = !UserLevel
'                User.Group1 = !USERADMNGR1
'                User.Group2 = !USERADMNGR2
'                User.Group3 = !USERADMNGR3
'                User.Group4 = !USERADMNGR4
'                User.Group5 = !USERADMNGR5
'                User.Group6 = !USERADMNGR6
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

'Create tables, indexes, columns, etc

Public Sub UpdateTables()
   If MDISect.bUnloading = 1 Then Exit Sub
   
   SaveSetting "Esi2000", "AppTitle", "admn", "ESI Administration"
   SysOpen.Show
   SysOpen.prg1.Visible = True
   SysOpen.pnl = "Configuration Settings."
   SysOpen.pnl.Refresh
   
   '2/15/01 number for activity table
   On Error Resume Next
   SysOpen.prg1.Value = 10
   Sleep 500
   
   'moved some to OldUpdates 12/5/02, 1/21/05, 10/6/06
   Err.Clear
   SysOpen.prg1.Value = 50
   'Moved from MOM 2/6/02
   CreateSystemEventLog
   
   '2/6/07 From MOM
   CustomReports
   InstallCustomReports
   
   'BuildKeys
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

'12/21/06 Removed Credit Managment and added 31-36 (from Administration)
'2/6/07 From MOM

Private Sub InstallCustomReports()
   Dim RdoTest As ADODB.Recordset
   Dim b As Byte
   Dim A As Integer
   Dim I As Integer
   Dim iList As Integer
   
   Dim sSection As String
   Dim sGroup As String
   On Error Resume Next
   sSql = "SELECT REPORT_INDEX FROM CustomReports WHERE REPORT_INDEX=25"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTest, ES_FORWARD)
   If Not bSqlRows Or Err > 0 Then
      'None, so we'll make some
      b = 0
      For I = 1 To 44
         Select Case I
            Case 1
               sSection = "Administration"
               sGroup = "System"
            Case 2
               sSection = "Administration"
               sGroup = "Sales"
            Case 3
               sSection = "Administration"
               sGroup = "Production Control"
            Case 4
               sSection = "Administration"
               sGroup = "Time Charges"
            Case 5
               sSection = "Administration"
               sGroup = "Inventory Management"
            Case 6
               sSection = "Sales"
               sGroup = "Order processing"
            Case 7
               sSection = "Sales"
               sGroup = "Packing Slips"
            Case 8
               sSection = "Sales"
               sGroup = "Bookings/Backlog"
            Case 9
               sSection = "Engineering"
               sGroup = "Routings"
            Case 10
               sSection = "Engineering"
               sGroup = "Bills Of Material"
            Case 11
               sSection = "Engineering"
               sGroup = "Document Control"
            Case 12
               sSection = "Engineering"
               sGroup = "Tooling"
            Case 13
               sSection = "Engineering"
               sGroup = "Estimating"
            Case 14
               sSection = "Production Control"
               sGroup = "Shop Floor Control"
            Case 15
               sSection = "Production Control"
               sGroup = "Capacity Planning"
            Case 16
               sSection = "Production Control"
               sGroup = "Purchasing"
            Case 17
               sSection = "Production Control"
               sGroup = "Time Charges"
            Case 18
               sSection = "Production Control"
               sGroup = "Material Requirements"
            Case 19
               sSection = "Production Control"
               sGroup = "Material Requisitions"
            Case 20
               sSection = "Inventory Control"
               sGroup = "Inventory"
            Case 21
               sSection = "Inventory Control"
               sGroup = "Material/Picks"
            Case 22
               sSection = "Inventory Control"
               sGroup = "Receiving"
            Case 23
               sSection = "Inventory Control"
               sGroup = "Inventory Management"
            Case 24
               sSection = "Inventory Control"
               sGroup = "Lot Tracking"
            Case 25
               sSection = "Quality Assurance"
               sGroup = "Inspection Reports"
            Case 26
               sSection = "Quality Assurance"
               sGroup = "First Article Inspection"
            Case 27
               sSection = "Quality Assurance"
               sGroup = "Statistical Process Control"
            Case 28
               sSection = "Quality Assurance"
               sGroup = "On Dock Inspection"
            Case 29
               sSection = "Financial Accounting"
               sGroup = "Accounts Receivable"
            Case 30
               sSection = "Financial Accounting"
               sGroup = "Accounts Payable"
            Case 31
               sSection = "Financial Accounting"
               sGroup = "Journals"
            Case 32
               sSection = "Financial Accounting"
               sGroup = "Job Costing"
            Case 33
               sSection = "Financial Accounting"
               sGroup = "Lot Costing"
            Case 34
               sSection = "Financial Accounting"
               sGroup = "General Ledger"
            Case 35
               sSection = "Financial Accounting"
               sGroup = "Closing"
            Case 36
               sSection = "Financial Accounting"
               sGroup = "Product Costing"
            Case Else
               b = 1
         End Select
         If b = 0 Then
            clsADOCon.BeginTrans
            For A = 1 To 10
               iList = iList + 1
               sSql = "INSERT INTO CustomReports (" _
                      & "REPORT_INDEX,REPORT_SECTION,REPORT_GROUP) " _
                      & "VALUES(" & iList & ",'" _
                      & sSection & "','" & sGroup & "')"
               clsADOCon.ExecuteSQL sSql
            Next
            clsADOCon.CommitTrans
         Else
            Exit For
         End If
      Next
   End If
   SaveSetting "ES2000", "system", "customreports", 1
   Set RdoTest = Nothing
End Sub

'12/21/06 Change REPORT_INDEX datatype and Indexes. Added Jump to Install
'2/6/07 From MOM

Public Sub CustomReports()
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT REPORT_INDEX FROM CustomReports WHERE REPORT_INDEX = 1"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum > 0 Then
      Err = 0
      clsADOCon.ADOErrNum = 0
      
      sSql = "CREATE TABLE CustomReports (" _
             & "REPORT_INDEX SMALLINT NOT NULL DEFAULT(0)," _
             & "REPORT_SECTION CHAR(30) NULL DEFAULT('')," _
             & "REPORT_GROUP CHAR(30) NULL DEFAULT('')," _
             & "REPORT_REF CHAR(12) NULL DEFAULT('')," _
             & "REPORT_NAME CHAR(12) NULL DEFAULT('')," _
             & "REPORT_CUSTOMREPORT CHAR(12) NULL DEFAULT(''))"
      clsADOCon.ExecuteSQL sSql
      
      If clsADOCon.ADOErrNum = 0 Then
         Err.Clear
         clsADOCon.ADOErrNum = 0
         sSql = "ALTER TABLE CustomReports ADD Constraint PK_CustomReports_REPORTREF PRIMARY KEY CLUSTERED (REPORT_INDEX) " _
                & "WITH FILLFACTOR=80 "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "CREATE INDEX ReportSection ON CustomReports(REPORT_SECTION) WITH FILLFACTOR = 80"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "CREATE INDEX ReportRef ON CustomReports(REPORT_REF) WITH FILLFACTOR = 80"
         clsADOCon.ExecuteSQL sSql
         
      End If
      InstallCustomReports
   End If
   
End Sub

'2/6/02 From MOM

Private Sub CreateSystemEventLog()
   'Dim sAttribs As String
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   
   sSql = "use msdb"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "select * from SystemEvents"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum > 0 Then
      Err.Clear
      clsADOCon.ADOErrNum = 0
      sSql = "CREATE TABLE SystemEvents (" _
             & "Event_ID int IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL," _
             & "Event_Date char(18) NULL DEFAULT('')," _
             & "Event_Section char(10) NULL DEFAULT('')," _
             & "Event_Form char(20) NULL DEFAULT('')," _
             & "Event_User char(30) NULL DEFAULT('')," _
             & "Event_Event int NULL DEFAULT(0)," _
             & "Event_Warning tinyint NULL DEFAULT(0)," _
             & "Event_Procedure char(20) NULL DEFAULT('')," _
             & "Event_Text char(60) NULL DEFAULT(''))"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "CREATE UNIQUE CLUSTERED INDEX EventRow ON " _
                & "SystemEvents(Event_ID) WITH FILLFACTOR = 80"
         clsADOCon.ExecuteSQL sSql
      End If
   End If
   'sAttribs = "Description=" _
   '           & "System Events" _
   '           & vbCr & "OemToAnsi=No" _
   '           & vbCr & "SERVER=" & sserver _
   '           & vbCr & "Database=msdb"
   ' Create new/revise registered DSN.
   'rdoEngine.rdoRegisterDataSource "EventLog", _
   '   "SQL Server", True, sAttribs
   
   sSql = "use " & sDataBase
   clsADOCon.ExecuteSQL sSql
   
End Sub


Public Sub OldUpdates()
   'Should be okay 8/17/02
   '    GoTo Sect4 'Should Have it 1/4/02
   '    GoTo Sect3 'Should Have it 10/11/01
   '    sSql = "Select DISTINCT INNUMBER FROM InvaTable WHERE INNUMBER=0"
   '    Set RdoNew = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   '        If Err > 0 Then
   '            If Left$(Err.Description, 5) = "S0022" Then
   '                clsAdocon.BeginTrans
   '                Err = 0
   '                sSql = "ALTER TABLE InvaTable ADD INNUMBER INT NULL " _
   '                     & "DEFAULT(0)"
   '                clsADOCon.ExecuteSQL sSQL
   '
   '                sSql = "CREATE INDEX InNum ON " _
   '                    & "dbo.InvaTable(INNUMBER) WITH FILLFACTOR = 90 "
   '                clsADOCon.ExecuteSQL sSQL
   '                    If Err = 0 Then
   '                        clsADOCon.CommitTrans
   '                        bType = 1
   '                        SysOpen.prg1.Value = 20
   '                    Else
   '                        bType = 0
   '                        clsAdocon.RollbackTrans
   '                    End If
   '            End If
   '        End If
   '
   '        If bType = 1 Then
   '            sSql = "SELECT INPART,INTYPE,INADATE FROM InvaTable " _
   '                & "ORDER BY INADATE"
   '            bSqlRows = GetDataSet(RdoNew, ES_FORWARD)
   '            If bSqlRows Then
   '                A = 30
   '                SysOpen.prg1.Value = A
   '                    With RdoNew
   '                        Do Until .EOF
   '                            I = I + 1
   '                            If I > 500 Then
   '                                I = 0
   '                                A = A + 10
   '                                If A > 90 Then A = 90
   '                                SysOpen.prg1.Value = A
   '                            End If
   '                            lCounter = lCounter + 1
   '                            If Not IsNull(!INADATE) Then sDate = Format(!INADATE, "mm/dd/yy")
   '                            If Not IsNull(!INTYPE) Then bType = !INTYPE
   '                            sPart = "" & Trim(!INPART)
   '                            sSql = "UPDATE InvaTable SET INNUMBER=" & lCounter & " WHERE " _
   '                                & "(INPART='" & sPart & "' AND INTYPE=" & bType & " AND " _
   '                                & "INADATE='" & sDate & "')"
   '                            clsADOCon.ExecuteSQL sSQL
   '                            .MoveNext
   '                        Loop
   '                        .Cancel
   '                        sSql = "UPDATE InvaTable SET INNUMBER=0 WHERE " _
   '                            & "INNUMBER is null"
   '                        clsADOCon.ExecuteSQL sSQL
   '                        SysOpen.prg1.Value = 90
   '                    End With
   '                End If
   '        End If
   'Sect1:
   '    sSql = "Select DEFFLOWTIME FROM Preferences WHERE PreRecord=1"
   '    Set RdoNew = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   '        If Err > 0 Then
   '            If Left$(Err.Description, 5) = "S0022" Then
   '                clsAdocon.BeginTrans
   '                Err = 0
   '                sSql = "ALTER TABLE Preferences ADD DEFFLOWTIME SMALLINT NULL " _
   '                     & "DEFAULT(0),DEFLEADTIME SMALLINT NULL DEFAULT(0)"
   '                clsADOCon.ExecuteSQL sSQL
   '                    If Err = 0 Then
   '                        clsADOCon.CommitTrans
   '                        SysOpen.prg1.Value = 50
   '                        sSql = "UPDATE Preferences SET DEFFLOWTIME=30, DEFLEADTIME=30 " _
   '                            & "WHERE PreRecord=1"
   '                        clsADOCon.ExecuteSQL sSQL
   '                    Else
   '                        clsAdocon.RollbackTrans
   '                    End If
   '            End If
   '        End If
   'Sect2:
   '    '6/7/01
   '    sSql = "Select EstMatlBurden FROM Preferences WHERE PreRecord=1"
   '    Set RdoNew = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   '    SysOpen.prg1.Value = 20
   '        If Err > 0 Then
   '            If Left$(Err.Description, 5) = "S0022" Then
   '                clsAdocon.BeginTrans
   '                Err = 0
   '                sSql = "ALTER TABLE Preferences ADD EstMatlBurden REAL NULL " _
   '                     & "DEFAULT(0),EstFactoryOverHead REAL NULL DEFAULT(0)," _
   '                     & "EstGenAdmnExp REAL NULL DEFAULT(0)," _
   '                     & "EstProfitOfSale REAL NULL DEFAULT(0)," _
   '                     & "EstUseWCOverhead TINYINT NULL DEFAULT(0) "
   '                clsADOCon.ExecuteSQL sSQL
   '                    If Err = 0 Then
   '                        clsADOCon.CommitTrans
   '                        SysOpen.prg1.Value = 30
   '                        sSql = "UPDATE Preferences SET EstMatlBurden=0, " _
   '                            & "EstFactoryOverHead=0," _
   '                            & "EstGenAdmnExp=0," _
   '                            & "EstProfitOfSale=0," _
   '                            & "EstUseWCOverhead=0 " _
   '                            & "WHERE PreRecord=1"
   '                        clsADOCon.ExecuteSQL sSQL
   '                    Else
   '                        clsAdocon.RollbackTrans
   '                    End If
   '            End If
   '        End If
   '   '6/11/01
   '    sSql = "Select EstLaborRate FROM Preferences WHERE PreRecord=1"
   '    Set RdoNew = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   '    SysOpen.prg1.Value = 20
   '        If Err > 0 Then
   '            If Left$(Err.Description, 5) = "S0022" Then
   '                clsAdocon.BeginTrans
   '                Err = 0
   '                sSql = "ALTER TABLE Preferences ADD EstLaborRate REAL NULL "
   '                clsADOCon.ExecuteSQL sSQL
   '                    If Err = 0 Then
   '                        clsADOCon.CommitTrans
   '                        SysOpen.prg1.Value = 30
   '                        sSql = "UPDATE Preferences SET EstLaborRate=0 " _
   '                            & "WHERE PreRecord=1"
   '                        clsADOCon.ExecuteSQL sSQL
   '                    Else
   '                        clsAdocon.RollbackTrans
   '                    End If
   '            End If
   '        End If
   'Sect3:
   '   '10/11/01 especially for SQL Server 2000
   '    On Error Resume Next
   '    If bVersion > 6 Then
   '        sSql = "DROP INDEX PclsTable.CcCode"
   '        clsADOCon.ExecuteSQL sSQL
   '            If Err = 0 Then
   '                sSql = "CREATE UNIQUE CLUSTERED INDEX ClassRef ON " _
   '                    & "PclsTable(CCREF) WITH FILLFACTOR = 80 "
   '                clsADOCon.ExecuteSQL sSQL
   '            End If
   '
   '        sSql = "DROP INDEX PcodTable.PcCode"
   '        clsADOCon.ExecuteSQL sSQL
   '            If Err = 0 Then
   '                sSql = "DELETE FROM PcodTable WHERE PCREF='BID'"
   '                clsADOCon.ExecuteSQL sSQL
   '
   '                sSql = "CREATE UNIQUE CLUSTERED INDEX CodeRef ON " _
   '                    & "PcodTable(PCREF) WITH FILLFACTOR = 80 "
   '                clsADOCon.ExecuteSQL sSQL
   '
   '                sSql = "INSERT INTO PcodTable (PCREF,PCCODE,PCDESC) " _
   '                    & "VALUES('BID','BID','Estimating Parts')"
   '                clsADOCon.ExecuteSQL sSQL
   '
   '            End If
   '    End If
   'Sect4:
   '    On Error Resume Next
   '    SysOpen.prg1.Value = 30
   '    sSql = "Select COAPDISC FROM ComnTable WHERE COREF=1"
   '    Set RdoNew = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   '    SysOpen.prg1.Value = 20
   '        If Err > 0 Then
   '            Err = 0
   '            sSql = "ALTER TABLE ComnTable ADD COAPDISC tinyint NULL DEFAULT(0)"
   '            clsADOCon.ExecuteSQL sSQL
   '                If Err = 0 Then
   '                    clsADOCon.CommitTrans
   '                    SysOpen.prg1.Value = 40
   '                    sSql = "UPDATE ComnTable SET COAPDISC=0 "
   '                    clsADOCon.ExecuteSQL sSQL
   '                Else
   '                    clsAdocon.RollbackTrans
   '                End If
   '        End If
   '    '3/12/02 Lots
   '    SysOpen.prg1.Value = 40
   '    sSql = "Select COLOTSACTIVE FROM ComnTable WHERE COREF=1"
   '    Set RdoNew = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   '        If Err > 0 Then
   '            Err = 0
   '            sSql = "ALTER TABLE ComnTable ADD COLOTSACTIVE tinyint NULL DEFAULT(0)," _
   '                & "COLOTSFIFO tinyint NULL default(0)"
   '            clsADOCon.ExecuteSQL sSQL
   '                If Err = 0 Then
   '                    clsADOCon.CommitTrans
   '                    SysOpen.prg1.Value = 50
   '                    sSql = "UPDATE ComnTable SET COLOTSACTIVE=0," _
   '                     & "COLOTSFIFO=1"
   '                    clsADOCon.ExecuteSQL sSQL
   '                Else
   '                    clsAdocon.RollbackTrans
   '                End If
   '        End If
   '9/18/02 Add Transfer fee
   '     sSql = "Select COTRANSFEEACCT FROM ComnTable WHERE COREF=1"
   '     Set RdoTest = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   '         If Err > 0 Then
   '             Err = 0
   '             sSql = "ALTER TABLE ComnTable ADD COTRANSFEEACCT CHAR(12) NULL DEFAULT('') "
   '             clsADOCon.ExecuteSQL sSQL
   '                 If Err = 0 Then
   '                     clsADOCon.CommitTrans
   '                     SysOpen.prg1.Value = 30
   '                     sSql = "UPDATE ComnTable SET COTRANSFEEACCT='' " _
   '                        & "WHERE COTRANSFEEACCT IS NULL"
   '                     clsADOCon.ExecuteSQL sSQL
   '                 Else
   '                     clsAdocon.RollbackTrans
   '                 End If
   '         End If
   '    SysOpen.prg1.Value = 40
   '    '11/05/02 update the TchdTable to include the actual date of the
   '    'time card.  We don't have it, so we'll use the next best thing
   '    sSql = "UPDATE TchdTable SET TMDATE=TMDAY " _
   '       & "WHERE TMDATE IS NULL"
   '    clsADOCon.ExecuteSQL sSQL
   '
   '    '2/14/03 Add SO and B&O Tax accounts
   '     SysOpen.prg1.Value = 60
   '     sSql = "SELECT CUBOTAXACCT FROM CustTable WHERE CUBOTAXACCT=''"
   '     Set RdoTest = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
   '         If Err > 0 Then
   '             Err = 0
   '             sSql = "ALTER TABLE CustTable ADD CUBOTAXACCT CHAR(12) NULL " _
   '                 & "DEFAULT('')," _
   '                 & "CUSALESACCT CHAR(12) NULL DEFAULT('')"
   '             clsADOCon.ExecuteSQL sSQL
   '             If Err = 0 Then
   '                sSql = "UPDATE CustTable SET CUBOTAXACCT=''," _
   '                    & "CUSALESACCT='' WHERE CUBOTAXACCT IS NULL "
   '                 clsADOCon.ExecuteSQL sSQL
   '             End If
   '        End If
   '        '7//9/03
   '        sSql = "SELECT TCOHFIXED FROM Tcitable WHERE TCOHFIXED=0"
   '        Set RdoTest = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
   '         If Err > 0 Then
   '             Err = 0
   '             sSql = "ALTER TABLE TcitTable ADD TCOHFIXED REAL NULL " _
   '                 & "DEFAULT(0)"
   '             clsADOCon.ExecuteSQL sSQL
   '             If Err = 0 Then
   '                sSql = "UPDATE TcitTable SET TCOHFIXED=0 " _
   '                    & "WHERE TCOHFIXED IS NULL "
   '                 clsADOCon.ExecuteSQL sSQL
   '             End If
   '        End If
   
   '   '10/06/03 Add Scrap and Rework info to MO's
   '    Err = 0
   '    sSql = "SELECT RUNSCRAP FROM RunsTable WHERE RUNSCRAP=999"
   '    bSqlRows = GetDataSet(RdoTest, ES_FORWARD)
   '        If Err = 40002 Then
   '            Err = 0
   '            clsAdocon.BeginTrans
   '            sSql = "ALTER TABLE RunsTable ADD " _
   '                & "RUNSCRAP REAL NULL DEFAULT(0)," _
   '                & "RUNREWORK REAL NULL DEFAULT(0)," _
   '                & "RUNREMAININGQTY REAL NULL DEFAULT(0) "
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "ALTER TABLE RnopTable ADD " _
   '                & "OPREWORK REAL NULL DEFAULT(0) "
   '            clsADOCon.ExecuteSQL sSQL
   '                If Err = 0 Then
   '                    clsADOCon.CommitTrans
   '                    sSql = "UPDATE RunsTable SET RUNSCRAP=0," _
   '                        & "RUNREWORK=0, RUNREMAININGQTY=(RUNQTY-RUNPARTIALQTY) " _
   '                        & "WHERE RUNREMAININGQTY IS NULL"
   '                    clsADOCon.ExecuteSQL sSQL
   '
   '                    sSql = "UPDATE RnopTable SET OPREWORK=0," _
   '                        & "OPACCEPT=OPYIELD WHERE OPREWORK IS NULL"
   '                    clsADOCon.ExecuteSQL sSQL
   '                Else
   '                    clsAdocon.RollbackTrans
   '                End If
   '        End If
   'End This Batch
   
   'Start 1/21/05
   '    '11/12/04 Verify Invoices for Closing MO's (those without accounting)
   '    Err = 0
   '    sSql = "SELECT COVERIFYINVOICES FROM ComnTable WHERE COVERIFYINVOICES=1"
   '    Set RdoTest = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
   '        If Err > 0 Then
   '            Err = 0
   '            sSql = "ALTER TABLE ComnTable ADD COVERIFYINVOICES tinyint NULL DEFAULT(1)"
   '            clsADOCon.ExecuteSQL sSQL
   '            If Err = 0 Then
   '                sSql = "UPDATE ComnTable SET COVERIFYINVOICES=1"
   '                clsADOCon.ExecuteSQL sSQL
   '            End If
   '        End If
   '    '1/20/05 Verify over/short acct (Larry - the only one found)
   '    Err.Clear
   '    sSql = "SELECT COADJACCT FROM ComnTable WHERE COADJACCT='foobar'"
   '    Set RdoTest = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
   '        If Err > 0 Then
   '            Err.Clear
   '            sSql = "ALTER TABLE ComnTable ADD COADJACCT char(12) NULL DEFAULT('')"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "ALTER TABLE PcodTable ADD PCADJACCT char(12) NULL DEFAULT('')"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "ALTER TABLE PartTable ADD PAADJACCT char(12) NULL DEFAULT('')"
   '            clsADOCon.ExecuteSQL sSQL
   '            If Err = 0 Then
   '                sSql = "UPDATE ComnTable SET COADJACCT='' WHERE COADJACCT IS NULL"
   '                clsADOCon.ExecuteSQL sSQL
   '
   '                sSql = "UPDATE PcodTable SET PCADJACCT='' WHERE PCADJACCT IS NULL"
   '                clsADOCon.ExecuteSQL sSQL
   '
   '                sSql = "UPDATE PartTable SET PAADJACCT='' WHERE PAADJACCT IS NULL"
   '                clsADOCon.ExecuteSQL sSQL
   '            End If
   '        End If
   '    '2/8/05 Splits defaults
   '    Err.Clear
   '    sSql = "SELECT COALLOWSPLITS FROM ComnTable WHERE COALLOWSPLITS=0 "
   '    clsADOCon.ExecuteSQL sSQL
   '        If Err > 0 Then
   '            Err.Clear
   '            sSql = "ALTER TABLE ComnTable ADD " _
   '                & "COALLOWSPLITS TINYINT DEFAULT(0)," _
   '                & "COSPLITSTARTRUN INT NULL DEFAULT(900) "
   '            clsADOCon.ExecuteSQL sSQL
   '            If Err = 0 Then
   '                sSql = "UPDATE ComnTable SET COALLOWSPLITS=0," _
   '                    & "COSPLITSTARTRUN=900 WHERE COALLOWSPLITS IS NULL"
   '                clsADOCon.ExecuteSQL sSQL
   '            End If
   '        End If
   '    '4/29/05 INTCOA
   '    Err.Clear
   '    sSql = "SELECT SellingPriceFormat FROM Preferences WHERE PreRecord=1"
   '    clsADOCon.ExecuteSQL sSQL
   '        If Err > 0 Then
   '            sSql = "ALTER TABLE Preferences ADD " _
   '                & "SellingPriceFormat CHAR(12) NULL DEFAULT ('#######0.000')"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "UPDATE Preferences SET SellingPriceFormat='#######0.000'"
   '            clsADOCon.ExecuteSQL sSQL
   '        End If
   '
   '    Err.Clear
   '    '7/15/05 Last Purchase Order
   '    sSql = "SELECT COLASTPURCHASEORDER From ComnTable WHERE COREF=1"
   '    clsADOCon.ExecuteSQL sSQL
   '        If Err > 0 Then
   '            sSql = "ALTER TABLE ComnTable ADD COLASTPURCHASEORDER " _
   '                & "INT NULL DEFAULT(0)"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "UPDATE ComnTable SET COLASTPURCHASEORDER=0 " _
   '                & "WHERE COLASTPURCHASEORDER IS NULL"
   '            clsADOCon.ExecuteSQL sSQL
   '        End If
   '
   '    '7/15/05 Last Sales Order
   '    sSql = "SELECT COLASTSALESORDER From ComnTable WHERE COREF=1"
   '    clsADOCon.ExecuteSQL sSQL
   '        If Err > 0 Then
   '            sSql = "ALTER TABLE ComnTable ADD COLASTSALESORDER " _
   '                & "CHAR(6) NULL DEFAULT('')"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "UPDATE ComnTable SET COLASTSALESORDER='S00000'" _
   '                & "WHERE COLASTSALESORDER IS NULL"
   '            clsADOCon.ExecuteSQL sSQL
   '        End If
   '
   '   '7/29/05 Transfers
   '    Err.Clear
   '    sSql = "SELECT TransferInvoice FROM Preferences WHERE PreRecord=1"
   '    clsADOCon.ExecuteSQL sSQL
   '        If Err > 0 Then
   '            sSql = "ALTER TABLE Preferences ADD TransferInvoice INT NULL DEFAULT(0)"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "UPDATE Preferences SET TransferInvoice=0 WHERE TransferInvoice IS NULL"
   '            clsADOCon.ExecuteSQL sSQL
   '        End If
   '    Err.Clear
   '    sSql = "SELECT AllowTransfers FROM Preferences WHERE PreRecord=1"
   '    clsADOCon.ExecuteSQL sSQL
   '        If Err > 0 Then
   '            sSql = "ALTER TABLE Preferences ADD AllowTransfers INT NULL DEFAULT(0)"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "UPDATE Preferences SET AllowTransfers=0 WHERE AllowTransfers IS NULL"
   '            clsADOCon.ExecuteSQL sSQL
   '        End If
   'End Transfers
   '10/6/06
   '    '4/12/06 Stored Procedures
   '            sSql = "CREATE PROCEDURE Qry_FillCommentsByIndex AS " & vbCr _
   '                & "SELECT COMMENT_CLASS FROM StchTable " & vbCr _
   '                & "ORDER BY COMMENT_LISTINDEX"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "CREATE PROCEDURE Qry_GetCommentClass" & vbCr _
   '                & "(@cmtclass char(20)) " & vbCr _
   '                & "as " & vbCr _
   '                & "SELECT COMMENT_REF,COMMENT_ID FROM StcdTable " & vbCr _
   '                & "WHERE COMMENT_CLASS=@cmtclass"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "CREATE PROCEDURE Qry_FillSalesPersons AS " & vbCr _
   '                & "SELECT SPNUMBER FROM SprsTable " & vbCr _
   '                & "ORDER BY SPNUMBER"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "CREATE PROCEDURE Qry_GetABCPreference AS " & vbCr _
   '                & "SELECT CycleCountInitialized FROM Preferences WHERE " _
   '                & "PreRecord=1"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "CREATE PROCEDURE Qry_FillABCCombo AS " & vbCr _
   '                & "SELECT COABCROW,COABCCODE,COABCUSED FROM CabcTable WHERE " _
   '                & "COABCUSED=1 ORDER BY COABCROW"
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "CREATE PROCEDURE Qry_FillABCUsed AS " & vbCr _
   '                & "SELECT COABCROW,COABCCODE,COABCFREQUENCY,COABCLOWCOST," & vbCr _
   '                & "COABCHIGHCOST,COABCUSED FROM CabcTable WHERE COABCUSED=1 " & vbCr _
   '                & "ORDER BY COABCROW "
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "CREATE PROCEDURE Qry_FillABCNotUsed AS " & vbCr _
   '                & "SELECT COABCROW,COABCCODE,COABCFREQUENCY,COABCLOWCOST," & vbCr _
   '                & "COABCHIGHCOST,COABCUSED FROM CabcTable WHERE COABCUSED=0 " & vbCr _
   '                & "ORDER BY COABCROW "
   '            clsADOCon.ExecuteSQL sSQL
   '
   '            sSql = "CREATE PROCEDURE Qry_GetProductCode" & vbCr _
   '                & "(@productcode char(6)) " & vbCr _
   '                & "as " & vbCr _
   '                & "SELECT PCCODE,PCDESC FROM PcodTable " & vbCr _
   '                & "WHERE PCREF=@productcode"
   '
   '            sSql = "CREATE PROCEDURE Qry_GetProductClass " & vbCr _
   '                & "(@productclass char(4)) " & vbCr _
   '                & "as " & vbCr _
   '                & "SELECT CCCODE,CCDESC FROM PclsTable " & vbCr _
   '                & "WHERE CCREF=@productclass"
   '            clsADOCon.ExecuteSQL sSQL
   '   '4/27/06 Queue and Move format
   '    sSql = "SELECT QueueMoveConversion FROM Preferences WHERE PreRecord=1"
   '    clsADOCon.ExecuteSQL sSQL
   '        If Err > 0 Then
   '            Err.Clear
   '            sSql = "ALTER TABLE Preferences ADD QueueMoveConversion tinyint NULL DEFAULT(24)"
   '            clsADOCon.ExecuteSQL sSQL
   '            If Err = 0 Then
   '                sSql = "UPDATE Preferences SET QueueMoveConversion=24 WHERE (PreRecord=1 AND " _
   '                    & "QueueMoveConversion IS NULL)"
   '                clsADOCon.ExecuteSQL sSQL
   '            End If
   '        End If
   'End 10/6/06
End Sub

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
