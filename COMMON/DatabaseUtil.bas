Attribute VB_Name = "DatabaseUtil"
Option Explicit

Public Const DB_VERSION = 202

Public sSaAdmin As String
Public sSaPassword As String
Public sserver As String
Public sDataBase As String
Public sSysCaption As String
Public sProgName As String
Public sDsn As String
Public sConStr As String
'version info
Private oldType As String     'live or test
Private NewType As String
Private oldRelease As Long
Private newRelease As Long
Private OldDbVersion As Integer
Private directory As String
Private newver As Long

Public ES_PARTCOUNT As Long 'Total Customer Part Numbers

Type ModuleErrors
   Number As Long
   Description As String
End Type

Public CurrError As ModuleErrors

' ES/2000 Company Info Datatype

Type CompanyInfo
   Name As String
   Addr(5) As String
   Phone As String
   Fax As String
   GlVerify As Byte
End Type

Global Co As CompanyInfo

Type CurrentSelections
   CurrentPart As String
   CurrentVendor As String
   CurrentCustomer As String
   CurrentShop As String
   CurrentRegion As String
   CurrentGroup As String
   CurrentUser As String
End Type

Public cUR As CurrentSelections
Private ver As Long

Function OpenDBServer(Optional bReStart As Boolean) As Boolean
   'return = True if successful
   
   Dim strWindowDir As String
   Dim strSaAdmin As String
   Dim strSaPassword As String
   Dim strServer As String
   Dim strDBName As String
   Dim strConStr As String
   Dim ErrNum    As Long
   Dim ErrDesc   As String
   
   'Dim RdoCheck As ADODB.Recordset
   Dim b As Byte
   Dim CloseExit As Long
   Dim iTimeOut As Integer
   Dim sWindows As String
   
   
   
   MouseCursor ccHourglass
   'strWindowDir = GetWindowsDir()
   'sserver = UCase$(GetSetting("Esi2000", "System", "ServerId", sserver))
   ' DNS sserver = UCase(GetUserSetting(USERSETTING_ServerName))
   sserver = UCase(GetConfUserSetting(USERSETTING_ServerName))
   
   sSaAdmin = Trim(GetSysLogon(True))
   sSaPassword = Trim(GetSysLogon(False))
   ' GetCurrentDatabase method fills sDatabase
   GetCurrentDatabase
   strDBName = sDataBase
   
   
'Public sSaAdmin As String
'Public sSaPassword As String
'Public sserver As String
   
   Set clsADOCon = New ClassFusionADO
   
'    strConStr = "Provider='sqloledb';Data Source='" & strServer & "';" & _
'        "Initial Catalog='" & strDBName & "';Integrated Security='SSPI';"

   strConStr = "Driver={SQL Server};Provider='sqloledb';UID=" & sSaAdmin & ";PWD=" & _
            sSaPassword & ";SERVER=" & sserver & ";DATABASE=" & strDBName & ";"
   
   'MsgBox strConStr
   
   If clsADOCon.OpenConnection(strConStr, ErrNum, ErrDesc) = False Then
      MsgBox "An error occured while trying to connect to the specified database"
'     MsgBox "An error occured while trying to connect to the specified database:" & Chr(13) & Chr(13) & _
'            "Error Number = " & CStr(ErrNum) & Chr(13) & _
'            "Error Description = " & ErrDesc, vbOKOnly + vbExclamation, "  DB Connection Error"
     GoTo CleanUp
   End If
   
   sConStr = strConStr
   SaveSetting "Esi2000", "System", "CloseSection", ""
   
   Dim TestDatabase As Boolean
   TestDatabase = IsADOTestDatabase()
      
   Dim RdoCheck As ADODB.Recordset
   If bReStart = 0 Then
      'Get Count of Parts to see how combo's are to be handled
      sSql = "SELECT COUNT(PARTREF) FROM PartTable"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCheck, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoCheck.Fields(0)) Then _
                       ES_PARTCOUNT = RdoCheck.Fields(0) Else _
                       ES_PARTCOUNT = 0
      End If
' Mohan Commented
      'GetDataBases
      'UpdateTables
      b = CheckSecuritySettings()
'      'If b = 0 Then GetSectionPermissions
      GetCompany 1
      On Error GoTo modErr1
      GetCustomerPermissions
      MDISect.Caption = GetSystemCaption
   Else
      'MouseCursor ccArrow
   End If
      
      
   OpenDBServer = True
   Exit Function
   
modErr1:
   Resume modErr2
modErr2:
   'Couldn't open msdb
   For b = 1 To 8
      bCustomerGroups(b) = 1
   Next
MsgBox "@@@ OpenDBServer Error " & Err.Description
   Set clsADOCon = Nothing
   Exit Function

CleanUp:
 'clsADOCon.CleanupRecordset RS
 Set clsADOCon = Nothing

End

End Function

Public Function IsADOTestDatabase() As Boolean
   'returns true if using a test database
   
   
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "select TestDatabase from Version"
   Dim rs As ADODB.Recordset
   Dim TestDatabase As Boolean
   If clsADOCon.GetDataSet(sSql, rs) Then
      If clsADOCon.ADOErrNum = 0 Then
         IsADOTestDatabase = IIf(rs(0) = 1, True, False)
      Else
         clsADOCon.ExecuteSql ("alter table Version add TestDatabase tinyint not null default 0")
      End If
   End If
   
End Function

'Private Function CheckDateVersion(OldDbVersion As Integer, NewDbVersion As Integer) As Boolean
'   'terminates if cannot proceed.
'   'returns false if no update required
'   'returns true if update required and authorized by and admin
'   Err.Clear
'
'   Dim strFulVer As String
'   clsADOCon.ADOErrNum = 0
'   sSql = "select * from Updates" & vbCrLf _
'      & "where UpdateID = (select max(UpdateID) from Updates)"
'   On Error Resume Next
'   Dim rdo As ADODB.Recordset
'   If clsADOCon.GetDataSet(sSql, rdo) Then
'      If clsADOCon.ADOErrNum = 0 Then
'         oldRelease = rdo!newRelease
'      End If
'   Else
'      oldRelease = 0
'   End If
'
'   If IsTestDatabase Then
'      oldType = "Test"
'   Else
'      oldType = "Live"
'   End If
'
'   If App.Minor < 10 Then
'    strFulVer = CStr(App.Major) & "0" & CStr(App.Minor)
'   Else
'    strFulVer = CStr(App.Major) & CStr(App.Minor)
'   End If
'
'   newRelease = CInt(strFulVer)
'
'   If InTestMode() Then
'      NewType = "Test"
'   Else
'      NewType = "Live"
'   End If
'
'   Dim msg As String
'   If oldRelease < newRelease Then
'      msg = "Old release (" & oldRelease & ") < New release (" & newRelease & ")" & vbCrLf
'   ElseIf oldRelease > newRelease Then
'
'      MsgBox "Old release (" & oldRelease & ") > New release (" & newRelease & ")" & vbCrLf _
'         & "You cannot proceed.", vbCritical
'      End
'   End If
'
'   If OldDbVersion < NewDbVersion Then
'      msg = msg & "Old db version (" & OldDbVersion & ") < New db version (" & NewDbVersion & ")" & vbCrLf
'   ElseIf OldDbVersion > NewDbVersion Then
'
'      MsgBox "Old db version (" & OldDbVersion & ") > New db version (" & NewDbVersion & ")" & vbCrLf _
'         & "You cannot proceed.", vbCritical
'      End
'   End If
'
'   If oldType <> NewType Then
'      msg = msg & "Db type (" & oldType & ") <> Application type (" & NewType & ")" & vbCrLf
'   End If
'
'   If msg = "" Then
'      CheckDateVersion = False
'      Exit Function
'
'   Else
'      If Secure.UserAdmn = 0 Then
'         msg = msg & "An administrator must perform an update before you can proceed."
'         MsgBox msg, vbCritical
'         End
'      Else
'         msg = msg & "Do you wish to update the database now?"
'         Select Case MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2)
'         Case vbYes
'            CheckDateVersion = True
'         Case Else
'            End
'         End Select
'      End If
'   End If
'
'End Function


Private Sub ExecuteScript(DisplayError As Boolean, sql As String)
   Err.Clear
   On Error Resume Next

   'display version being updated to
   Dim display As String
   display = "Updating database to version " & newver
   If Left(SysMsgBox.msg, Len(display)) = display Then
      SysMsgBox.msg = SysMsgBox.msg & "."
   Else
      SysMsgBox.msg = display
   End If
   SysMsgBox.Refresh
   DoEvents
   
   Dim saveSql As String
   saveSql = sql
   
   clsADOCon.ExecuteSql saveSql
   'RdoCon.Execute saveSql
   
   DoEvents

   'display error if required
   'always display a timeout error
   If Err Then
      Debug.Print CStr(Err.Number) & ": " & "  " & Err.Description & vbCrLf & " SQL: " & sql
      Dim msg As String
      If RunningInIDE() Or DisplayError Or InStr(1, Err.Description, "timeout expired", vbTextCompare) > 0 Then
         msg = "Database version " & ver & " update failed with error " & CStr(Err.Number) & vbCrLf _
            & Err.Description & vbCrLf
         If InStr(1, Err.Description, "timeout expired", vbTextCompare) > 0 Then
            msg = msg & clsADOCon.CommandTimeout & " second timeout occurred performing database update." & vbCrLf
         End If
         msg = msg & "It is important that this information be captured and reported." & vbCrLf & vbCrLf _
            & "SQL: " & sql
         MsgBox msg, vbCritical, "Critical Database Version " & ver & " Update Failure"
      End If
   End If
End Sub


Private Sub AlterNumericColumn(TableName As String, ColumnName As String, NewType As String)
   'remove default constraint if any, alter column, and add a default of zero
   'example:
   'AlterNumericColumn "RunsTable", "RUNEXP", "decimal(12,4) null"
   
   'drop defaults created by alter table, if any
   ExecuteScript False, "exec DropColumnDefault '" & TableName & "', '" & ColumnName & "'"
   'drop defaults created by sp_binddefault (DEFZERO and DEFBLANK)
   ExecuteScript False, "exec sp_unbindefault '" & TableName & "." & ColumnName & "'"
   'alter the column
   ExecuteScript False, "alter table " & TableName & " alter column " & ColumnName _
      & " " & NewType
   'add a zero constraint back
   ExecuteScript False, "alter table " & TableName & " add constraint DF_" & TableName & "_" _
      & ColumnName & " default 0 for " & ColumnName
   
End Sub

Private Sub AlterStringColumn(TableName As String, ColumnName As String, NewType As String)
   'remove default constraint if any, alter column, and add a default of blank
   'example:
   'AlterStringColumn "RndlTable", "RUNDLSDOCREF", "varchar(30)"
   
   'drop defaults created by alter table, if any
   ExecuteScript False, "exec DropColumnDefault '" & TableName & "', '" & ColumnName & "'"
   'drop defaults created by sp_binddefault (DEFZERO and DEFBLANK)
   ExecuteScript False, "exec sp_unbindefault '" & TableName & "." & ColumnName & "'"
   'alter the column
   ExecuteScript False, "alter table " & TableName & " alter column " & ColumnName _
      & " " & NewType & " NULL"
   'add a blank constraint back
   ExecuteScript False, "alter table " & TableName & " add constraint DF_" & TableName & "_" _
      & ColumnName & " default '' for " & ColumnName
   
End Sub

Private Sub DropColumnDefault(TableName As String, ColumnName As String)
   'remove default constraint, if any, from a column
   
   'drop defaults created by alter table, if any
   ExecuteScript False, "exec DropColumnDefault '" & TableName & "', '" & ColumnName & "'"
   'drop defaults created by sp_binddefault (DEFZERO and DEFBLANK)
   ExecuteScript False, "exec sp_unbindefault '" & TableName & "." & ColumnName & "'"
   
End Sub

Public Function ColumnExists(TableName As String, ColumnName As String) As Boolean
   'returns True if column exists
   
   sSql = "SELECT 1 FROM information_schema.Columns" & vbCrLf _
      & "WHERE COLUMN_NAME = '" & ColumnName & "'" & "AND TABLE_NAME = '" & TableName & "'"
   Dim rdo As ADODB.Recordset
   If clsADOCon.GetDataSet(sSql, rdo) Then
      ColumnExists = True
   End If
End Function

Public Function TableExists(strTableName As String) As Boolean
   'returns True if column exists
   'On Error Resume Next
   Err.Clear
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT * FROM INFORMATION_SCHEMA.TABLES where TABLE_NAME = '" & strTableName & "'"
   Dim rdo As ADODB.Recordset
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If bSqlRows Then
      TableExists = True
   Else
      TableExists = False
   End If
   Set rdo = Nothing
   'clsADOCon.ADOErrNum = 0
End Function

Public Function StoreProcedureExists(strSPName As String) As Boolean
   'returns True if store procedure exists
   On Error Resume Next
   Err.Clear
   sSql = "Select * From SysObjects Where Type = 'P' and Name='" & strSPName & "'"
   Dim rdo As ADODB.Recordset
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If bSqlRows Then
      StoreProcedureExists = True
   Else
      StoreProcedureExists = False
   End If
End Function

Public Function GetSysLogon(bGetLogin As Byte) As String
   'get database login info
   'if bGetLogin = 0 then get the password
   'if bGetLogin = 1 then get the login
   
   Dim A As Integer
   Dim iLen As Integer
   Dim sTest As String
   Dim sNewString As String
   Dim sPassword As String
   
   If bGetLogin <> 0 Then
      'GetSysLogon = GetSetting("UserObjects", "System", "NoReg", GetSysLogon)
      ' DNS GetSysLogon = GetUserSetting(USERSETTING_SqlLogin)
      GetSysLogon = GetConfUserSetting(USERSETTING_SqlLogin)
      
      If Trim(GetSysLogon) = "" Then GetSysLogon = "sa"
   Else
      ' DNS sPassword = GetDatabasePassword
      sPassword = GetConfigDBPass
      GetSysLogon = sPassword
   End If
   
End Function

' Calls the windows API to get the windows directory and
' ensures that a trailing dir separator is present
' Returns: The windows directory

Public Function GetWindowsDir()
   Dim intZeroPos As Integer
   Dim gintMAX_SIZE As Integer
   Dim strBuf As String
   gintMAX_SIZE = 255 'Maximum buffer size
   
   strBuf = Space$(gintMAX_SIZE)
   'Get the windows directory and then trim the buffer to the exact length
   'returned and add a dir sep (backslash) if the API didn't return one
   If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
      intZeroPos = InStr(strBuf, Chr$(0))
      If intZeroPos > 0 Then strBuf = Left$(strBuf, intZeroPos - 1)
      GetWindowsDir = strBuf
   Else
      GetWindowsDir = ""
   End If
   
End Function

Public Function GetSystemCaption() As String
   If InTestMode() Then
      GetSystemCaption = "TEST MODE "
   End If
   'GetSystemCaption = GetSystemCaption & "ES/2000 ERP"
   If sProgName = "" Then
      GetSystemCaption = GetSystemCaption & "Fusion ERP"
   Else
      GetSystemCaption = GetSystemCaption & sProgName
   End If
   If sDataBase <> "" Then
      GetSystemCaption = GetSystemCaption & " - " & sDataBase
   End If
   
   ' add version
   GetSystemCaption = GetSystemCaption & " - v " & App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Sub GetCurrentDatabase()
'Database
'sDataBase = GetSetting("Esi2000", "System", "CurDatabase", sDataBase)
' DNS sDataBase = GetUserSetting(USERSETTING_DatabaseName)

    sDataBase = GetConfUserSetting(USERSETTING_DatabaseName)
    If Trim(sDataBase = "") Then sDataBase = "Esi2000Db"

End Sub


'Note: Skips over KeySets and Dynamic Cursors

Public Sub ClearResultSet(RdoDataSet As ADODB.Recordset)
   
   ' Don't have Updatable property in ADODV Recordset.
   RdoDataSet.Cancel
'   If Not RdoDataSet.Updatable Then
'      Do While RdoDataSet.MoreResults
'      Loop
'      RdoDataSet.Cancel
'   End If
   
End Sub


Sub GetCompany(Optional bWantAddress As Byte)
   Dim ActRs As ADODB.Recordset
   Dim bByte As Byte
   Dim A As Integer
   Dim b As Integer
   Dim C As Integer
   Dim d As Integer
   Dim sAddress As String
   
   On Error GoTo modErr1
   If bWantAddress Then
      sSql = "SELECT COREF,CONAME,COPHONE,COFAX,COGLVERIFY,COADR FROM ComnTable " _
             & "WHERE COREF=1"
   Else
      sSql = "SELECT COREF,CONAME,COPHONE,COFAX,COGLVERIFY FROM ComnTable " _
             & "WHERE COREF=1"
   End If
   'bSqlRows = GetDataSet(ActRs, ES_STATIC)
   bSqlRows = clsADOCon.GetDataSet(sSql, ActRs, ES_STATIC)
   If bSqlRows Then
      With ActRs
         Co.Name = "" & Trim(!CONAME)
         Co.Phone = "" & Trim(!COPHONE)
         Co.Fax = "" & Trim(!COFAX)
         Co.GlVerify = !COGLVERIFY
         If bWantAddress Then sAddress = "" & Trim(!COADR)
      End With
   End If
   'have to parse CfLf if we want address. Crystal Reports formulae only
   If bWantAddress Then
      On Error Resume Next
      A = InStr(1, sAddress, Chr(13) & Chr(10))
      Co.Addr(1) = Left(sAddress, A - 1)
      
      sAddress = Right(sAddress, Len(sAddress) - (A + 1))
      b = InStr(1, sAddress, Chr(13) & Chr(10))
      If b = 0 Then
         bByte = 1
         b = Len(sAddress)
         Co.Addr(2) = Left(sAddress, b)
      Else
         Co.Addr(2) = Left(sAddress, b - 1)
      End If
      
      If bByte = 0 Then
         sAddress = Right(sAddress, Len(sAddress) - (b + 1))
         C = InStr(1, sAddress, Chr(13) & Chr(10))
         If C = 0 Then
            bByte = 1
            C = Len(sAddress)
            Co.Addr(3) = Left(sAddress, C)
         Else
            Co.Addr(3) = Left(sAddress, C - 1)
         End If
      End If
      
      If bByte = 0 Then
         sAddress = Right(sAddress, Len(sAddress) - (C + 1))
         d = InStr(1, sAddress, Chr(13) & Chr(10))
         If d = 0 Then
            bByte = 1
            d = Len(sAddress)
            Co.Addr(4) = Left(sAddress, d)
         Else
            Co.Addr(4) = Left(sAddress, d - 1)
         End If
      End If
   End If
   sFacility = Co.Name
   Set ActRs = Nothing
   Exit Sub
   
modErr1:
   Resume modErr2
modErr2:
   On Error GoTo 0
   
End Sub




'Make sure that the user's DSN is pointed to the
'correct server. If none is registered, then build it

'Public Function RegisterSqlDsn(sDataSource As String) As String
'   Dim sAttribs As String
'   If sDataSource = "" Then sDataSource = "ESI2000"
'   sAttribs = "Description=" _
'              & "ES/2000ERP SQL Server Data " _
'              & vbCrLf & "OemToAnsi=No" _
'              & vbCrLf & "SERVER=" & sserver _
'              & vbCrLf & "Database=" & sDataBase
'   'Create new DSN or revise registered DSN.
'   rdoEngine.rdoRegisterDataSource sDataSource, _
'      "SQL Server", True, sAttribs
'   RegisterSqlDsn = sDataSource
'   Exit Function
'
'modErr1:
'   On Error GoTo 0
'   RegisterSqlDsn = sDataSource
'
'End Function
'


'Code  GETSERVERDATETIME() = Format(GetServerDateTime,"mm/dd/yy") etal
'11/21/06 Revised for clarity

Public Function GetServerDateTime() As Variant
   Dim RdoTme As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT GETDATE() AS ServerTime"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTme, ES_FORWARD)
   If bSqlRows Then
      GetServerDateTime = RdoTme!ServerTime
   Else
      Dim cs As String
      cs = clsADOCon.ConnectionString
      clsADOCon.CleanupConnection clsADOCon.conConnection
      clsADOCon.CheckConnection

      GetServerDateTime = Now
   End If
   Set RdoTme = Nothing
   Exit Function
   
modErr1:
   GetServerDateTime = Format(Now, "mm/dd/yy")
   
End Function

Public Function GetServerDate() As Variant
   Dim RdoTme As ADODB.Recordset
   On Error Resume Next
   GetServerDate = Format(Now, "mm/dd/yy")
   sSql = "select getdate() AS ServerTime"
   If clsADOCon.GetDataSet(sSql, RdoTme, ES_FORWARD) Then
      GetServerDate = Format(RdoTme!ServerTime, "mm/dd/yy")
   Else
      Dim cs As String
      cs = clsADOCon.ConnectionString
      clsADOCon.CleanupConnection clsADOCon.conConnection
      clsADOCon.CheckConnection

      GetServerDate = Format(Now, "mm/dd/yy")
   End If
End Function

Private Sub AddNonNullColumnWithDefault(TableName As String, ColumnName As String, _
   TypeName As String, DefaultValue As String)

   'if column already exists, just return
   If ColumnExists(TableName, ColumnName) Then
      Exit Sub
   End If
   
   'add a non-null column and do the gyrations to give it a default
   ExecuteScript True, "ALTER TABLE " & TableName & " ADD " & ColumnName & " " & TypeName & " NULL"
   ExecuteScript True, "UPDATE " & TableName & " SET " & ColumnName & " = " & DefaultValue
   ExecuteScript True, "ALTER TABLE " & TableName & " ALTER COLUMN " & ColumnName & " " & TypeName & " NOT NULL"
   ExecuteScript True, "ALTER TABLE " & TableName & " ADD CONSTRAINT DF_" & TableName & "_" & ColumnName & " DEFAULT " & DefaultValue & " FOR " & ColumnName


End Sub

Sub ExecSQL(sql As String)
    Err.Clear
    On Error Resume Next
    clsADOCon.ExecuteSql sql
    Err.Clear
End Sub

Public Function IsTestDatabase() As Boolean
   'returns true if using a test database
   
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "select TestDatabase from Version"
   Dim rdo As ADODB.Recordset
   Dim TestDatabase As Boolean
   If clsADOCon.GetDataSet(sSql, rdo) Then
      If clsADOCon.ADOErrNum = 0 Then
         IsTestDatabase = IIf(rdo.Fields(0) = 1, True, False)
      Else
         clsADOCon.ExecuteSql "alter table Version add TestDatabase tinyint not null default 0"
      End If
   End If
   
End Function

Private Sub DropIndex(TableName As String, IndexName As String)
   'drop index that works with SQL2000, SQL2005, and SQL2008
   
      ExecuteScript False, "DROP INDEX " & TableName & "." & IndexName & " -- SQL2000"
      ExecuteScript False, "DROP INDEX " & IndexName & " ON " & TableName & " -- SQL2005 & SQL2008"
End Sub

Public Function CrystalDate(dt As Variant) As String
   'returns date in format suitable for Crystal Reports SQL
   CrystalDate = " Date(" & Format(dt, "yyyy,mm,dd") & ") "
End Function


Public Sub UpdateDatabase()
   
   Dim rdo As ADODB.Recordset
   
   'if no version table, create it and set it to version 0
   'sSql = "select max(Version) as Version from Version"
   sSql = "select Version from Version"
   On Error Resume Next 'need to attempt all steps even if they fail
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If Err.Number <> 0 Then
      bSqlRows = False
      Err.Clear
   End If
   If bSqlRows Then
      OldDbVersion = rdo!Version
      ver = rdo!Version
   Else
      ExecuteScript False, ("create table Version ( Version  int )")
      ExecuteScript False, ("insert Version ( Version ) values ( 0 )")
      ver = 0
   End If
   Set rdo = Nothing
   
   'Continue if:
   '1. Update required and user is an admin
   '2. User is a developer running in VB IDE (to debug db updates)
   Dim UpdateReqd As Boolean
   If IsUpdateRequired(OldDbVersion, DB_VERSION) Then
      UpdateReqd = True
   Else
      'If Not RunningInIDE Then
         Exit Sub
      'End If
   End If
   
   Err.Clear
   On Error Resume Next
   
   MouseCursor ccHourglass
   SysMsgBox.tmr1.Enabled = False
   SysMsgBox.msg = "Updating database."
   SysMsgBox.Show
   
   'allow really long timeouts (normal limit is 60 sec)
   Dim timeout As Integer
   'TODO: get the query timeout
   'timeout = RdoCon.QueryTimeout
   'RdoCon.QueryTimeout = 1200    '20 minutes

  
   ' 11/07/2016
   UpdateDatabase78
   
   ' 1/5/2017 - Terry
   UpdateDatabase79
   
   ' v 17.4.0 3/15/2017 - Terry
   UpdateDatabase80
   
   ' v 17.5.1 4/11/2017 - Terry
   UpdateDatabase81
   
   ' v 17.5.4 4/?/2017 - Terry - update WC Queue Status stored procedures
   UpdateDatabase82
   
   ' v 17.5.5 5/?/2017 - Terry
   UpdateDatabase83
   
   ' v 17.5.6 5/18/2017 - Terry
   UpdateDatabase84
   
   ' v 17.5.8 6/24/2017 - Terry
   UpdateDatabase85
   
   ' v 17.5.9 - Terry
   UpdateDatabase86
   
   ' v 17.6.0 - Terry
   UpdateDatabase87
   
   ' v 17.6.1 - Terry
   UpdateDatabase88
   
   ' v 17.6.2 - Terry
   UpdateDatabase89
   
   ' v 17.6.3 9/18/17 - Terry
   UpdateDatabase90
   
   ' v 17.6.4 db v 166 9/21/17 - Terry
   UpdateDatabase91
   
   ' v 17.6.4 db v 167 9/28/17 - Terry
   UpdateDatabase92
   
   ' v 17.6.5 db v 168 10/17/17 - Terry
   UpdateDatabase93
   
   ' v 17.6.6 db v 169 10/23/17 - Terry
   UpdateDatabase94
   
   ' v 17.7.0 db v 170 11/13/17 - Terry
   UpdateDatabase95
   
   ' v 17.17.1 db v 171 12/?/17 - Terry
   UpdateDatabase96
   
   ' v 17.17.2 db v 172 12/15/17 - Terry
   UpdateDatabase97
   
   ' v 18.0.1 db v 173 1/23/2018 - Terry
   UpdateDatabase98
   
   ' v 18.0.2 db v 174 2/6/2018 - Terry
   UpdateDatabase99
   
   ' v 18.0.3 db v 175 2/24/2018 - Terry
   UpdateDatabase100
   
   ' v 18.0.4 db v 176 3/?/2018 - Terry
   UpdateDatabase101
   
   ' v 18.0.5 db v 177 4/25/2018 - Terry
   UpdateDatabase102
   
   ' v 18.0.6 db v 178 6/25/2018 - Terry
   UpdateDatabase103
   
   ' v 18.0.7 db v 179 7/18/2018 - Terry
   UpdateDatabase104
   
   ' v 18.0.8 db v 180 8/8/2018 - Terry
   UpdateDatabase105
   
   ' v 18.1.0,1,2,3 db v 181, 182, 183 11/1/2018 - Terry
   UpdateDatabase106
   
   ' v 19.0..0 11/4/2019 - Terry
   UpdateDatabase107
   
   ' v 19.0.1 1/27/2018 - Terry
   UpdateDatabase108
   
   ' v 19.0.3 3/14/2018 - Terry
   UpdateDatabase109
   
   ' v 19.0.4/5 3/19/19 & 3/21/19 - Terry
   UpdateDatabase110
   
   ' v 19.1.0 5/17/19 - Terry
   UpdateDatabase111
   
   ' v 19.1.1 6/12/19 - Terry
   UpdateDatabase112
   
   ' v 19.1.2 6/?/19 - Terry
   UpdateDatabase113
   
   ' v 19.1.4 - Terry
   UpdateDatabase193
   
   ' v 19.1.6 11/21/19 - Terry
   UpdateDatabase194
   
   ' v 19.1.7 12/3/19 - Terry
   UpdateDatabase195
   
   ' v 20.0 - 1/5/20 - Terry
   UpdateDatabase196
   
   ' v 20.1 - 1/?/20 - Terry
   UpdateDatabase197
   
   ' v 20.3 - 3/25/20 - Terry
   UpdateDatabase198
   
   ' v 20.4 - 3/31/20 - Terry
   UpdateDatabase199
   
   ' v 20.5 - 7/20/20 - Terry
   UpdateDatabase200
   
   ' v 20.6 - 8/?/20 - Terry
   UpdateDatabase201
   
   ' v 20.7 - 8/?/20 - Terry
   UpdateDatabase202
   
   'record update
   'don't do this if running in VB IDE and no update was required
   If UpdateReqd Then
      sSql = "insert Updates" & vbCrLf _
         & "(AppDirectory,OldRelease,NewRelease,UserInitials," & vbCrLf _
         & "OldDbVersion, NewDbVersion,OldDbType,NewDbType)" & vbCrLf _
         & "values('" & App.Path & "'," & oldRelease & "," & newRelease & ",'" & sInitials & "'," & vbCrLf _
         & OldDbVersion & "," & DB_VERSION & ",'" & oldType & "','" & NewType & "')"
         
      ExecuteScript True, sSql
      
      'update test/live indicator if necessary
      If oldType <> NewType Then
         sSql = "update Version set TestDatabase = " & IIf(NewType = "Live", 0, 1)
         ExecuteScript True, sSql
      End If
   End If
   
   Unload SysMsgBox
   MouseCursor ccDefault

End Sub

   
Private Function UpdateDatabase78()
   'updates for ROLT by class and product code
   '

   Dim sql As String
   sql = ""
   
   newver = 153
   If ver < newver Then
   
      clsADOCon.ADOErrNum = 0
      
      If Not ColumnExists("RthdTable", "RTINACTIVE") Then
        sSql = "ALTER TABLE RthdTable ADD RTINACTIVE tinyint not null default 0"
        clsADOCon.ExecuteSql sSql
      End If
         
      If Not ColumnExists("CihdTable", "INVREASONS") Then
        sSql = "ALTER TABLE CihdTable add INVREASONS varchar(2048) null default('')"
        clsADOCon.ExecuteSql sSql
      End If
      

      sql = "" & vbCrLf
      sql = sql & "ALTER PROCEDURE [dbo].[Qry_FillRoutings]  as" & vbCrLf
      sql = sql & "SELECT RTREF,RTNUM From RthdTable WHERE RTINACTIVE = 0  ORDER BY RTREF"
      sql = sql & "" & vbCrLf
      
      ExecuteScript False, sql
 

      If Not ColumnExists("ComnTable", "COWARNSERVICEOPOPEN") Then
          sSql = "ALTER table dbo.ComnTable add COWARNSERVICEOPOPEN tinyint NULL CONSTRAINT DF_ComnTable_COWARNSERVICEOPOPEN DEFAULT 0"
          ExecuteScript False, sSql
      End If

      If Not ColumnExists("CCitTable", "CILOTLOCATION") Then
          sSql = "Alter table CCitTable Add CILOTLOCATION [char](4) NULL"
          ExecuteScript False, sSql
      End If

      If (Not TableExists("LtTrkTable")) Then
         sSql = "CREATE TABLE [dbo].[LtTrkTable](" & vbCrLf
         sSql = sSql & " [LOTNUMBER] [varchar](15) NULL," & vbCrLf
         sSql = sSql & " [LOTUSERLOTID] [varchar](40) NULL" & vbCrLf
         sSql = sSql & " ) ON [PRIMARY]"
         
         ExecuteScript False, sSql
      End If

      If Not ColumnExists("rtopTable", "OPFILLREF") Then
          sSql = "ALTER TABLE rtopTable ADD OPFILLREF Varchar(30) NULL"
          ExecuteScript False, sSql
      End If


 
      sSql = "ALTER PROCEDURE [dbo].[RptInvMovFromWIP] " & vbCrLf
      sSql = sSql & "    @StartDate as varchar(16), @EndDate as Varchar(16),  " & vbCrLf
      sSql = sSql & "    @PartType1 as Integer, @PartType2 as Integer, @PartType3 as Integer, @PartType4 as Integer  " & vbCrLf
      sSql = sSql & "   AS  " & vbCrLf
      sSql = sSql & "   BEGIN " & vbCrLf
      sSql = sSql & "    " & vbCrLf
      sSql = sSql & "      IF (@PartType1 = 1)      " & vbCrLf
      sSql = sSql & "         SET @PartType1 = 1     " & vbCrLf
      sSql = sSql & "      Else                     " & vbCrLf
      sSql = sSql & "         SET @PartType1 = 0     " & vbCrLf
      sSql = sSql & "    " & vbCrLf
      sSql = sSql & "      IF (@PartType2 = 1)      " & vbCrLf
      sSql = sSql & "         SET @PartType2 = 2     " & vbCrLf
      sSql = sSql & "      Else                     " & vbCrLf
      sSql = sSql & "         SET @PartType2 = 0     " & vbCrLf
      sSql = sSql & "       " & vbCrLf
      sSql = sSql & "   IF (@PartType3 = 1)     " & vbCrLf
      sSql = sSql & "      SET @PartType3 = 3  " & vbCrLf
      sSql = sSql & "   Else                    " & vbCrLf
      sSql = sSql & "     SET @PartType3 = 0  " & vbCrLf
      sSql = sSql & "                              " & vbCrLf
      sSql = sSql & "      IF (@PartType4 = 1)     " & vbCrLf
      sSql = sSql & "         SET @PartType4 = 4  " & vbCrLf
      sSql = sSql & "      Else                    " & vbCrLf
      sSql = sSql & "         SET @PartType4 = 0  " & vbCrLf
      sSql = sSql & "    " & vbCrLf
      sSql = sSql & "      SELECT Inadate, InvaTable.INTYPE,InvaTable.INLOTNUMBER,InvaTable.INPART, InvaTable.INREF2,  " & vbCrLf
      sSql = sSql & "         InvaTable.INAQTY, InvaTable.INAMT, LohdTable.LOTORIGINALQTY, LohdTable.LOTTOTMATL, INTOTMATL,  " & vbCrLf
      sSql = sSql & "         LohdTable.LOTTOTLABOR, INTOTLABOR, LohdTable.LOTTOTEXP, INTOTEXP, LohdTable.LOTTOTOH, INTOTOH, " & vbCrLf
      sSql = sSql & "         LOTDATECOSTED, INDEBITACCT, INCREDITACCT " & vbCrLf
      sSql = sSql & "      FROM " & vbCrLf
      sSql = sSql & "         (PartTable PartTable INNER JOIN InvaTable InvaTable ON " & vbCrLf
      sSql = sSql & "            PartTable.PARTREF = InvaTable.INPART) " & vbCrLf
      sSql = sSql & "          LEFT OUTER JOIN LohdTable LohdTable ON " & vbCrLf
      sSql = sSql & "            InvaTable.INLOTNUMBER = LohdTable.LOTNUMBER " & vbCrLf
      sSql = sSql & "      WHERE " & vbCrLf
      sSql = sSql & "         InvaTable.INTYPE IN (6,12) and  " & vbCrLf
      sSql = sSql & "         Convert(DateTime, Inadate, 101) between @StartDate and @EndDate  " & vbCrLf
      sSql = sSql & "            AND PALEVEL IN (@PartType1, @PartType2, @PartType3, @PartType4) " & vbCrLf
      sSql = sSql & "      --    AND INPART = 'MX000075' " & vbCrLf
      sSql = sSql & "         AND INLOTNUMBER NOT IN (SELECT a.INLOTNUMBER  FROM InvaTable a where  " & vbCrLf
      sSql = sSql & "            a.INPART = InvaTable.INPART " & vbCrLf
      sSql = sSql & "            --a.INPART = 'MX000075'  " & vbCrLf
      sSql = sSql & "            AND a.INTYPE IN (6,38,12) AND Convert(DateTime, a.Inadate, 101) between @StartDate and @EndDate " & vbCrLf
      sSql = sSql & "      GROUP BY INLOTNUMBER  " & vbCrLf
      sSql = sSql & "      HAVING COUNT(INLOTNUMBER) > 1) " & vbCrLf
      sSql = sSql & "      ORDER BY " & vbCrLf
      sSql = sSql & "         PartTable.PALEVEL ASC, " & vbCrLf
      sSql = sSql & "         PartTable.PACLASS ASC " & vbCrLf
      sSql = sSql & "    " & vbCrLf
      sSql = sSql & "   END "
      
Debug.Print sSql

      ExecuteScript False, sSql
 
      sSql = "ALTER procedure [dbo].[UpdateTimeCardTotals]" & vbCrLf
      sSql = sSql & " @EmpNo int," & vbCrLf
      sSql = sSql & " @Date datetime" & vbCrLf
      sSql = sSql & "as " & vbCrLf
      sSql = sSql & "/* test" & vbCrLf
      sSql = sSql & "    UpdateTimeCardTotals 52, '9/8/2008'" & vbCrLf
      sSql = sSql & "*/" & vbCrLf
      sSql = sSql & "update TchdTable " & vbCrLf
      sSql = sSql & "set TMSTART = (select top 1 TCSTART From TcitTable " & vbCrLf
      sSql = sSql & " join TchdTable on TCCARD = TMCARD" & vbCrLf
      sSql = sSql & " where TCEMP = @EmpNo and TMDAY = @Date and TCSTOP <> ''" & vbCrLf
      sSql = sSql & " and ISDATE(TCSTART + 'm') = 1 and ISDATE(TCSTOP + 'm') = 1" & vbCrLf
      sSql = sSql & " order by tcstarttime)," & vbCrLf
      sSql = sSql & "TMSTOP = (select top 1 TCSTOP From TcitTable" & vbCrLf
      sSql = sSql & " join TchdTable on TCCARD = TMCARD" & vbCrLf
      sSql = sSql & " where TCEMP = @EmpNo and TMDAY = @Date and TCSTOP <> ''" & vbCrLf
      sSql = sSql & " and ISDATE(TCSTART + 'm') = 1 and ISDATE(TCSTOP + 'm') = 1" & vbCrLf
      sSql = sSql & " order by case when datediff( n, cast(rtrim(TCSTART) + 'm' as datetime)," & vbCrLf
      sSql = sSql & " cast(rtrim(TCSTOP) + 'm' as datetime) ) >= 0  " & vbCrLf
      sSql = sSql & " then cast(rtrim(TCSTOP) + 'm' as datetime) " & vbCrLf
      sSql = sSql & " else dateadd(day, 1, cast(rtrim(TCSTOP) + 'm' as datetime)) end desc)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "where TMEMP = @EmpNo and TMDAY = @Date" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "update TchdTable " & vbCrLf
      sSql = sSql & " set TMREGHRS = (select isnull(sum(hrs), 0.000) from viewTimeCardHours" & vbCrLf
      sSql = sSql & " where EmpNo = @EmpNo" & vbCrLf
      sSql = sSql & " and chgDay = @Date" & vbCrLf
      sSql = sSql & " and type = 'R')," & vbCrLf
      sSql = sSql & " TMOVTHRS = (select isnull(sum(hrs), 0.000) from viewTimeCardHours" & vbCrLf
      sSql = sSql & " where EmpNo = @EmpNo" & vbCrLf
      sSql = sSql & " and chgDay = @Date" & vbCrLf
      sSql = sSql & " and type = 'O')," & vbCrLf
      sSql = sSql & " TMDBLHRS = (select isnull(sum(hrs), 0.000) from viewTimeCardHours" & vbCrLf
      sSql = sSql & " where EmpNo = @EmpNo" & vbCrLf
      sSql = sSql & " and chgDay = @Date" & vbCrLf
      sSql = sSql & " and type = 'D')" & vbCrLf
      sSql = sSql & "where TMEMP = @EmpNo and TMDAY = @Date"
      
      ExecuteScript False, sSql
      
      If (Not TableExists("TlitTableNew")) Then
      
         sSql = "CREATE TABLE [dbo].[TlitTableNew](" & vbCrLf
         sSql = sSql & " [TOOL_NUM] [char](30) NOT NULL CONSTRAINT [DF_TlitTableNew_TOOL_NUM]  DEFAULT ('')," & vbCrLf
         sSql = sSql & " [TOOL_PARTREF] [char](30) NOT NULL CONSTRAINT [DF_TlitTableNew_TOOL_PARTREF]  DEFAULT ('')," & vbCrLf
         sSql = sSql & " [TOOL_CLASS] [char](12) NULL CONSTRAINT [DF_TlitTableNew_TOOL_CLASS]  DEFAULT ('')" & vbCrLf
         sSql = sSql & ") ON [PRIMARY]" & vbCrLf
         
         ExecuteScript False, sSql
      End If

      If (Not TableExists("TlitTableNew")) Then    'WRONG -- corrected in UpdateDatabase84
         sSql = "CREATE TABLE [dbo].[TlnhdTableNew](" & vbCrLf
         sSql = sSql & "   [TOOL_NUM] [char](30) NOT NULL," & vbCrLf
         sSql = sSql & "   [TOOL_DTADDED] [char](12) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_CGPONUM] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_CUSTPONUM] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_CGSOPONUM] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_GOVOWNED] [char](10) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_CLASS] [char](12) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_HOMEBLDG] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_HOMEAISLE] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_SHELFNUM] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_GRID] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_LOCNUM] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_COMMENTS] [varchar](1020) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_OWNER] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_ACCTTO] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_SN] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_MAKEPN] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_CAVNUM] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_DIM] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_BLANKPONUM] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_MONUM] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_TOOLMATSTAT] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_SRVSTAT] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_DISPSTAT] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_STORAGESTAT] [char](20) NULL" & vbCrLf
         sSql = sSql & " CONSTRAINT [PK_TlnhdTableNew_TOOL_NUM] PRIMARY KEY CLUSTERED" & vbCrLf
         sSql = sSql & "(" & vbCrLf
         sSql = sSql & "   [TOOL_NUM] Asc" & vbCrLf
         sSql = sSql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 80) ON [PRIMARY]" & vbCrLf
         sSql = sSql & ") ON [PRIMARY]"

         ExecuteScript False, sSql
         
         sSql = "ALTER TABLE TlnhdTableNew Add TOOL_ITAR int NULL CONSTRAINT DF__TlnhdTabl__TOOL_ITAR default(0)"
         ExecuteScript False, sSql

      End If
      
      If StoreProcedureExists("RptHistoryWCQListStats") Then
         sSql = "DROP PROCEDURE RptHistoryWCQListStats"
         ExecuteScript False, sSql
      End If
      
      
      If StoreProcedureExists("RptWCQListStats") Then
         sSql = "DROP PROCEDURE RptWCQListStats"
         ExecuteScript False, sSql
      End If
      
      sSql = "CREATE PROCEDURE [dbo].[RptHistoryWCQListStats]" & vbCrLf
      sSql = sSql & " @BeginDate  as varchar(30),@EdnDate as varchar(10)," & vbCrLf
      sSql = sSql & " @OpShop as varchar(12), @OpCenter as varchar(12)  " & vbCrLf
      sSql = sSql & " AS" & vbCrLf
      sSql = sSql & " BEGIN " & vbCrLf
      sSql = sSql & "   IF (@OpShop = '')  " & vbCrLf
      sSql = sSql & "   BEGIN   " & vbCrLf
      sSql = sSql & "      SET @OpShop = '%'    " & vbCrLf
      sSql = sSql & "   End " & vbCrLf
      sSql = sSql & "                " & vbCrLf
      sSql = sSql & "   IF (@OpCenter = '')    " & vbCrLf
      sSql = sSql & "   BEGIN   " & vbCrLf
      sSql = sSql & "      SET @OpCenter = '%'        " & vbCrLf
      sSql = sSql & "   End                       " & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   select distinct runstable.Runref, runstable.runno, OPCENTER, RUNSTATUS, RUNQTY," & vbCrLf
      sSql = sSql & "      runopcur,RnopTable.OPQDATE,  RnopTable.OPSCHEDDATE," & vbCrLf
      sSql = sSql & "      f.PrevOPNO, f.OPQDATE PREVQDATE, f.OPCOMPDATE PREVCOMPDATE," & vbCrLf
      sSql = sSql & "      DATEDIFF(day,f.OPCOMPDATE, RnopTable.OPQDATE) DtDiffQue," & vbCrLf
      sSql = sSql & "      DATEDIFF(day, f.OPCOMPDATE, GETDATE()) DtDiffNow" & vbCrLf
      sSql = sSql & "      from runstable,RnopTable," & vbCrLf
      sSql = sSql & "      (select a.runref, a.runno, b.OPNO PrevOPNO, OPQDATE, OPCOMPDATE," & vbCrLf
      sSql = sSql & "          ROW_NUMBER() OVER (PARTITION BY opref, oprun" & vbCrLf
      sSql = sSql & "                     ORDER BY opref DESC, oprun DESC, OPNO desc) as rn" & vbCrLf
      sSql = sSql & "        from runstable a,rnopTable b" & vbCrLf
      sSql = sSql & "        where a.runref = b.opref and" & vbCrLf
      sSql = sSql & "          a.RunNO = b.oprun And b.opno < a.runopcur" & vbCrLf
      sSql = sSql & "      ) as f" & vbCrLf
      sSql = sSql & "   where RnopTable.OPSCHEDDATE between @BeginDate and @EdnDate" & vbCrLf
      sSql = sSql & "      and RunsTable.runref =  f.runref AND RunsTable.runno = f.runno" & vbCrLf
      sSql = sSql & "      and RnopTable.opref = RunsTable.runref AND RunsTable.runno = RnopTable.OPRUN" & vbCrLf
      sSql = sSql & "      and RunsTable.runopcur = RnopTable.Opno" & vbCrLf
      sSql = sSql & "      and RnopTable.OPSHOP LIKE @OpShop AND RnopTable.OPCENTER LIKE @OpCenter" & vbCrLf
      sSql = sSql & "       --AND RnopTable.OPCOMPLETE = 0" & vbCrLf
      sSql = sSql & "      --and f.rn = 1" & vbCrLf
      sSql = sSql & " End"
         
      ExecuteScript False, sSql
      
      
      sSql = "CREATE PROCEDURE [dbo].[RptWCQListStats]" & vbCrLf
      sSql = sSql & " @BeginDate  as varchar(30),@EdnDate as varchar(10)," & vbCrLf
      sSql = sSql & " @OpShop as varchar(12), @OpCenter as varchar(12)  " & vbCrLf
      sSql = sSql & " AS" & vbCrLf
      sSql = sSql & " BEGIN " & vbCrLf
      sSql = sSql & " IF (@OpShop = '')  " & vbCrLf
      sSql = sSql & " BEGIN   " & vbCrLf
      sSql = sSql & "     SET @OpShop = '%'    " & vbCrLf
      sSql = sSql & " End " & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & " IF (@OpCenter = '')    " & vbCrLf
      sSql = sSql & " BEGIN   " & vbCrLf
      sSql = sSql & "     SET @OpCenter = '%'        " & vbCrLf
      sSql = sSql & " End                       " & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & " select distinct runstable.Runref, runstable.runno, OPCENTER, RUNSTATUS, RUNQTY," & vbCrLf
      sSql = sSql & "     runopcur,RnopTable.OPQDATE,  RnopTable.OPSCHEDDATE," & vbCrLf
      sSql = sSql & "     f.PrevOPNO, f.OPQDATE PREVQDATE, f.OPSCHEDDATE PREVSCHEDDATE," & vbCrLf
      sSql = sSql & "     DATEDIFF(day,f.OPSCHEDDATE, RnopTable.OPQDATE) DtDiffQue," & vbCrLf
      sSql = sSql & "     DATEDIFF(day, f.OPSCHEDDATE, GETDATE()) DtDiffNow" & vbCrLf
      sSql = sSql & "    from runstable,RnopTable," & vbCrLf
      sSql = sSql & "    (select a.runref, a.runno, b.OPNO PrevOPNO, OPQDATE, OPSCHEDDATE," & vbCrLf
      sSql = sSql & "          ROW_NUMBER() OVER (PARTITION BY opref, oprun" & vbCrLf
      sSql = sSql & "                        ORDER BY opref DESC, oprun DESC, OPNO desc) as rn" & vbCrLf
      sSql = sSql & "       from runstable a,rnopTable b" & vbCrLf
      sSql = sSql & "       where a.runref = b.opref and" & vbCrLf
      sSql = sSql & "          a.RunNO = b.oprun And b.opno < a.runopcur" & vbCrLf
      sSql = sSql & "    ) as f" & vbCrLf
      sSql = sSql & " where RunsTable.RunSched between @BeginDate and @EdnDate" & vbCrLf
      sSql = sSql & "    and RunsTable.runref =  f.runref AND RunsTable.runno = f.runno" & vbCrLf
      sSql = sSql & "    and RnopTable.opref = RunsTable.runref AND RunsTable.runno = RnopTable.OPRUN" & vbCrLf
      sSql = sSql & "    and RunsTable.runopcur = RnopTable.Opno" & vbCrLf
      sSql = sSql & "    and RnopTable.OPSHOP LIKE @OpShop AND RnopTable.OPCENTER LIKE @OpCenter" & vbCrLf
      sSql = sSql & "    and f.rn = 1" & vbCrLf
      sSql = sSql & " End"
      
      ExecuteScript False, sSql
      
      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver
   
   End If
End Function
   
   
'''''''''''''''''''''''''''''''''

Private Function UpdateDatabase79()

   Dim sql As String
   sql = ""
   
   newver = 154
   If ver < newver Then
   
        clsADOCon.ADOErrNum = 0
        
'''''''''''''''''''''''''''''''''''''''''''''''''''
      
' required for Vendor Statement Report
sql = "alter table EsReportVendorStmt add Journal varchar(12) NULL" & vbCrLf
ExecuteScript False, sql

' column does not exist in AWJ database, so just add it in this script
sql = "alter table ComnTable add COWARNSERVICEOPOPEN int NOT NULL DEFAULT(0)" & vbCrLf
ExecuteScript False, sql

' add flag to turn on sheete inventory functions
sql = "alter table ComnTable add COUSESHEETINVENTORY int NOT NULL DEFAULT(0) " & vbCrLf
ExecuteScript False, sql

' lot header additions
sql = "alter table LohdTable add LOTRESERVEDBY varchar(4)" & vbCrLf
ExecuteScript False, sql
sql = "alter table LohdTable add LOTRESERVEDON DATETIME" & vbCrLf
ExecuteScript False, sql
sql = "alter table LohdTable add LOTCERT varchar(40)" & vbCrLf
ExecuteScript False, sql

' lot item additions
sql = "alter table LoitTable add LOIHEIGHT decimal(12,4)" & vbCrLf
ExecuteScript False, sql
sql = "alter table LoitTable add LOILENGTH decimal(12,4)" & vbCrLf
ExecuteScript False, sql
sql = "alter table LoitTable add LOIAREA as cast(isnull(LOIHEIGHT,0) * isnull(LOILENGTH,0) as decimal(12,4))" & vbCrLf
ExecuteScript False, sql
sql = "alter table LoitTable add LOIUSER varchar(4)" & vbCrLf
ExecuteScript False, sql
sql = "alter table LoitTable add LOIPARENTREC int" & vbCrLf
ExecuteScript False, sql
sql = "alter table LoitTable add LOISONUMBER int" & vbCrLf
ExecuteScript False, sql
sql = "alter table LoitTable add LOISHEETACTTYPE char(2)        -- PK,RS" & vbCrLf
ExecuteScript False, sql
sql = "alter table LoitTable add LOIINACTIVE bit                -- = 1 to exclude from future sheet processes" & vbCrLf
ExecuteScript False, sql

sql = "if object_id('SheetPick') IS NOT NULL" & vbCrLf
sql = sql & "    drop procedure SheetPick" & vbCrLf
ExecuteScript False, sql

sql = "create procedure dbo.SheetPick" & vbCrLf
sql = sql & "   @UserLotNo varchar(40),     -- lotuserlotid" & vbCrLf
sql = sql & "   @User varchar(4)," & vbCrLf
sql = sql & "   @SO int" & vbCrLf
sql = sql & "   " & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "Pick a sheet" & vbCrLf
sql = sql & "exec SheetPick '029373-1-A','MGR',222222" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- find all rectangles in lot" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "create table #rect" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "   id int identity," & vbCrLf
sql = sql & "   ParentRecord int," & vbCrLf
sql = sql & "   Height decimal(12,4)," & vbCrLf
sql = sql & "   Length decimal(12,4)," & vbCrLf
sql = sql & "   NewRecord int," & vbCrLf
sql = sql & "   Qty decimal(12,4)" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @LotNo char(15)" & vbCrLf
sql = sql & "select @LotNo = LOTNUMBER from LohdTable where LOTUSERLOTID = @UserLotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert #rect (ParentRecord,Height,Length,NewRecord,Qty)" & vbCrLf
sql = sql & "select li.LOIRECORD,li.LOIHEIGHT,li.LOILENGTH,0,cast(li.LOIHEIGHT*li.LOILENGTH as decimal(12,4)) from LoitTable li" & vbCrLf
sql = sql & "   join LohdTable lh on lh.LOTNUMBER = li.LOINUMBER" & vbCrLf
sql = sql & "   where lh.LOTNUMBER = @LotNo and LOIQUANTITY > 0 " & vbCrLf
sql = sql & "   --and LOTRESERVEDON is not null" & vbCrLf
sql = sql & "   and (li.LOIINACTIVE is null or li.LOIINACTIVE = 0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- assign new recordnumbers" & vbCrLf
sql = sql & "declare @max int" & vbCrLf
sql = sql & "select @max = max(LOIRECORD) from LoitTable where LOINUMBER = @LotNo" & vbCrLf
sql = sql & "update #rect set NewRecord = @max + id" & vbCrLf
sql = sql & "   " & vbCrLf
sql = sql & "--select * from #rect  " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @partRef varchar(30), @currentTime datetime, @uom char(2), @type int" & vbCrLf
sql = sql & "declare @unitCost decimal(12,4), @nextINNUMBER int, @totalQty decimal(12,4)" & vbCrLf
sql = sql & "select @partRef = LOTPARTREF, @unitCost = LOTUNITCOST, @totalQty = LOTREMAININGQTY from LohdTable where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "select @uom = PAUNITS from PartTable where PARTREF = @partRef" & vbCrLf
sql = sql & "set @currentTime = getdate()" & vbCrLf
sql = sql & "select @nextINNUMBER = max(INNUMBER) + 1 from InvaTable" & vbCrLf
sql = sql & "set @type = 19     -- manual adjustment" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- deactivate rectangles being picked so they won't show up again" & vbCrLf
sql = sql & "Update li" & vbCrLf
sql = sql & "   set li.LOIINACTIVE = 1" & vbCrLf
sql = sql & "from LoitTable li " & vbCrLf
sql = sql & "join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD = r.ParentRecord" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from LoitTable li " & vbCrLf
sql = sql & "--join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD = r.ParentRecord" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new LoitTable record to zero lot quantity" & vbCrLf
sql = sql & "INSERT INTO [dbo].[LoitTable]" & vbCrLf
sql = sql & "           ([LOINUMBER]" & vbCrLf
sql = sql & "           ,[LOIRECORD]" & vbCrLf
sql = sql & "           ,[LOITYPE]" & vbCrLf
sql = sql & "           ,[LOIPARTREF]" & vbCrLf
sql = sql & "           ,[LOIADATE]" & vbCrLf
sql = sql & "           ,[LOIPDATE]" & vbCrLf
sql = sql & "           ,[LOIQUANTITY]" & vbCrLf
sql = sql & "           ,[LOIMOPARTREF]" & vbCrLf
sql = sql & "           ,[LOIMORUNNO]" & vbCrLf
sql = sql & "           ,[LOIPONUMBER]" & vbCrLf
sql = sql & "           ,[LOIPOITEM]" & vbCrLf
sql = sql & "           ,[LOIPOREV]" & vbCrLf
sql = sql & "           ,[LOIPSNUMBER]" & vbCrLf
sql = sql & "           ,[LOIPSITEM]" & vbCrLf
sql = sql & "           ,[LOICUSTINVNO]" & vbCrLf
sql = sql & "           ,[LOICUST]" & vbCrLf
sql = sql & "           ,[LOIVENDINVNO]" & vbCrLf
sql = sql & "           ,[LOIVENDOR]" & vbCrLf
sql = sql & "           ,[LOIACTIVITY]" & vbCrLf
sql = sql & "           ,[LOICOMMENT]" & vbCrLf
sql = sql & "           ,[LOIUNITS]" & vbCrLf
sql = sql & "           ,[LOIMOPKCANCEL]" & vbCrLf
sql = sql & "           ,[LOIHEIGHT]" & vbCrLf
sql = sql & "           ,[LOILENGTH]" & vbCrLf
sql = sql & "           ,[LOIUSER]" & vbCrLf
sql = sql & "          ,LOIPARENTREC" & vbCrLf
sql = sql & "           ,[LOISONUMBER]" & vbCrLf
sql = sql & "          ,LOISHEETACTTYPE" & vbCrLf
sql = sql & "          )" & vbCrLf
sql = sql & "     SELECT" & vbCrLf
sql = sql & "           @LotNo" & vbCrLf
sql = sql & "           ,r.NewRecord" & vbCrLf
sql = sql & "           ,@type              -- manual adjustment" & vbCrLf
sql = sql & "           ,@partRef" & vbCrLf
sql = sql & "           ,@currentTime" & vbCrLf
sql = sql & "           ,@currentTime" & vbCrLf
sql = sql & "           ,-r.Qty" & vbCrLf
sql = sql & "           ,null   --LOIMOPARTREF, char(30),>" & vbCrLf
sql = sql & "           ,null   --<LOIMORUNNO, int,>" & vbCrLf
sql = sql & "           ,null   --<LOIPONUMBER, int,>" & vbCrLf
sql = sql & "           ,null   --<LOIPOITEM, smallint,>" & vbCrLf
sql = sql & "           ,null   --<LOIPOREV, char(2),>" & vbCrLf
sql = sql & "           ,null   --<LOIPSNUMBER, char(8),>" & vbCrLf
sql = sql & "           ,null   --<LOIPSITEM, smallint,>" & vbCrLf
sql = sql & "           ,null   --<LOICUSTINVNO, int,>" & vbCrLf
sql = sql & "           ,null   --<LOICUST, char(10),>" & vbCrLf
sql = sql & "           ,null   --<LOIVENDINVNO, char(20),>" & vbCrLf
sql = sql & "           ,null   --<LOIVENDOR, char(10),>" & vbCrLf
sql = sql & "           ,null   --<LOIACTIVITY, int,> points to InvaTable.INNO when IA is created" & vbCrLf
sql = sql & "           ,'sheet pick'   --<LOICOMMENT, varchar(40),>" & vbCrLf
sql = sql & "           ,@uom   --<LOIUNITS, char(2),>" & vbCrLf
sql = sql & "           ,null   --<LOIMOPKCANCEL, smallint,>" & vbCrLf
sql = sql & "           ,r.HEIGHT   --<LOIHEIGHT, decimal(12,4),>" & vbCrLf
sql = sql & "           ,r.LENGTH   --<LOILENGTH, decimal(12,4),>" & vbCrLf
sql = sql & "           ,@User          --<LOIUSER, varchar(4),>" & vbCrLf
sql = sql & "          ,r.ParentRecord  --LOIPARENTREC" & vbCrLf
sql = sql & "           ,@SO                --<LOISONUMBER, int,>" & vbCrLf
sql = sql & "          ,'PK'            --LOISHEETACTTYPE" & vbCrLf
sql = sql & "       from #rect r" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from LoitTable li " & vbCrLf
sql = sql & "--join #rect r on li.LOINUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- set lot quantity = 0" & vbCrLf
sql = sql & "update LohdTable set LOTREMAININGQTY = 0 where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new IA record to reduce part quantity" & vbCrLf
sql = sql & "INSERT INTO [dbo].[InvaTable]" & vbCrLf
sql = sql & "           ([INTYPE]" & vbCrLf
sql = sql & "           ,[INPART]" & vbCrLf
sql = sql & "           ,[INREF1]" & vbCrLf
sql = sql & "           ,[INREF2]" & vbCrLf
sql = sql & "           ,[INPDATE]" & vbCrLf
sql = sql & "           ,[INADATE]" & vbCrLf
sql = sql & "           ,[INPQTY]" & vbCrLf
sql = sql & "           ,[INAQTY]" & vbCrLf
sql = sql & "           ,[INAMT]" & vbCrLf
sql = sql & "           ,[INTOTMATL]" & vbCrLf
sql = sql & "           ,[INTOTLABOR]" & vbCrLf
sql = sql & "           ,[INTOTEXP]" & vbCrLf
sql = sql & "           ,[INTOTOH]" & vbCrLf
sql = sql & "           ,[INTOTHRS]" & vbCrLf
sql = sql & "           ,[INCREDITACCT]" & vbCrLf
sql = sql & "           ,[INDEBITACCT]" & vbCrLf
sql = sql & "           ,[INGLJOURNAL]" & vbCrLf
sql = sql & "           ,[INGLPOSTED]" & vbCrLf
sql = sql & "           ,[INGLDATE]" & vbCrLf
sql = sql & "           ,[INMOPART]" & vbCrLf
sql = sql & "           ,[INMORUN]" & vbCrLf
sql = sql & "           ,[INSONUMBER]" & vbCrLf
sql = sql & "           ,[INSOITEM]" & vbCrLf
sql = sql & "           ,[INSOREV]" & vbCrLf
sql = sql & "           ,[INPONUMBER]" & vbCrLf
sql = sql & "           ,[INPORELEASE]" & vbCrLf
sql = sql & "           ,[INPOITEM]" & vbCrLf
sql = sql & "           ,[INPOREV]" & vbCrLf
sql = sql & "           ,[INPSNUMBER]" & vbCrLf
sql = sql & "           ,[INPSITEM]" & vbCrLf
sql = sql & "           ,[INWIPLABACCT]" & vbCrLf
sql = sql & "           ,[INWIPMATACCT]" & vbCrLf
sql = sql & "           ,[INWIPOHDACCT]" & vbCrLf
sql = sql & "           ,[INWIPEXPACCT]" & vbCrLf
sql = sql & "           ,[INNUMBER]" & vbCrLf
sql = sql & "           ,[INLOTNUMBER]" & vbCrLf
sql = sql & "           ,[INUSER]" & vbCrLf
sql = sql & "           ,[INUNITS]" & vbCrLf
sql = sql & "           ,[INDRLABACCT]" & vbCrLf
sql = sql & "           ,[INDRMATACCT]" & vbCrLf
sql = sql & "           ,[INDREXPACCT]" & vbCrLf
sql = sql & "           ,[INDROHDACCT]" & vbCrLf
sql = sql & "           ,[INCRLABACCT]" & vbCrLf
sql = sql & "           ,[INCRMATACCT]" & vbCrLf
sql = sql & "           ,[INCREXPACCT]" & vbCrLf
sql = sql & "           ,[INCROHDACCT]" & vbCrLf
sql = sql & "           ,[INLOTTRACK]" & vbCrLf
sql = sql & "           ,[INUSEACTUALCOST]" & vbCrLf
sql = sql & "           ,[INCOSTEDBY]" & vbCrLf
sql = sql & "           ,[INMAINTCOSTED])" & vbCrLf
sql = sql & "     select" & vbCrLf
sql = sql & "           @type               --<INTYPE, int,>" & vbCrLf
sql = sql & "           ,@partRef           --char(30),>" & vbCrLf
sql = sql & "           ,'Manual Adjustment'    --<INREF1, char(20),>" & vbCrLf
sql = sql & "           ,'Sheet Inventory'  --<INREF2, char(40),>" & vbCrLf
sql = sql & "           ,@currentTime       --<INPDATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,@currentTime       --<INADATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,-r.Qty         --<INPQTY, decimal(12,4),>" & vbCrLf
sql = sql & "           ,-r.Qty         --<INAQTY, decimal(12,4),>" & vbCrLf
sql = sql & "           ,@unitCost          -- <INAMT, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTMATL, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTLABOR, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTEXP, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTOH, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTHRS, decimal(12,4),>" & vbCrLf
sql = sql & "           ,''                 --<INCREDITACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDEBITACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INGLJOURNAL, char(12),>" & vbCrLf
sql = sql & "           ,0                  --<INGLPOSTED, tinyint,>" & vbCrLf
sql = sql & "           ,null               --<INGLDATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,''                 --<INMOPART, char(30),>" & vbCrLf
sql = sql & "           ,0                  --<INMORUN, int,>" & vbCrLf
sql = sql & "           ,0                  --<INSONUMBER, int,>" & vbCrLf
sql = sql & "           ,0                  --<INSOITEM, int,>" & vbCrLf
sql = sql & "           ,''                 --<INSOREV, char(2),>" & vbCrLf
sql = sql & "           ,0                  --<INPONUMBER, int,>" & vbCrLf
sql = sql & "           ,0                  --<INPORELEASE, smallint,>" & vbCrLf
sql = sql & "           ,0                  --<INPOITEM, smallint,>" & vbCrLf
sql = sql & "           ,''                 --<INPOREV, char(2),>" & vbCrLf
sql = sql & "           ,''                 --<INPSNUMBER, char(8),>" & vbCrLf
sql = sql & "           ,0                  --<INPSITEM, smallint,>" & vbCrLf
sql = sql & "           ,''                 --<INWIPLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPOHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPEXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,@nextINNUMBER      --<INNUMBER, int,>" & vbCrLf
sql = sql & "           ,@LotNo             --<INLOTNUMBER, char(15),>" & vbCrLf
sql = sql & "           ,@User              --<INUSER, char(4),>" & vbCrLf
sql = sql & "           ,@uom               --<INUNITS, char(2),>" & vbCrLf
sql = sql & "           ,''                 --<INDRLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDRMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDREXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDROHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCRLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCRMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCREXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCROHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,null               --<INLOTTRACK, bit,>" & vbCrLf
sql = sql & "           ,null               --<INUSEACTUALCOST, bit,>" & vbCrLf
sql = sql & "           ,null               --<INCOSTEDBY, char(4),>" & vbCrLf
sql = sql & "           ,0                  --<INMAINTCOSTED, int,>)" & vbCrLf
sql = sql & "       from #rect r" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now update part QOH" & vbCrLf
sql = sql & "update PartTable set PAQOH = PAQOH - @totalQty where PARTREF = @partRef" & vbCrLf
sql = sql & "commit tran" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql

sql = "if object_id('SheetRestock') IS NOT NULL" & vbCrLf
sql = sql & "    drop procedure SheetRestock" & vbCrLf
ExecuteScript False, sql

sql = "create procedure dbo.SheetRestock" & vbCrLf
sql = sql & "   @UserLotNo varchar(40)," & vbCrLf
sql = sql & "   @User varchar(4)," & vbCrLf
sql = sql & "   @Comments varchar(2048)," & vbCrLf
sql = sql & "   @Location varchar(4)," & vbCrLf
sql = sql & "   @Params varchar(2000)   -- LOIRECORD,NEWHT,NEWLEN,... repeat (include comma at end)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "restock sheet LOI records" & vbCrLf
sql = sql & "exec SheetRestock '029373-1-A', 'MGR', '2,36,96,0,12,72,'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "create table #rect" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "   ParentRecord int," & vbCrLf
sql = sql & "   NewHeight decimal(12,4)," & vbCrLf
sql = sql & "   NewLength decimal(12,4)," & vbCrLf
sql = sql & "   NewRecord int," & vbCrLf
sql = sql & "   Qty decimal(12,4)," & vbCrLf
sql = sql & "   NewIANumber int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @nextRecord int, @nextIANumber int" & vbCrLf
sql = sql & "declare @LotNo char(15), @id int, @SO int" & vbCrLf
sql = sql & "select @LotNo = LOTNUMBER from LohdTable where LOTUSERLOTID = @UserLotNo" & vbCrLf
sql = sql & "select @nextRecord = max(LOIRECORD) + 1 , @SO = max(LOISONUMBER) from LoitTable where LOINUMBER = @LotNo" & vbCrLf
sql = sql & "select @nextIANumber = max(INNUMBER) + 1 from InvaTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- first update the lot header" & vbCrLf
sql = sql & "update LohdTable set LOTCOMMENTS = @Comments, LOTLOCATION = @Location where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- extract parameters" & vbCrLf
sql = sql & "DECLARE @start INT, @end INT" & vbCrLf
sql = sql & "declare @stringId varchar(10), @stringHt varchar(10), @stringLen varchar(10)" & vbCrLf
sql = sql & "SELECT @start = 1, @end = CHARINDEX(',', @Params) " & vbCrLf
sql = sql & "WHILE @start < LEN(@Params) + 1 BEGIN " & vbCrLf
sql = sql & "    IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    set @stringId = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "    SET @start = @end + 1 " & vbCrLf
sql = sql & "    SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    set @stringHt = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "    SET @start = @end + 1 " & vbCrLf
sql = sql & "    SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    set @stringLen = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "    SET @start = @end + 1 " & vbCrLf
sql = sql & "    SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "   " & vbCrLf
sql = sql & "   declare @ht decimal(12,4), @len decimal(12,4)" & vbCrLf
sql = sql & "   set @ht = cast(@stringHt as decimal(12,4))" & vbCrLf
sql = sql & "   set @len = cast(@stringLen as decimal(12,4))" & vbCrLf
sql = sql & "   insert into #rect (ParentRecord, NewHeight, NewLength, NewRecord, NewIANumber)" & vbCrLf
sql = sql & "   values (cast(@stringId as int), @ht, @len," & vbCrLf
sql = sql & "       case when @ht*@len = 0 then 0 else @nextRecord end," & vbCrLf
sql = sql & "       case when @ht*@len = 0 then 0 else @nextIANumber end)" & vbCrLf
sql = sql & "   if (@ht*@len) <> 0 " & vbCrLf
sql = sql & "   begin" & vbCrLf
sql = sql & "       set @nextRecord = @nextRecord + 1" & vbCrLf
sql = sql & "       set @nextIANumber = @nextIANumber + 1" & vbCrLf
sql = sql & "   end" & vbCrLf
sql = sql & "END " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #rect set Qty = NewHeight * NewLength" & vbCrLf
sql = sql & "update #rect set ParentRecord = isnull((select top 1 ParentRecord from #rect where ParentRecord <> 0 ),0)" & vbCrLf
sql = sql & "where ParentRecord = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from #rect" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @partRef varchar(30), @currentTime datetime, @uom char(2), @type int" & vbCrLf
sql = sql & "declare @unitCost decimal(12,4), @nextINNUMBER int, @totalQty decimal(12,4)" & vbCrLf
sql = sql & "select @partRef = LOTPARTREF, @unitCost = LOTUNITCOST, @totalQty = LOTREMAININGQTY from LohdTable where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "select @uom = PAUNITS from PartTable where PARTREF = @partRef" & vbCrLf
sql = sql & "set @currentTime = getdate()" & vbCrLf
sql = sql & "select @nextINNUMBER = max(INNUMBER) + 1 from InvaTable" & vbCrLf
sql = sql & "set @type = 19     -- manual adjustment" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update existing lot items" & vbCrLf
sql = sql & "Update li" & vbCrLf
sql = sql & "   set LOIINACTIVE = 1" & vbCrLf
sql = sql & "from LoitTable li " & vbCrLf
sql = sql & "join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD = r.ParentRecord" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select distinct li.* from LoitTable li " & vbCrLf
sql = sql & "--join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD = r.ParentRecord" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- remove rectangles with no quantity remaining.  These items will not be restocked." & vbCrLf
sql = sql & "delete from #rect where Qty = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new lot items" & vbCrLf
sql = sql & "INSERT INTO [dbo].[LoitTable]" & vbCrLf
sql = sql & "           ([LOINUMBER]" & vbCrLf
sql = sql & "           ,[LOIRECORD]" & vbCrLf
sql = sql & "           ,[LOITYPE]" & vbCrLf
sql = sql & "           ,[LOIPARTREF]" & vbCrLf
sql = sql & "           ,[LOIADATE]" & vbCrLf
sql = sql & "           ,[LOIPDATE]" & vbCrLf
sql = sql & "           ,[LOIQUANTITY]" & vbCrLf
sql = sql & "           ,[LOIMOPARTREF]" & vbCrLf
sql = sql & "           ,[LOIMORUNNO]" & vbCrLf
sql = sql & "           ,[LOIPONUMBER]" & vbCrLf
sql = sql & "           ,[LOIPOITEM]" & vbCrLf
sql = sql & "           ,[LOIPOREV]" & vbCrLf
sql = sql & "           ,[LOIPSNUMBER]" & vbCrLf
sql = sql & "           ,[LOIPSITEM]" & vbCrLf
sql = sql & "           ,[LOICUSTINVNO]" & vbCrLf
sql = sql & "           ,[LOICUST]" & vbCrLf
sql = sql & "           ,[LOIVENDINVNO]" & vbCrLf
sql = sql & "           ,[LOIVENDOR]" & vbCrLf
sql = sql & "           ,[LOIACTIVITY]" & vbCrLf
sql = sql & "           ,[LOICOMMENT]" & vbCrLf
sql = sql & "           ,[LOIUNITS]" & vbCrLf
sql = sql & "           ,[LOIMOPKCANCEL]" & vbCrLf
sql = sql & "           ,[LOIHEIGHT]" & vbCrLf
sql = sql & "           ,[LOILENGTH]" & vbCrLf
sql = sql & "           ,[LOIUSER]" & vbCrLf
sql = sql & "          ,LOIPARENTREC" & vbCrLf
sql = sql & "           ,[LOISONUMBER]" & vbCrLf
sql = sql & "          ,LOISHEETACTTYPE" & vbCrLf
sql = sql & "          )" & vbCrLf
sql = sql & "     SELECT" & vbCrLf
sql = sql & "           @LotNo" & vbCrLf
sql = sql & "           ,r.NewRecord" & vbCrLf
sql = sql & "           ,@type              -- manual adjustment" & vbCrLf
sql = sql & "           ,@partRef" & vbCrLf
sql = sql & "           ,@currentTime" & vbCrLf
sql = sql & "           ,@currentTime" & vbCrLf
sql = sql & "           ,r.Qty" & vbCrLf
sql = sql & "           ,null   --LOIMOPARTREF, char(30),>" & vbCrLf
sql = sql & "           ,null   --<LOIMORUNNO, int,>" & vbCrLf
sql = sql & "           ,null   --<LOIPONUMBER, int,>" & vbCrLf
sql = sql & "           ,null   --<LOIPOITEM, smallint,>" & vbCrLf
sql = sql & "           ,null   --<LOIPOREV, char(2),>" & vbCrLf
sql = sql & "           ,null   --<LOIPSNUMBER, char(8),>" & vbCrLf
sql = sql & "           ,null   --<LOIPSITEM, smallint,>" & vbCrLf
sql = sql & "           ,null   --<LOICUSTINVNO, int,>" & vbCrLf
sql = sql & "           ,null   --<LOICUST, char(10),>" & vbCrLf
sql = sql & "           ,null   --<LOIVENDINVNO, char(20),>" & vbCrLf
sql = sql & "           ,null   --<LOIVENDOR, char(10),>" & vbCrLf
sql = sql & "           ,null   --<LOIACTIVITY, int,> points to InvaTable.INNO when IA is created" & vbCrLf
sql = sql & "           ,'sheet restock'    --<LOICOMMENT, varchar(40),>" & vbCrLf
sql = sql & "           ,@uom   --<LOIUNITS, char(2),>" & vbCrLf
sql = sql & "           ,null   --<LOIMOPKCANCEL, smallint,>" & vbCrLf
sql = sql & "           ,r.NewHeight    --<LOIHEIGHT, decimal(12,4),>" & vbCrLf
sql = sql & "           ,r.NewLength    --<LOILENGTH, decimal(12,4),>" & vbCrLf
sql = sql & "           ,@User      --<LOIUSER, varchar(4),>" & vbCrLf
sql = sql & "          ,r.ParentRecord  --LOIPICKEDFROMREC" & vbCrLf
sql = sql & "           ,@SO        --<LOISONUMBER, int,>" & vbCrLf
sql = sql & "          ,'RS'    --LOISHEETACTTYPE" & vbCrLf
sql = sql & "       from #rect r" & vbCrLf
sql = sql & "       where r.Qty <> 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select distinct li.* from LoitTable li " & vbCrLf
sql = sql & "--join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD > 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- set LOTREMAININGQTY and remove reservation" & vbCrLf
sql = sql & "declare @sum decimal(12,4)" & vbCrLf
sql = sql & "select @sum = sum(Qty) from #rect" & vbCrLf
sql = sql & "update LohdTable " & vbCrLf
sql = sql & "   set LOTREMAININGQTY = LOTREMAININGQTY + @sum," & vbCrLf
sql = sql & "   LOTRESERVEDBY = NULL," & vbCrLf
sql = sql & "   LOTRESERVEDON = NULL" & vbCrLf
sql = sql & "   where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new ia items" & vbCrLf
sql = sql & "INSERT INTO [dbo].[InvaTable]" & vbCrLf
sql = sql & "           ([INTYPE]" & vbCrLf
sql = sql & "           ,[INPART]" & vbCrLf
sql = sql & "           ,[INREF1]" & vbCrLf
sql = sql & "           ,[INREF2]" & vbCrLf
sql = sql & "           ,[INPDATE]" & vbCrLf
sql = sql & "           ,[INADATE]" & vbCrLf
sql = sql & "           ,[INPQTY]" & vbCrLf
sql = sql & "           ,[INAQTY]" & vbCrLf
sql = sql & "           ,[INAMT]" & vbCrLf
sql = sql & "           ,[INTOTMATL]" & vbCrLf
sql = sql & "           ,[INTOTLABOR]" & vbCrLf
sql = sql & "           ,[INTOTEXP]" & vbCrLf
sql = sql & "           ,[INTOTOH]" & vbCrLf
sql = sql & "           ,[INTOTHRS]" & vbCrLf
sql = sql & "           ,[INCREDITACCT]" & vbCrLf
sql = sql & "           ,[INDEBITACCT]" & vbCrLf
sql = sql & "           ,[INGLJOURNAL]" & vbCrLf
sql = sql & "           ,[INGLPOSTED]" & vbCrLf
sql = sql & "           ,[INGLDATE]" & vbCrLf
sql = sql & "           ,[INMOPART]" & vbCrLf
sql = sql & "           ,[INMORUN]" & vbCrLf
sql = sql & "           ,[INSONUMBER]" & vbCrLf
sql = sql & "           ,[INSOITEM]" & vbCrLf
sql = sql & "           ,[INSOREV]" & vbCrLf
sql = sql & "           ,[INPONUMBER]" & vbCrLf
sql = sql & "           ,[INPORELEASE]" & vbCrLf
sql = sql & "           ,[INPOITEM]" & vbCrLf
sql = sql & "           ,[INPOREV]" & vbCrLf
sql = sql & "           ,[INPSNUMBER]" & vbCrLf
sql = sql & "           ,[INPSITEM]" & vbCrLf
sql = sql & "           ,[INWIPLABACCT]" & vbCrLf
sql = sql & "           ,[INWIPMATACCT]" & vbCrLf
sql = sql & "           ,[INWIPOHDACCT]" & vbCrLf
sql = sql & "           ,[INWIPEXPACCT]" & vbCrLf
sql = sql & "           ,[INNUMBER]" & vbCrLf
sql = sql & "           ,[INLOTNUMBER]" & vbCrLf
sql = sql & "           ,[INUSER]" & vbCrLf
sql = sql & "           ,[INUNITS]" & vbCrLf
sql = sql & "           ,[INDRLABACCT]" & vbCrLf
sql = sql & "           ,[INDRMATACCT]" & vbCrLf
sql = sql & "           ,[INDREXPACCT]" & vbCrLf
sql = sql & "           ,[INDROHDACCT]" & vbCrLf
sql = sql & "           ,[INCRLABACCT]" & vbCrLf
sql = sql & "           ,[INCRMATACCT]" & vbCrLf
sql = sql & "           ,[INCREXPACCT]" & vbCrLf
sql = sql & "           ,[INCROHDACCT]" & vbCrLf
sql = sql & "           ,[INLOTTRACK]" & vbCrLf
sql = sql & "           ,[INUSEACTUALCOST]" & vbCrLf
sql = sql & "           ,[INCOSTEDBY]" & vbCrLf
sql = sql & "           ,[INMAINTCOSTED])" & vbCrLf
sql = sql & "     select" & vbCrLf
sql = sql & "           @type               --<INTYPE, int,>" & vbCrLf
sql = sql & "           ,@partRef           --char(30),>" & vbCrLf
sql = sql & "           ,'Manual Adjustment'    --<INREF1, char(20),>" & vbCrLf
sql = sql & "           ,'Sheet Inventory'  --<INREF2, char(40),>" & vbCrLf
sql = sql & "           ,@currentTime       --<INPDATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,@currentTime       --<INADATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,r.Qty          --<INPQTY, decimal(12,4),>" & vbCrLf
sql = sql & "           ,r.Qty          --<INAQTY, decimal(12,4),>" & vbCrLf
sql = sql & "           ,@unitCost          -- <INAMT, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTMATL, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTLABOR, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTEXP, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTOH, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTHRS, decimal(12,4),>" & vbCrLf
sql = sql & "           ,''                 --<INCREDITACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDEBITACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INGLJOURNAL, char(12),>" & vbCrLf
sql = sql & "           ,0                  --<INGLPOSTED, tinyint,>" & vbCrLf
sql = sql & "           ,null               --<INGLDATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,''                 --<INMOPART, char(30),>" & vbCrLf
sql = sql & "           ,0                  --<INMORUN, int,>" & vbCrLf
sql = sql & "           ,0                  --<INSONUMBER, int,>" & vbCrLf
sql = sql & "           ,0                  --<INSOITEM, int,>" & vbCrLf
sql = sql & "           ,''                 --<INSOREV, char(2),>" & vbCrLf
sql = sql & "           ,0                  --<INPONUMBER, int,>" & vbCrLf
sql = sql & "           ,0                  --<INPORELEASE, smallint,>" & vbCrLf
sql = sql & "           ,0                  --<INPOITEM, smallint,>" & vbCrLf
sql = sql & "           ,''                 --<INPOREV, char(2),>" & vbCrLf
sql = sql & "           ,''                 --<INPSNUMBER, char(8),>" & vbCrLf
sql = sql & "           ,0                  --<INPSITEM, smallint,>" & vbCrLf
sql = sql & "           ,''                 --<INWIPLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPOHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPEXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,r.NewIANumber      --<INNUMBER, int,>" & vbCrLf
sql = sql & "           ,@LotNo             --<INLOTNUMBER, char(15),>" & vbCrLf
sql = sql & "           ,@User              --<INUSER, char(4),>" & vbCrLf
sql = sql & "           ,@uom               --<INUNITS, char(2),>" & vbCrLf
sql = sql & "           ,''                 --<INDRLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDRMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDREXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDROHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCRLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCRMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCREXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCROHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,null               --<INLOTTRACK, bit,>" & vbCrLf
sql = sql & "           ,null               --<INUSEACTUALCOST, bit,>" & vbCrLf
sql = sql & "           ,null               --<INCOSTEDBY, char(4),>" & vbCrLf
sql = sql & "           ,0                  --<INMAINTCOSTED, int,>)" & vbCrLf
sql = sql & "       from #rect r" & vbCrLf
sql = sql & "       where r.Qty <> 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now update part QOH" & vbCrLf
sql = sql & "update PartTable set PAQOH = PAQOH - @totalQty where PARTREF = @partRef" & vbCrLf
sql = sql & "commit tran" & vbCrLf
sql = sql & "" & vbCrLf
Clipboard.Clear
Clipboard.SetText sql
ExecuteScript False, sql

sql = "if object_id('SheetUnitConversion') IS NOT NULL" & vbCrLf
sql = sql & "    drop procedure SheetUnitConversion" & vbCrLf
ExecuteScript False, sql

sql = "create procedure dbo.SheetUnitConversion" & vbCrLf
sql = sql & "   @LotNo as char(15)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "/* convert lot from purchasing units to inventory units" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "exec SheetUnitConversion '42701-729201-79'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get purchase item unit cost, lot dimensions, and units" & vbCrLf
sql = sql & "declare @purchaseUnitCost decimal(12,4), @height decimal(12,4), @length decimal(12,4), " & vbCrLf
sql = sql & "   @purchUnit char(2), @invUnit char(2), @lotQtyAfter decimal(12,4), @unitCost decimal(12,4), " & vbCrLf
sql = sql & "   @purchQty decimal(12,4), @partRef varchar(30), @lotQtyBefore decimal(12,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select @purchaseUnitCost = poi.PIAMT, @height = loh.LOTMATHEIGHT, @length = loh.LOTMATLENGTH," & vbCrLf
sql = sql & "   @purchUnit = pt.PAPUNITS, @invUnit = pt.PAUNITS, @purchQty = poi.PIAQTY," & vbCrLf
sql = sql & "   @lotQtyAfter = loh.LOTMATHEIGHT * loh.LOTMATLENGTH * ROUND((loh.LOTTOTMATL + .1) / POI.PIAMT,0)," & vbCrLf
sql = sql & "   @partRef = pt.PARTREF, @lotQtyBefore = loh.LOTORIGINALQTY " & vbCrLf
sql = sql & "from PoitTable poi" & vbCrLf
sql = sql & "join PohdTable poh on poh.PONUMBER = poi.PINUMBER" & vbCrLf
sql = sql & "join LohdTable loh on loh.LOTPO = poi.PINUMBER and loh.LOTPOITEM = poi.PIITEM and loh.LOTPOITEMREV = poi.PIREV" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = loh.LOTPARTREF" & vbCrLf
sql = sql & "where loh.LOTNUMBER = @LotNo and  pt.PAPUNITS = 'SH' and pt.PAUNITS <> pt.PAPUNITS and loh.LOTMATHEIGHT > 1 and loh.LOTMATLENGTH > 1 " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--print 'before = ' + cast(@lotQtyBefore as varchar(18))" & vbCrLf
sql = sql & "--print 'after  = ' + cast(@lotQtyAfter as varchar(18))" & vbCrLf
sql = sql & "--print cast(@purchaseUnitCost as varchar(12))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if data missing or lot alreadhy converted, do not do it again" & vbCrLf
sql = sql & "if @purchaseUnitCost is null or @lotQtyBefore = @lotQtyAfter or isnull(@lotQtyAfter, 0) = 0" & vbCrLf
sql = sql & "   return" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @unitCost = (@purchaseUnitCost * @lotQtyBefore) / @lotQtyAfter" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--print cast(@height as varchar(12))" & vbCrLf
sql = sql & "--print cast(@length as varchar(12))" & vbCrLf
sql = sql & "--print cast(@purchUnit as varchar(12))" & vbCrLf
sql = sql & "--print cast(@invUnit as varchar(12))" & vbCrLf
sql = sql & "--print cast(@unitCost as varchar(12))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update lot header record" & vbCrLf
sql = sql & "update LohdTable set LOTORIGINALQTY = @lotQtyAfter, LOTREMAININGQTY = @lotQtyAfter, LOTUNITCOST = @unitCost" & vbCrLf
sql = sql & "where LOTNUMBER = @LotNo " & vbCrLf
sql = sql & "--select * from LohdTable where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update first lot item (receipt)" & vbCrLf
sql = sql & "update LoitTable set LOIQUANTITY = @lotQtyAfter, LOIUNITS = @invUnit, LOIHEIGHT = @height, LOILENGTH = @length" & vbCrLf
sql = sql & "where LOINUMBER = @LotNo and LOIRECORD = 1" & vbCrLf
sql = sql & "--select * from LoitTable where LOINUMBER = @LotNo and LOIRECORD = 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update inventory activity for PO receipt" & vbCrLf
sql = sql & "update invatable set INPQTY = @lotQtyAfter, INAQTY = @lotQtyAfter, INAMT = @unitCost, INUNITS = @invUnit" & vbCrLf
sql = sql & "where INLOTNUMBER = @LotNo and INTYPE = 15" & vbCrLf
sql = sql & "--select * from InvaTable where INLOTNUMBER = @LotNo and INTYPE = 15" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update part QOH" & vbCrLf
sql = sql & "select PAQOH from PartTable where PARTREF = @partRef" & vbCrLf
sql = sql & "update PartTable set PAQOH = PAQOH + @lotQtyAfter - @lotQtyBefore where PARTREF = @partRef" & vbCrLf
sql = sql & "--select PAQOH from PartTable where PARTREF = @partRef" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "commit tran" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql

sql = "create table dbo.SSRSInfo" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "   SSRSID int IDENTITY(1,1) NOT NULL," & vbCrLf
sql = sql & "   SSRSFolderUrl varchar(255) NULL" & vbCrLf
sql = sql & ")" & vbCrLf
ExecuteScript False, sql

      
''''''''''''''''''''''''''''''''''''''''''''''''''
        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver
   
   End If
End Function
   

   
'''''''''''''''''''''''''''''''''

Private Function UpdateDatabase80()

   Dim sql As String
   sql = ""
   
   newver = 155
   If ver < newver Then
   
        clsADOCon.ADOErrNum = 0
        
'''''''''''''''''''''''''''''''''''''''''''''''''''
      
        ' separate column for sql causing errors
        sql = "ALTER TABLE SystemEvents ADD Event_SQL VARCHAR(MAX) NULL"
        ExecuteScript False, sql

        sql = "ALTER TABLE SystemEvents ADD CONSTRAINT DF_Event_Date DEFAULT getdate() FOR Event_Date"
        ExecuteScript False, sql
        
        
        'from UpdateDatabase32 with syntax error fixed (set Set @RowCount = 1)
      If StoreProcedureExists("RptIncomeStatement") Then
         sSql = "DROP PROCEDURE RptIncomeStatement"
         ExecuteScript False, sSql
      End If
      
      
      sSql = "CREATE PROCEDURE [dbo].[RptIncomeStatement]" & vbCrLf
      sSql = sSql & "   @StartDate as varchar(12),@EndDate as varchar(12)," & vbCrLf
      sSql = sSql & "   @YearBeginDate as varchar(12), @InclIncAcct as varchar(1)" & vbCrLf
      sSql = sSql & "AS " & vbCrLf
      sSql = sSql & "BEGIN" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   declare @glAcctRef as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @glMsAcct as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @SumCurBal decimal(15,4)" & vbCrLf
      sSql = sSql & "   declare @SumYTD decimal(15,4)" & vbCrLf
      sSql = sSql & "   declare @SumPrevBal as decimal(15,4)" & vbCrLf
      sSql = sSql & "   declare @level as integer" & vbCrLf
      sSql = sSql & "   declare @TopLevelDesc as varchar(30)" & vbCrLf
      sSql = sSql & "   declare @InclInAcct as Integer" & vbCrLf
      sSql = sSql & "   declare @TopLevAcct as varchar(20)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   declare @PrevMaster as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @RowCount as integer" & vbCrLf
      sSql = sSql & "   declare @GlMasterAcc as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @GlChildAcct as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @ChildKey as varchar(1024)" & vbCrLf
      sSql = sSql & "   declare @GLSortKey as varchar(1024)" & vbCrLf
      sSql = sSql & "   declare @SortKey as varchar(1024)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   DELETE FROM EsReportIncStatement" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   if (@InclIncAcct = '1')" & vbCrLf
      sSql = sSql & "      SET @InclInAcct = ''" & vbCrLf
      sSql = sSql & "   Else" & vbCrLf
      sSql = sSql & "      SET @InclInAcct = '0'" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   DECLARE balAcctStruc CURSOR  FOR" & vbCrLf
      sSql = sSql & "      SELECT '4', COINCMACCT, COINCMDESC FROM GlmsTable" & vbCrLf
      sSql = sSql & "      Union All" & vbCrLf
      sSql = sSql & "      SELECT '5', COCOGSACCT, COCOGSDESC FROM GlmsTable" & vbCrLf
      sSql = sSql & "      Union All" & vbCrLf
      sSql = sSql & "      SELECT '6', COEXPNACCT, COEXPNDESC FROM GlmsTable" & vbCrLf
      sSql = sSql & "      Union All" & vbCrLf
      sSql = sSql & "      SELECT '7', COOINCACCT, COOINCDESC FROM GlmsTable" & vbCrLf
      sSql = sSql & "      Union All" & vbCrLf
      sSql = sSql & "      SELECT '8', COOEXPACCT, COOEXPDESC FROM GlmsTable" & vbCrLf
      sSql = sSql & "      Union All" & vbCrLf
      sSql = sSql & "      SELECT '9', COFDTXACCT, COFDTXDESC FROM GlmsTable" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   OPEN balAcctStruc" & vbCrLf
      sSql = sSql & "   FETCH NEXT FROM balAcctStruc INTO @level,@TopLevAcct, @TopLevelDesc" & vbCrLf
      sSql = sSql & "   WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "   BEGIN" & vbCrLf
      sSql = sSql & "      IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "      BEGIN" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "        With cte" & vbCrLf
      sSql = sSql & "        as " & vbCrLf
      sSql = sSql & "        (select GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL, 1 as level," & vbCrLf
      sSql = sSql & "          cast(cast(@level as varchar(4))+ char(36)+ GLACCTREF as varchar(max)) as SortKey" & vbCrLf
      sSql = sSql & "        From GlacTable" & vbCrLf
      sSql = sSql & "        where GLMASTER = cast(@TopLevAcct as varchar(20)) AND GLINACTIVE LIKE @InclInAcct" & vbCrLf
      sSql = sSql & "        Union All" & vbCrLf
      sSql = sSql & "        select a.GLACCTREF, a.GLDESCR, a.GLMASTER, a.GLFSLEVEL, level + 1," & vbCrLf
      sSql = sSql & "         cast(COALESCE(SortKey,'') + char(36) + COALESCE(a.GLACCTREF,'') as varchar(max))as SortKey" & vbCrLf
      sSql = sSql & "        From cte" & vbCrLf
      sSql = sSql & "          inner join GlacTable a" & vbCrLf
      sSql = sSql & "            on cte.GLACCTREF = a.GLMASTER" & vbCrLf
      sSql = sSql & "          WHERE GLINACTIVE LIKE @InclInAcct" & vbCrLf
      sSql = sSql & "        )" & vbCrLf
      sSql = sSql & "        INSERT INTO EsReportIncStatement(GLTOPMaster, GLTOPMASTERDESC, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL,SORTKEYLEVEL,GLACCSORTKEY)" & vbCrLf
      sSql = sSql & "        select @level, @TopLevelDesc, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL, level, SortKey" & vbCrLf
      sSql = sSql & "        from cte order by SortKey" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "      End" & vbCrLf
      sSql = sSql & "      FETCH NEXT FROM balAcctStruc INTO @level,@TopLevAcct, @TopLevelDesc" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & "   Close balAcctStruc" & vbCrLf
      sSql = sSql & "   DEALLOCATE balAcctStruc" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   UPDATE EsReportIncStatement SET CurrentBal = foo.Balance--, SUMCURBAL = foo.Balance" & vbCrLf
      sSql = sSql & "   From" & vbCrLf
      sSql = sSql & "       (SELECT SUM(GjitTable.JICRD) - SUM(GjitTable.JIDEB) as Balance, JIACCOUNT" & vbCrLf
      sSql = sSql & "         FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME" & vbCrLf
      sSql = sSql & "      WHERE GJPOST BETWEEN @StartDate AND @EndDate" & vbCrLf
      sSql = sSql & "         AND GjhdTable.GJPOSTED = 1" & vbCrLf
      sSql = sSql & "      GROUP BY JIACCOUNT) as foo" & vbCrLf
      sSql = sSql & "   Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   UPDATE EsReportIncStatement SET YTD = foo.Balance--, SUMYTD = foo.Balance" & vbCrLf
      sSql = sSql & "   From" & vbCrLf
      sSql = sSql & "      (SELECT JIACCOUNT,SUM(GjitTable.JICRD) - SUM(GjitTable.JIDEB) AS Balance" & vbCrLf
      sSql = sSql & "         FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME" & vbCrLf
      sSql = sSql & "      WHERE (GJPOST BETWEEN @YearBeginDate AND @EndDate)" & vbCrLf
      sSql = sSql & "         AND GjhdTable.GJPOSTED = 1" & vbCrLf
      sSql = sSql & "      GROUP BY JIACCOUNT) as foo" & vbCrLf
      sSql = sSql & "   Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   UPDATE EsReportIncStatement SET PreviousBal = foo.Balance--, SUMPREVBAL = foo.Balance" & vbCrLf
      sSql = sSql & "   From" & vbCrLf
      sSql = sSql & "      (SELECT JIACCOUNT,SUM(GjitTable.JICRD) - SUM(GjitTable.JIDEB) AS Balance" & vbCrLf
      sSql = sSql & "         FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME" & vbCrLf
      sSql = sSql & "      WHERE (GJPOST BETWEEN DATEADD(year, -1, @YearBeginDate) AND DATEADD(year, -1, @EndDate))" & vbCrLf
      sSql = sSql & "         AND GjhdTable.GJPOSTED = 1" & vbCrLf
      sSql = sSql & "      GROUP BY JIACCOUNT) as foo" & vbCrLf
      sSql = sSql & "   Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   SELECT @level =  MAX(SORTKEYLEVEL) FROM EsReportIncStatement" & vbCrLf
      sSql = sSql & "   --set @level = 9" & vbCrLf
      sSql = sSql & "   WHILE (@level >= 1 )" & vbCrLf
      sSql = sSql & "   BEGIN" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "      DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
      sSql = sSql & "      SELECT GLMASTER, SUM(ISNULL(SUMCURBAL,0) + (ISNULL(CurrentBal,0))) ," & vbCrLf
      sSql = sSql & "         Sum (IsNull(SUMYTD, 0) + (IsNull(YTD, 0))), Sum(IsNull(SUMPREVBAL, 0) + (IsNull(PreviousBal, 0)))" & vbCrLf
      sSql = sSql & "      From" & vbCrLf
      sSql = sSql & "         (SELECT DISTINCT GLACCTREF, GLDESCR, GLMASTER," & vbCrLf
      sSql = sSql & "         CurrentBal , YTD, PreviousBal, SUMCURBAL, SUMYTD, SUMPREVBAL" & vbCrLf
      sSql = sSql & "         FROM EsReportIncStatement WHERE SORTKEYLEVEL = @level) as foo" & vbCrLf
      sSql = sSql & "      group by GLMASTER" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "      OPEN curAcctStruc" & vbCrLf
      sSql = sSql & "      FETCH NEXT FROM curAcctStruc INTO @glMsAcct, @SumCurBal, @SumYTD, @SumPrevBal" & vbCrLf
      sSql = sSql & "      WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "      BEGIN" & vbCrLf
      sSql = sSql & "         IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "         BEGIN" & vbCrLf
      sSql = sSql & "            UPDATE EsReportIncStatement SET SUMCURBAL = @SumCurBal, SUMYTD = @SumYTD," & vbCrLf
      sSql = sSql & "               SUMPREVBAL = @SumPrevBal, GLDESCR = 'TOTAL ' + LTRIM(GLDESCR)," & vbCrLf
      sSql = sSql & "            HASCHILD = 1" & vbCrLf
      sSql = sSql & "            WHERE GLACCTREF = @glMsAcct" & vbCrLf
      sSql = sSql & "         End" & vbCrLf
      sSql = sSql & "         FETCH NEXT FROM curAcctStruc INTO @glMsAcct, @SumCurBal, @SumYTD, @SumPrevBal" & vbCrLf
      sSql = sSql & "      End" & vbCrLf
      sSql = sSql & "      Close curAcctStruc" & vbCrLf
      sSql = sSql & "      DEALLOCATE curAcctStruc" & vbCrLf
      sSql = sSql & "      SET @level = @level - 1" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   UPDATE EsReportIncStatement SET SUMCURBAL = CurrentBal WHERE SUMCURBAL IS NULL" & vbCrLf
      sSql = sSql & "   UPDATE EsReportIncStatement SET SUMPREVBAL = PreviousBal WHERE SUMPREVBAL IS NULL" & vbCrLf
      sSql = sSql & "   UPDATE EsReportIncStatement SET SUMYTD = YTD  WHERE SUMYTD IS NULL" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   set @level = 0 " & vbCrLf
      sSql = sSql & "   set @RowCount = 1" & vbCrLf
      sSql = sSql & "   SET @PrevMaster = ''" & vbCrLf
      sSql = sSql & "   DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
      sSql = sSql & "      SELECT GLMASTER, GLACCSORTKEY" & vbCrLf
      sSql = sSql & "      FROM EsReportIncStatement " & vbCrLf
      sSql = sSql & "         WHERE HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLMASTER ='AA10'--GLTOPMaster = 1 AND " & vbCrLf
      sSql = sSql & "      ORDER BY GLACCSORTKEY" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   OPEN curAcctStruc" & vbCrLf
      sSql = sSql & "   FETCH NEXT FROM curAcctStruc INTO @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "   WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "   BEGIN" & vbCrLf
      sSql = sSql & "    IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "    BEGIN" & vbCrLf
      sSql = sSql & "      if (@PrevMaster <> @GlMasterAcc)" & vbCrLf
      sSql = sSql & "      BEGIN" & vbCrLf
      sSql = sSql & "         UPDATE EsReportIncStatement SET " & vbCrLf
      sSql = sSql & "            SortKeyRev = RIGHT ('0000'+ Cast(@RowCount as varchar), 4) + char(36)+ @GlMasterAcc" & vbCrLf
      sSql = sSql & "         WHERE GLMASTER = @GlMasterAcc AND HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLTOPMaster = 1 " & vbCrLf
      sSql = sSql & "         SET @RowCount = @RowCount + 1" & vbCrLf
      sSql = sSql & "         SET @PrevMaster = @GlMasterAcc" & vbCrLf
      sSql = sSql & "      END" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "    End" & vbCrLf
      sSql = sSql & "    FETCH NEXT FROM curAcctStruc INTO @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & "       " & vbCrLf
      sSql = sSql & "   Close curAcctStruc" & vbCrLf
      sSql = sSql & "   DEALLOCATE curAcctStruc" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "    SELECT @level = MAX(SORTKEYLEVEL) FROM EsReportIncStatement" & vbCrLf
      sSql = sSql & "   --set @level = 7" & vbCrLf
      sSql = sSql & "   WHILE (@level >= 1 )" & vbCrLf
      sSql = sSql & "   BEGIN" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      SET @PrevMaster = ''" & vbCrLf
      sSql = sSql & "        DECLARE curAcctStruc1 CURSOR  FOR" & vbCrLf
      sSql = sSql & "         SELECT DISTINCT GLACCTREF, GLMASTER, GLACCSORTKEY" & vbCrLf
      sSql = sSql & "         FROM EsReportIncStatement " & vbCrLf
      sSql = sSql & "            WHERE SORTKEYLEVEL = @level AND HASCHILD IS NOT NULL--GLTOPMaster = 1 AND " & vbCrLf
      sSql = sSql & "         order by GLACCSORTKEY" & vbCrLf
      sSql = sSql & "    " & vbCrLf
      sSql = sSql & "        OPEN curAcctStruc1" & vbCrLf
      sSql = sSql & "        FETCH NEXT FROM curAcctStruc1 INTO @GlChildAcct, @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "        WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "        BEGIN" & vbCrLf
      sSql = sSql & "          IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "          BEGIN" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "            if (@PrevMaster <> @GlChildAcct)" & vbCrLf
      sSql = sSql & "            BEGIN" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   print 'Record' + @GlChildAcct + ':' + @GlMasterAcc" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "               SELECT TOP 1 @ChildKey = SortKeyRev,@SortKey = GLACCSORTKEY" & vbCrLf
      sSql = sSql & "               FROM EsReportIncStatement " & vbCrLf
      sSql = sSql & "                  WHERE SORTKEYLEVEL > @level AND GLMASTER = @GlChildAcct --GLTOPMaster = 1 AND " & vbCrLf
      sSql = sSql & "               order by GLACCSORTKEY desc" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "               UPDATE EsReportIncStatement SET " & vbCrLf
      sSql = sSql & "                  SortKeyRev = Cast(@ChildKey as varchar(512)) + char(36)+ @GlMasterAcc" & vbCrLf
      sSql = sSql & "               WHERE GLACCTREF = @GlChildAcct AND GLMASTER = @GlMasterAcc " & vbCrLf
      sSql = sSql & "                  AND SORTKEYLEVEL = @level --GLTOPMaster = 1 AND " & vbCrLf
      sSql = sSql & "               SET @PrevMaster = @GlChildAcct" & vbCrLf
      sSql = sSql & "            END" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "          End" & vbCrLf
      sSql = sSql & "          FETCH NEXT FROM curAcctStruc1 INTO @GlChildAcct, @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "        End" & vbCrLf
      sSql = sSql & "               " & vbCrLf
      sSql = sSql & "        Close curAcctStruc1" & vbCrLf
      sSql = sSql & "        DEALLOCATE curAcctStruc1" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "        SET @level = @level - 1" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "  SELECT GLTOPMaster, GLTOPMASTERDESC, GLACCTREF, GLACCTNO, GLDESCR, GLMASTER, GLTYPE,GLINACTIVE, GLFSLEVEL," & vbCrLf
      sSql = sSql & "      SUMCURBAL , CurrentBal, SUMYTD, YTD, SUMPREVBAL, PreviousBal, SORTKEYLEVEL, GLACCSORTKEY" & vbCrLf
      sSql = sSql & "   FROM EsReportIncStatement ORDER BY SortKeyRev --GLTOPMASTER, GLACCSORTKEY desc, SortKeyLevel" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "End"
      
      ExecuteScript False, sSql
      
      sSql = "ALTER procedure [dbo].[UpdateTimeCardTotals]" & vbCrLf
      sSql = sSql & " @EmpNo int," & vbCrLf
      sSql = sSql & " @Date datetime" & vbCrLf
      sSql = sSql & "as " & vbCrLf
      sSql = sSql & "/* test" & vbCrLf
      sSql = sSql & "    UpdateTimeCardTotals 52, '9/8/2008'" & vbCrLf
      sSql = sSql & "*/" & vbCrLf
      sSql = sSql & "update TchdTable " & vbCrLf
      sSql = sSql & "set TMSTART = (select top 1 TCSTART From TcitTable " & vbCrLf
      sSql = sSql & "join TchdTable on TCCARD = TMCARD" & vbCrLf
      sSql = sSql & "where TCEMP = @EmpNo and TMDAY = @Date and TCSTOP <> ''" & vbCrLf
      sSql = sSql & "and ISDATE(TCSTART + 'm') = 1 and ISDATE(TCSTOP + 'm') = 1" & vbCrLf
      sSql = sSql & "order by tcstarttime)," & vbCrLf
      sSql = sSql & "TMSTOP = (select top 1 TCSTOP From TcitTable" & vbCrLf
      sSql = sSql & "join TchdTable on TCCARD = TMCARD" & vbCrLf
      sSql = sSql & "where TCEMP = @EmpNo and TMDAY = @Date and TCSTOP <> ''" & vbCrLf
      sSql = sSql & "and ISDATE(TCSTART + 'm') = 1 and ISDATE(TCSTOP + 'm') = 1" & vbCrLf
      sSql = sSql & "order by TCSTOPTIME desc)" & vbCrLf
      sSql = sSql & "where TMEMP = @EmpNo and TMDAY = @Date" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "update TchdTable " & vbCrLf
      sSql = sSql & "set TMREGHRS = (select isnull(sum(hrs), 0.000) from viewTimeCardHours" & vbCrLf
      sSql = sSql & "where EmpNo = @EmpNo" & vbCrLf
      sSql = sSql & "and chgDay = @Date" & vbCrLf
      sSql = sSql & "and type = 'R')," & vbCrLf
      sSql = sSql & "TMOVTHRS = (select isnull(sum(hrs), 0.000) from viewTimeCardHours" & vbCrLf
      sSql = sSql & "where EmpNo = @EmpNo" & vbCrLf
      sSql = sSql & "and chgDay = @Date" & vbCrLf
      sSql = sSql & "and type = 'O')," & vbCrLf
      sSql = sSql & "TMDBLHRS = (select isnull(sum(hrs), 0.000) from viewTimeCardHours" & vbCrLf
      sSql = sSql & "where EmpNo = @EmpNo" & vbCrLf
      sSql = sSql & "and chgDay = @Date" & vbCrLf
      sSql = sSql & "and type = 'D')" & vbCrLf
      sSql = sSql & "where TMEMP = @EmpNo and TMDAY = @Date"
      
      ExecuteScript False, sSql
      
     
''''''''''''''''''''''''''''''''''''''''''''''''''
        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver
   
   End If
End Function
   
Private Function UpdateDatabase81()

   Dim sql As String
   sql = ""

   newver = 156
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

'SQL = "alter table ComnTable add COBACKGROUNDCOLORRGB char(6) null"
'ExecuteScript False, SQL

sql = "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'CompleteAllOps')" & vbCrLf
sql = sql & "DROP PROCEDURE CompleteAllOps" & vbCrLf
ExecuteScript False, sql

sql = "create procedure dbo.CompleteAllOps" & vbCrLf
sql = sql & "   @PartRef varchar(30)," & vbCrLf
sql = sql & "   @RunNo int," & vbCrLf
sql = sql & "   @NeedJournal int OUTPUT" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "close all operations for an MO" & vbCrLf
sql = sql & "NeedJournal  = 0 if successful" & vbCrLf
sql = sql & "            = 1 if TJ Journal needs to be created" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "   declare @NeedJournal int" & vbCrLf
sql = sql & "   exec CompleteAllOps '111A340113', 236, @NeedJournal OUT" & vbCrLf
sql = sql & "   print @NeedJournal" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create a table of all open time charges" & vbCrLf
sql = sql & "Set @NeedJournal = 0" & vbCrLf
sql = sql & "set nocount on" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "select " & vbCrLf
sql = sql & "   cast(null as varchar(11)) as TCCARD," & vbCrLf
sql = sql & "   ISEMPLOYEE as TCEMP," & vbCrLf
sql = sql & "   cast(null as varchar(20)) as TCSTART," & vbCrLf
sql = sql & "   cast(null as varchar(20)) as TCSTOP," & vbCrLf
sql = sql & "   ISMOSTART as TCSTARTTIME," & vbCrLf
sql = sql & "   DATEADD(minute, DATEDIFF(minute, 0, getdate()), 0) as TCSTOPTIME," & vbCrLf
sql = sql & "   cast(null as REAL) AS TCHOURS," & vbCrLf
sql = sql & "   cast(null as smalldatetime) as TCTIME," & vbCrLf
sql = sql & "   'RG' as TCCODE," & vbCrLf
sql = sql & "   cast(null as real) as TCRATE," & vbCrLf
sql = sql & "   cast(null as real) as TCOHRATE," & vbCrLf
sql = sql & "   cast(1 as smallint) as TCRATENO," & vbCrLf
sql = sql & "   cast('' as varchar(12)) as TCACCT," & vbCrLf
sql = sql & "   cast('' as varchar(12)) as TCACCOUNT," & vbCrLf
sql = sql & "   op.OPSHOP as TCSHOP," & vbCrLf
sql = sql & "   op.OPCENTER as TCWC," & vbCrLf
sql = sql & "   cast (0 as tinyint) as TCPAYTYPE," & vbCrLf
sql = sql & "   isnull(ISSURUN,'R') as TCSURUN," & vbCrLf
sql = sql & "   cast(0 as real) as TCYIELD," & vbCrLf
sql = sql & "   OPREF AS TCPARTREF," & vbCrLf
sql = sql & "   OPRUN AS TCRUNNO," & vbCrLf
sql = sql & "    OPNO AS TCOPNO," & vbCrLf
sql = sql & "   0 TCSORT," & vbCrLf
sql = sql & "   cast(0 as real) as TCOHFIXED," & vbCrLf
sql = sql & "   dbo.fnGetOpenJournalID('TJ', ISMOSTART) as TCGLJOURNAL," & vbCrLf
sql = sql & "   --cast(null as varchar(12)) as TCGLJOURNAL," & vbCrLf
sql = sql & "   0 as TCGLREF," & vbCrLf
sql = sql & "   'MOCLS' AS TCSOURCE," & vbCrLf
sql = sql & "   0 as TCMULTIJOB," & vbCrLf
sql = sql & "   0 as TCACCEPT," & vbCrLf
sql = sql & "   0 as TCREJECT," & vbCrLf
sql = sql & "   0 as TCSCRAP," & vbCrLf
sql = sql & "   'Closure forced from Close All Ops' as TCCOMMENTS," & vbCrLf
sql = sql & "   cast(null as varchar(6)) as _ShiftRef," & vbCrLf
sql = sql & "   cast(null as datetime) as _ShiftStartDate" & vbCrLf  ' 7/19/17 - was cast as Date.  failed in sql 2005
sql = sql & "into #temp" & vbCrLf
sql = sql & "from IstcTable istc" & vbCrLf
sql = sql & "join RnopTable op on op.OPREF = istc.ISMO and op.OPRUN = istc.ISRUN and op.OPNO = istc.ISOP" & vbCrLf
sql = sql & "where ISMO = @PartRef and ISRUN = @RunNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- return if no TJ journal found for these start dates" & vbCrLf
sql = sql & "if exists (select * from #temp where isnull(TCGLJOURNAL, '') = '')" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "   rollback tran" & vbCrLf
sql = sql & "   set @NeedJournal = 1" & vbCrLf
sql = sql & "   return" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- do not allow charge over 16 hours" & vbCrLf
sql = sql & "update #temp set TCSTOPTIME = dateadd(MINUTE,16*60,TCSTARTTIME) where datediff(MINUTE,TCSTARTTIME,TCSTOPTIME) > 16 * 60" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- construct hh:mma/p version of times" & vbCrLf
sql = sql & "update #temp set TCSTART = REPLACE(SUBSTRING(CONVERT(VARCHAR(20),TCSTARTTIME,0),13,6),' ','0')" & vbCrLf
sql = sql & "update #temp set TCSTOP = REPLACE(SUBSTRING(CONVERT(VARCHAR(20),TCSTOPTIME,0),13,6),' ','0')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- calculate hours" & vbCrLf
sql = sql & "update #temp set tchours = datediff(MINUTE,TCSTARTTIME, TCSTOPTIME) / 60." & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- calculate # minutes" & vbCrLf
sql = sql & "update #temp set TCTIME = dateadd(minute,datediff(MINUTE,TCSTARTTIME, TCSTOPTIME),'1/1/1900')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- determine shift code and shift start date" & vbCrLf
sql = sql & "select a.SFREF, A.PREMNUMBER, cast(SFSTHR + 'm' as time) as _START, cast(SFENHR + 'm' as time) as _END, cast(null as int) as _MINUTES" & vbCrLf
sql = sql & "into #shifts" & vbCrLf
sql = sql & "FROM dbo.sfempTable a" & vbCrLf
sql = sql & "join sfcdTable b on a.SFREF = b.SFREF" & vbCrLf
sql = sql & "where premnumber in (select tcemp from #temp)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #shifts set _MINUTES = DATEDIFF(MINUTE,_START,_END)" & vbCrLf
sql = sql & "update #shifts set _MINUTES = _MINUTES + 1440 where _MINUTES < 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- case 1: charge starts up to 1 hour before shift start to end of shift, same day" & vbCrLf
sql = sql & "update t set _ShiftRef = s.SFREF," & vbCrLf
sql = sql & "   _ShiftStartDate = cast(TCSTARTTIME as date)" & vbCrLf
sql = sql & "FROM #shifts s join #temp t on t.TCEMP = s.PREMNUMBER" & vbCrLf
sql = sql & "where (DATEDIFF(minute, _START, cast(t.TCSTART + 'm' as time)) BETWEEN -60 AND _MINUTES)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- case 2: time charge starts after midnight -- use prior day as shift start day" & vbCrLf
sql = sql & "update t set _ShiftRef = s.SFREF," & vbCrLf
sql = sql & "   _ShiftStartDate = dateadd(day,-1,cast(TCSTARTTIME as date))" & vbCrLf
sql = sql & "FROM #shifts s join #temp t on t.TCEMP = s.PREMNUMBER" & vbCrLf
sql = sql & "where (DATEDIFF(minute, _START, cast(t.TCSTART + 'm' as time))  + 1440 BETWEEN -30 AND _MINUTES)" & vbCrLf
sql = sql & "and _ShiftStartDate is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- case 3: outside of shift hours.  Use shift code (if one assigned) and time charge start date as shift start date" & vbCrLf
sql = sql & "update t set _ShiftRef = s.SFREF,      -- null if no shift defined for employee" & vbCrLf
sql = sql & "   _ShiftStartDate = cast(TCSTARTTIME as date)" & vbCrLf
sql = sql & "FROM #temp t  " & vbCrLf
sql = sql & "left join #shifts s on t.TCEMP = s.PREMNUMBER" & vbCrLf
sql = sql & "where _ShiftStartDate is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get time card references where they exist" & vbCrLf
sql = sql & "update t" & vbCrLf
sql = sql & "set TCCARD = TMCARD" & vbCrLf
sql = sql & "from #temp t join TchdTable h on h.TMEMP = t.TCEMP AND h.TMDAY = t._ShiftStartDate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- for charges where there is no current time card, create them" & vbCrLf
sql = sql & "-- first create base parameters to construct timecard ID" & vbCrLf
sql = sql & "declare @now as datetime, @nowDays int, @nowMs int" & vbCrLf
sql = sql & "set @now = getdate()" & vbCrLf
sql = sql & "set @nowDays = DATEDIFF(DAY,'1/1/1900',cast(@now as date))" & vbCrLf
sql = sql & "set @nowMs = 1000000.0 *cast(DATEDIFF(MILLISECOND,'1/1/1900',cast(@now as time)) as float)/(3600.0*24*1000)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO TchdTable (TMCARD,TMEMP,TMDAY) " & vbCrLf
sql = sql & "select cast(@nowDays as varchar(5)) + cast(@nowMs + ROW_NUMBER() over( order by TCEMP, TCSTARTTIME) - 1 as varchar(6)), " & vbCrLf
sql = sql & "   TCEMP, _ShiftStartDate" & vbCrLf
sql = sql & "FROM #temp where TCCARD is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- add new timecards where needed" & vbCrLf
sql = sql & "update t" & vbCrLf
sql = sql & "set t.TCCARD = TMCARD" & vbCrLf
sql = sql & "from #temp t join TchdTable h on h.TMEMP = t.TCEMP AND h.TMDAY = t._ShiftStartDate" & vbCrLf
sql = sql & "where t.TCCARD is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- find a regular timecode" & vbCrLf
sql = sql & "update #temp set TCCODE = isnull((select top 1 TYPECODE from TmcdTable where TYPETYPE = 'R' ORDER BY TYPESEQ),'RT')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get rates and accounts" & vbCrLf
sql = sql & "update t " & vbCrLf
sql = sql & "set TCACCT = RTRIM(COALESCE(NULLIF(w.WCNACCT,''), NULLIF(s.SHPACCT,''), c.WIPLABACCT))," & vbCrLf
sql = sql & "   TCRATE = case when e.PREMPAYRATE <> 0 then e.PREMPAYRATE" & vbCrLf
sql = sql & "       when w.WCNSTDRATE <> 0 then w.WCNSTDRATE" & vbCrLf
sql = sql & "       else s.SHPRATE end," & vbCrLf
sql = sql & "   TCOHRATE = case when w.WCNOHPCT <> 0 then w.WCNOHPCT " & vbCrLf
sql = sql & "       else s.SHPOHRATE end," & vbCrLf
sql = sql & "   TCOHFIXED = case when w.WCNOHFIXED <> 0 then w.WCNOHFIXED" & vbCrLf
sql = sql & "       else s.SHPOHTOTAL end" & vbCrLf
sql = sql & "       from RnopTable r" & vbCrLf
sql = sql & "join PartTable p on r.OPREF = p.PARTREF" & vbCrLf
sql = sql & "join WcntTable w on r.OPCENTER = w.WCNREF" & vbCrLf
sql = sql & "join ShopTable s on r.OPSHOP = s.SHPREF" & vbCrLf
sql = sql & "join #temp t on t.TCPARTREF = r.OPREF and t.TCRUNNO = r.OPRUN and t.TCOPNO = r.OPNO" & vbCrLf
sql = sql & "join ComnTable c on 1 = 1" & vbCrLf
sql = sql & "join EmplTable e on E.PREMNUMBER = T.TCEMP" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #temp " & vbCrLf
sql = sql & "set TCOHRATE = TCRATE * TCOHRATE / 100" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create time charge in TcitTable" & vbCrLf
sql = sql & "INSERT INTO TcitTable (TCCARD,TCEMP,TCSTART,TCSTOP,TCSTARTTIME,TCSTOPTIME," & vbCrLf
sql = sql & "      TCHOURS,TCTIME,TCCODE,TCRATE,TCOHRATE,TCRATENO,TCACCT,TCACCOUNT," & vbCrLf
sql = sql & "      TCSHOP,TCWC,TCPAYTYPE,TCSURUN,TCYIELD,TCPARTREF,TCRUNNO," & vbCrLf
sql = sql & "      TCOPNO,TCSORT,TCOHFIXED,TCGLJOURNAL,TCGLREF,TCSOURCE," & vbCrLf
sql = sql & "      TCMULTIJOB,TCACCEPT,TCREJECT,TCSCRAP,TCCOMMENTS)" & vbCrLf
sql = sql & "select TCCARD,TCEMP,TCSTART,TCSTOP,TCSTARTTIME,TCSTOPTIME," & vbCrLf
sql = sql & "      TCHOURS,TCTIME,TCCODE,TCRATE,TCOHRATE,TCRATENO,TCACCT,TCACCOUNT," & vbCrLf
sql = sql & "      TCSHOP,TCWC,TCPAYTYPE,TCSURUN,TCYIELD,TCPARTREF,TCRUNNO," & vbCrLf
sql = sql & "      TCOPNO,TCSORT,TCOHFIXED,TCGLJOURNAL,TCGLREF,TCSOURCE," & vbCrLf
sql = sql & "      TCMULTIJOB,TCACCEPT,TCREJECT,TCSCRAP,TCCOMMENTS" & vbCrLf
sql = sql & "from #temp" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- set min and max times for charges in affected in timecards" & vbCrLf
sql = sql & "update TchdTable set TMSTART = (select top 1 TCSTART From TcitTable " & vbCrLf
sql = sql & "where TCCARD = TMCARD and TCSTOP <> ''" & vbCrLf
sql = sql & "and ISDATE(TCSTART + 'm') = 1 and ISDATE(TCSTOP + 'm') = 1" & vbCrLf
sql = sql & "order by tcstarttime)," & vbCrLf
sql = sql & "TMSTOP = (select top 1 TCSTOP From TcitTable" & vbCrLf
sql = sql & "where TCCARD = TMCARD and TCSTOP <> ''" & vbCrLf
sql = sql & "and ISDATE(TCSTART + 'm') = 1 and ISDATE(TCSTOP + 'm') = 1" & vbCrLf
sql = sql & "order by TCSTOPTIME desc)" & vbCrLf
sql = sql & "where TMCARD in (select distinct TCCARD FROM #temp)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- total regular, overhead and doubletime hours" & vbCrLf
sql = sql & "select TCCARD, TYPETYPE, ISNULL(SUM(TCHOURS),0.0) as HRS" & vbCrLf
sql = sql & "into #totals" & vbCrLf
sql = sql & "from TcitTable ti" & vbCrLf
sql = sql & "JOIN TmcdTable tc ON ti.TCCODE = tc.TYPECODE" & vbCrLf
sql = sql & "where TCCARD in (select TCCARD FROM #temp)" & vbCrLf
sql = sql & "group by TCCARD, TYPETYPE" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update tc set TMREGHRS = isnull(HRS,0.0)" & vbCrLf
sql = sql & "from tchdtable tc" & vbCrLf
sql = sql & "join #totals tot on tot.TCCARD = tc.TMCARD and tot.TYPETYPE = 'R'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update tc set TMOVTHRS = isnull(HRS,0.0)" & vbCrLf
sql = sql & "from tchdtable tc" & vbCrLf
sql = sql & "join #totals tot on tot.TCCARD = tc.TMCARD and tot.TYPETYPE = 'O'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update tc set TMDBLHRS = isnull(HRS,0.0)" & vbCrLf
sql = sql & "from tchdtable tc" & vbCrLf
sql = sql & "join #totals tot on tot.TCCARD = tc.TMCARD and tot.TYPETYPE = 'D'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from TcitTable tc" & vbCrLf
sql = sql & "--join #temp tmp on tc.TCPARTREF = tmp.TCPARTREF and tc.TCRUNNO = tmp.TCRUNNO and tc.TCOPNO = tmp.TCOPNO " & vbCrLf
sql = sql & "-- and tc.TCEMP = tmp.TCEMP and tc.TCSTARTTIME = tmp.TCSTARTTIME" & vbCrLf
sql = sql & "--select * from TchdTable where TMCARD in (select tccard from #temp)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "delete istc" & vbCrLf
sql = sql & "from IstcTable istc join #temp t on t.TCPARTREF = istc.ISMO and t.TCRUNNO = istc.ISRUN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
sql = sql & "drop table #shifts" & vbCrLf
sql = sql & "drop table #totals" & vbCrLf
sql = sql & "commit tran" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "return" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''

Private Function UpdateDatabase82()

'update database version template
'set version at top of this file
'add call to this function
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 157
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'RptWCQListStats')" & vbCrLf
sql = sql & "DROP PROCEDURE RptWCQListStats" & vbCrLf
ExecuteScript False, sql

sql = "create procedure dbo.RptWCQListStats" & vbCrLf
sql = sql & " @BeginDate  as varchar(30),@EdnDate as varchar(10)," & vbCrLf
sql = sql & " @OpShop as varchar(12), @OpCenter as varchar(12)  " & vbCrLf
sql = sql & " AS" & vbCrLf
sql = sql & " BEGIN " & vbCrLf
sql = sql & "   IF (@OpShop = '')  " & vbCrLf
sql = sql & "   BEGIN   " & vbCrLf
sql = sql & "       SET @OpShop = '%'    " & vbCrLf
sql = sql & "   End " & vbCrLf
sql = sql & "                " & vbCrLf
sql = sql & "   IF (@OpCenter = '')    " & vbCrLf
sql = sql & "   BEGIN   " & vbCrLf
sql = sql & "       SET @OpCenter = '%'        " & vbCrLf
sql = sql & "   End                       " & vbCrLf
sql = sql & "    " & vbCrLf
sql = sql & "   select distinct runstable.Runref, runstable.runno, OPCENTER, RUNSTATUS, RUNQTY," & vbCrLf
sql = sql & "       runopcur,RnopTable.OPQDATE,  RnopTable.OPSCHEDDATE,RunsTable.RUNPRIORITY," & vbCrLf
sql = sql & "       f.PrevOPNO, f.OPQDATE PREVQDATE, f.OPCOMPDATE PREVCOMPDATE," & vbCrLf
sql = sql & "       DATEDIFF(day,f.OPCOMPDATE, RnopTable.OPQDATE) DtDiffQue," & vbCrLf
sql = sql & "       DATEDIFF(day, f.OPCOMPDATE, GETDATE()) DtDiffNow" & vbCrLf
sql = sql & "      from runstable,RnopTable," & vbCrLf
sql = sql & "      (select a.runref, a.runno, b.OPNO PrevOPNO, OPQDATE, OPCOMPDATE," & vbCrLf
sql = sql & "            ROW_NUMBER() OVER (PARTITION BY opref, oprun" & vbCrLf
sql = sql & "                          ORDER BY opref DESC, oprun DESC, OPNO desc) as rn" & vbCrLf
sql = sql & "         from runstable a,rnopTable b" & vbCrLf
sql = sql & "         where a.runref = b.opref and" & vbCrLf
sql = sql & "            a.RunNO = b.oprun And b.opno < a.runopcur" & vbCrLf
sql = sql & "      ) as f" & vbCrLf
sql = sql & "   where RnopTable.OPSCHEDDATE between @BeginDate and @EdnDate" & vbCrLf
sql = sql & "      and RunsTable.runref =  f.runref AND RunsTable.runno = f.runno" & vbCrLf
sql = sql & "      and RnopTable.opref = RunsTable.runref AND RunsTable.runno = RnopTable.OPRUN" & vbCrLf
sql = sql & "      and RunsTable.runopcur = RnopTable.Opno" & vbCrLf
sql = sql & "      and RnopTable.OPSHOP LIKE @OpShop AND RnopTable.OPCENTER LIKE @OpCenter" & vbCrLf
sql = sql & "       AND RnopTable.OPCOMPLETE = 0" & vbCrLf
sql = sql & "      and f.rn = 1" & vbCrLf
sql = sql & "   ORDER BY RunsTable.RUNPRIORITY, RnopTable.OPSCHEDDATE" & vbCrLf
sql = sql & "    " & vbCrLf
sql = sql & " End" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql

'''''''''''''''''''''''''''''''''''

sql = "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'RptHistoryWCQListStats')" & vbCrLf
sql = sql & "DROP PROCEDURE RptHistoryWCQListStats" & vbCrLf
ExecuteScript False, sql

sql = "CREATE PROCEDURE dbo.RptHistoryWCQListStats" & vbCrLf
sql = sql & " @BeginDate  as varchar(30),@EdnDate as varchar(10)," & vbCrLf
sql = sql & " @OpShop as varchar(12), @OpCenter as varchar(12)  " & vbCrLf
sql = sql & " AS" & vbCrLf
sql = sql & " BEGIN " & vbCrLf
sql = sql & "   IF (@OpShop = '')  " & vbCrLf
sql = sql & "   BEGIN   " & vbCrLf
sql = sql & "      SET @OpShop = '%'    " & vbCrLf
sql = sql & "   End;" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   WITH CTE as " & vbCrLf
sql = sql & "   (select b.opref, b.oprun, b.OPNO , OPQDATE, " & vbCrLf
sql = sql & "       OPCOMPDATE, OPCOMPLETE," & vbCrLf
sql = sql & "       OPSHOP,OPCENTER,OPSCHEDDATE, " & vbCrLf
sql = sql & "       ROW_NUMBER() OVER (PARTITION BY opref, oprun" & vbCrLf
sql = sql & "                   ORDER BY opref , oprun , OPNO) as rn" & vbCrLf
sql = sql & "   from rnopTable b" & vbCrLf
sql = sql & "   where b.opcomplete = 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   )" & vbCrLf
sql = sql & "   SELECT DISTINCT CTE.opref AS Runref, CTE.oprun AS runno, CTE.OPCENTER, " & vbCrLf
sql = sql & "       RunsTable.RUNSTATUS, RunsTable.RUNQTY,RunsTable.runopcur,RunsTable.RUNPRIORITY," & vbCrLf
sql = sql & "       CTE.OPQDATE, CTE.OPSCHEDDATE, CTE.OPNO, CTE.OPCOMPDATE, " & vbCrLf
sql = sql & "       prev.OPCOMPDATE as PrevCOMDATE, " & vbCrLf
sql = sql & "       DATEDIFF(day,prev.OPCOMPDATE, CTE.OPCOMPDATE) DtDiffQue," & vbCrLf
sql = sql & "       DATEDIFF(day, prev.OPCOMPDATE, GETDATE()) DtDiffNow" & vbCrLf
sql = sql & "   FROM CTE" & vbCrLf
sql = sql & "   LEFT JOIN CTE prev ON  CTE.OPREF = prev.OPREF" & vbCrLf
sql = sql & "       AND CTE.OPRUN = prev.OPRUN" & vbCrLf
sql = sql & "       AND prev.rn = CTE.rn - 1" & vbCrLf
sql = sql & "   JOIN RunsTable ON RunsTable.runref =  CTE.opref " & vbCrLf
sql = sql & "       AND RunsTable.runno = CTE.oprun" & vbCrLf
sql = sql & "   WHERE CTE.OPSHOP LIKE @OpShop AND CTE.OPCENTER LIKE (@OpCenter + '%')" & vbCrLf
sql = sql & "   AND CTE.OPCOMPDATE >= @BeginDate " & vbCrLf
sql = sql & "   AND CTE.OPCOMPDATE < DATEADD(d, 1, @EdnDate)" & vbCrLf
sql = sql & "   ORDER BY RunsTable.RUNPRIORITY,CTE.OPSCHEDDATE" & vbCrLf
sql = sql & " End" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function UpdateDatabase83()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 158     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

'not all customers got this column addition in 2011
If Not ColumnExists("EsReportVendorStmt", "CHKACCT") Then ExecuteScript False, "ALTER TABLE EsReportVendorStmt ADD CHKACCT Char(12)"

'some customers did not get this update, which added a second parameter
If StoreProcedureExists("UpdatePackingSlipCosts") Then
    sSql = "DROP PROCEDURE UpdatePackingSlipCosts"
    ExecuteScript False, sSql
End If
      
sql = "create procedure dbo.UpdatePackingSlipCosts" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & " @PackingSlip varchar(8)," & vbCrLf
sql = sql & " @UpdateIaEvenIfJournalClosed bit    -- = 1 to update items for closed journals" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "update InvaTable" & vbCrLf
sql = sql & "set INAMT = LOTUNITCOST," & vbCrLf
sql = sql & "INTOTMATL = cast ( abs( INAQTY ) * LOTTOTMATL / LOTORIGINALQTY as decimal(12,4))," & vbCrLf
sql = sql & "INTOTLABOR = cast ( abs( INAQTY ) * LOTTOTLABOR / LOTORIGINALQTY as decimal(12,4))," & vbCrLf
sql = sql & "INTOTEXP = cast ( abs( INAQTY ) * LOTTOTEXP / LOTORIGINALQTY as decimal(12,4)), " & vbCrLf
sql = sql & "INTOTOH = cast ( abs( INAQTY ) * LOTTOTOH / LOTORIGINALQTY as decimal(12,4))," & vbCrLf
sql = sql & "INTOTHRS = cast ( abs( INAQTY ) * LOTTOTHRS / LOTORIGINALQTY as decimal(12,4))" & vbCrLf
sql = sql & "from LoitTable " & vbCrLf
sql = sql & "join LohdTable ON LOINUMBER = LOTNUMBER" & vbCrLf
sql = sql & "join InvaTable ia2 ON INNUMBER = LOIACTIVITY" & vbCrLf
sql = sql & "where ia2.INPSNUMBER = @PackingSlip" & vbCrLf
sql = sql & "and LOTORIGINALQTY <> 0" & vbCrLf
sql = sql & "and (@UpdateIaEvenIfJournalClosed = 1 or ia2.INGLPOSTED = 0)" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''''

' this should have happened in UpdateDatabase27, but that did not happen for all users.
If StoreProcedureExists("RptChartOfAccount") Then
   sSql = "DROP PROCEDURE RptChartOfAccount"
   ExecuteScript False, sSql
End If

sql = "CREATE  PROCEDURE [dbo].[RptChartOfAccount]  " & vbCrLf
sql = sql & "    @InclIncAcct as varchar(1)" & vbCrLf
sql = sql & "AS " & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "  exec RptChartOfAccount 0" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN " & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   declare @glAcctRef as varchar(10) " & vbCrLf
sql = sql & "   declare @glMsAcct as varchar(10) " & vbCrLf
sql = sql & "   declare @TopLevelDesc as varchar(30)" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   declare @level as varchar(12)" & vbCrLf
sql = sql & "   declare @InclInAcct as Integer" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   if (@InclIncAcct = '1')" & vbCrLf
sql = sql & "      SET @InclInAcct = ''" & vbCrLf
sql = sql & "   else" & vbCrLf
sql = sql & "      SET @InclInAcct = '0'" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   CREATE TABLE #tempChartOfAcct(   " & vbCrLf
sql = sql & "   [TOPLEVEL] [varchar](12) NULL,  " & vbCrLf
sql = sql & "   [TOPLEVELDESC] [varchar](30) NULL, " & vbCrLf
sql = sql & "   [GLACCTREF] [varchar](112) NULL,         " & vbCrLf
sql = sql & "   [GLDESCR] [varchar](120) NULL,  " & vbCrLf
sql = sql & "   [GLMASTER] [varchar](12) NULL,   " & vbCrLf
sql = sql & "   [GLFSLEVEL] [INT] NULL," & vbCrLf
sql = sql & "   [GLINACTIVE] [int] NULL," & vbCrLf
sql = sql & "   [SORTKEYLEVEL] [int] NULL,           " & vbCrLf
sql = sql & "   [GLACCSORTKEY] [varchar](512) NULL           " & vbCrLf
sql = sql & ")                             " & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   DECLARE balAcctStruc CURSOR  FOR " & vbCrLf
sql = sql & "      SELECT COASSTACCT, COASSTDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COLIABACCT, COLIABDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COINCMACCT, COINCMDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL " & vbCrLf
sql = sql & "      SELECT COEQTYACCT, COEQTYDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COCOGSACCT, COCOGSDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COEXPNACCT, COEXPNDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COOINCACCT, COOINCDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COFDTXACCT, COFDTXDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "     SELECT COOEXPACCT, COOEXPDESC FROM GlmsTable" & vbCrLf
sql = sql & "      --UNION ALL" & vbCrLf
sql = sql & "      --SELECT COFDTXACCT, COFDTXDESC FROM GlmsTable" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   OPEN balAcctStruc" & vbCrLf
sql = sql & "   FETCH NEXT FROM balAcctStruc INTO @level, @TopLevelDesc" & vbCrLf
sql = sql & "   WHILE (@@FETCH_STATUS <> -1) " & vbCrLf
sql = sql & "   BEGIN " & vbCrLf
sql = sql & "      IF (@@FETCH_STATUS <> -2) " & vbCrLf
sql = sql & "      BEGIN " & vbCrLf
sql = sql & "         " & vbCrLf
sql = sql & "         INSERT INTO #tempChartOfAcct(TOPLEVEL, TOPLEVELDESC, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL,GLINACTIVE, SORTKEYLEVEL,GLACCSORTKEY)" & vbCrLf
sql = sql & "         select @level as TopLevel, @TopLevelDesc as TopLevelDesc, @level as GLACCTREF, " & vbCrLf
sql = sql & "            @TopLevelDesc as GLDESCR, '' as GLMASTER, 0 as GLFSLEVEL, 0,0 as level, @level as SortKey;" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "         with cte" & vbCrLf
sql = sql & "         as" & vbCrLf
sql = sql & "         (select GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL, GLINACTIVE, 0 as level," & vbCrLf
sql = sql & "            cast(cast(@level as varchar(12))+ char(36)+ GLACCTREF as varchar(max)) as SortKey" & vbCrLf
sql = sql & "         from GlacTable" & vbCrLf
sql = sql & "         where GLMASTER = cast(@level as varchar(12)) AND GLINACTIVE LIKE @InclInAcct" & vbCrLf
sql = sql & "         union all" & vbCrLf
sql = sql & "         select a.GLACCTREF, a.GLDESCR, a.GLMASTER, a.GLFSLEVEL, a.GLINACTIVE, level + 1," & vbCrLf
sql = sql & "          cast(COALESCE(SortKey,'') + char(36) + COALESCE(a.GLACCTREF,'') as varchar(max))as SortKey" & vbCrLf
sql = sql & "         from cte" & vbCrLf
sql = sql & "            inner join GlacTable a" & vbCrLf
sql = sql & "               on cte.GLACCTREF = a.GLMASTER" & vbCrLf
sql = sql & "            WHERE a.GLINACTIVE LIKE @InclInAcct" & vbCrLf
sql = sql & "         )" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "         INSERT INTO #tempChartOfAcct(TOPLEVEL, TOPLEVELDESC, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL,GLINACTIVE, SORTKEYLEVEL,GLACCSORTKEY)" & vbCrLf
sql = sql & "         select @level as TopLevel, @TopLevelDesc as TopLevelDesc, " & vbCrLf
sql = sql & "               Replicate('  ', level) + GLACCTREF as GLACCTREF, " & vbCrLf
sql = sql & "               Replicate('  ', level) + GLDESCR as GLDESCR, GLMASTER, " & vbCrLf
sql = sql & "               GLFSLEVEL, GLINACTIVE,level, SortKey" & vbCrLf
sql = sql & "         from cte order by SortKey" & vbCrLf
sql = sql & "         " & vbCrLf
sql = sql & "      END" & vbCrLf
sql = sql & "      FETCH NEXT FROM balAcctStruc INTO @level, @TopLevelDesc" & vbCrLf
sql = sql & "   END         " & vbCrLf
sql = sql & "   CLOSE balAcctStruc" & vbCrLf
sql = sql & "   DEALLOCATE balAcctStruc" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   SELECT TOPLEVEL, TOPLEVELDESC, GLACCTREF, GLDESCR, GLMASTER, " & vbCrLf
sql = sql & "      GLFSLEVEL,GLINACTIVE, SORTKEYLEVEL,GLACCSORTKEY " & vbCrLf
sql = sql & "   FROM #tempChartOfAcct ORDER BY GLACCSORTKEY" & vbCrLf
sql = sql & "                                           " & vbCrLf
sql = sql & "   DROP table #tempChartOfAcct            " & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql

''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function

Private Function UpdateDatabase84()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 159     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

      If (Not TableExists("TlitTableNew")) Then
      
         sSql = "CREATE TABLE [dbo].[TlitTableNew](" & vbCrLf
         sSql = sSql & " [TOOL_NUM] [char](30) NOT NULL CONSTRAINT [DF_TlitTableNew_TOOL_NUM]  DEFAULT ('')," & vbCrLf
         sSql = sSql & " [TOOL_PARTREF] [char](30) NOT NULL CONSTRAINT [DF_TlitTableNew_TOOL_PARTREF]  DEFAULT ('')," & vbCrLf
         sSql = sSql & " [TOOL_CLASS] [char](12) NULL CONSTRAINT [DF_TlitTableNew_TOOL_CLASS]  DEFAULT ('')" & vbCrLf
         sSql = sSql & ") ON [PRIMARY]" & vbCrLf
         
         ExecuteScript False, sSql
      End If
      
      'this code did not execute in previous update
      If (Not TableExists("TlnhdTableNew")) Then
         sSql = "CREATE TABLE [dbo].[TlnhdTableNew](" & vbCrLf
         sSql = sSql & "   [TOOL_NUM] [char](30) NOT NULL," & vbCrLf
         sSql = sSql & "   [TOOL_DTADDED] [char](12) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_CGPONUM] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_CUSTPONUM] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_CGSOPONUM] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_GOVOWNED] [char](10) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_CLASS] [char](12) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_HOMEBLDG] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_HOMEAISLE] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_SHELFNUM] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_GRID] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_LOCNUM] [char](30) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_COMMENTS] [varchar](1020) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_OWNER] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_ACCTTO] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_SN] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_MAKEPN] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_CAVNUM] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_DIM] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_BLANKPONUM] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_MONUM] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_TOOLMATSTAT] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_SRVSTAT] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_DISPSTAT] [char](20) NULL," & vbCrLf
         sSql = sSql & "   [TOOL_STORAGESTAT] [char](20) NULL" & vbCrLf
         sSql = sSql & " CONSTRAINT [PK_TlnhdTableNew_TOOL_NUM] PRIMARY KEY CLUSTERED" & vbCrLf
         sSql = sSql & "(" & vbCrLf
         sSql = sSql & "   [TOOL_NUM] Asc" & vbCrLf
         sSql = sSql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 80) ON [PRIMARY]" & vbCrLf
         sSql = sSql & ") ON [PRIMARY]"

         ExecuteScript False, sSql
      End If
      
'      sSql = "ALTER TABLE TlnhdTableNew Add TOOL_ITAR int NULL CONSTRAINT DF__TlnhdTabl__TOOL_ITAR default(0)"
'      ExecuteScript False, sSql
      
If Not ColumnExists("TlnhdTableNew", "TOOL_ITAR") Then
   ExecuteScript False, "ALTER TABLE TlnhdTableNew Add TOOL_ITAR int NULL CONSTRAINT DF__TlnhdTabl__TOOL_ITAR default(0)"
End If
      


'      sSql = "ALTER TABLE TlnhdTableNew DROP COLUMN TOOL_MONUM"
'      ExecuteScript False, sSql
'
      sSql = "ALTER TABLE TlnhdTableNew Add TOOL_WEIGHT VARCHAR(20) NULL CONSTRAINT DF__TlnhdTabl__TOOL_WEIGHT default('')"
      ExecuteScript False, sSql

sql = "IF object_id('fnCompress') IS NOT NULL" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "    DROP FUNCTION [dbo].fnCompress" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript False, sql

sql = "create function dbo.fnCompress( @in varchar(30))" & vbCrLf
sql = sql & "returns varchar(30)" & vbCrLf
sql = sql & "-- remove tabs, crs, lfs, blanks, commas, single-quotes and dashes to create a key field" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "  return REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(@in,CHAR(9),''),CHAR(10),''),CHAR(13),''),' ',''),'''',''),'-','')" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql

sql = "ALTER TABLE [dbo].[TlnhdTableNew] DROP CONSTRAINT [PK_TlnhdTableNew_TOOL_NUM]" & vbCrLf
ExecuteScript False, sql

sql = "ALTER TABLE [TlnhdTableNew] ADD TOOL_NUMREF varchar(30) NULL" & vbCrLf
ExecuteScript False, sql

sql = "UPDATE TlnhdTableNew SET TOOL_NUMREF = dbo.fnCompress(TOOL_NUM)," & vbCrLf
sql = sql & "  TOOL_WEIGHT = ''" & vbCrLf
ExecuteScript False, sql

sql = "ALTER TABLE [TlnhdTableNew] ALTER COLUMN TOOL_weight varchar(20) NOT NULL" & vbCrLf
ExecuteScript False, sql

sql = "ALTER TABLE [TlnhdTableNew] ALTER COLUMN TOOL_NUMREF varchar(30) NOT NULL" & vbCrLf
ExecuteScript False, sql

sql = "ALTER TABLE [dbo].[TlnhdTableNew] ADD  CONSTRAINT [PK_TlnhdTableNew_TOOL_NUMREF] PRIMARY KEY CLUSTERED " & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "  [TOOL_NUMREF] ASC" & vbCrLf
sql = sql & ")" & vbCrLf
ExecuteScript False, sql

sql = "EXEC sp_rename 'TlitTableNew.TOOL_NUM', 'TOOL_NUMREF'" & vbCrLf
ExecuteScript False, sql




''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function

Private Function UpdateDatabase85()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 160     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

      If StoreProcedureExists("BackLogByPartNumber") Then
         sSql = "DROP PROCEDURE BackLogByPartNumber"
         ExecuteScript False, sSql
      End If

sql = "create procedure dbo.BackLogByPartNumber" & vbCrLf
sql = sql & "  @BegDate as varchar(16), " & vbCrLf
sql = sql & "  @EndDate as varchar(16), " & vbCrLf
sql = sql & "  @PartNumber as Varchar(30)," & vbCrLf
sql = sql & "  @Customer as varchar(10)," & vbCrLf
sql = sql & "  @PartClass as Varchar(16)," & vbCrLf
sql = sql & "  @PartCode as varchar(8)" & vbCrLf
sql = sql & " AS" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & " test" & vbCrLf
sql = sql & " BackLogByPartNumber '4/1/2017', '4/30/2017', '1', 'K', 'ALL', 'ALL'" & vbCrLf
sql = sql & "*/    " & vbCrLf
sql = sql & " declare @SoType as varchar(1)   " & vbCrLf
sql = sql & " declare @SoText as varchar(6)   " & vbCrLf
sql = sql & " declare @ItSo as int            " & vbCrLf
sql = sql & " declare @ItRev as char(2)       " & vbCrLf
sql = sql & " declare @ItNum as int           " & vbCrLf
sql = sql & " declare @ItQty as decimal(12,4) " & vbCrLf
sql = sql & " declare @PaLotRemQty as decimal(12,4)   " & vbCrLf
sql = sql & " declare @PartRem as decimal(12,4)       " & vbCrLf
sql = sql & " declare @RunningTot as decimal(12,4)    " & vbCrLf
sql = sql & " declare @ItDollars as decimal(12,4)     " & vbCrLf
sql = sql & " declare @ItSched as smalldatetime       " & vbCrLf
sql = sql & " declare @CusName as varchar(10)         " & vbCrLf
sql = sql & " declare @PartNum as varchar(30)         " & vbCrLf
sql = sql & " declare @CurPartNum as varchar(30)      " & vbCrLf
sql = sql & " declare @PartDesc as varchar(30)        " & vbCrLf
sql = sql & " declare @PartLoc as varchar(4)          " & vbCrLf
sql = sql & " declare @PartExDesc as varchar(3072)    " & vbCrLf
sql = sql & " declare @ItCanceled as tinyint          " & vbCrLf
sql = sql & " declare @ItPSNum as varchar(8)          " & vbCrLf
sql = sql & " declare @ItInvoice as int declare @ItPSShipped as tinyint   " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & " IF (@PartNumber = 'ALL')                      " & vbCrLf
sql = sql & "   SET @PartNumber = '%'                      " & vbCrLf
sql = sql & " IF (@Customer = 'ALL')                      " & vbCrLf
sql = sql & "   SET @Customer = ''                      " & vbCrLf
sql = sql & " IF (@PartClass = 'ALL')                     " & vbCrLf
sql = sql & "     SET @PartClass = ''                     " & vbCrLf
sql = sql & " IF (@PartCode = 'ALL')                      " & vbCrLf
sql = sql & "    SET @PartCode = ''                      " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CREATE TABLE #tempBackLogInfo               " & vbCrLf
sql = sql & "    (SOTYPE varchar(1) NULL,        " & vbCrLf
sql = sql & "    SOTEXT varchar(6) NULL,         " & vbCrLf
sql = sql & "    ITSO Int NULL,                  " & vbCrLf
sql = sql & "    ITREV char(2) NULL,              " & vbCrLf
sql = sql & "    ITNUMBER int NULL,              " & vbCrLf
sql = sql & "    ITQTY decimal(12,4) NULL,       " & vbCrLf
sql = sql & "    PALOTQTYREMAINING decimal(12,4) NULL,   " & vbCrLf
sql = sql & "    RUNQTYTOT decimal(12,4) NULL,   " & vbCrLf
sql = sql & "    ITDOLLARS decimal(12,4) NULL,   " & vbCrLf
sql = sql & "    ITSCHED smalldatetime NULL,     " & vbCrLf
sql = sql & "    CUNICKNAME varchar(10) NULL,    " & vbCrLf
sql = sql & "    PARTNUM varchar(30) NULL,       " & vbCrLf
sql = sql & "    PADESC varchar(30) NULL,     " & vbCrLf
sql = sql & "    PAEXTDESC varchar(3072) NULL,   " & vbCrLf
sql = sql & "    PALOCATION varchar(4) NULL,     " & vbCrLf
sql = sql & "    ITCANCELED tinyint NULL,        " & vbCrLf
sql = sql & "    ITPSNUMBER varchar(8) NULL, ITINVOICE int NULL, " & vbCrLf
sql = sql & "    ITPSSHIPPED tinyint NULL)       " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "      DECLARE curbackLog CURSOR   FOR                             " & vbCrLf
sql = sql & "       SELECT SohdTable.SOTYPE, SohdTable.SOTEXT,                  " & vbCrLf
sql = sql & "          SoitTable.ITSO, SoitTable.ITREV, SoitTable.ITNUMBER,    " & vbCrLf
sql = sql & "          SoitTable.ITQTY, PartTable.PALOTQTYREMAINING,           " & vbCrLf
sql = sql & "          SoitTable.ITDOLLARS,SoitTable.ITSCHED, CustTable.CUNICKNAME,    " & vbCrLf
sql = sql & "          PartTable.PARTNUM, PartTable.PADESC, PartTable.PAEXTDESC,       " & vbCrLf
sql = sql & "          PartTable.PALOCATION, SoitTable.ITCANCELED,                     " & vbCrLf
sql = sql & "          SoitTable.ITPSNUMBER , SoitTable.ITINVOICE, SoitTable.ITPSSHIPPED   " & vbCrLf
sql = sql & "       From SohdTable, SoitTable, CustTable, PartTable             " & vbCrLf
sql = sql & "       WHERE SohdTable.SOCUST = CustTable.CUREF AND                " & vbCrLf
sql = sql & "          SohdTable.SONUMBER =SoitTable.ITSO AND                  " & vbCrLf
sql = sql & "          SoitTable.ITPART=PartTable.PARTREF AND                  " & vbCrLf
sql = sql & "          SoitTable.ITCANCELED=0 AND SoitTable.ITPSNUMBER=''      " & vbCrLf
sql = sql & "          AND SoitTable.ITINVOICE=0 AND SoitTable.ITPSSHIPPED=0" & vbCrLf
sql = sql & "          AND SoitTable.ITPART LIKE @PartNumber + '%'" & vbCrLf
sql = sql & "          AND CUREF LIKE '%' + @Customer + '%'                    " & vbCrLf
sql = sql & "          AND SoitTable.ITSCHED BETWEEN @BegDate AND @EndDate     " & vbCrLf
sql = sql & "          AND PartTable.PACLASS LIKE '%' + @PartClass + '%'       " & vbCrLf
sql = sql & "          AND PartTable.PAPRODCODE LIKE '%' + @PartCode + '%'     " & vbCrLf
sql = sql & "       ORDER BY partnum, ITSCHED                                   " & vbCrLf
sql = sql & "    " & vbCrLf
sql = sql & "    OPEN curbackLog                                                " & vbCrLf
sql = sql & "    FETCH NEXT FROM curbackLog INTO @SoType, @SoText, @ItSo,  @ItRev, @ItNum, @ItQty, @PaLotRemQty, " & vbCrLf
sql = sql & "                     @ItDollars,@ItSched, @CusName, @PartNum,        " & vbCrLf
sql = sql & "                     @PartDesc, @PartExDesc, @PartLoc, @ItCanceled,  " & vbCrLf
sql = sql & "                     @ItPSNum, @ItInvoice, @ItPSShipped              " & vbCrLf
sql = sql & "     SET @CurPartNum = @PartNum                                      " & vbCrLf
sql = sql & "     SET @RunningTot = 0                                             " & vbCrLf
sql = sql & "     WHILE (@@FETCH_STATUS <> -1)                                    " & vbCrLf
sql = sql & "     BEGIN                                                           " & vbCrLf
sql = sql & "         IF (@@FETCH_STATUS <> -2)                                   " & vbCrLf
sql = sql & "         BEGIN                                                       " & vbCrLf
sql = sql & "             IF  @CurPartNum <> @PartNum                             " & vbCrLf
sql = sql & "            BEGIN                                                    " & vbCrLf
sql = sql & "                 SET @RunningTot = @ItQty                            " & vbCrLf
sql = sql & "                 set @CurPartNum = @PartNum                          " & vbCrLf
sql = sql & "             End                                                     " & vbCrLf
sql = sql & "             Else                                                    " & vbCrLf
sql = sql & "             BEGIN                                                   " & vbCrLf
sql = sql & "                 SET @RunningTot = @RunningTot + @ItQty              " & vbCrLf
sql = sql & "             End                                                     " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "             SET @PartRem = @PaLotRemQty - @RunningTot                       " & vbCrLf
sql = sql & "             INSERT INTO #tempBackLogInfo (SOTYPE, SOTEXT, ITSO, ITREV,      " & vbCrLf
sql = sql & "                 ITNUMBER,ITQTY, PALOTQTYREMAINING,RUNQTYTOT, ITDOLLARS,     " & vbCrLf
sql = sql & "                 ITSCHED,CUNICKNAME, PARTNUM, PADESC,PAEXTDESC,PALOCATION,   " & vbCrLf
sql = sql & "                      ITCANCELED, ITPSNUMBER, ITINVOICE, ITPSSHIPPED)             " & vbCrLf
sql = sql & "             VALUES (@SoType, @SoText, @ItSo, @ItRev,@ItNum, @ItQty,@PaLotRemQty,@PartRem, @ItDollars,@ItSched,@CusName, " & vbCrLf
sql = sql & "                 @PartNum,@PartDesc,@PartExDesc,@PartLoc, @ItCanceled,@ItPSNum,@ItInvoice,@ItPSShipped)  " & vbCrLf
sql = sql & "         End                                                                 " & vbCrLf
sql = sql & "         FETCH NEXT FROM curbackLog INTO @SoType, @SoText, @ItSo,            " & vbCrLf
sql = sql & "             @ItRev, @ItNum, @ItQty, @PaLotRemQty,                           " & vbCrLf
sql = sql & "             @ItDollars,@ItSched, @CusName, @PartNum,                        " & vbCrLf
sql = sql & "             @PartDesc, @PartExDesc, @PartLoc, @ItCanceled,                  " & vbCrLf
sql = sql & "             @ItPSNum, @ItInvoice, @ItPSShipped                              " & vbCrLf
sql = sql & "     End                                                                     " & vbCrLf
sql = sql & "     CLOSE curbackLog   --// close the cursor                                " & vbCrLf
sql = sql & "     DEALLOCATE curbackLog                                                   " & vbCrLf
sql = sql & "     -- select data for the report                                           " & vbCrLf
sql = sql & "     SELECT SOTYPE, SOTEXT, ITSO, ITREV,                                     " & vbCrLf
sql = sql & "         ITNUMBER,ITQTY, PALOTQTYREMAINING,RUNQTYTOT, ITDOLLARS,             " & vbCrLf
sql = sql & "         ITSCHED,CUNICKNAME, PARTNUM, PADESC,PAEXTDESC, PALOCATION,          " & vbCrLf
sql = sql & "         ITCANCELED , ITPSNUMBER, ITINVOICE, ITPSSHIPPED                     " & vbCrLf
sql = sql & "     FROM #tempBackLogInfo                                                   " & vbCrLf
sql = sql & "     ORDER BY ITSCHED                                                        " & vbCrLf
sql = sql & "     -- drop the temp table                                                  " & vbCrLf
sql = sql & "     DROP table #tempBackLogInfo                                             " & vbCrLf
ExecuteScript False, sql

''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function



Private Function UpdateDatabase86()

   Dim sql As String
   sql = ""

   newver = 161     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "if object_id('SheetUnitConversion') IS NOT NULL" & vbCrLf
sql = sql & "    drop procedure SheetUnitConversion" & vbCrLf
ExecuteScript False, sql

sql = "create procedure dbo.SheetUnitConversion" & vbCrLf
sql = sql & "   @LotNo as char(15)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "/* convert lot from purchasing units to inventory units" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "exec SheetUnitConversion '42701-729201-79'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get purchase item unit cost, lot dimensions, and units" & vbCrLf
sql = sql & "declare @purchaseUnitCost decimal(12,4), @height decimal(12,4), @length decimal(12,4), " & vbCrLf
sql = sql & "   @purchUnit char(2), @invUnit char(2), @lotQtyAfter decimal(12,4), @unitCost decimal(12,4), " & vbCrLf
sql = sql & "   @purchQty decimal(12,4), @partRef varchar(30), @lotQtyBefore decimal(12,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select @purchaseUnitCost = poi.PIAMT, @height = loh.LOTMATHEIGHT, @length = loh.LOTMATLENGTH," & vbCrLf
sql = sql & "   @purchUnit = pt.PAPUNITS, @invUnit = pt.PAUNITS, @purchQty = poi.PIAQTY," & vbCrLf
sql = sql & "   @lotQtyAfter = loh.LOTMATHEIGHT * loh.LOTMATLENGTH * ROUND((loh.LOTTOTMATL + .1) / POI.PIAMT,0)," & vbCrLf
sql = sql & "   @partRef = pt.PARTREF, @lotQtyBefore = loh.LOTORIGINALQTY " & vbCrLf
sql = sql & "from PoitTable poi" & vbCrLf
sql = sql & "join PohdTable poh on poh.PONUMBER = poi.PINUMBER" & vbCrLf
sql = sql & "join LohdTable loh on loh.LOTPO = poi.PINUMBER and loh.LOTPOITEM = poi.PIITEM and loh.LOTPOITEMREV = poi.PIREV" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = loh.LOTPARTREF" & vbCrLf
sql = sql & "where loh.LOTNUMBER = @LotNo and  pt.PAPUNITS = 'SH' and pt.PAUNITS <> pt.PAPUNITS and loh.LOTMATHEIGHT > 1 and loh.LOTMATLENGTH > 1 " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--print 'before = ' + cast(@lotQtyBefore as varchar(18))" & vbCrLf
sql = sql & "--print 'after  = ' + cast(@lotQtyAfter as varchar(18))" & vbCrLf
sql = sql & "--print cast(@purchaseUnitCost as varchar(12))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if data missing or lot alreadhy converted, do not do it again" & vbCrLf
sql = sql & "if @purchaseUnitCost is null or @lotQtyBefore = @lotQtyAfter or isnull(@lotQtyAfter, 0) = 0" & vbCrLf
sql = sql & "   return" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @unitCost = (@purchaseUnitCost * @lotQtyBefore) / @lotQtyAfter" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--print cast(@height as varchar(12))" & vbCrLf
sql = sql & "--print cast(@length as varchar(12))" & vbCrLf
sql = sql & "--print cast(@purchUnit as varchar(12))" & vbCrLf
sql = sql & "--print cast(@invUnit as varchar(12))" & vbCrLf
sql = sql & "--print cast(@unitCost as varchar(12))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update lot header record" & vbCrLf
sql = sql & "update LohdTable set LOTORIGINALQTY = @lotQtyAfter, LOTREMAININGQTY = @lotQtyAfter, LOTUNITCOST = @unitCost," & vbCrLf
sql = sql & "   LOTTOTMATL = @lotQtyAfter * @unitCost" & vbCrLf
sql = sql & "where LOTNUMBER = @LotNo " & vbCrLf
sql = sql & "--select * from LohdTable where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update first lot item (receipt)" & vbCrLf
sql = sql & "update LoitTable set LOIQUANTITY = @lotQtyAfter, LOIUNITS = @invUnit, LOIHEIGHT = @height, LOILENGTH = @length" & vbCrLf
sql = sql & "where LOINUMBER = @LotNo and LOIRECORD = 1" & vbCrLf
sql = sql & "--select * from LoitTable where LOINUMBER = @LotNo and LOIRECORD = 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update inventory activity for PO receipt" & vbCrLf
sql = sql & "update invatable set INPQTY = @lotQtyAfter, INAQTY = @lotQtyAfter, INAMT = @unitCost, INUNITS = @invUnit" & vbCrLf
sql = sql & "where INLOTNUMBER = @LotNo and INTYPE = 15" & vbCrLf
sql = sql & "--select * from InvaTable where INLOTNUMBER = @LotNo and INTYPE = 15" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update part QOH" & vbCrLf
sql = sql & "select PAQOH from PartTable where PARTREF = @partRef" & vbCrLf
sql = sql & "update PartTable set PAQOH = PAQOH + @lotQtyAfter - @lotQtyBefore where PARTREF = @partRef" & vbCrLf
sql = sql & "--select PAQOH from PartTable where PARTREF = @partRef" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "commit tran" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql


' fix sheet inventory rounding errors
sql = "update lh" & vbCrLf
sql = sql & "set LOTTOTMATL = LOTORIGINALQTY * LOTUNITCOST" & vbCrLf
sql = sql & "from LohdTable lh" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = lh.LOTPARTREF" & vbCrLf
sql = sql & "join ComnTable com on COUSESHEETINVENTORY = 1" & vbCrLf
sql = sql & "where PAPUNITS = 'SH' and PAUNITS <> PAPUNITS" & vbCrLf
sql = sql & "and LOTTOTMATL <> LOTORIGINALQTY * LOTUNITCOST" & vbCrLf
ExecuteScript False, sql

sql = "if object_id('SheetRestock') IS NOT NULL" & vbCrLf
sql = sql & "    drop procedure SheetRestock" & vbCrLf
ExecuteScript False, sql

sql = "create procedure [dbo].[SheetRestock]" & vbCrLf
sql = sql & "   @UserLotNo varchar(40)," & vbCrLf
sql = sql & "   @User varchar(4)," & vbCrLf
sql = sql & "   @Comments varchar(2048)," & vbCrLf
sql = sql & "   @Location varchar(4)," & vbCrLf
sql = sql & "   @Params varchar(2000)   -- LOIRECORD,NEWHT,NEWLEN,... repeat (include comma at end)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "restock sheet LOI records" & vbCrLf
sql = sql & "exec SheetRestock '029373-1-A', 'MGR', '2,36,96,0,12,72,'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "create table #rect" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "   ParentRecord int," & vbCrLf
sql = sql & "   NewHeight decimal(12,4)," & vbCrLf
sql = sql & "   NewLength decimal(12,4)," & vbCrLf
sql = sql & "   NewRecord int," & vbCrLf
sql = sql & "   Qty decimal(12,4)," & vbCrLf
sql = sql & "   NewIANumber int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @nextRecord int, @nextIANumber int" & vbCrLf
sql = sql & "declare @LotNo char(15), @id int, @SO int" & vbCrLf
sql = sql & "select @LotNo = LOTNUMBER from LohdTable where LOTUSERLOTID = @UserLotNo" & vbCrLf
sql = sql & "select @nextRecord = max(LOIRECORD) + 1 , @SO = max(LOISONUMBER) from LoitTable where LOINUMBER = @LotNo" & vbCrLf
sql = sql & "select @nextIANumber = max(INNUMBER) + 1 from InvaTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- first update the lot header" & vbCrLf
sql = sql & "update LohdTable set LOTCOMMENTS = @Comments, LOTLOCATION = @Location where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- extract parameters" & vbCrLf
sql = sql & "DECLARE @start INT, @end INT" & vbCrLf
sql = sql & "declare @stringId varchar(10), @stringHt varchar(10), @stringLen varchar(10)" & vbCrLf
sql = sql & "SELECT @start = 1, @end = CHARINDEX(',', @Params) " & vbCrLf
sql = sql & "WHILE @start < LEN(@Params) + 1 BEGIN " & vbCrLf
sql = sql & "    IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    set @stringId = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "    SET @start = @end + 1 " & vbCrLf
sql = sql & "    SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    set @stringHt = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "    SET @start = @end + 1 " & vbCrLf
sql = sql & "    SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    set @stringLen = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "    SET @start = @end + 1 " & vbCrLf
sql = sql & "    SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "   " & vbCrLf
sql = sql & "   declare @ht decimal(12,4), @len decimal(12,4)" & vbCrLf
sql = sql & "   set @ht = cast(@stringHt as decimal(12,4))" & vbCrLf
sql = sql & "   set @len = cast(@stringLen as decimal(12,4))" & vbCrLf
sql = sql & "   insert into #rect (ParentRecord, NewHeight, NewLength, NewRecord, NewIANumber)" & vbCrLf
sql = sql & "   values (cast(@stringId as int), @ht, @len," & vbCrLf
sql = sql & "       case when @ht*@len = 0 then 0 else @nextRecord end," & vbCrLf
sql = sql & "       case when @ht*@len = 0 then 0 else @nextIANumber end)" & vbCrLf
sql = sql & "   if (@ht*@len) <> 0 " & vbCrLf
sql = sql & "   begin" & vbCrLf
sql = sql & "       set @nextRecord = @nextRecord + 1" & vbCrLf
sql = sql & "       set @nextIANumber = @nextIANumber + 1" & vbCrLf
sql = sql & "   end" & vbCrLf
sql = sql & "END " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #rect set Qty = NewHeight * NewLength" & vbCrLf
sql = sql & "update #rect set ParentRecord = isnull((select top 1 ParentRecord from #rect where ParentRecord <> 0 ),0)" & vbCrLf
sql = sql & "where ParentRecord = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from #rect" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @partRef varchar(30), @currentTime datetime, @uom char(2), @type int" & vbCrLf
sql = sql & "declare @unitCost decimal(12,4), @nextINNUMBER int, @totalQty decimal(12,4)" & vbCrLf
sql = sql & "select @partRef = LOTPARTREF, @unitCost = LOTUNITCOST, @totalQty = LOTREMAININGQTY from LohdTable where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "select @uom = PAUNITS from PartTable where PARTREF = @partRef" & vbCrLf
sql = sql & "set @currentTime = getdate()" & vbCrLf
sql = sql & "select @nextINNUMBER = max(INNUMBER) + 1 from InvaTable" & vbCrLf
sql = sql & "set @type = 19     -- manual adjustment" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update existing lot items" & vbCrLf
sql = sql & "Update li" & vbCrLf
sql = sql & "   set LOIINACTIVE = 1" & vbCrLf
sql = sql & "from LoitTable li " & vbCrLf
sql = sql & "join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD = r.ParentRecord" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- remove rectangles with no quantity remaining.  These items will not be restocked." & vbCrLf
sql = sql & "delete from #rect where Qty = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new lot items" & vbCrLf
sql = sql & "INSERT INTO [dbo].[LoitTable]" & vbCrLf
sql = sql & "           ([LOINUMBER]" & vbCrLf
sql = sql & "           ,[LOIRECORD]" & vbCrLf
sql = sql & "           ,[LOITYPE]" & vbCrLf
sql = sql & "           ,[LOIPARTREF]" & vbCrLf
sql = sql & "           ,[LOIADATE]" & vbCrLf
sql = sql & "           ,[LOIPDATE]" & vbCrLf
sql = sql & "           ,[LOIQUANTITY]" & vbCrLf
sql = sql & "           ,[LOIMOPARTREF]" & vbCrLf
sql = sql & "           ,[LOIMORUNNO]" & vbCrLf
sql = sql & "           ,[LOIPONUMBER]" & vbCrLf
sql = sql & "           ,[LOIPOITEM]" & vbCrLf
sql = sql & "           ,[LOIPOREV]" & vbCrLf
sql = sql & "           ,[LOIPSNUMBER]" & vbCrLf
sql = sql & "           ,[LOIPSITEM]" & vbCrLf
sql = sql & "           ,[LOICUSTINVNO]" & vbCrLf
sql = sql & "           ,[LOICUST]" & vbCrLf
sql = sql & "           ,[LOIVENDINVNO]" & vbCrLf
sql = sql & "           ,[LOIVENDOR]" & vbCrLf
sql = sql & "           ,[LOIACTIVITY]" & vbCrLf
sql = sql & "           ,[LOICOMMENT]" & vbCrLf
sql = sql & "           ,[LOIUNITS]" & vbCrLf
sql = sql & "           ,[LOIMOPKCANCEL]" & vbCrLf
sql = sql & "           ,[LOIHEIGHT]" & vbCrLf
sql = sql & "           ,[LOILENGTH]" & vbCrLf
sql = sql & "           ,[LOIUSER]" & vbCrLf
sql = sql & "          ,LOIPARENTREC" & vbCrLf
sql = sql & "           ,[LOISONUMBER]" & vbCrLf
sql = sql & "          ,LOISHEETACTTYPE" & vbCrLf
sql = sql & "          )" & vbCrLf
sql = sql & "     SELECT" & vbCrLf
sql = sql & "           @LotNo" & vbCrLf
sql = sql & "           ,r.NewRecord" & vbCrLf
sql = sql & "           ,@type              -- manual adjustment" & vbCrLf
sql = sql & "           ,@partRef" & vbCrLf
sql = sql & "           ,@currentTime" & vbCrLf
sql = sql & "           ,@currentTime" & vbCrLf
sql = sql & "           ,r.Qty" & vbCrLf
sql = sql & "           ,null   --LOIMOPARTREF, char(30),>" & vbCrLf
sql = sql & "           ,null   --<LOIMORUNNO, int,>" & vbCrLf
sql = sql & "           ,null   --<LOIPONUMBER, int,>" & vbCrLf
sql = sql & "           ,null   --<LOIPOITEM, smallint,>" & vbCrLf
sql = sql & "           ,null   --<LOIPOREV, char(2),>" & vbCrLf
sql = sql & "           ,null   --<LOIPSNUMBER, char(8),>" & vbCrLf
sql = sql & "           ,null   --<LOIPSITEM, smallint,>" & vbCrLf
sql = sql & "           ,null   --<LOICUSTINVNO, int,>" & vbCrLf
sql = sql & "           ,null   --<LOICUST, char(10),>" & vbCrLf
sql = sql & "           ,null   --<LOIVENDINVNO, char(20),>" & vbCrLf
sql = sql & "           ,null   --<LOIVENDOR, char(10),>" & vbCrLf
sql = sql & "           ,null   --<LOIACTIVITY, int,> points to InvaTable.INNO when IA is created" & vbCrLf
sql = sql & "           ,'sheet restock'    --<LOICOMMENT, varchar(40),>" & vbCrLf
sql = sql & "           ,@uom   --<LOIUNITS, char(2),>" & vbCrLf
sql = sql & "           ,null   --<LOIMOPKCANCEL, smallint,>" & vbCrLf
sql = sql & "           ,r.NewHeight    --<LOIHEIGHT, decimal(12,4),>" & vbCrLf
sql = sql & "           ,r.NewLength    --<LOILENGTH, decimal(12,4),>" & vbCrLf
sql = sql & "           ,@User      --<LOIUSER, varchar(4),>" & vbCrLf
sql = sql & "          ,r.ParentRecord  --LOIPICKEDFROMREC" & vbCrLf
sql = sql & "           ,@SO        --<LOISONUMBER, int,>" & vbCrLf
sql = sql & "          ,'RS'    --LOISHEETACTTYPE" & vbCrLf
sql = sql & "       from #rect r" & vbCrLf
sql = sql & "       where r.Qty <> 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select distinct li.* from LoitTable li " & vbCrLf
sql = sql & "--join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD > 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- set LOTREMAININGQTY and remove reservation" & vbCrLf
sql = sql & "declare @sum decimal(12,4)" & vbCrLf
sql = sql & "select @sum = sum(Qty) from #rect" & vbCrLf
sql = sql & "update LohdTable " & vbCrLf
sql = sql & "   set LOTREMAININGQTY = LOTREMAININGQTY + @sum," & vbCrLf
sql = sql & "   LOTRESERVEDBY = NULL," & vbCrLf
sql = sql & "   LOTRESERVEDON = NULL" & vbCrLf
sql = sql & "   where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new ia items" & vbCrLf
sql = sql & "INSERT INTO [dbo].[InvaTable]" & vbCrLf
sql = sql & "           ([INTYPE]" & vbCrLf
sql = sql & "           ,[INPART]" & vbCrLf
sql = sql & "           ,[INREF1]" & vbCrLf
sql = sql & "           ,[INREF2]" & vbCrLf
sql = sql & "           ,[INPDATE]" & vbCrLf
sql = sql & "           ,[INADATE]" & vbCrLf
sql = sql & "           ,[INPQTY]" & vbCrLf
sql = sql & "           ,[INAQTY]" & vbCrLf
sql = sql & "           ,[INAMT]" & vbCrLf
sql = sql & "           ,[INTOTMATL]" & vbCrLf
sql = sql & "           ,[INTOTLABOR]" & vbCrLf
sql = sql & "           ,[INTOTEXP]" & vbCrLf
sql = sql & "           ,[INTOTOH]" & vbCrLf
sql = sql & "           ,[INTOTHRS]" & vbCrLf
sql = sql & "           ,[INCREDITACCT]" & vbCrLf
sql = sql & "           ,[INDEBITACCT]" & vbCrLf
sql = sql & "           ,[INGLJOURNAL]" & vbCrLf
sql = sql & "           ,[INGLPOSTED]" & vbCrLf
sql = sql & "           ,[INGLDATE]" & vbCrLf
sql = sql & "           ,[INMOPART]" & vbCrLf
sql = sql & "           ,[INMORUN]" & vbCrLf
sql = sql & "           ,[INSONUMBER]" & vbCrLf
sql = sql & "           ,[INSOITEM]" & vbCrLf
sql = sql & "           ,[INSOREV]" & vbCrLf
sql = sql & "           ,[INPONUMBER]" & vbCrLf
sql = sql & "           ,[INPORELEASE]" & vbCrLf
sql = sql & "           ,[INPOITEM]" & vbCrLf
sql = sql & "           ,[INPOREV]" & vbCrLf
sql = sql & "           ,[INPSNUMBER]" & vbCrLf
sql = sql & "           ,[INPSITEM]" & vbCrLf
sql = sql & "           ,[INWIPLABACCT]" & vbCrLf
sql = sql & "           ,[INWIPMATACCT]" & vbCrLf
sql = sql & "           ,[INWIPOHDACCT]" & vbCrLf
sql = sql & "           ,[INWIPEXPACCT]" & vbCrLf
sql = sql & "           ,[INNUMBER]" & vbCrLf
sql = sql & "           ,[INLOTNUMBER]" & vbCrLf
sql = sql & "           ,[INUSER]" & vbCrLf
sql = sql & "           ,[INUNITS]" & vbCrLf
sql = sql & "           ,[INDRLABACCT]" & vbCrLf
sql = sql & "           ,[INDRMATACCT]" & vbCrLf
sql = sql & "           ,[INDREXPACCT]" & vbCrLf
sql = sql & "           ,[INDROHDACCT]" & vbCrLf
sql = sql & "           ,[INCRLABACCT]" & vbCrLf
sql = sql & "           ,[INCRMATACCT]" & vbCrLf
sql = sql & "           ,[INCREXPACCT]" & vbCrLf
sql = sql & "           ,[INCROHDACCT]" & vbCrLf
sql = sql & "           ,[INLOTTRACK]" & vbCrLf
sql = sql & "           ,[INUSEACTUALCOST]" & vbCrLf
sql = sql & "           ,[INCOSTEDBY]" & vbCrLf
sql = sql & "           ,[INMAINTCOSTED])" & vbCrLf
sql = sql & "     select" & vbCrLf
sql = sql & "           @type               --<INTYPE, int,>" & vbCrLf
sql = sql & "           ,@partRef           --char(30),>" & vbCrLf
sql = sql & "           ,'Manual Adjustment'    --<INREF1, char(20),>" & vbCrLf
sql = sql & "           ,'Sheet Inventory'  --<INREF2, char(40),>" & vbCrLf
sql = sql & "           ,@currentTime       --<INPDATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,@currentTime       --<INADATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,r.Qty          --<INPQTY, decimal(12,4),>" & vbCrLf
sql = sql & "           ,r.Qty          --<INAQTY, decimal(12,4),>" & vbCrLf
sql = sql & "           ,@unitCost          -- <INAMT, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTMATL, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTLABOR, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTEXP, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTOH, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTHRS, decimal(12,4),>" & vbCrLf
sql = sql & "           ,''                 --<INCREDITACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDEBITACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INGLJOURNAL, char(12),>" & vbCrLf
sql = sql & "           ,0                  --<INGLPOSTED, tinyint,>" & vbCrLf
sql = sql & "           ,null               --<INGLDATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,''                 --<INMOPART, char(30),>" & vbCrLf
sql = sql & "           ,0                  --<INMORUN, int,>" & vbCrLf
sql = sql & "           ,0                  --<INSONUMBER, int,>" & vbCrLf
sql = sql & "           ,0                  --<INSOITEM, int,>" & vbCrLf
sql = sql & "           ,''                 --<INSOREV, char(2),>" & vbCrLf
sql = sql & "           ,0                  --<INPONUMBER, int,>" & vbCrLf
sql = sql & "           ,0                  --<INPORELEASE, smallint,>" & vbCrLf
sql = sql & "           ,0                  --<INPOITEM, smallint,>" & vbCrLf
sql = sql & "           ,''                 --<INPOREV, char(2),>" & vbCrLf
sql = sql & "           ,''                 --<INPSNUMBER, char(8),>" & vbCrLf
sql = sql & "           ,0                  --<INPSITEM, smallint,>" & vbCrLf
sql = sql & "           ,''                 --<INWIPLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPOHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPEXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,r.NewIANumber      --<INNUMBER, int,>" & vbCrLf
sql = sql & "           ,@LotNo             --<INLOTNUMBER, char(15),>" & vbCrLf
sql = sql & "           ,@User              --<INUSER, char(4),>" & vbCrLf
sql = sql & "           ,@uom               --<INUNITS, char(2),>" & vbCrLf
sql = sql & "           ,''                 --<INDRLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDRMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDREXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDROHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCRLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCRMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCREXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCROHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,null               --<INLOTTRACK, bit,>" & vbCrLf
sql = sql & "           ,null               --<INUSEACTUALCOST, bit,>" & vbCrLf
sql = sql & "           ,null               --<INCOSTEDBY, char(4),>" & vbCrLf
sql = sql & "           ,0                  --<INMAINTCOSTED, int,>)" & vbCrLf
sql = sql & "       from #rect r" & vbCrLf
sql = sql & "       where r.Qty <> 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now update part QOH" & vbCrLf
sql = sql & "update PartTable set PAQOH = PAQOH + @totalQty where PARTREF = @partRef" & vbCrLf
sql = sql & "commit tran" & vbCrLf
ExecuteScript False, sql

sql = "update pt set PAQOH = (select sum(lotremainingqty) from LohdTable where LOTPARTREF = partref)" & vbCrLf
sql = sql & "from PartTable pt " & vbCrLf
sql = sql & "join ComnTable on COUSESHEETINVENTORY = 1" & vbCrLf
sql = sql & "where PAPUNITS = 'SH'" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function

Private Function UpdateDatabase87()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 162     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "CREATE TABLE [dbo].[MrpPartComments](" & vbCrLf
sql = sql & "  [CommentID] [int] IDENTITY(1,1) NOT NULL," & vbCrLf
sql = sql & "  [MrpPart] [char](30) NOT NULL," & vbCrLf
sql = sql & "  [CreatedOn] [datetime] NOT NULL," & vbCrLf
sql = sql & "  [CreatedBy] [varchar](3) NOT NULL," & vbCrLf
sql = sql & "  [Comment] [varchar](max) NOT NULL," & vbCrLf
sql = sql & " CONSTRAINT [PK_MrpPartComments] PRIMARY KEY CLUSTERED " & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "  [CommentID] ASC" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql

sql = "ALTER TABLE [dbo].[MrpPartComments] ADD  CONSTRAINT [DF_Table_1_CommentDate]  DEFAULT (getdate()) FOR [CreatedOn]" & vbCrLf
ExecuteScript False, sql

sql = "ALTER TABLE [dbo].[MrpPartComments]  WITH CHECK ADD  CONSTRAINT [FK_MrpPartComments_PartTable] FOREIGN KEY([MrpPart])" & vbCrLf
sql = sql & "REFERENCES [dbo].[PartTable] ([PARTREF])" & vbCrLf
ExecuteScript False, sql

sql = "ALTER TABLE [dbo].[MrpPartComments] CHECK CONSTRAINT [FK_MrpPartComments_PartTable]" & vbCrLf
ExecuteScript False, sql


sql = "if object_id('RptMRPMOQtyShortage') IS NOT NULL" & vbCrLf
sql = sql & "    drop procedure RptMRPMOQtyShortage" & vbCrLf
ExecuteScript False, sql


sql = "create PROCEDURE [dbo].[RptMRPMOQtyShortage]" & vbCrLf
sql = sql & "      @InMOPart as varchar(30), @StartDate as datetime, @EndDate as datetime" & vbCrLf
sql = sql & "AS " & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   declare @ChildKey as varchar(1024)" & vbCrLf
sql = sql & "   declare @GLSortKey as varchar(1024)" & vbCrLf
sql = sql & "   declare @SortKey as varchar(1024)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   declare @MOPart as varchar(30)" & vbCrLf
sql = sql & "   declare @MORun as Integer" & vbCrLf
sql = sql & "   declare @MOQtyRqd as decimal(12,4)" & vbCrLf
sql = sql & "   declare @MOPartRqDt as datetime " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   declare @Part as varchar(30)" & vbCrLf
sql = sql & "   declare @PAQOH as decimal(12,4)" & vbCrLf
sql = sql & "   declare @RunTot as decimal(12,4)" & vbCrLf
sql = sql & "   declare @AssyPart as varchar(30)" & vbCrLf
sql = sql & "   declare @BMQtyReq as decimal(12,4)" & vbCrLf
sql = sql & "   declare @RunQtyReq as decimal (12, 4)" & vbCrLf
sql = sql & "   declare @PartDateQrd as datetime" & vbCrLf
sql = sql & "   " & vbCrLf
sql = sql & "   --DROP TABLE #tempMOPartsDetail " & vbCrLf
sql = sql & "   DELETE FROM tempMrplPartShort" & vbCrLf
sql = sql & "   " & vbCrLf
sql = sql & "  BEGIN" & vbCrLf
sql = sql & "   IF (@InMOPart = '')" & vbCrLf
sql = sql & "      SET @InMOPart = @InMOPart + '%'" & vbCrLf
sql = sql & "  " & vbCrLf
sql = sql & "  DECLARE curMrpExp CURSOR  FOR" & vbCrLf
sql = sql & "   SELECT MRP_PARTREF,0 as RUNNO, MRP_PARTQTYRQD, MRP_ACTIONDATE" & vbCrLf
sql = sql & "   FROM MrplTable, PartTable   " & vbCrLf
sql = sql & "   WHERE MRP_PARTREF = PartRef   " & vbCrLf
sql = sql & "      AND MrplTable.MRP_PARTREF LIKE @InMOPart" & vbCrLf
sql = sql & "      AND MrplTable.MRP_PARTPRODCODE LIKE '%'  " & vbCrLf
sql = sql & "      AND MrplTable.MRP_PARTCLASS LIKE '%'  " & vbCrLf
sql = sql & "      AND MrplTable.MRP_POBUYER LIKE '%'  " & vbCrLf
sql = sql & "      AND MrplTable.MRP_PARTDATERQD BETWEEN @StartDate AND @EndDate" & vbCrLf
sql = sql & "      AND MrplTable.MRP_TYPE IN (6, 5)   " & vbCrLf
sql = sql & "      AND PartTable.PAMAKEBUY ='M'" & vbCrLf
sql = sql & "   UNION" & vbCrLf
sql = sql & "      SELECT DISTINCT RUNREF, RUNNO, RUNQTY,runpkstart  as MRP_ACTIONDATE FROM RunsTable WHERE " & vbCrLf
sql = sql & "         RUNREF LIKE @InMOPart AND RUNSTATUS = 'SC'" & vbCrLf
sql = sql & "         AND RUNPKSTART BETWEEN @StartDate  AND @EndDate + ' 23:00' order by MRP_ACTIONDATE" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "  OPEN curMrpExp" & vbCrLf
sql = sql & "  FETCH NEXT FROM curMrpExp INTO @MOPart, @MORun, @MOQtyRqd, @MOPartRqDt" & vbCrLf
sql = sql & "  WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "  BEGIN" & vbCrLf
sql = sql & "     IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "     BEGIN" & vbCrLf
sql = sql & "--print 'MO:' + @MOPart + '; RUN:' + Convert(varchar(10), @MORun) + '; Date:' + Convert(varchar(24), @MOPartRqDt, 101);" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "      with cte" & vbCrLf
sql = sql & "      as (select BMASSYPART,BMPARTREF,BMPARTREV, BMQTYREQD , RTrim(BMUNITS) BMUNITS, " & vbCrLf
sql = sql & "            BMCONVERSION, BMSEQUENCE, 0 as level," & vbCrLf
sql = sql & "            cast(LTRIM(RTrim(BMASSYPART)) + char(36)+ COALESCE(cast(BMSEQUENCE as varchar(4)), '') + LTRIM(RTrim(BMPARTREF)) as varchar(max)) as SortKey" & vbCrLf
sql = sql & "         from BmplTable" & vbCrLf
sql = sql & "         where BMASSYPART = @MOPart" & vbCrLf
sql = sql & "         union all" & vbCrLf
sql = sql & "         select a.BMASSYPART,a.BMPARTREF,a.BMPARTREV, a.BMQTYREQD , RTrim(a.BMUNITS) BMUNITS, " & vbCrLf
sql = sql & "            a.BMCONVERSION, a.BMSEQUENCE, level + 1," & vbCrLf
sql = sql & "            cast(COALESCE(SortKey,'') + char(36) + COALESCE(cast(a.BMSEQUENCE as varchar(4)), '') + COALESCE(LTRIM(RTrim(a.BMPARTREF)) ,'') as varchar(max))as SortKey" & vbCrLf
sql = sql & "         from cte" & vbCrLf
sql = sql & "            inner join BmplTable a" & vbCrLf
sql = sql & "on cte.BMPARTREF = a.BMASSYPART" & vbCrLf
sql = sql & "     ) " & vbCrLf
sql = sql & "     INSERT INTO tempMrplPartShort(BMASSYPART,BMPARTREF,BMQTYREQD," & vbCrLf
sql = sql & "      SORTKEYLEVEL,BMSEQUENCE, SortKey, PAQOH, RUNNO,MRP_PARTQTYRQD, MRP_ACTIONDATE)" & vbCrLf
sql = sql & "     select BMASSYPART, BMPARTREF,BMQTYREQD,level,BMSEQUENCE, SortKey, PAQOH, @MORun, @MOQtyRqd, @MOPartRqDt" & vbCrLf
sql = sql & "       from cte, PartTable WHERE PARTREF = BMPARTREF  AND BMPARTREF <> 'NULL' order by SortKey,BMSEQUENCE" & vbCrLf
sql = sql & "        " & vbCrLf
sql = sql & "   End" & vbCrLf
sql = sql & "   FETCH NEXT FROM curMrpExp INTO @MOPart, @MORun, @MOQtyRqd, @MOPartRqDt" & vbCrLf
sql = sql & "  End" & vbCrLf
sql = sql & "  Close curMrpExp" & vbCrLf
sql = sql & "  DEALLOCATE curMrpExp" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "  DECLARE curRunTot CURSOR  FOR" & vbCrLf
sql = sql & "  select DISTINCT BMPARTREF, PAQOH from tempMrplPartShort order by BMPARTREF" & vbCrLf
sql = sql & "  OPEN curRunTot" & vbCrLf
sql = sql & "  FETCH NEXT FROM curRunTot INTO @Part, @PAQOH" & vbCrLf
sql = sql & "  WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "  BEGIN" & vbCrLf
sql = sql & "     IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "     BEGIN" & vbCrLf
sql = sql & "        SET @RunTot = 0.0000" & vbCrLf
sql = sql & "        SET @RunTot = @PAQOH" & vbCrLf
sql = sql & "        DECLARE curRunTot1 CURSOR  FOR" & vbCrLf
sql = sql & "         select DISTINCT BMASSYPART, BMQTYREQD, MRP_PARTQTYRQD, MRP_ACTIONDATE from tempMrplPartShort " & vbCrLf
sql = sql & "            WHERE BMPARTREF = @Part AND sortkeylevel = 0 " & vbCrLf
sql = sql & "         order by MRP_ACTIONDATE  -- BMASSYPART, " & vbCrLf
sql = sql & "        OPEN curRunTot1" & vbCrLf
sql = sql & "        FETCH NEXT FROM curRunTot1 INTO @AssyPart, @BMQtyReq, @RunQtyReq, @PartDateQrd" & vbCrLf
sql = sql & "        WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "        BEGIN" & vbCrLf
sql = sql & "          IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "          BEGIN" & vbCrLf
sql = sql & "            --Set @RunTot = ROUND(@RunTot,4)" & vbCrLf
sql = sql & "            Set @RunTot = @RunTot -  ( @BMQtyReq * @RunQtyReq)" & vbCrLf
sql = sql & "            UPDATE tempMrplPartShort SET PAQRUNTOT = @RunTot WHERE " & vbCrLf
sql = sql & "BMASSYPART = @AssyPart AND BMPARTREF = @Part AND MRP_ACTIONDATE = @PartDateQrd" & vbCrLf
sql = sql & "          END" & vbCrLf
sql = sql & "          FETCH NEXT FROM curRunTot1 INTO @AssyPart, @BMQtyReq, @RunQtyReq,@PartDateQrd" & vbCrLf
sql = sql & "        End" & vbCrLf
sql = sql & "        Close curRunTot1" & vbCrLf
sql = sql & "        DEALLOCATE curRunTot1" & vbCrLf
sql = sql & "     END" & vbCrLf
sql = sql & "     FETCH NEXT FROM curRunTot INTO @Part, @PAQOH" & vbCrLf
sql = sql & "  End" & vbCrLf
sql = sql & "  Close curRunTot" & vbCrLf
sql = sql & "  DEALLOCATE curRunTot" & vbCrLf
sql = sql & " END" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function

Private Function UpdateDatabase88()

   Dim sql As String
   sql = ""

   newver = 163
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "alter table SSRSInfo add WebUrl varchar(255) null" & vbCrLf
ExecuteScript False, sql

sql = "update SSRSInfo set WebUrl = 'http://localhost/Fusion/' where WebUrl is null" & vbCrLf
ExecuteScript False, sql

sql = "update SSRSInfo set SSRSFolderUrl = REPLACE(SSRSFolderUrl,'/Reports','/ReportServer') " & vbCrLf
sql = sql & "where SSRSFolderUrl not like '%/ReportServer%'" & vbCrLf
ExecuteScript False, sql

sql = "update SSRSInfo set SSRSFolderUrl = REPLACE(SSRSFolderUrl,'/Report.aspx?ItemPath=','/ReportViewer.aspx?') " & vbCrLf
sql = sql & "where SSRSFolderUrl not like '%/ReportViewer.aspx%'" & vbCrLf
ExecuteScript False, sql

''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function


Private Function UpdateDatabase89()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 164     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''


' TOOL_MONUM may have been dropped in a prior update.  attempt add it back just in case
If Not ColumnExists("TlnhdTableNew", "TOOL_MONUM") Then
   ExecuteScript False, "ALTER TABLE TlnhdTableNew ADD TOOL_MONUM varchar(20)"
End If

sql = "alter table TlnhdTableNew add TOOL_CODE varchar(6)" & vbCrLf
ExecuteScript False, sql

sql = "alter table TlnhdTableNew add TOOL_UNITNUM varchar(1)" & vbCrLf
ExecuteScript False, sql

sql = "alter table TlnhdTableNew add TOOL_GOVPRIMECONTRACT varchar(20)" & vbCrLf
ExecuteScript False, sql

sql = "alter table TlnhdTableNew add TOOL_CATEGORY varchar(1)" & vbCrLf
ExecuteScript False, sql

sql = "alter table TlnhdTableNew add TOOL_LASTINVDATE datetime null" & vbCrLf
ExecuteScript False, sql

sql = "create table ToolNewCategories" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "  ToolCategory varchar(2) not null primary key clustered" & vbCrLf
sql = sql & ")" & vbCrLf
ExecuteScript False, sql

sql = "insert ToolNewCategories (ToolCategory) values( '' )" & vbCrLf
sql = sql & "insert ToolNewCategories (ToolCategory) values( '1' )" & vbCrLf
sql = sql & "insert ToolNewCategories (ToolCategory) values( '2' )" & vbCrLf
sql = sql & "insert ToolNewCategories (ToolCategory) values( '3' )" & vbCrLf
ExecuteScript False, sql

' change data yes/no/true/false fields to bits
sql = "update TlnhdTableNew set TOOL_GOVOWNED = '1' where TOOL_GOVOWNED = 'True'" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_GOVOWNED = '0' where ISNULL(TOOL_GOVOWNED,'0') <> '1'" & vbCrLf
ExecuteScript False, sql
sql = "alter table tlnhdtablenew alter column TOOL_GOVOWNED bit not null" & vbCrLf
ExecuteScript False, sql
sql = "ALTER TABLE TlnhdTableNew ADD CONSTRAINT DF__TlnhdTabl__TOOL_GOVOWNED DEFAULT 0 FOR TOOL_GOVOWNED" & vbCrLf
ExecuteScript False, sql

sql = "alter table TlnhdTableNew drop constraint DF__TlnhdTabl__TOOL_ITAR" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_ITAR bit not null" & vbCrLf
ExecuteScript False, sql
sql = "ALTER TABLE TlnhdTableNew ADD CONSTRAINT DF__TlnhdTabl__TOOL_ITAR DEFAULT 0 FOR TOOL_ITAR" & vbCrLf
ExecuteScript False, sql



sql = "alter table TlnhdTableNew alter column TOOL_ACCTTO varchar(20)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_BLANKPONUM varchar(20)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_CAVNUM varchar(20)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_CGPONUM varchar(30)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_CGSOPONUM varchar(30)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_CLASS varchar(12)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_CUSTPONUM varchar(30)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_DIM varchar(20)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_DISPSTAT varchar(20)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_DTADDED varchar(12)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_GRID varchar(30)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_HOMEAISLE varchar(30)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_HOMEBLDG varchar(30)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_LOCNUM varchar(30)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_MAKEPN varchar(20)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_MONUM varchar(20)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_NUM varchar(30)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_OWNER varchar(20)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_SHELFNUM varchar(30)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_SN varchar(20)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_SRVSTAT varchar(20)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_STORAGESTAT varchar(20)" & vbCrLf
ExecuteScript False, sql
sql = "alter table TlnhdTableNew alter column TOOL_TOOLMATSTAT varchar(20)" & vbCrLf
ExecuteScript False, sql

sql = "update TlnhdTableNew set TOOL_ACCTTO = RTRIM(TOOL_ACCTTO)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_BLANKPONUM = RTRIM(TOOL_BLANKPONUM)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_CATEGORY = RTRIM(TOOL_CATEGORY)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_CAVNUM = RTRIM(TOOL_CAVNUM)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_CGPONUM = RTRIM(TOOL_CGPONUM)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_CGSOPONUM = RTRIM(TOOL_CGSOPONUM)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_CLASS = RTRIM(TOOL_CLASS)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_CODE = RTRIM(TOOL_CODE)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_COMMENTS = RTRIM(TOOL_COMMENTS)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_CUSTPONUM = RTRIM(TOOL_CUSTPONUM)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_DIM = RTRIM(TOOL_DIM)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_DISPSTAT = RTRIM(TOOL_DISPSTAT)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_DTADDED = RTRIM(TOOL_DTADDED)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_GOVPRIMECONTRACT = RTRIM(TOOL_GOVPRIMECONTRACT)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_GRID = RTRIM(TOOL_GRID)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_HOMEAISLE = RTRIM(TOOL_HOMEAISLE)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_HOMEBLDG = RTRIM(TOOL_HOMEBLDG)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_LOCNUM = RTRIM(TOOL_LOCNUM)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_MAKEPN = RTRIM(TOOL_MAKEPN)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_MONUM = RTRIM(TOOL_MONUM)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_NUM = RTRIM(TOOL_NUM)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_NUMREF = RTRIM(TOOL_NUMREF)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_OWNER = RTRIM(TOOL_OWNER)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_SHELFNUM = RTRIM(TOOL_SHELFNUM)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_SN = RTRIM(TOOL_SN)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_SRVSTAT = RTRIM(TOOL_SRVSTAT)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_STORAGESTAT = RTRIM(TOOL_STORAGESTAT)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_TOOLMATSTAT = RTRIM(TOOL_TOOLMATSTAT)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_UNITNUM = RTRIM(TOOL_UNITNUM)" & vbCrLf
ExecuteScript False, sql
sql = "update TlnhdTableNew set TOOL_WEIGHT = RTRIM(TOOL_WEIGHT)" & vbCrLf
ExecuteScript False, sql

'fix problem with WIP Report
If Not ColumnExists("EsReportWIP", "WIPRUNQTY") Then
   sql = "alter table EsReportWIP add WIPRUNQTY decimal(12,4) null" & vbCrLf
   ExecuteScript False, sql
End If

If Not ColumnExists("EsReportWIP", "WIPRUNPARTIALQTY") Then
   sql = "alter table EsReportWIP add WIPRUNPARTIALQTY decimal(12,4) null" & vbCrLf
   ExecuteScript False, sql
End If

' change tool reference
sql = "alter table TlitTableNew drop constraint DF_TlitTableNew_TOOL_NUM" & vbCrLf
ExecuteScript False, sql

sql = "ALTER TABLE [TlitTableNew] ALTER COLUMN TOOL_NUMREF varchar(30) NOT NULL" & vbCrLf
ExecuteScript False, sql

sql = "update TlitTableNew set TOOL_NUMREF = RTRIM(tool_numref)" & vbCrLf
ExecuteScript False, sql

''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function


Private Function UpdateDatabase90()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 165     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "alter table Preferences add EngineeringLaborRate decimal(12,2) null" & vbCrLf
ExecuteScript False, sql

sql = "alter table EmplTable add PREMENGINEER bit not null constraint DF_EmplTable_PREMENGINEER default 0" & vbCrLf
ExecuteScript False, sql

sql = "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'AddEngrTimeCharge')" & vbCrLf
sql = sql & "  DROP PROCEDURE AddEngrTimeCharge" & vbCrLf
ExecuteScript False, sql

sql = "create procedure AddEngrTimeCharge" & vbCrLf
sql = sql & "  @EmpNo as int," & vbCrLf
sql = sql & "  @Date as datetime," & vbCrLf
sql = sql & "  @MoPartRef as varchar(30)," & vbCrLf
sql = sql & "  @RunNo as int," & vbCrLf
sql = sql & "  @OpNo as int," & vbCrLf
sql = sql & "  @Hours as decimal(12,2)," & vbCrLf
sql = sql & "  @Comment as varchar(1024)," & vbCrLf
sql = sql & "  @journalId as varchar(12)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "/* add a time charge, creating a new time card if required" & vbCrLf
sql = sql & "9/15/17 TEL - created for CASGAS Engineering Time Charges" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if time card does not exist for employee, add it" & vbCrLf
sql = sql & "declare @card char(11)" & vbCrLf
sql = sql & "select @card = TMCARD from TchdTable where TMEMP = @EmpNo and TMDAY = @date" & vbCrLf
sql = sql & "if @card is null " & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "  declare @now as datetime, @nowDays int, @nowMs int" & vbCrLf
sql = sql & "  set @now = getdate()" & vbCrLf
sql = sql & "  set @nowDays = DATEDIFF(DAY,'1/1/1900',cast(@now as date))" & vbCrLf
sql = sql & "  set @nowMs = 1000000.0 *cast(DATEDIFF(MILLISECOND,'1/1/1900',cast(@now as time)) as float)/(3600.0*24*1000)" & vbCrLf
sql = sql & "  set @card = cast(@nowDays as varchar(5)) + cast(@nowMs as varchar(6))" & vbCrLf
sql = sql & "  insert into TchdTable (TMCARD,TMEMP,TMDAY) values (@card, @EmpNo, @Date)" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- calculate fields needed for time charge" & vbCrLf
sql = sql & "declare @startTime smalldatetime, @stopTime smalldatetime" & vbCrLf
sql = sql & "set @startTime = DATEADD(dd, DATEDIFF(dd, 0, @Date), 0)    -- truncate time portion" & vbCrLf
sql = sql & "set @stopTime = DATEADD(MINUTE, @Hours * 60, @startTime)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- construct hh:mma/p version of times" & vbCrLf
sql = sql & "declare @start varchar(6), @stop varchar(6)" & vbCrLf
sql = sql & "set @start = REPLACE(SUBSTRING(CONVERT(VARCHAR(20),@startTime,0),13,6),' ','0')" & vbCrLf
sql = sql & "set @stop = REPLACE(SUBSTRING(CONVERT(VARCHAR(20),@stopTime,0),13,6),' ','0')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- calculate # minutes" & vbCrLf
sql = sql & "declare @time datetime" & vbCrLf
sql = sql & "set @time = dateadd(minute,datediff(MINUTE,@startTime, @stopTime),'1/1/1900')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get regular timecode" & vbCrLf
sql = sql & "declare @timecode varchar(2)" & vbCrLf
sql = sql & "select @timecode = TYPECODE from TmcdTable where typetype = 'R'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get engineering rate" & vbCrLf
sql = sql & "declare @rate decimal(10,2)" & vbCrLf
sql = sql & "select @rate = EngineeringLaborRate from Preferences " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get employee account number" & vbCrLf
sql = sql & "declare @acct varchar(12)" & vbCrLf
sql = sql & "select @acct = EmplTable.PREMACCTS from EmplTable where PREMNUMBER = @EmpNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get shop and wc for operation" & vbCrLf
sql = sql & "declare @shop varchar(12), @wc varchar(12)" & vbCrLf
sql = sql & "select @shop = OPSHOP, @wc = OPCENTER " & vbCrLf
sql = sql & "from RnopTable where opref = @MoPartRef and oprun = @RunNo and OPNO = @OpNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get journal id" & vbCrLf
sql = sql & "--declare @journalId varchar(12)" & vbCrLf
sql = sql & "--set @journalId = dbo.fnGetOpenJournalID('TJ', @Date)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now insert the new time charge" & vbCrLf
sql = sql & "INSERT INTO TcitTable (TCCARD,TCEMP,TCSTART,TCSTOP,TCSTARTTIME,TCSTOPTIME," & vbCrLf
sql = sql & "      TCHOURS,TCTIME,TCCODE,TCRATE,TCOHRATE,TCRATENO,TCACCT,TCACCOUNT," & vbCrLf
sql = sql & "      TCSHOP,TCWC,TCPAYTYPE,TCSURUN,TCYIELD,TCPARTREF,TCRUNNO," & vbCrLf
sql = sql & "      TCOPNO,TCSORT,TCOHFIXED,TCGLJOURNAL,TCGLREF,TCSOURCE," & vbCrLf
sql = sql & "      TCMULTIJOB,TCACCEPT,TCREJECT,TCSCRAP,TCCOMMENTS)" & vbCrLf
sql = sql & "values( @card,@EmpNo, @start,@stop,@startTime, @stopTime," & vbCrLf
sql = sql & "      @Hours,@time,@timecode,@rate,@rate,1,@acct,@acct," & vbCrLf
sql = sql & "      @shop,@wc,0,'I',0,@MoPartRef,@RunNo," & vbCrLf
sql = sql & "    @OpNo,0,0,@journalId,0,'Engr'," & vbCrLf
sql = sql & "      0,0,0,0,@Comment)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now roll up totals for this timecard" & vbCrLf
sql = sql & "EXECUTE UpdateTimeCardTotals @EmpNo, @Date" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function


Private Function UpdateDatabase91()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 166     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

' convert sheet inventory to use closed date for lot items instead of the LOIINACTIVE flag.
' this will allow you to run a 7/1 inventory on some other date

' convert inactive to datetime
sql = "alter table LoitTable add LOICLOSED datetime" & vbCrLf
ExecuteScript False, sql

' set closed date = date of child activity
sql = "update parent set LOICLOSED = cast(CONVERT(VARCHAR(10),child.LOIADATE,101) as datetime)" & vbCrLf
sql = sql & "from LoitTable parent join LoitTable child on child.LOINUMBER = parent.LOINUMBER and child.LOIPARENTREC = parent.LOIRECORD" & vbCrLf
sql = sql & "and parent.LOIINACTIVE = 1      -- 314" & vbCrLf
ExecuteScript False, sql

' for any remaining inactive items, set closed date = date of the activity
sql = "update LoitTable set LOICLOSED = cast(CONVERT(VARCHAR(10),LOIADATE,101) as datetime) where LOIINACTIVE = 1 and LOICLOSED is null" & vbCrLf
ExecuteScript False, sql

sql = "alter table LoitTable drop column LOIINACTIVE" & vbCrLf
ExecuteScript False, sql

sql = "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'SheetPick')" & vbCrLf
sql = sql & "  DROP PROCEDURE SheetPick" & vbCrLf
ExecuteScript False, sql

sql = "create procedure SheetPick" & vbCrLf
sql = sql & "   @UserLotNo varchar(40),     -- lotuserlotid" & vbCrLf
sql = sql & "   @User varchar(4)," & vbCrLf
sql = sql & "   @SO int," & vbCrLf
sql = sql & "   @PickDate datetime" & vbCrLf
sql = sql & "   " & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "Pick a sheet" & vbCrLf
sql = sql & "exec SheetPick '029373-1-A','MGR',222222" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- remove time from pick date" & vbCrLf
sql = sql & "set @PickDate = cast(convert(varchar(10),@PickDate,101)as datetime)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- find all rectangles in lot" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "create table #rect" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "   id int identity," & vbCrLf
sql = sql & "   ParentRecord int," & vbCrLf
sql = sql & "   Height decimal(12,4)," & vbCrLf
sql = sql & "   Length decimal(12,4)," & vbCrLf
sql = sql & "   NewRecord int," & vbCrLf
sql = sql & "   Qty decimal(12,4)" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @LotNo char(15)" & vbCrLf
sql = sql & "select @LotNo = LOTNUMBER from LohdTable where LOTUSERLOTID = @UserLotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert #rect (ParentRecord,Height,Length,NewRecord,Qty)" & vbCrLf
sql = sql & "select li.LOIRECORD,li.LOIHEIGHT,li.LOILENGTH,0,cast(li.LOIHEIGHT*li.LOILENGTH as decimal(12,4)) from LoitTable li" & vbCrLf
sql = sql & "   join LohdTable lh on lh.LOTNUMBER = li.LOINUMBER" & vbCrLf
sql = sql & "   where lh.LOTNUMBER = @LotNo and LOIQUANTITY > 0 " & vbCrLf
sql = sql & "   and li.LOICLOSED is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- assign new recordnumbers" & vbCrLf
sql = sql & "declare @max int" & vbCrLf
sql = sql & "select @max = max(LOIRECORD) from LoitTable where LOINUMBER = @LotNo" & vbCrLf
sql = sql & "update #rect set NewRecord = @max + id" & vbCrLf
sql = sql & "   " & vbCrLf
sql = sql & "declare @partRef varchar(30), @currentTime datetime, @uom char(2), @type int" & vbCrLf
sql = sql & "declare @unitCost decimal(12,4), @nextINNUMBER int, @totalQty decimal(12,4)" & vbCrLf
sql = sql & "select @partRef = LOTPARTREF, @unitCost = LOTUNITCOST, @totalQty = LOTREMAININGQTY from LohdTable where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "select @uom = PAUNITS from PartTable where PARTREF = @partRef" & vbCrLf
sql = sql & "set @currentTime = getdate()" & vbCrLf
sql = sql & "select @nextINNUMBER = max(INNUMBER) + 1 from InvaTable" & vbCrLf
sql = sql & "set @type = 19     -- manual adjustment" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- deactivate rectangles being picked so they won't show up again" & vbCrLf
sql = sql & "Update li " & vbCrLf
sql = sql & "set li.LOICLOSED = @PickDate" & vbCrLf
sql = sql & "from LoitTable li " & vbCrLf
sql = sql & "join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD = r.ParentRecord" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new LoitTable record to zero lot quantity" & vbCrLf
sql = sql & "INSERT INTO [dbo].[LoitTable]" & vbCrLf
sql = sql & "           ([LOINUMBER]" & vbCrLf
sql = sql & "           ,[LOIRECORD]" & vbCrLf
sql = sql & "           ,[LOITYPE]" & vbCrLf
sql = sql & "           ,[LOIPARTREF]" & vbCrLf
sql = sql & "           ,[LOIADATE]" & vbCrLf
sql = sql & "           ,[LOIPDATE]" & vbCrLf
sql = sql & "           ,[LOIQUANTITY]" & vbCrLf
sql = sql & "           ,[LOIMOPARTREF]" & vbCrLf
sql = sql & "           ,[LOIMORUNNO]" & vbCrLf
sql = sql & "           ,[LOIPONUMBER]" & vbCrLf
sql = sql & "           ,[LOIPOITEM]" & vbCrLf
sql = sql & "           ,[LOIPOREV]" & vbCrLf
sql = sql & "           ,[LOIPSNUMBER]" & vbCrLf
sql = sql & "           ,[LOIPSITEM]" & vbCrLf
sql = sql & "           ,[LOICUSTINVNO]" & vbCrLf
sql = sql & "           ,[LOICUST]" & vbCrLf
sql = sql & "           ,[LOIVENDINVNO]" & vbCrLf
sql = sql & "           ,[LOIVENDOR]" & vbCrLf
sql = sql & "           ,[LOIACTIVITY]" & vbCrLf
sql = sql & "           ,[LOICOMMENT]" & vbCrLf
sql = sql & "           ,[LOIUNITS]" & vbCrLf
sql = sql & "           ,[LOIMOPKCANCEL]" & vbCrLf
sql = sql & "           ,[LOIHEIGHT]" & vbCrLf
sql = sql & "           ,[LOILENGTH]" & vbCrLf
sql = sql & "           ,[LOIUSER]" & vbCrLf
sql = sql & "          ,LOIPARENTREC" & vbCrLf
sql = sql & "           ,[LOISONUMBER]" & vbCrLf
sql = sql & "          ,LOISHEETACTTYPE" & vbCrLf
sql = sql & "          )" & vbCrLf
sql = sql & "     SELECT" & vbCrLf
sql = sql & "           @LotNo" & vbCrLf
sql = sql & "           ,r.NewRecord" & vbCrLf
sql = sql & "           ,@type              -- manual adjustment" & vbCrLf
sql = sql & "           ,@partRef" & vbCrLf
sql = sql & "           ,@PickDate" & vbCrLf
sql = sql & "           ,@PickDate" & vbCrLf
sql = sql & "           ,-r.Qty" & vbCrLf
sql = sql & "           ,null   --LOIMOPARTREF, char(30),>" & vbCrLf
sql = sql & "           ,null   --<LOIMORUNNO, int,>" & vbCrLf
sql = sql & "           ,null   --<LOIPONUMBER, int,>" & vbCrLf
sql = sql & "           ,null   --<LOIPOITEM, smallint,>" & vbCrLf
sql = sql & "           ,null   --<LOIPOREV, char(2),>" & vbCrLf
sql = sql & "           ,null   --<LOIPSNUMBER, char(8),>" & vbCrLf
sql = sql & "           ,null   --<LOIPSITEM, smallint,>" & vbCrLf
sql = sql & "           ,null   --<LOICUSTINVNO, int,>" & vbCrLf
sql = sql & "           ,null   --<LOICUST, char(10),>" & vbCrLf
sql = sql & "           ,null   --<LOIVENDINVNO, char(20),>" & vbCrLf
sql = sql & "           ,null   --<LOIVENDOR, char(10),>" & vbCrLf
sql = sql & "           ,null   --<LOIACTIVITY, int,> points to InvaTable.INNO when IA is created" & vbCrLf
sql = sql & "           ,'sheet pick'   --<LOICOMMENT, varchar(40),>" & vbCrLf
sql = sql & "           ,@uom   --<LOIUNITS, char(2),>" & vbCrLf
sql = sql & "           ,null   --<LOIMOPKCANCEL, smallint,>" & vbCrLf
sql = sql & "           ,r.HEIGHT   --<LOIHEIGHT, decimal(12,4),>" & vbCrLf
sql = sql & "           ,r.LENGTH   --<LOILENGTH, decimal(12,4),>" & vbCrLf
sql = sql & "           ,@User          --<LOIUSER, varchar(4),>" & vbCrLf
sql = sql & "          ,r.ParentRecord  --LOIPARENTREC" & vbCrLf
sql = sql & "           ,@SO                --<LOISONUMBER, int,>" & vbCrLf
sql = sql & "          ,'PK'            --LOISHEETACTTYPE" & vbCrLf
sql = sql & "       from #rect r" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- set lot quantity = 0" & vbCrLf
sql = sql & "update LohdTable set LOTREMAININGQTY = 0 where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new IA record to reduce part quantity" & vbCrLf
sql = sql & "INSERT INTO [dbo].[InvaTable]" & vbCrLf
sql = sql & "           ([INTYPE]" & vbCrLf
sql = sql & "           ,[INPART]" & vbCrLf
sql = sql & "           ,[INREF1]" & vbCrLf
sql = sql & "           ,[INREF2]" & vbCrLf
sql = sql & "           ,[INPDATE]" & vbCrLf
sql = sql & "           ,[INADATE]" & vbCrLf
sql = sql & "           ,[INPQTY]" & vbCrLf
sql = sql & "           ,[INAQTY]" & vbCrLf
sql = sql & "           ,[INAMT]" & vbCrLf
sql = sql & "           ,[INTOTMATL]" & vbCrLf
sql = sql & "           ,[INTOTLABOR]" & vbCrLf
sql = sql & "           ,[INTOTEXP]" & vbCrLf
sql = sql & "           ,[INTOTOH]" & vbCrLf
sql = sql & "           ,[INTOTHRS]" & vbCrLf
sql = sql & "           ,[INCREDITACCT]" & vbCrLf
sql = sql & "           ,[INDEBITACCT]" & vbCrLf
sql = sql & "           ,[INGLJOURNAL]" & vbCrLf
sql = sql & "           ,[INGLPOSTED]" & vbCrLf
sql = sql & "           ,[INGLDATE]" & vbCrLf
sql = sql & "           ,[INMOPART]" & vbCrLf
sql = sql & "           ,[INMORUN]" & vbCrLf
sql = sql & "           ,[INSONUMBER]" & vbCrLf
sql = sql & "           ,[INSOITEM]" & vbCrLf
sql = sql & "           ,[INSOREV]" & vbCrLf
sql = sql & "           ,[INPONUMBER]" & vbCrLf
sql = sql & "           ,[INPORELEASE]" & vbCrLf
sql = sql & "           ,[INPOITEM]" & vbCrLf
sql = sql & "           ,[INPOREV]" & vbCrLf
sql = sql & "           ,[INPSNUMBER]" & vbCrLf
sql = sql & "           ,[INPSITEM]" & vbCrLf
sql = sql & "           ,[INWIPLABACCT]" & vbCrLf
sql = sql & "           ,[INWIPMATACCT]" & vbCrLf
sql = sql & "           ,[INWIPOHDACCT]" & vbCrLf
sql = sql & "           ,[INWIPEXPACCT]" & vbCrLf
sql = sql & "           ,[INNUMBER]" & vbCrLf
sql = sql & "           ,[INLOTNUMBER]" & vbCrLf
sql = sql & "           ,[INUSER]" & vbCrLf
sql = sql & "           ,[INUNITS]" & vbCrLf
sql = sql & "           ,[INDRLABACCT]" & vbCrLf
sql = sql & "           ,[INDRMATACCT]" & vbCrLf
sql = sql & "           ,[INDREXPACCT]" & vbCrLf
sql = sql & "           ,[INDROHDACCT]" & vbCrLf
sql = sql & "           ,[INCRLABACCT]" & vbCrLf
sql = sql & "           ,[INCRMATACCT]" & vbCrLf
sql = sql & "           ,[INCREXPACCT]" & vbCrLf
sql = sql & "           ,[INCROHDACCT]" & vbCrLf
sql = sql & "           ,[INLOTTRACK]" & vbCrLf
sql = sql & "           ,[INUSEACTUALCOST]" & vbCrLf
sql = sql & "           ,[INCOSTEDBY]" & vbCrLf
sql = sql & "           ,[INMAINTCOSTED])" & vbCrLf
sql = sql & "     select" & vbCrLf
sql = sql & "           @type               --<INTYPE, int,>" & vbCrLf
sql = sql & "           ,@partRef           --char(30),>" & vbCrLf
sql = sql & "           ,'Manual Adjustment'    --<INREF1, char(20),>" & vbCrLf
sql = sql & "           ,'Sheet Inventory'  --<INREF2, char(40),>" & vbCrLf
sql = sql & "           ,@PickDate       --<INPDATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,@PickDate       --<INADATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,-r.Qty         --<INPQTY, decimal(12,4),>" & vbCrLf
sql = sql & "           ,-r.Qty         --<INAQTY, decimal(12,4),>" & vbCrLf
sql = sql & "           ,@unitCost          -- <INAMT, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTMATL, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTLABOR, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTEXP, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTOH, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTHRS, decimal(12,4),>" & vbCrLf
sql = sql & "           ,''                 --<INCREDITACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDEBITACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INGLJOURNAL, char(12),>" & vbCrLf
sql = sql & "           ,0                  --<INGLPOSTED, tinyint,>" & vbCrLf
sql = sql & "           ,null               --<INGLDATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,''                 --<INMOPART, char(30),>" & vbCrLf
sql = sql & "           ,0                  --<INMORUN, int,>" & vbCrLf
sql = sql & "           ,0                  --<INSONUMBER, int,>" & vbCrLf
sql = sql & "           ,0                  --<INSOITEM, int,>" & vbCrLf
sql = sql & "           ,''                 --<INSOREV, char(2),>" & vbCrLf
sql = sql & "           ,0                  --<INPONUMBER, int,>" & vbCrLf
sql = sql & "           ,0                  --<INPORELEASE, smallint,>" & vbCrLf
sql = sql & "           ,0                  --<INPOITEM, smallint,>" & vbCrLf
sql = sql & "           ,''                 --<INPOREV, char(2),>" & vbCrLf
sql = sql & "           ,''                 --<INPSNUMBER, char(8),>" & vbCrLf
sql = sql & "           ,0                  --<INPSITEM, smallint,>" & vbCrLf
sql = sql & "           ,''                 --<INWIPLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPOHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPEXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,@nextINNUMBER      --<INNUMBER, int,>" & vbCrLf
sql = sql & "           ,@LotNo             --<INLOTNUMBER, char(15),>" & vbCrLf
sql = sql & "           ,@User              --<INUSER, char(4),>" & vbCrLf
sql = sql & "           ,@uom               --<INUNITS, char(2),>" & vbCrLf
sql = sql & "           ,''                 --<INDRLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDRMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDREXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDROHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCRLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCRMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCREXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCROHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,null               --<INLOTTRACK, bit,>" & vbCrLf
sql = sql & "           ,null               --<INUSEACTUALCOST, bit,>" & vbCrLf
sql = sql & "           ,null               --<INCOSTEDBY, char(4),>" & vbCrLf
sql = sql & "           ,0                  --<INMAINTCOSTED, int,>)" & vbCrLf
sql = sql & "       from #rect r" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now update part QOH" & vbCrLf
sql = sql & "update PartTable set PAQOH = PAQOH - @totalQty where PARTREF = @partRef" & vbCrLf
sql = sql & "commit tran" & vbCrLf
ExecuteScript False, sql

sql = "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'SheetRestock')" & vbCrLf
sql = sql & "  DROP PROCEDURE SheetRestock" & vbCrLf
ExecuteScript False, sql

sql = "create procedure SheetRestock" & vbCrLf
sql = sql & "   @UserLotNo varchar(40)," & vbCrLf
sql = sql & "   @User varchar(4)," & vbCrLf
sql = sql & "   @Comments varchar(2048)," & vbCrLf
sql = sql & "   @Location varchar(4)," & vbCrLf
sql = sql & "   @Params varchar(2000)   -- LOIRECORD,NEWHT,NEWLEN,... repeat (include comma at end)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "restock sheet LOI records" & vbCrLf
sql = sql & "exec SheetRestock '029373-1-A', 'MGR', '2,36,96,0,12,72,'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "create table #rect" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "   ParentRecord int," & vbCrLf
sql = sql & "   NewHeight decimal(12,4)," & vbCrLf
sql = sql & "   NewLength decimal(12,4)," & vbCrLf
sql = sql & "   NewRecord int," & vbCrLf
sql = sql & "   Qty decimal(12,4)," & vbCrLf
sql = sql & "   NewIANumber int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @nextRecord int, @nextIANumber int" & vbCrLf
sql = sql & "declare @LotNo char(15), @id int, @SO int" & vbCrLf
sql = sql & "select @LotNo = LOTNUMBER from LohdTable where LOTUSERLOTID = @UserLotNo" & vbCrLf
sql = sql & "select @nextRecord = max(LOIRECORD) + 1 , @SO = max(LOISONUMBER) from LoitTable where LOINUMBER = @LotNo" & vbCrLf
sql = sql & "select @nextIANumber = max(INNUMBER) + 1 from InvaTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- first update the lot header" & vbCrLf
sql = sql & "update LohdTable set LOTCOMMENTS = @Comments, LOTLOCATION = @Location where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- extract parameters" & vbCrLf
sql = sql & "DECLARE @start INT, @end INT" & vbCrLf
sql = sql & "declare @stringId varchar(10), @stringHt varchar(10), @stringLen varchar(10)" & vbCrLf
sql = sql & "SELECT @start = 1, @end = CHARINDEX(',', @Params) " & vbCrLf
sql = sql & "WHILE @start < LEN(@Params) + 1 BEGIN " & vbCrLf
sql = sql & "    IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    set @stringId = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "    SET @start = @end + 1 " & vbCrLf
sql = sql & "    SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    set @stringHt = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "    SET @start = @end + 1 " & vbCrLf
sql = sql & "    SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    set @stringLen = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "    SET @start = @end + 1 " & vbCrLf
sql = sql & "    SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "   " & vbCrLf
sql = sql & "   declare @ht decimal(12,4), @len decimal(12,4)" & vbCrLf
sql = sql & "   set @ht = cast(@stringHt as decimal(12,4))" & vbCrLf
sql = sql & "   set @len = cast(@stringLen as decimal(12,4))" & vbCrLf
sql = sql & "   insert into #rect (ParentRecord, NewHeight, NewLength, NewRecord, NewIANumber)" & vbCrLf
sql = sql & "   values (cast(@stringId as int), @ht, @len," & vbCrLf
sql = sql & "       case when @ht*@len = 0 then 0 else @nextRecord end," & vbCrLf
sql = sql & "       case when @ht*@len = 0 then 0 else @nextIANumber end)" & vbCrLf
sql = sql & "   if (@ht*@len) <> 0 " & vbCrLf
sql = sql & "   begin" & vbCrLf
sql = sql & "       set @nextRecord = @nextRecord + 1" & vbCrLf
sql = sql & "       set @nextIANumber = @nextIANumber + 1" & vbCrLf
sql = sql & "   end" & vbCrLf
sql = sql & "END " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #rect set Qty = NewHeight * NewLength" & vbCrLf
sql = sql & "update #rect set ParentRecord = isnull((select top 1 ParentRecord from #rect where ParentRecord <> 0 ),0)" & vbCrLf
sql = sql & "where ParentRecord = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from #rect" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @partRef varchar(30), @currentTime datetime, @uom char(2), @type int" & vbCrLf
sql = sql & "declare @unitCost decimal(12,4), @nextINNUMBER int, @totalQty decimal(12,4)" & vbCrLf
sql = sql & "select @partRef = LOTPARTREF, @unitCost = LOTUNITCOST, @totalQty = LOTREMAININGQTY from LohdTable where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "select @uom = PAUNITS from PartTable where PARTREF = @partRef" & vbCrLf
sql = sql & "set @currentTime = getdate()" & vbCrLf
sql = sql & "select @nextINNUMBER = max(INNUMBER) + 1 from InvaTable" & vbCrLf
sql = sql & "set @type = 19     -- manual adjustment" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- deactivate picked rectangles so they won't show up again" & vbCrLf
sql = sql & "Update li " & vbCrLf
sql = sql & "set li.LOICLOSED = cast(convert(varchar(10),@currentTime,101) as datetime)" & vbCrLf
sql = sql & "from LoitTable li " & vbCrLf
sql = sql & "join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD = r.ParentRecord" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- remove rectangles with no quantity remaining.  These items will not be restocked." & vbCrLf
sql = sql & "delete from #rect where Qty = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new lot items" & vbCrLf
sql = sql & "INSERT INTO [dbo].[LoitTable]" & vbCrLf
sql = sql & "           ([LOINUMBER]" & vbCrLf
sql = sql & "           ,[LOIRECORD]" & vbCrLf
sql = sql & "           ,[LOITYPE]" & vbCrLf
sql = sql & "           ,[LOIPARTREF]" & vbCrLf
sql = sql & "           ,[LOIADATE]" & vbCrLf
sql = sql & "           ,[LOIPDATE]" & vbCrLf
sql = sql & "           ,[LOIQUANTITY]" & vbCrLf
sql = sql & "           ,[LOIMOPARTREF]" & vbCrLf
sql = sql & "           ,[LOIMORUNNO]" & vbCrLf
sql = sql & "           ,[LOIPONUMBER]" & vbCrLf
sql = sql & "           ,[LOIPOITEM]" & vbCrLf
sql = sql & "           ,[LOIPOREV]" & vbCrLf
sql = sql & "           ,[LOIPSNUMBER]" & vbCrLf
sql = sql & "           ,[LOIPSITEM]" & vbCrLf
sql = sql & "           ,[LOICUSTINVNO]" & vbCrLf
sql = sql & "           ,[LOICUST]" & vbCrLf
sql = sql & "           ,[LOIVENDINVNO]" & vbCrLf
sql = sql & "           ,[LOIVENDOR]" & vbCrLf
sql = sql & "           ,[LOIACTIVITY]" & vbCrLf
sql = sql & "           ,[LOICOMMENT]" & vbCrLf
sql = sql & "           ,[LOIUNITS]" & vbCrLf
sql = sql & "           ,[LOIMOPKCANCEL]" & vbCrLf
sql = sql & "           ,[LOIHEIGHT]" & vbCrLf
sql = sql & "           ,[LOILENGTH]" & vbCrLf
sql = sql & "           ,[LOIUSER]" & vbCrLf
sql = sql & "          ,LOIPARENTREC" & vbCrLf
sql = sql & "           ,[LOISONUMBER]" & vbCrLf
sql = sql & "          ,LOISHEETACTTYPE" & vbCrLf
sql = sql & "          )" & vbCrLf
sql = sql & "     SELECT" & vbCrLf
sql = sql & "           @LotNo" & vbCrLf
sql = sql & "           ,r.NewRecord" & vbCrLf
sql = sql & "           ,@type              -- manual adjustment" & vbCrLf
sql = sql & "           ,@partRef" & vbCrLf
sql = sql & "           ,@currentTime" & vbCrLf
sql = sql & "           ,@currentTime" & vbCrLf
sql = sql & "           ,r.Qty" & vbCrLf
sql = sql & "           ,null   --LOIMOPARTREF, char(30),>" & vbCrLf
sql = sql & "           ,null   --<LOIMORUNNO, int,>" & vbCrLf
sql = sql & "           ,null   --<LOIPONUMBER, int,>" & vbCrLf
sql = sql & "           ,null   --<LOIPOITEM, smallint,>" & vbCrLf
sql = sql & "           ,null   --<LOIPOREV, char(2),>" & vbCrLf
sql = sql & "           ,null   --<LOIPSNUMBER, char(8),>" & vbCrLf
sql = sql & "           ,null   --<LOIPSITEM, smallint,>" & vbCrLf
sql = sql & "           ,null   --<LOICUSTINVNO, int,>" & vbCrLf
sql = sql & "           ,null   --<LOICUST, char(10),>" & vbCrLf
sql = sql & "           ,null   --<LOIVENDINVNO, char(20),>" & vbCrLf
sql = sql & "           ,null   --<LOIVENDOR, char(10),>" & vbCrLf
sql = sql & "           ,null   --<LOIACTIVITY, int,> points to InvaTable.INNO when IA is created" & vbCrLf
sql = sql & "           ,'sheet restock'    --<LOICOMMENT, varchar(40),>" & vbCrLf
sql = sql & "           ,@uom   --<LOIUNITS, char(2),>" & vbCrLf
sql = sql & "           ,null   --<LOIMOPKCANCEL, smallint,>" & vbCrLf
sql = sql & "           ,r.NewHeight    --<LOIHEIGHT, decimal(12,4),>" & vbCrLf
sql = sql & "           ,r.NewLength    --<LOILENGTH, decimal(12,4),>" & vbCrLf
sql = sql & "           ,@User      --<LOIUSER, varchar(4),>" & vbCrLf
sql = sql & "          ,r.ParentRecord  --LOIPICKEDFROMREC" & vbCrLf
sql = sql & "           ,@SO        --<LOISONUMBER, int,>" & vbCrLf
sql = sql & "          ,'RS'    --LOISHEETACTTYPE" & vbCrLf
sql = sql & "       from #rect r" & vbCrLf
sql = sql & "       where r.Qty <> 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select distinct li.* from LoitTable li " & vbCrLf
sql = sql & "--join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD > 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- set LOTREMAININGQTY and remove reservation" & vbCrLf
sql = sql & "declare @sum decimal(12,4)" & vbCrLf
sql = sql & "select @sum = sum(Qty) from #rect" & vbCrLf
sql = sql & "update LohdTable " & vbCrLf
sql = sql & "   set LOTREMAININGQTY = LOTREMAININGQTY + @sum," & vbCrLf
sql = sql & "   LOTRESERVEDBY = NULL," & vbCrLf
sql = sql & "   LOTRESERVEDON = NULL" & vbCrLf
sql = sql & "   where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new ia items" & vbCrLf
sql = sql & "INSERT INTO [dbo].[InvaTable]" & vbCrLf
sql = sql & "           ([INTYPE]" & vbCrLf
sql = sql & "           ,[INPART]" & vbCrLf
sql = sql & "           ,[INREF1]" & vbCrLf
sql = sql & "           ,[INREF2]" & vbCrLf
sql = sql & "           ,[INPDATE]" & vbCrLf
sql = sql & "           ,[INADATE]" & vbCrLf
sql = sql & "           ,[INPQTY]" & vbCrLf
sql = sql & "           ,[INAQTY]" & vbCrLf
sql = sql & "           ,[INAMT]" & vbCrLf
sql = sql & "           ,[INTOTMATL]" & vbCrLf
sql = sql & "           ,[INTOTLABOR]" & vbCrLf
sql = sql & "           ,[INTOTEXP]" & vbCrLf
sql = sql & "           ,[INTOTOH]" & vbCrLf
sql = sql & "           ,[INTOTHRS]" & vbCrLf
sql = sql & "           ,[INCREDITACCT]" & vbCrLf
sql = sql & "           ,[INDEBITACCT]" & vbCrLf
sql = sql & "           ,[INGLJOURNAL]" & vbCrLf
sql = sql & "           ,[INGLPOSTED]" & vbCrLf
sql = sql & "           ,[INGLDATE]" & vbCrLf
sql = sql & "           ,[INMOPART]" & vbCrLf
sql = sql & "           ,[INMORUN]" & vbCrLf
sql = sql & "           ,[INSONUMBER]" & vbCrLf
sql = sql & "           ,[INSOITEM]" & vbCrLf
sql = sql & "           ,[INSOREV]" & vbCrLf
sql = sql & "           ,[INPONUMBER]" & vbCrLf
sql = sql & "           ,[INPORELEASE]" & vbCrLf
sql = sql & "           ,[INPOITEM]" & vbCrLf
sql = sql & "           ,[INPOREV]" & vbCrLf
sql = sql & "           ,[INPSNUMBER]" & vbCrLf
sql = sql & "           ,[INPSITEM]" & vbCrLf
sql = sql & "           ,[INWIPLABACCT]" & vbCrLf
sql = sql & "           ,[INWIPMATACCT]" & vbCrLf
sql = sql & "           ,[INWIPOHDACCT]" & vbCrLf
sql = sql & "           ,[INWIPEXPACCT]" & vbCrLf
sql = sql & "           ,[INNUMBER]" & vbCrLf
sql = sql & "           ,[INLOTNUMBER]" & vbCrLf
sql = sql & "           ,[INUSER]" & vbCrLf
sql = sql & "           ,[INUNITS]" & vbCrLf
sql = sql & "           ,[INDRLABACCT]" & vbCrLf
sql = sql & "           ,[INDRMATACCT]" & vbCrLf
sql = sql & "           ,[INDREXPACCT]" & vbCrLf
sql = sql & "           ,[INDROHDACCT]" & vbCrLf
sql = sql & "           ,[INCRLABACCT]" & vbCrLf
sql = sql & "           ,[INCRMATACCT]" & vbCrLf
sql = sql & "           ,[INCREXPACCT]" & vbCrLf
sql = sql & "           ,[INCROHDACCT]" & vbCrLf
sql = sql & "           ,[INLOTTRACK]" & vbCrLf
sql = sql & "           ,[INUSEACTUALCOST]" & vbCrLf
sql = sql & "           ,[INCOSTEDBY]" & vbCrLf
sql = sql & "           ,[INMAINTCOSTED])" & vbCrLf
sql = sql & "     select" & vbCrLf
sql = sql & "           @type               --<INTYPE, int,>" & vbCrLf
sql = sql & "           ,@partRef           --char(30),>" & vbCrLf
sql = sql & "           ,'Manual Adjustment'    --<INREF1, char(20),>" & vbCrLf
sql = sql & "           ,'Sheet Inventory'  --<INREF2, char(40),>" & vbCrLf
sql = sql & "           ,@currentTime       --<INPDATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,@currentTime       --<INADATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,r.Qty          --<INPQTY, decimal(12,4),>" & vbCrLf
sql = sql & "           ,r.Qty          --<INAQTY, decimal(12,4),>" & vbCrLf
sql = sql & "           ,@unitCost          -- <INAMT, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTMATL, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTLABOR, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTEXP, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTOH, decimal(12,4),>" & vbCrLf
sql = sql & "           ,0.0                    --<INTOTHRS, decimal(12,4),>" & vbCrLf
sql = sql & "           ,''                 --<INCREDITACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDEBITACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INGLJOURNAL, char(12),>" & vbCrLf
sql = sql & "           ,0                  --<INGLPOSTED, tinyint,>" & vbCrLf
sql = sql & "           ,null               --<INGLDATE, smalldatetime,>" & vbCrLf
sql = sql & "           ,''                 --<INMOPART, char(30),>" & vbCrLf
sql = sql & "           ,0                  --<INMORUN, int,>" & vbCrLf
sql = sql & "           ,0                  --<INSONUMBER, int,>" & vbCrLf
sql = sql & "           ,0                  --<INSOITEM, int,>" & vbCrLf
sql = sql & "           ,''                 --<INSOREV, char(2),>" & vbCrLf
sql = sql & "           ,0                  --<INPONUMBER, int,>" & vbCrLf
sql = sql & "           ,0                  --<INPORELEASE, smallint,>" & vbCrLf
sql = sql & "           ,0                  --<INPOITEM, smallint,>" & vbCrLf
sql = sql & "           ,''                 --<INPOREV, char(2),>" & vbCrLf
sql = sql & "           ,''                 --<INPSNUMBER, char(8),>" & vbCrLf
sql = sql & "           ,0                  --<INPSITEM, smallint,>" & vbCrLf
sql = sql & "           ,''                 --<INWIPLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPOHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INWIPEXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,r.NewIANumber      --<INNUMBER, int,>" & vbCrLf
sql = sql & "           ,@LotNo             --<INLOTNUMBER, char(15),>" & vbCrLf
sql = sql & "           ,@User              --<INUSER, char(4),>" & vbCrLf
sql = sql & "           ,@uom               --<INUNITS, char(2),>" & vbCrLf
sql = sql & "           ,''                 --<INDRLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDRMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDREXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INDROHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCRLABACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCRMATACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCREXPACCT, char(12),>" & vbCrLf
sql = sql & "           ,''                 --<INCROHDACCT, char(12),>" & vbCrLf
sql = sql & "           ,null               --<INLOTTRACK, bit,>" & vbCrLf
sql = sql & "           ,null               --<INUSEACTUALCOST, bit,>" & vbCrLf
sql = sql & "           ,null               --<INCOSTEDBY, char(4),>" & vbCrLf
sql = sql & "           ,0                  --<INMAINTCOSTED, int,>)" & vbCrLf
sql = sql & "       from #rect r" & vbCrLf
sql = sql & "       where r.Qty <> 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now update part QOH" & vbCrLf
sql = sql & "update PartTable set PAQOH = PAQOH + @totalQty where PARTREF = @partRef" & vbCrLf
sql = sql & "commit tran" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function

Private Function UpdateDatabase92()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 167     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'AddOrUpdateColumn')" & vbCrLf
sql = sql & "  DROP PROCEDURE AddOrUpdateColumn" & vbCrLf
ExecuteScript False, sql

sql = "create procedure dbo.AddOrUpdateColumn " & vbCrLf
sql = sql & "  @Table varchar(40)," & vbCrLf
sql = sql & "  @Column varchar(40)," & vbCrLf
sql = sql & "  @Properties varchar(80)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* create or update a column" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "  AddOrUpdateColumn 'ComnTable', 'DenyLoginIfPriorOpOpen', 'tinyint null'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "  declare @sql varchar(120)" & vbCrLf
sql = sql & "  if exists (select COLUMN_NAME from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = @Table " & vbCrLf
sql = sql & "  and COLUMN_NAME = @Column )" & vbCrLf
sql = sql & "     set @sql = 'ALTER TABLE ' + @Table + ' ALTER COLUMN ' + @Column + ' ' + @Properties" & vbCrLf
sql = sql & "  else" & vbCrLf
sql = sql & "     set @sql = 'ALTER TABLE ' + @Table + ' ADD ' + @Column + ' ' + @Properties" & vbCrLf
sql = sql & "  execute (@sql)" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql

sql = "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'DeleteStoredProcedureIfExists')" & vbCrLf
sql = sql & "  DROP PROCEDURE DeleteStoredProcedureIfExists" & vbCrLf
ExecuteScript False, sql

sql = "create procedure DeleteStoredProcedureIfExists" & vbCrLf
sql = sql & "  @Proc_Name varchar(50)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "  IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = @Proc_Name)" & vbCrLf
sql = sql & "  begin" & vbCrLf
sql = sql & "     declare @sql varchar(100)" & vbCrLf
sql = sql & "     set @sql = 'DROP PROCEDURE ' + @Proc_Name" & vbCrLf
sql = sql & "     execute(@sql)" & vbCrLf
sql = sql & "  end" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql

sql = "AddOrUpdateColumn 'ComnTable', 'DenyLoginIfPriorOpOpen', 'tinyint null'" & vbCrLf
ExecuteScript False, sql

' Inactive Inventory (MANSERV)

'Rename the stored procedure.
sql = "EXEC sp_rename 'InventoryExcessReport', 'InventoryExcessReport_Old';" & vbCrLf
ExecuteScript False, sql

sql = "EXEC DeleteStoredProcedureIfExists 'InventoryExcessReport'" & vbCrLf
ExecuteScript False, sql

sql = "create PROCEDURE dbo.InventoryExcessReport" & vbCrLf
sql = sql & "          @BeginDate as varchar(16), @EndDate as varchar(16), @PartClass as Varchar(16), " & vbCrLf
sql = sql & "          @PartCode as varchar(8), @InclZQty as Integer, @PartType1 as Integer, " & vbCrLf
sql = sql & "          @PartType2 as Integer, @PartType3 as Integer, @PartType4 as Integer" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "exec dbo.InventoryExcessReport '1/1/2015','12/31/2015','','',1,1,1,1,1" & vbCrLf
sql = sql & "*/                                    " & vbCrLf
sql = sql & "BEGIN                                 " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @sqlZQty as varchar(12)   " & vbCrLf
sql = sql & "                                            " & vbCrLf
sql = sql & "IF (@PartClass = 'ALL')           " & vbCrLf
sql = sql & "BEGIN                             " & vbCrLf
sql = sql & "  SET @PartClass = ''           " & vbCrLf
sql = sql & "End                               " & vbCrLf
sql = sql & "IF (@PartCode = 'ALL')            " & vbCrLf
sql = sql & "BEGIN                             " & vbCrLf
sql = sql & "  SET @PartCode = ''            " & vbCrLf
sql = sql & "End                               " & vbCrLf
sql = sql & "                                            " & vbCrLf
sql = sql & "IF (@PartType1 = 1)               " & vbCrLf
sql = sql & "  SET @PartType1 = 1            " & vbCrLf
sql = sql & "Else                              " & vbCrLf
sql = sql & "  SET @PartType1 = 0            " & vbCrLf
sql = sql & "                                            " & vbCrLf
sql = sql & "IF (@PartType2 = 1)               " & vbCrLf
sql = sql & "  SET @PartType2 = 2            " & vbCrLf
sql = sql & "Else                              " & vbCrLf
sql = sql & "  SET @PartType2 = 0                " & vbCrLf
sql = sql & "                                          " & vbCrLf
sql = sql & "IF (@PartType3 = 1)                   " & vbCrLf
sql = sql & "  SET @PartType3 = 3                " & vbCrLf
sql = sql & "Else                                  " & vbCrLf
sql = sql & "  SET @PartType3 = 0                " & vbCrLf
sql = sql & "                                          " & vbCrLf
sql = sql & "IF (@PartType4 = 1)                   " & vbCrLf
sql = sql & "  SET @PartType4 = 4                " & vbCrLf
sql = sql & "Else                                  " & vbCrLf
sql = sql & "  SET @PartType4 = 0   " & vbCrLf
sql = sql & "     " & vbCrLf
sql = sql & "-- create a list of matching parts" & vbCrLf
sql = sql & "select PARTNUM, PARTREF, PACLASS, PAPRODCODE, PALEVEL,PADESC, PAEXTDESC, PAQOH AS QOH_Then, " & vbCrLf
sql = sql & "  PAQOH AS QOH_Now, cast('' as char(1)) AS MRP_Activity into #tempParts from PartTable" & vbCrLf
sql = sql & "where PACLASS LIKE '%' + @PartClass + '%'                     " & vbCrLf
sql = sql & "AND PAPRODCODE LIKE '%' + @PartCode + '%'                       " & vbCrLf
sql = sql & "AND PALEVEL IN (@PartType1, @PartType2, @PartType3, @PartType4)  " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete parts that did not exist until after end date" & vbCrLf
sql = sql & "delete from #tempParts where not exists " & vbCrLf
sql = sql & "(select INADATE from InvaTable where INPART = #tempParts.PARTREF and INADATE <= @EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- determine which parts have past-due MRP activity" & vbCrLf
sql = sql & "update #tempParts set MRP_Activity = 'X' where exists" & vbCrLf
sql = sql & "  (SELECT mrp_Partref FROM dbo.MrplTable      " & vbCrLf
sql = sql & "    WHERE MRP_PARTREF = #tempParts.PARTREF and mrp_type IN (2, 3, 4, 11, 12, 17)" & vbCrLf
sql = sql & "    AND mrp_partDateRQD < DATEADD(dd, +1 , @EndDate))    " & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete parts where there has been inventory activity in the date range" & vbCrLf
sql = sql & "delete from #tempParts" & vbCrLf
sql = sql & "where exists (SELECT INPART FROM invaTable where INPART = PARTREF " & vbCrLf
sql = sql & "  and INADATE BETWEEN @BeginDate and @EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- determine quantity at end date, if the date is different" & vbCrLf
sql = sql & "declare @today datetime" & vbCrLf
sql = sql & "set @today = cast(convert(varchar(10), getdate(), 101) as datetime)" & vbCrLf
sql = sql & "if @today <> @EndDate" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "  update #tempParts set QOH_Then = QOH_Then " & vbCrLf
sql = sql & "     - ISNULL((select sum(INAQTY) from InvaTable where INPART = PARTREF and INADATE > @EndDate),0)" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if zero quantity parts are not included, remove those with zero quantity at the end date" & vbCrLf
sql = sql & "if @InclZQty = 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "  delete from #tempParts where QOH_Then = 0 and QOH_Now = 0 and MRP_Activity = ''" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- return results" & vbCrLf
sql = sql & "select * from #tempParts order by PARTREF" & vbCrLf
sql = sql & "drop table #tempParts" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function

Private Function UpdateDatabase93()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 168     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = 'DropStoredProcedureIfExists')" & vbCrLf
sql = sql & "  DROP PROCEDURE DeleteStoredProcedureIfExists" & vbCrLf
ExecuteScript False, sql

sql = "create procedure DropStoredProcedureIfExists" & vbCrLf
sql = sql & "  @Proc_Name varchar(50)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "  IF EXISTS (SELECT * FROM sys.objects WHERE type = 'P' AND name = @Proc_Name)" & vbCrLf
sql = sql & "  begin" & vbCrLf
sql = sql & "     declare @sql varchar(100)" & vbCrLf
sql = sql & "     set @sql = 'DROP PROCEDURE ' + @Proc_Name" & vbCrLf
sql = sql & "     execute(@sql)" & vbCrLf
sql = sql & "  end" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql

' script to correctly associate sheet inventory lots and invatable rows

' first update duplicate INNUMBERS resulting from errors in SheetPick and SheetRestock sp's
sql = "if exists (select 1 from ComnTable where COUSESHEETINVENTORY = 1)" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "declare @dups TABLE (INNUMBER int)" & vbCrLf
sql = sql & "insert @dups" & vbCrLf
sql = sql & "select innumber from invatable" & vbCrLf
sql = sql & "where INPART in (select partref from PartTable where PAPUNITS = 'SH')" & vbCrLf
sql = sql & "group by innumber having count(*) > 1" & vbCrLf
sql = sql & "declare @maxINNUMBER int" & vbCrLf
sql = sql & "select @maxINNUMBER = max(INNUMBER) FROM InvaTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @dups2 table (INNUMBER int, Qty decimal(12,2), newINNUMBER int)" & vbCrLf
sql = sql & "insert @dups2" & vbCrLf
sql = sql & "select ia.INNUMBER, INAQTY, @maxINNUMBER + ROW_NUMBER() OVER (ORDER BY ia.INNUMBER, ia.INAQTY)" & vbCrLf
sql = sql & "from @dups join InvaTable ia on ia.INNUMBER = [@dups].INNUMBER" & vbCrLf
sql = sql & "where INAQTY <> (select min(INAQTY) from @dups join InvaTable ia2 on ia2.INNUMBER = [@dups].INNUMBER where ia2.INNUMBER = ia.INNUMBER)" & vbCrLf
sql = sql & "order by ia.INNUMBER, INAQTY" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ia" & vbCrLf
sql = sql & "set ia.INNUMBER = d.newINNUMBER" & vbCrLf
sql = sql & "from InvaTable ia join @dups2 d on d.INNUMBER = ia.INNUMBER and d.Qty = ia.INAQTY" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now find matching lot item records and point them at the appropriate INNUMBERS" & vbCrLf
sql = sql & "update  loi" & vbCrLf
sql = sql & "set LOIACTIVITY = INNUMBER" & vbCrLf
sql = sql & "from LoitTable loi join InvaTable ia on ia.INLOTNUMBER = loi.LOINUMBER" & vbCrLf
sql = sql & "and ia.INAQTY = loi.LOIQUANTITY" & vbCrLf
sql = sql & "and ia.INADATE = loi.LOIADATE" & vbCrLf
sql = sql & "where LOIPARTREF in (select partref from PartTable where PAPUNITS = 'SH')" & vbCrLf
sql = sql & "and LOIACTIVITY is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "commit tran" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql

sql = "DropStoredProcedureIfExists 'SheetCancelPick'" & vbCrLf
ExecuteScript False, sql

sql = "create procedure SheetCancelPick" & vbCrLf
sql = sql & "@LotUserLotID varchar(40)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get the lot id" & vbCrLf
sql = sql & "declare @LotID varchar(15)" & vbCrLf
sql = sql & "select @LotID = LOTNUMBER from LohdTable where LOTUSERLOTID = @LotUserLotID" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- remember the parent records" & vbCrLf
sql = sql & "declare @parents TABLE" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "Record int," & vbCrLf
sql = sql & "Area decimal(12,4)," & vbCrLf
sql = sql & "INNUMBER int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT @parents" & vbCrLf
sql = sql & "SELECT LOIPARENTREC, LOIAREA, LOIACTIVITY" & vbCrLf
sql = sql & "from LoitTable" & vbCrLf
sql = sql & "where LOINUMBER = @LotID AND LOICLOSED is NULL" & vbCrLf
sql = sql & "and LOISHEETACTTYPE = 'PK'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete the pick records" & vbCrLf
sql = sql & "delete LoitTable" & vbCrLf
sql = sql & "where LOINUMBER = @LotID AND LOICLOSED is NULL" & vbCrLf
sql = sql & "and LOISHEETACTTYPE = 'PK'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete inventory activity records" & vbCrLf
sql = sql & "delete InvaTable" & vbCrLf
sql = sql & "from InvaTable" & vbCrLf
sql = sql & "where INLOTNUMBER = @LotID" & vbCrLf
sql = sql & "and INNUMBER in (select INNUMBER from @parents)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- restore the parent records to pickable status" & vbCrLf
sql = sql & "update li" & vbCrLf
sql = sql & "set LOICLOSED = NULL" & vbCrLf
sql = sql & "from LoitTable li" & vbCrLf
sql = sql & "where LOINUMBER = @LotID AND LOIRECORD in (select Record from @parents)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- restore LOTREMAININGQTY" & vbCrLf
sql = sql & "declare @area decimal(12,4)" & vbCrLf
sql = sql & "select @area = sum(Area) from @parents" & vbCrLf
sql = sql & "update LohdTable set LOTREMAININGQTY = LOTREMAININGQTY + @area where LOTNUMBER = @LotID" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- restore part QOH" & vbCrLf
sql = sql & "update PartTable set PAQOH = PAQOH + @area" & vbCrLf
sql = sql & "from PartTable join LohdTable on PARTREF = LOTPARTREF" & vbCrLf
sql = sql & "where LOTNUMBER = @LotID" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "commit tran" & vbCrLf
ExecuteScript False, sql

sql = "DropStoredProcedureIfExists 'SheetPick'" & vbCrLf
ExecuteScript False, sql

sql = "create procedure [dbo].[SheetPick]" & vbCrLf
sql = sql & "@UserLotNo varchar(40),     -- lotuserlotid" & vbCrLf
sql = sql & "@User varchar(4)," & vbCrLf
sql = sql & "@SO int," & vbCrLf
sql = sql & "@PickDate datetime" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "Pick a sheet" & vbCrLf
sql = sql & "exec SheetPick '029373-1-A','MGR',222222" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- remove time from pick date" & vbCrLf
sql = sql & "set @PickDate = cast(convert(varchar(10),@PickDate,101)as datetime)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- find all rectangles in lot" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "create table #rect" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "id int identity," & vbCrLf
sql = sql & "ParentRecord int," & vbCrLf
sql = sql & "Height decimal(12,4)," & vbCrLf
sql = sql & "Length decimal(12,4)," & vbCrLf
sql = sql & "NewRecord int," & vbCrLf
sql = sql & "Qty decimal(12,4)," & vbCrLf
sql = sql & "NewINNUMBER int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @LotNo char(15)" & vbCrLf
sql = sql & "select @LotNo = LOTNUMBER from LohdTable where LOTUSERLOTID = @UserLotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert #rect (ParentRecord,Height,Length,NewRecord,Qty)" & vbCrLf
sql = sql & "select li.LOIRECORD,li.LOIHEIGHT,li.LOILENGTH,0,cast(li.LOIHEIGHT*li.LOILENGTH as decimal(12,4)) from LoitTable li" & vbCrLf
sql = sql & "join LohdTable lh on lh.LOTNUMBER = li.LOINUMBER" & vbCrLf
sql = sql & "where lh.LOTNUMBER = @LotNo and LOIQUANTITY > 0" & vbCrLf
sql = sql & "and li.LOICLOSED is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- assign new LoitTable record numbers" & vbCrLf
sql = sql & "declare @max int" & vbCrLf
sql = sql & "select @max = max(LOIRECORD) from LoitTable where LOINUMBER = @LotNo" & vbCrLf
sql = sql & "update #rect set NewRecord = @max + id" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @partRef varchar(30), @uom char(2), @type int" & vbCrLf
sql = sql & "declare @unitCost decimal(12,4), @maxINNUMBER int, @totalQty decimal(12,4)" & vbCrLf
sql = sql & "select @partRef = LOTPARTREF, @unitCost = LOTUNITCOST, @totalQty = LOTREMAININGQTY from LohdTable where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "select @uom = PAUNITS from PartTable where PARTREF = @partRef" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- assign" & vbCrLf
sql = sql & "select @maxINNUMBER = max(INNUMBER) from InvaTable" & vbCrLf
sql = sql & "update #rect set NewINNUMBER = @maxINNUMBER + id" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @type = 19     -- manual adjustment" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- deactivate rectangles being picked so they won't show up again" & vbCrLf
sql = sql & "Update li" & vbCrLf
sql = sql & "set li.LOICLOSED = @PickDate" & vbCrLf
sql = sql & "from LoitTable li" & vbCrLf
sql = sql & "join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD = r.ParentRecord" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new LoitTable record to zero lot quantity" & vbCrLf
sql = sql & "INSERT INTO [dbo].[LoitTable]" & vbCrLf
sql = sql & "([LOINUMBER]" & vbCrLf
sql = sql & ",[LOIRECORD]" & vbCrLf
sql = sql & ",[LOITYPE]" & vbCrLf
sql = sql & ",[LOIPARTREF]" & vbCrLf
sql = sql & ",[LOIADATE]" & vbCrLf
sql = sql & ",[LOIPDATE]" & vbCrLf
sql = sql & ",[LOIQUANTITY]" & vbCrLf
sql = sql & ",[LOIMOPARTREF]" & vbCrLf
sql = sql & ",[LOIMORUNNO]" & vbCrLf
sql = sql & ",[LOIPONUMBER]" & vbCrLf
sql = sql & ",[LOIPOITEM]" & vbCrLf
sql = sql & ",[LOIPOREV]" & vbCrLf
sql = sql & ",[LOIPSNUMBER]" & vbCrLf
sql = sql & ",[LOIPSITEM]" & vbCrLf
sql = sql & ",[LOICUSTINVNO]" & vbCrLf
sql = sql & ",[LOICUST]" & vbCrLf
sql = sql & ",[LOIVENDINVNO]" & vbCrLf
sql = sql & ",[LOIVENDOR]" & vbCrLf
sql = sql & ",[LOIACTIVITY]" & vbCrLf
sql = sql & ",[LOICOMMENT]" & vbCrLf
sql = sql & ",[LOIUNITS]" & vbCrLf
sql = sql & ",[LOIMOPKCANCEL]" & vbCrLf
sql = sql & ",[LOIHEIGHT]" & vbCrLf
sql = sql & ",[LOILENGTH]" & vbCrLf
sql = sql & ",[LOIUSER]" & vbCrLf
sql = sql & ",LOIPARENTREC" & vbCrLf
sql = sql & ",[LOISONUMBER]" & vbCrLf
sql = sql & ",LOISHEETACTTYPE" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "SELECT" & vbCrLf
sql = sql & "@LotNo" & vbCrLf
sql = sql & ",r.NewRecord" & vbCrLf
sql = sql & ",@type              -- manual adjustment" & vbCrLf
sql = sql & ",@partRef" & vbCrLf
sql = sql & ",@PickDate" & vbCrLf
sql = sql & ",@PickDate" & vbCrLf
sql = sql & ",-r.Qty" & vbCrLf
sql = sql & ",null   --LOIMOPARTREF, char(30),>" & vbCrLf
sql = sql & ",null   --<LOIMORUNNO, int,>" & vbCrLf
sql = sql & ",null   --<LOIPONUMBER, int,>" & vbCrLf
sql = sql & ",null   --<LOIPOITEM, smallint,>" & vbCrLf
sql = sql & ",null   --<LOIPOREV, char(2),>" & vbCrLf
sql = sql & ",null   --<LOIPSNUMBER, char(8),>" & vbCrLf
sql = sql & ",null   --<LOIPSITEM, smallint,>" & vbCrLf
sql = sql & ",null   --<LOICUSTINVNO, int,>" & vbCrLf
sql = sql & ",null   --<LOICUST, char(10),>" & vbCrLf
sql = sql & ",null   --<LOIVENDINVNO, char(20),>" & vbCrLf
sql = sql & ",null   --<LOIVENDOR, char(10),>" & vbCrLf
sql = sql & ",r.NewINNUMBER   --<LOIACTIVITY, int,> points to InvaTable.INNUMBER when IA is created" & vbCrLf
sql = sql & ",'sheet pick'   --<LOICOMMENT, varchar(40),>" & vbCrLf
sql = sql & ",@uom   --<LOIUNITS, char(2),>" & vbCrLf
sql = sql & ",null   --<LOIMOPKCANCEL, smallint,>" & vbCrLf
sql = sql & ",r.HEIGHT   --<LOIHEIGHT, decimal(12,4),>" & vbCrLf
sql = sql & ",r.LENGTH   --<LOILENGTH, decimal(12,4),>" & vbCrLf
sql = sql & ",@User          --<LOIUSER, varchar(4),>" & vbCrLf
sql = sql & ",r.ParentRecord  --LOIPARENTREC" & vbCrLf
sql = sql & ",@SO                --<LOISONUMBER, int,>" & vbCrLf
sql = sql & ",'PK'            --LOISHEETACTTYPE" & vbCrLf
sql = sql & "from #rect r" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- set lot quantity = 0" & vbCrLf
sql = sql & "update LohdTable set LOTREMAININGQTY = 0 where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new IA record to reduce part quantity" & vbCrLf
sql = sql & "INSERT INTO [dbo].[InvaTable]" & vbCrLf
sql = sql & "([INTYPE]" & vbCrLf
sql = sql & ",[INPART]" & vbCrLf
sql = sql & ",[INREF1]" & vbCrLf
sql = sql & ",[INREF2]" & vbCrLf
sql = sql & ",[INPDATE]" & vbCrLf
sql = sql & ",[INADATE]" & vbCrLf
sql = sql & ",[INPQTY]" & vbCrLf
sql = sql & ",[INAQTY]" & vbCrLf
sql = sql & ",[INAMT]" & vbCrLf
sql = sql & ",[INTOTMATL]" & vbCrLf
sql = sql & ",[INTOTLABOR]" & vbCrLf
sql = sql & ",[INTOTEXP]" & vbCrLf
sql = sql & ",[INTOTOH]" & vbCrLf
sql = sql & ",[INTOTHRS]" & vbCrLf
sql = sql & ",[INCREDITACCT]" & vbCrLf
sql = sql & ",[INDEBITACCT]" & vbCrLf
sql = sql & ",[INGLJOURNAL]" & vbCrLf
sql = sql & ",[INGLPOSTED]" & vbCrLf
sql = sql & ",[INGLDATE]" & vbCrLf
sql = sql & ",[INMOPART]" & vbCrLf
sql = sql & ",[INMORUN]" & vbCrLf
sql = sql & ",[INSONUMBER]" & vbCrLf
sql = sql & ",[INSOITEM]" & vbCrLf
sql = sql & ",[INSOREV]" & vbCrLf
sql = sql & ",[INPONUMBER]" & vbCrLf
sql = sql & ",[INPORELEASE]" & vbCrLf
sql = sql & ",[INPOITEM]" & vbCrLf
sql = sql & ",[INPOREV]" & vbCrLf
sql = sql & ",[INPSNUMBER]" & vbCrLf
sql = sql & ",[INPSITEM]" & vbCrLf
sql = sql & ",[INWIPLABACCT]" & vbCrLf
sql = sql & ",[INWIPMATACCT]" & vbCrLf
sql = sql & ",[INWIPOHDACCT]" & vbCrLf
sql = sql & ",[INWIPEXPACCT]" & vbCrLf
sql = sql & ",[INNUMBER]" & vbCrLf
sql = sql & ",[INLOTNUMBER]" & vbCrLf
sql = sql & ",[INUSER]" & vbCrLf
sql = sql & ",[INUNITS]" & vbCrLf
sql = sql & ",[INDRLABACCT]" & vbCrLf
sql = sql & ",[INDRMATACCT]" & vbCrLf
sql = sql & ",[INDREXPACCT]" & vbCrLf
sql = sql & ",[INDROHDACCT]" & vbCrLf
sql = sql & ",[INCRLABACCT]" & vbCrLf
sql = sql & ",[INCRMATACCT]" & vbCrLf
sql = sql & ",[INCREXPACCT]" & vbCrLf
sql = sql & ",[INCROHDACCT]" & vbCrLf
sql = sql & ",[INLOTTRACK]" & vbCrLf
sql = sql & ",[INUSEACTUALCOST]" & vbCrLf
sql = sql & ",[INCOSTEDBY]" & vbCrLf
sql = sql & ",[INMAINTCOSTED])" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "@type               --<INTYPE, int,>" & vbCrLf
sql = sql & ",@partRef           --char(30),>" & vbCrLf
sql = sql & ",'Manual Adjustment'    --<INREF1, char(20),>" & vbCrLf
sql = sql & ",'Sheet Inventory'  --<INREF2, char(40),>" & vbCrLf
sql = sql & ",@PickDate       --<INPDATE, smalldatetime,>" & vbCrLf
sql = sql & ",@PickDate       --<INADATE, smalldatetime,>" & vbCrLf
sql = sql & ",-r.Qty         --<INPQTY, decimal(12,4),>" & vbCrLf
sql = sql & ",-r.Qty         --<INAQTY, decimal(12,4),>" & vbCrLf
sql = sql & ",@unitCost          -- <INAMT, decimal(12,4),>" & vbCrLf
sql = sql & ",0.0                    --<INTOTMATL, decimal(12,4),>" & vbCrLf
sql = sql & ",0.0                    --<INTOTLABOR, decimal(12,4),>" & vbCrLf
sql = sql & ",0.0                    --<INTOTEXP, decimal(12,4),>" & vbCrLf
sql = sql & ",0.0                    --<INTOTOH, decimal(12,4),>" & vbCrLf
sql = sql & ",0.0                    --<INTOTHRS, decimal(12,4),>" & vbCrLf
sql = sql & ",''                 --<INCREDITACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INDEBITACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INGLJOURNAL, char(12),>" & vbCrLf
sql = sql & ",0                  --<INGLPOSTED, tinyint,>" & vbCrLf
sql = sql & ",null               --<INGLDATE, smalldatetime,>" & vbCrLf
sql = sql & ",''                 --<INMOPART, char(30),>" & vbCrLf
sql = sql & ",0                  --<INMORUN, int,>" & vbCrLf
sql = sql & ",0                  --<INSONUMBER, int,>" & vbCrLf
sql = sql & ",0                  --<INSOITEM, int,>" & vbCrLf
sql = sql & ",''                 --<INSOREV, char(2),>" & vbCrLf
sql = sql & ",0                  --<INPONUMBER, int,>" & vbCrLf
sql = sql & ",0                  --<INPORELEASE, smallint,>" & vbCrLf
sql = sql & ",0                  --<INPOITEM, smallint,>" & vbCrLf
sql = sql & ",''                 --<INPOREV, char(2),>" & vbCrLf
sql = sql & ",''                 --<INPSNUMBER, char(8),>" & vbCrLf
sql = sql & ",0                  --<INPSITEM, smallint,>" & vbCrLf
sql = sql & ",''                 --<INWIPLABACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INWIPMATACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INWIPOHDACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INWIPEXPACCT, char(12),>" & vbCrLf
sql = sql & ",R.NewINNUMBER      --<INNUMBER, int,>" & vbCrLf
sql = sql & ",@LotNo             --<INLOTNUMBER, char(15),>" & vbCrLf
sql = sql & ",@User              --<INUSER, char(4),>" & vbCrLf
sql = sql & ",@uom               --<INUNITS, char(2),>" & vbCrLf
sql = sql & ",''                 --<INDRLABACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INDRMATACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INDREXPACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INDROHDACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INCRLABACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INCRMATACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INCREXPACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INCROHDACCT, char(12),>" & vbCrLf
sql = sql & ",null               --<INLOTTRACK, bit,>" & vbCrLf
sql = sql & ",null               --<INUSEACTUALCOST, bit,>" & vbCrLf
sql = sql & ",null               --<INCOSTEDBY, char(4),>" & vbCrLf
sql = sql & ",0                  --<INMAINTCOSTED, int,>)" & vbCrLf
sql = sql & "from #rect r" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now update part QOH" & vbCrLf
sql = sql & "update PartTable set PAQOH = PAQOH - @totalQty where PARTREF = @partRef" & vbCrLf
sql = sql & "commit tran" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql

sql = "DropStoredProcedureIfExists 'SheetRestock'" & vbCrLf
ExecuteScript False, sql

sql = "create procedure SheetRestock" & vbCrLf
sql = sql & "@UserLotNo varchar(40)," & vbCrLf
sql = sql & "@User varchar(4)," & vbCrLf
sql = sql & "@Comments varchar(2048)," & vbCrLf
sql = sql & "@Location varchar(4)," & vbCrLf
sql = sql & "@RestockDate datetime," & vbCrLf
sql = sql & "@Params varchar(2000)   -- LOIRECORD,NEWHT,NEWLEN,... repeat (include comma at end)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "restock sheet LOI records" & vbCrLf
sql = sql & "exec SheetRestock '029373-1-A', 'MGR', '2,36,96,0,12,72,'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "create table #rect" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "ParentRecord int," & vbCrLf
sql = sql & "NewHeight decimal(12,4)," & vbCrLf
sql = sql & "NewLength decimal(12,4)," & vbCrLf
sql = sql & "NewRecord int," & vbCrLf
sql = sql & "Qty decimal(12,4)," & vbCrLf
sql = sql & "NewINNUMBER int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @nextRecord int, @nextIANumber int" & vbCrLf
sql = sql & "declare @LotNo char(15), @id int, @SO int" & vbCrLf
sql = sql & "select @LotNo = LOTNUMBER from LohdTable where LOTUSERLOTID = @UserLotNo" & vbCrLf
sql = sql & "select @nextRecord = max(LOIRECORD) + 1 , @SO = max(LOISONUMBER) from LoitTable where LOINUMBER = @LotNo" & vbCrLf
sql = sql & "select @nextIANumber = max(INNUMBER) + 1 from InvaTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- first update the lot header" & vbCrLf
sql = sql & "update LohdTable set LOTCOMMENTS = @Comments, LOTLOCATION = @Location where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- extract parameters" & vbCrLf
sql = sql & "DECLARE @start INT, @end INT" & vbCrLf
sql = sql & "declare @stringId varchar(10), @stringHt varchar(10), @stringLen varchar(10)" & vbCrLf
sql = sql & "SELECT @start = 1, @end = CHARINDEX(',', @Params)" & vbCrLf
sql = sql & "WHILE @start < LEN(@Params) + 1 BEGIN" & vbCrLf
sql = sql & "IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @stringId = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "SET @start = @end + 1" & vbCrLf
sql = sql & "SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @stringHt = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "SET @start = @end + 1" & vbCrLf
sql = sql & "SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF @end = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @stringLen = SUBSTRING(@Params, @start, @end - @start)" & vbCrLf
sql = sql & "SET @start = @end + 1" & vbCrLf
sql = sql & "SET @end = CHARINDEX(',', @Params, @start)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @ht decimal(12,4), @len decimal(12,4)" & vbCrLf
sql = sql & "set @ht = cast(@stringHt as decimal(12,4))" & vbCrLf
sql = sql & "set @len = cast(@stringLen as decimal(12,4))" & vbCrLf
sql = sql & "insert into #rect (ParentRecord, NewHeight, NewLength, NewRecord, NewINNUMBER)" & vbCrLf
sql = sql & "values (cast(@stringId as int), @ht, @len," & vbCrLf
sql = sql & "case when @ht*@len = 0 then 0 else @nextRecord end," & vbCrLf
sql = sql & "case when @ht*@len = 0 then 0 else @nextIANumber end)" & vbCrLf
sql = sql & "if (@ht*@len) <> 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "set @nextRecord = @nextRecord + 1" & vbCrLf
sql = sql & "set @nextIANumber = @nextIANumber + 1" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #rect set Qty = NewHeight * NewLength" & vbCrLf
sql = sql & "update #rect set ParentRecord = isnull((select top 1 ParentRecord from #rect where ParentRecord <> 0 ),0)" & vbCrLf
sql = sql & "where ParentRecord = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from #rect" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @partRef varchar(30), @uom char(2), @type int" & vbCrLf
sql = sql & "declare @unitCost decimal(12,4), @nextINNUMBER int, @totalQty decimal(12,4)" & vbCrLf
sql = sql & "select @partRef = LOTPARTREF, @unitCost = LOTUNITCOST, @totalQty = LOTREMAININGQTY from LohdTable where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "select @uom = PAUNITS from PartTable where PARTREF = @partRef" & vbCrLf
sql = sql & "select @nextINNUMBER = max(INNUMBER) + 1 from InvaTable" & vbCrLf
sql = sql & "set @type = 19     -- manual adjustment" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- deactivate picked rectangles so they won't show up again" & vbCrLf
sql = sql & "Update li" & vbCrLf
sql = sql & "set li.LOICLOSED = cast(convert(varchar(10),@RestockDate,101) as datetime)" & vbCrLf
sql = sql & "from LoitTable li" & vbCrLf
sql = sql & "join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD = r.ParentRecord" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- remove rectangles with no quantity remaining.  These items will not be restocked." & vbCrLf
sql = sql & "delete from #rect where Qty = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new lot items" & vbCrLf
sql = sql & "INSERT INTO [dbo].[LoitTable]" & vbCrLf
sql = sql & "([LOINUMBER]" & vbCrLf
sql = sql & ",[LOIRECORD]" & vbCrLf
sql = sql & ",[LOITYPE]" & vbCrLf
sql = sql & ",[LOIPARTREF]" & vbCrLf
sql = sql & ",[LOIADATE]" & vbCrLf
sql = sql & ",[LOIPDATE]" & vbCrLf
sql = sql & ",[LOIQUANTITY]" & vbCrLf
sql = sql & ",[LOIMOPARTREF]" & vbCrLf
sql = sql & ",[LOIMORUNNO]" & vbCrLf
sql = sql & ",[LOIPONUMBER]" & vbCrLf
sql = sql & ",[LOIPOITEM]" & vbCrLf
sql = sql & ",[LOIPOREV]" & vbCrLf
sql = sql & ",[LOIPSNUMBER]" & vbCrLf
sql = sql & ",[LOIPSITEM]" & vbCrLf
sql = sql & ",[LOICUSTINVNO]" & vbCrLf
sql = sql & ",[LOICUST]" & vbCrLf
sql = sql & ",[LOIVENDINVNO]" & vbCrLf
sql = sql & ",[LOIVENDOR]" & vbCrLf
sql = sql & ",[LOIACTIVITY]" & vbCrLf
sql = sql & ",[LOICOMMENT]" & vbCrLf
sql = sql & ",[LOIUNITS]" & vbCrLf
sql = sql & ",[LOIMOPKCANCEL]" & vbCrLf
sql = sql & ",[LOIHEIGHT]" & vbCrLf
sql = sql & ",[LOILENGTH]" & vbCrLf
sql = sql & ",[LOIUSER]" & vbCrLf
sql = sql & ",LOIPARENTREC" & vbCrLf
sql = sql & ",[LOISONUMBER]" & vbCrLf
sql = sql & ",LOISHEETACTTYPE" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "SELECT" & vbCrLf
sql = sql & "@LotNo" & vbCrLf
sql = sql & ",r.NewRecord" & vbCrLf
sql = sql & ",@type              -- manual adjustment" & vbCrLf
sql = sql & ",@partRef" & vbCrLf
sql = sql & ",@RestockDate" & vbCrLf
sql = sql & ",@RestockDate" & vbCrLf
sql = sql & ",r.Qty" & vbCrLf
sql = sql & ",null   --LOIMOPARTREF, char(30),>" & vbCrLf
sql = sql & ",null   --<LOIMORUNNO, int,>" & vbCrLf
sql = sql & ",null   --<LOIPONUMBER, int,>" & vbCrLf
sql = sql & ",null   --<LOIPOITEM, smallint,>" & vbCrLf
sql = sql & ",null   --<LOIPOREV, char(2),>" & vbCrLf
sql = sql & ",null   --<LOIPSNUMBER, char(8),>" & vbCrLf
sql = sql & ",null   --<LOIPSITEM, smallint,>" & vbCrLf
sql = sql & ",null   --<LOICUSTINVNO, int,>" & vbCrLf
sql = sql & ",null   --<LOICUST, char(10),>" & vbCrLf
sql = sql & ",null   --<LOIVENDINVNO, char(20),>" & vbCrLf
sql = sql & ",null   --<LOIVENDOR, char(10),>" & vbCrLf
sql = sql & ",r.NewINNUMBER   --<LOIACTIVITY, int,> points to InvaTable.INNUMBER when IA is created" & vbCrLf
sql = sql & ",'sheet restock'    --<LOICOMMENT, varchar(40),>" & vbCrLf
sql = sql & ",@uom   --<LOIUNITS, char(2),>" & vbCrLf
sql = sql & ",null   --<LOIMOPKCANCEL, smallint,>" & vbCrLf
sql = sql & ",r.NewHeight    --<LOIHEIGHT, decimal(12,4),>" & vbCrLf
sql = sql & ",r.NewLength    --<LOILENGTH, decimal(12,4),>" & vbCrLf
sql = sql & ",@User      --<LOIUSER, varchar(4),>" & vbCrLf
sql = sql & ",r.ParentRecord  --LOIPICKEDFROMREC" & vbCrLf
sql = sql & ",@SO        --<LOISONUMBER, int,>" & vbCrLf
sql = sql & ",'RS'    --LOISHEETACTTYPE" & vbCrLf
sql = sql & "from #rect r" & vbCrLf
sql = sql & "where r.Qty <> 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select distinct li.* from LoitTable li" & vbCrLf
sql = sql & "--join #rect r on li.LOINUMBER = @LotNo and li.LOIRECORD > 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- set LOTREMAININGQTY and remove reservation" & vbCrLf
sql = sql & "declare @sum decimal(12,4)" & vbCrLf
sql = sql & "select @sum = sum(Qty) from #rect" & vbCrLf
sql = sql & "update LohdTable" & vbCrLf
sql = sql & "set LOTREMAININGQTY = LOTREMAININGQTY + @sum," & vbCrLf
sql = sql & "LOTRESERVEDBY = NULL," & vbCrLf
sql = sql & "LOTRESERVEDON = NULL" & vbCrLf
sql = sql & "where LOTNUMBER = @LotNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create new ia items" & vbCrLf
sql = sql & "INSERT INTO [dbo].[InvaTable]" & vbCrLf
sql = sql & "([INTYPE]" & vbCrLf
sql = sql & ",[INPART]" & vbCrLf
sql = sql & ",[INREF1]" & vbCrLf
sql = sql & ",[INREF2]" & vbCrLf
sql = sql & ",[INPDATE]" & vbCrLf
sql = sql & ",[INADATE]" & vbCrLf
sql = sql & ",[INPQTY]" & vbCrLf
sql = sql & ",[INAQTY]" & vbCrLf
sql = sql & ",[INAMT]" & vbCrLf
sql = sql & ",[INTOTMATL]" & vbCrLf
sql = sql & ",[INTOTLABOR]" & vbCrLf
sql = sql & ",[INTOTEXP]" & vbCrLf
sql = sql & ",[INTOTOH]" & vbCrLf
sql = sql & ",[INTOTHRS]" & vbCrLf
sql = sql & ",[INCREDITACCT]" & vbCrLf
sql = sql & ",[INDEBITACCT]" & vbCrLf
sql = sql & ",[INGLJOURNAL]" & vbCrLf
sql = sql & ",[INGLPOSTED]" & vbCrLf
sql = sql & ",[INGLDATE]" & vbCrLf
sql = sql & ",[INMOPART]" & vbCrLf
sql = sql & ",[INMORUN]" & vbCrLf
sql = sql & ",[INSONUMBER]" & vbCrLf
sql = sql & ",[INSOITEM]" & vbCrLf
sql = sql & ",[INSOREV]" & vbCrLf
sql = sql & ",[INPONUMBER]" & vbCrLf
sql = sql & ",[INPORELEASE]" & vbCrLf
sql = sql & ",[INPOITEM]" & vbCrLf
sql = sql & ",[INPOREV]" & vbCrLf
sql = sql & ",[INPSNUMBER]" & vbCrLf
sql = sql & ",[INPSITEM]" & vbCrLf
sql = sql & ",[INWIPLABACCT]" & vbCrLf
sql = sql & ",[INWIPMATACCT]" & vbCrLf
sql = sql & ",[INWIPOHDACCT]" & vbCrLf
sql = sql & ",[INWIPEXPACCT]" & vbCrLf
sql = sql & ",[INNUMBER]" & vbCrLf
sql = sql & ",[INLOTNUMBER]" & vbCrLf
sql = sql & ",[INUSER]" & vbCrLf
sql = sql & ",[INUNITS]" & vbCrLf
sql = sql & ",[INDRLABACCT]" & vbCrLf
sql = sql & ",[INDRMATACCT]" & vbCrLf
sql = sql & ",[INDREXPACCT]" & vbCrLf
sql = sql & ",[INDROHDACCT]" & vbCrLf
sql = sql & ",[INCRLABACCT]" & vbCrLf
sql = sql & ",[INCRMATACCT]" & vbCrLf
sql = sql & ",[INCREXPACCT]" & vbCrLf
sql = sql & ",[INCROHDACCT]" & vbCrLf
sql = sql & ",[INLOTTRACK]" & vbCrLf
sql = sql & ",[INUSEACTUALCOST]" & vbCrLf
sql = sql & ",[INCOSTEDBY]" & vbCrLf
sql = sql & ",[INMAINTCOSTED])" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "@type               --<INTYPE, int,>" & vbCrLf
sql = sql & ",@partRef           --char(30),>" & vbCrLf
sql = sql & ",'Manual Adjustment'    --<INREF1, char(20),>" & vbCrLf
sql = sql & ",'Sheet Inventory'  --<INREF2, char(40),>" & vbCrLf
sql = sql & ",@RestockDate       --<INPDATE, smalldatetime,>" & vbCrLf
sql = sql & ",@RestockDate       --<INADATE, smalldatetime,>" & vbCrLf
sql = sql & ",r.Qty          --<INPQTY, decimal(12,4),>" & vbCrLf
sql = sql & ",r.Qty          --<INAQTY, decimal(12,4),>" & vbCrLf
sql = sql & ",@unitCost          -- <INAMT, decimal(12,4),>" & vbCrLf
sql = sql & ",0.0                    --<INTOTMATL, decimal(12,4),>" & vbCrLf
sql = sql & ",0.0                    --<INTOTLABOR, decimal(12,4),>" & vbCrLf
sql = sql & ",0.0                    --<INTOTEXP, decimal(12,4),>" & vbCrLf
sql = sql & ",0.0                    --<INTOTOH, decimal(12,4),>" & vbCrLf
sql = sql & ",0.0                    --<INTOTHRS, decimal(12,4),>" & vbCrLf
sql = sql & ",''                 --<INCREDITACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INDEBITACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INGLJOURNAL, char(12),>" & vbCrLf
sql = sql & ",0                  --<INGLPOSTED, tinyint,>" & vbCrLf
sql = sql & ",null               --<INGLDATE, smalldatetime,>" & vbCrLf
sql = sql & ",''                 --<INMOPART, char(30),>" & vbCrLf
sql = sql & ",0                  --<INMORUN, int,>" & vbCrLf
sql = sql & ",0                  --<INSONUMBER, int,>" & vbCrLf
sql = sql & ",0                  --<INSOITEM, int,>" & vbCrLf
sql = sql & ",''                 --<INSOREV, char(2),>" & vbCrLf
sql = sql & ",0                  --<INPONUMBER, int,>" & vbCrLf
sql = sql & ",0                  --<INPORELEASE, smallint,>" & vbCrLf
sql = sql & ",0                  --<INPOITEM, smallint,>" & vbCrLf
sql = sql & ",''                 --<INPOREV, char(2),>" & vbCrLf
sql = sql & ",''                 --<INPSNUMBER, char(8),>" & vbCrLf
sql = sql & ",0                  --<INPSITEM, smallint,>" & vbCrLf
sql = sql & ",''                 --<INWIPLABACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INWIPMATACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INWIPOHDACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INWIPEXPACCT, char(12),>" & vbCrLf
sql = sql & ",r.NewINNUMBER      --<INNUMBER, int,>" & vbCrLf
sql = sql & ",@LotNo             --<INLOTNUMBER, char(15),>" & vbCrLf
sql = sql & ",@User              --<INUSER, char(4),>" & vbCrLf
sql = sql & ",@uom               --<INUNITS, char(2),>" & vbCrLf
sql = sql & ",''                 --<INDRLABACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INDRMATACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INDREXPACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INDROHDACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INCRLABACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INCRMATACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INCREXPACCT, char(12),>" & vbCrLf
sql = sql & ",''                 --<INCROHDACCT, char(12),>" & vbCrLf
sql = sql & ",null               --<INLOTTRACK, bit,>" & vbCrLf
sql = sql & ",null               --<INUSEACTUALCOST, bit,>" & vbCrLf
sql = sql & ",null               --<INCOSTEDBY, char(4),>" & vbCrLf
sql = sql & ",0                  --<INMAINTCOSTED, int,>)" & vbCrLf
sql = sql & "from #rect r" & vbCrLf
sql = sql & "where r.Qty <> 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now update part QOH" & vbCrLf
sql = sql & "update PartTable set PAQOH = PAQOH + @totalQty where PARTREF = @partRef" & vbCrLf
sql = sql & "commit tran" & vbCrLf
ExecuteScript False, sql



''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function


Private Function UpdateDatabase94()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 169     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "AddOrUpdateColumn 'EsReportVendorStmt', 'Journal', 'varchar(12) NULL'" & vbCrLf
ExecuteScript False, sql

''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function

Private Function UpdateDatabase95()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 170     ' set actual version
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "DropStoredProcedureIfExists 'InventoryExcessReport'" & vbCrLf
ExecuteScript False, sql

sql = "create PROCEDURE [dbo].[InventoryExcessReport]" & vbCrLf
sql = sql & "@BeginDate as varchar(16), @EndDate as varchar(16), @PartClass as Varchar(16)," & vbCrLf
sql = sql & "@PartCode as varchar(8), @InclZQty as Integer, @PartType1 as Integer," & vbCrLf
sql = sql & "@PartType2 as Integer, @PartType3 as Integer, @PartType4 as Integer" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "exec dbo.InventoryExcessReport '1/1/2015','12/31/2015','','',1,1,1,1,1" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @sqlZQty as varchar(12)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartClass = 'ALL')" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @PartClass = ''" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "IF (@PartCode = 'ALL')" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @PartCode = ''" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartType1 = 1)" & vbCrLf
sql = sql & "SET @PartType1 = 1" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType1 = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartType2 = 1)" & vbCrLf
sql = sql & "SET @PartType2 = 2" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType2 = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartType3 = 1)" & vbCrLf
sql = sql & "SET @PartType3 = 3" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType3 = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartType4 = 1)" & vbCrLf
sql = sql & "SET @PartType4 = 4" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType4 = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create a list of matching parts" & vbCrLf
sql = sql & "select PARTNUM, PARTREF, PACLASS, PAPRODCODE, PALEVEL,PADESC, PAEXTDESC, PAQOH AS QOH_Then," & vbCrLf
sql = sql & "PAQOH AS QOH_Now, cast('' as char(1)) AS MRP_Activity into #tempParts from PartTable" & vbCrLf
sql = sql & "where PACLASS LIKE '%' + @PartClass + '%'" & vbCrLf
sql = sql & "AND PAPRODCODE LIKE '%' + @PartCode + '%'" & vbCrLf
sql = sql & "AND PALEVEL IN (@PartType1, @PartType2, @PartType3, @PartType4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete parts that did not exist until after end date" & vbCrLf
sql = sql & "delete from #tempParts where not exists" & vbCrLf
sql = sql & "(select INADATE from InvaTable where INPART = #tempParts.PARTREF and INADATE <= @EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- determine which parts have past-due MRP activity" & vbCrLf
sql = sql & "update #tempParts set MRP_Activity = 'X' where exists" & vbCrLf
sql = sql & "(SELECT mrp_Partref FROM dbo.MrplTable" & vbCrLf
sql = sql & "WHERE MRP_PARTREF = #tempParts.PARTREF and mrp_type IN (2, 3, 4, 11, 12, 17)" & vbCrLf
sql = sql & "AND mrp_partDateRQD < DATEADD(dd, +1 , @EndDate))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete parts where there has been inventory activity in the date range" & vbCrLf
sql = sql & "delete from #tempParts" & vbCrLf
sql = sql & "where exists (SELECT INPART FROM invaTable where INPART = PARTREF" & vbCrLf
sql = sql & "and INADATE BETWEEN @BeginDate and @EndDate" & vbCrLf
sql = sql & "and INTYPE not in (19,30))   -- ignore manual adjustments and ABC cycle counts" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- determine quantity at end date, if the date is different" & vbCrLf
sql = sql & "declare @today datetime" & vbCrLf
sql = sql & "set @today = cast(convert(varchar(10), getdate(), 101) as datetime)" & vbCrLf
sql = sql & "if @today <> @EndDate" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "update #tempParts set QOH_Then = QOH_Then" & vbCrLf
sql = sql & "- ISNULL((select sum(INAQTY) from InvaTable where INPART = PARTREF and INADATE > @EndDate),0)" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if zero quantity parts are not included, remove those with zero quantity at the end date" & vbCrLf
sql = sql & "if @InclZQty = 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "delete from #tempParts where QOH_Then = 0 and QOH_Now = 0 and MRP_Activity = ''" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- return results" & vbCrLf
sql = sql & "select * from #tempParts order by PARTREF" & vbCrLf
sql = sql & "drop table #tempParts" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql

' BOM Fix
sql = "update BmhdTable set BMHRELEASED = 1 where BMHRELEASED = 2"
ExecuteScript False, sql

' CGS Detail Report fixes
sql = "DropStoredProcedureIfExists 'RptMOCostDetail'" & vbCrLf
ExecuteScript False, sql

sql = "CREATE PROCEDURE RptMOCostDetail" & vbCrLf
sql = sql & "@MOPart as varchar(30),@MORun as int, @MOQty as decimal(15,4)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @SumTotMat decimal(15,4)" & vbCrLf
sql = sql & "declare @SumTotLabor decimal(15,4)" & vbCrLf
sql = sql & "declare @SumTotExp decimal(15,4)" & vbCrLf
sql = sql & "declare @SumTotOH decimal(15,4)" & vbCrLf
sql = sql & "declare @level as integer" & vbCrLf
sql = sql & "declare @Part as varchar(30)" & vbCrLf
sql = sql & "declare @PrevParent  as varchar(30)" & vbCrLf
sql = sql & "declare @RowCount as integer" & vbCrLf
sql = sql & "declare @ChildPart as varchar(30)" & vbCrLf
sql = sql & "declare @ParentPart as varchar(30)" & vbCrLf
sql = sql & "declare @MOPart1 as varchar(30)" & vbCrLf
sql = sql & "declare @MoRun1 as varchar(20)" & vbCrLf
sql = sql & "declare @Part1 as varchar(30)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @ChildKey as varchar(1024)" & vbCrLf
sql = sql & "declare @GLSortKey as varchar(1024)" & vbCrLf
sql = sql & "declare @SortKey as varchar(1024)" & vbCrLf
sql = sql & "declare @ParentLotNum as varchar(15)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @Maxlevel as int" & vbCrLf
sql = sql & "declare @LotRunNo as int" & vbCrLf
sql = sql & "declare @LotOrgQty as decimal(15,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @LotUSpMat decimal(15,4)" & vbCrLf
sql = sql & "declare @LotUSpLabor decimal(15,4)" & vbCrLf
sql = sql & "declare @LotUSpExp decimal(15,4)" & vbCrLf
sql = sql & "declare @LotUSpOH decimal(15,4)" & vbCrLf
sql = sql & "declare @LotMatl decimal(15,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @MOLotNum as varchar(15)" & vbCrLf
sql = sql & "declare @SplitLot as varchar(15)" & vbCrLf
sql = sql & "declare @cnt  as int" & vbCrLf
sql = sql & "declare @sumQty decimal(15,4)" & vbCrLf
sql = sql & "declare @MOPartRunKey as Varchar(30)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--DROP TABLE #tempMOPartsDetail" & vbCrLf
sql = sql & "-- DELETE FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @MOPartRunKey = RTRIM(@MOPart) + '_' + Convert(varchar(10), @MORun)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CREATE TABLE #tempMOPartsDetail" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "LOTMOPARTRUNKEY varchar(50) NULL," & vbCrLf
sql = sql & "INMOPART Varchar(30) NULL," & vbCrLf
sql = sql & "INMORUN int NULL ," & vbCrLf
sql = sql & "INPART varchar(30) NULL ," & vbCrLf
sql = sql & "LOTNUMBER varchar(15) NULL," & vbCrLf
sql = sql & "LOTUSERLOTID varchar(40) NULL," & vbCrLf
sql = sql & "INTOTMATL decimal(12,4) NULL," & vbCrLf
sql = sql & "INTOTLABOR decimal(12,4) NULL," & vbCrLf
sql = sql & "INTOTEXP decimal(12,4) NULL," & vbCrLf
sql = sql & "INTOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTMATL decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTLABOR decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTEXP decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTMAL decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTLABOR decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTEXP decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTDATECOSTED smalldatetime NULL," & vbCrLf
sql = sql & "SortKey varchar(512) NULL," & vbCrLf
sql = sql & "HASCHILD int NULL," & vbCrLf
sql = sql & "SORTKEYLEVEL tinyint NULL," & vbCrLf
sql = sql & "SortKeyRev varchar(512)," & vbCrLf
sql = sql & "PARTSUM varchar(40)," & vbCrLf
sql = sql & "BMQTYREQD decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTORGQTY decimal(12,4) NULL," & vbCrLf
sql = sql & "BMTOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTSPLITFROMSYS varchar(15)," & vbCrLf
sql = sql & "INVNO int NULL," & vbCrLf
sql = sql & "ITPSNUMBER varchar(8) NULL," & vbCrLf
sql = sql & "ITPSITEM smallint NULL," & vbCrLf
sql = sql & "PICKQTY decimal(12,4) NULL" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "with cte" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "(select BMASSYPART, BMPARTREF,  BMQTYREQD,0 as level, cast('1' + char(36)+ BMPARTREF as varchar(max)) as SortKey" & vbCrLf
sql = sql & "from BmplTable" & vbCrLf
sql = sql & "where BMASSYPART = @MOPart" & vbCrLf
sql = sql & "union all" & vbCrLf
sql = sql & "select a.BMASSYPART, a.BMPARTREF, a.BMQTYREQD, level + 1," & vbCrLf
sql = sql & "cast(COALESCE(SortKey,'') + char(36) + COALESCE(a.BMPARTREF,'') as varchar(max))as SortKey" & vbCrLf
sql = sql & "from cte" & vbCrLf
sql = sql & "inner join BmplTable a" & vbCrLf
sql = sql & "on cte.BMPARTREF = a.BMASSYPART" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "INSERT INTO #tempMOPartsDetail(INMOPART,INPART,BMQTYREQD,SORTKEYLEVEL,SortKey)" & vbCrLf
sql = sql & "select BMASSYPART, BMPARTREF,BMQTYREQD,level,SortKey" & vbCrLf
sql = sql & "from cte order by SortKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @cnt = 0" & vbCrLf
sql = sql & "print 'Update Start:' + cast(getdate() as char(25))" & vbCrLf
sql = sql & "print 'Count :' + Convert(varchar(10), @cnt)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET INMORUN = @MORun, LOTNUMBER = b.INLOTNUMBER," & vbCrLf
sql = sql & "LOTSPLITFROMSYS = c.LOTSPLITFROMSYS" & vbCrLf
sql = sql & "FROM dbo.LohdTable AS c INNER JOIN" & vbCrLf
sql = sql & "dbo.InvaTable AS b ON c.LOTNUMBER = b.INLOTNUMBER AND c.LOTPARTREF = b.INPART LEFT OUTER JOIN" & vbCrLf
sql = sql & "dbo.#tempMOPartsDetail AS d ON b.INMOPART = d.INMOPART AND b.INPART = d.INPART" & vbCrLf
sql = sql & "WHERE     (b.INMOPART = @MOPart) AND (b.INMORUN = @MORun) AND (b.INTYPE = 10) AND (d.SORTKEYLEVEL = 0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET PICKQTY = sumqty * -1" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail," & vbCrLf
sql = sql & "(SELECT SUM(b.INAQTY) sumqty, d.INMOPART mopart, d.INMORUN morun, d.LOTNUMBER lotnum, d.INPART subpart" & vbCrLf
sql = sql & "FROM dbo.LohdTable AS c INNER JOIN" & vbCrLf
sql = sql & "dbo.InvaTable AS b ON c.LOTNUMBER = b.INLOTNUMBER AND c.LOTPARTREF = b.INPART LEFT OUTER JOIN" & vbCrLf
sql = sql & "dbo.#tempMOPartsDetail AS d ON b.INMOPART = d.INMOPART AND b.INPART = d.INPART" & vbCrLf
sql = sql & "WHERE     b.INMOPART = @MOPart AND b.INMORUN  = @MORun AND (b.INTYPE = 10) AND (d.SORTKEYLEVEL = 0)" & vbCrLf
sql = sql & "GROUP BY d.INMOPART, d.INMORUN, d.LOTNUMBER, d.INPART" & vbCrLf
sql = sql & ") as f" & vbCrLf
sql = sql & "WHERE INMOPART = f.mopart AND INMORUN = f.morun" & vbCrLf
sql = sql & "AND LOTNUMBER = lotnum AND INPART = subpart" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// Update the" & vbCrLf
sql = sql & "--// update the top level" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET INMORUN = @MORun, LOTNUMBER = b.INLOTNUMBER," & vbCrLf
sql = sql & "LOTUSERLOTID = c.LOTUSERLOTID," & vbCrLf
sql = sql & "INTOTMATL = b.INTOTMATL, INTOTLABOR = b.INTOTLABOR," & vbCrLf
sql = sql & "INTOTEXP = b.INTOTEXP, INTOTOH = b.INTOTOH," & vbCrLf
sql = sql & "LOTTOTMATL = (c.LOTTOTMATL * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTLABOR = (c.LOTTOTLABOR * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTEXP = (c.LOTTOTEXP * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTOH = (c.LOTTOTOH * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTDATECOSTED = c.LOTDATECOSTED, LOTORGQTY = c.LOTORIGINALQTY," & vbCrLf
sql = sql & "BMTOTOH = (c.LOTTOTOH * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTSPLITFROMSYS = c.LOTSPLITFROMSYS" & vbCrLf
sql = sql & "FROM dbo.LohdTable AS c INNER JOIN" & vbCrLf
sql = sql & "dbo.InvaTable AS b ON c.LOTNUMBER = b.INLOTNUMBER AND c.LOTPARTREF = b.INPART LEFT OUTER JOIN" & vbCrLf
sql = sql & "dbo.#tempMOPartsDetail AS d ON b.INMOPART = d.INMOPART AND b.INPART = d.INPART" & vbCrLf
sql = sql & "WHERE     (b.INMOPART = @MOPart) AND (b.INMORUN = @MORun) AND (b.INTYPE = 10) AND (d.SORTKEYLEVEL = 0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "print 'Update 2:' + cast(getdate() as char(25))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--FROM #tempMOPartsDetail d, InvaTable b, LohdTable c" & vbCrLf
sql = sql & "--WHERE d.INMOPART = b.INMOPART AND d.INPART = b.INPART" & vbCrLf
sql = sql & "-- AND b.INLOTNUMBER = c.LOTnumber" & vbCrLf
sql = sql & "-- and c.lotpartref = b.INPART" & vbCrLf
sql = sql & "-- and b.INMOPART = @MOPart AND b.INMORUN  = @MORun" & vbCrLf
sql = sql & "-- AND b.INTYPE = 10 AND SORTKEYLEVEL = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// set the totals for" & vbCrLf
sql = sql & "SELECT @Maxlevel =  MAX(SORTKEYLEVEL) FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "SET @level  = 1" & vbCrLf
sql = sql & "WHILE (@level <= @Maxlevel )" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curMORun CURSOR  FOR" & vbCrLf
sql = sql & "SELECT DISTINCT INMOPART,INPART" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail WHERE SORTKEYLEVEL = @level" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curMORun" & vbCrLf
sql = sql & "FETCH NEXT FROM curMORun INTO @MOPart, @Part" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @ParentLotNum = LOTNUMBER FROM #tempMOPartsDetail WHERE" & vbCrLf
sql = sql & "INPART = @MOPart AND SORTKEYLEVEL = @level - 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @LotRunNo = LOTMORUNNO, @LotOrgQty = LOTORIGINALQTY" & vbCrLf
sql = sql & "FROM lohdTable where LOTNUMBER = @ParentLotNum" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET INMORUN = @LotRunNo, LOTNUMBER = b.INLOTNUMBER," & vbCrLf
sql = sql & "LOTSPLITFROMSYS = c.LOTSPLITFROMSYS" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail d, InvaTable b, LohdTable c" & vbCrLf
sql = sql & "WHERE d.INMOPART = b.INMOPART AND d.INPART = b.INPART" & vbCrLf
sql = sql & "AND b.INLOTNUMBER = c.LOTnumber" & vbCrLf
sql = sql & "and c.lotpartref = b.INPART" & vbCrLf
sql = sql & "and b.INMOPART = @MOPart AND b.INMORUN  = @LotRunNo" & vbCrLf
sql = sql & "AND b.INTYPE = 10 AND SORTKEYLEVEL = @level" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET PICKQTY = sumqty * -1" & vbCrLf
sql = sql & "FROM" & vbCrLf
sql = sql & "(SELECT SUM(b.INAQTY) sumqty, d.INMOPART mopart, d.INMORUN morun, d.LOTNUMBER lotnum, d.INPART subpart" & vbCrLf
sql = sql & "FROM dbo.LohdTable AS c INNER JOIN" & vbCrLf
sql = sql & "dbo.InvaTable AS b ON c.LOTNUMBER = b.INLOTNUMBER AND c.LOTPARTREF = b.INPART LEFT OUTER JOIN" & vbCrLf
sql = sql & "dbo.#tempMOPartsDetail AS d ON b.INMOPART = d.INMOPART AND b.INPART = d.INPART" & vbCrLf
sql = sql & "WHERE     b.INMOPART = @MOPart AND b.INMORUN  = @LotRunNo AND (b.INTYPE = 10) AND (d.SORTKEYLEVEL = @level)" & vbCrLf
sql = sql & "GROUP BY d.INMOPART, d.INMORUN, d.LOTNUMBER, d.INPART" & vbCrLf
sql = sql & ") as f" & vbCrLf
sql = sql & "WHERE INMOPART = mopart AND INMORUN = morun" & vbCrLf
sql = sql & "AND LOTNUMBER = lotnum AND INPART = subpart" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// update the top level" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET INMORUN = @LotRunNo, LOTNUMBER = b.INLOTNUMBER," & vbCrLf
sql = sql & "LOTUSERLOTID = c.LOTUSERLOTID," & vbCrLf
sql = sql & "INTOTMATL = b.INTOTMATL, INTOTLABOR = b.INTOTLABOR," & vbCrLf
sql = sql & "INTOTEXP = b.INTOTEXP, INTOTOH = b.INTOTOH," & vbCrLf
sql = sql & "LOTTOTMATL = (c.LOTTOTMATL * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTLABOR = (c.LOTTOTLABOR * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTEXP = (c.LOTTOTEXP * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTOH = (c.LOTTOTOH * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTDATECOSTED = c.LOTDATECOSTED, LOTORGQTY = @LotOrgQty," & vbCrLf
sql = sql & "BMTOTOH = (c.LOTTOTOH * PICKQTY) / @LotOrgQty," & vbCrLf
sql = sql & "LOTSPLITFROMSYS = c.LOTSPLITFROMSYS" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail d, InvaTable b, LohdTable c" & vbCrLf
sql = sql & "WHERE d.INMOPART = b.INMOPART AND d.INPART = b.INPART" & vbCrLf
sql = sql & "AND b.INLOTNUMBER = c.LOTnumber" & vbCrLf
sql = sql & "and c.lotpartref = b.INPART" & vbCrLf
sql = sql & "and b.INMOPART = @MOPart AND b.INMORUN  = @LotRunNo" & vbCrLf
sql = sql & "AND b.INTYPE = 10 AND SORTKEYLEVEL = @level" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curMORun INTO @MOPart, @Part" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curMORun" & vbCrLf
sql = sql & "DEALLOCATE curMORun" & vbCrLf
sql = sql & "SET @level = @level + 1" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "print 'Update 2:'+ cast(getdate() as char(25))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curMOSplit CURSOR  FOR" & vbCrLf
sql = sql & "SELECT DISTINCT LOTNUMBER, LOTSPLITFROMSYS, LOTTOTMATL--, LOTTOTLABOR, LOTTOTEXP, LOTTOTOH" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail WHERE LOTSPLITFROMSYS <> ''" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curMOSplit" & vbCrLf
sql = sql & "FETCH NEXT FROM curMOSplit INTO @MOLotNum, @SplitLot, @LotMatl" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "print 'LotSplit LotNum:' + @SplitLot" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF @LotMatl = 0" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SELECT @LotUSpMat = (LOTTOTMATL / LOTORIGINALQTY)," & vbCrLf
sql = sql & "@LotUSpLabor = (LOTTOTLABOR / LOTORIGINALQTY)," & vbCrLf
sql = sql & "@LotUSpExp = (LOTTOTEXP / LOTORIGINALQTY)," & vbCrLf
sql = sql & "@LotUSpOH = (LOTTOTOH / LOTORIGINALQTY)" & vbCrLf
sql = sql & "FROM Lohdtable WHERE LOTNUMBER = @SplitLot" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET LOTTOTMATL = (@LotUSpMat * PICKQTY)," & vbCrLf
sql = sql & "LOTTOTLABOR = (@LotUSpLabor * PICKQTY)," & vbCrLf
sql = sql & "LOTTOTEXP = (@LotUSpExp * PICKQTY)," & vbCrLf
sql = sql & "LOTTOTOH = (@LotUSpOH * PICKQTY)" & vbCrLf
sql = sql & "WHERE LOTNUMBER = @MOLotNum AND LOTSPLITFROMSYS = @SplitLot" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curMOSplit INTO @MOLotNum, @SplitLot, @LotMatl" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curMOSplit" & vbCrLf
sql = sql & "DEALLOCATE curMOSplit" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @level =  MAX(SORTKEYLEVEL) FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "WHILE (@level >= 0 )" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curMODet CURSOR  FOR" & vbCrLf
sql = sql & "--SELECT INPART, LOTTOTMATL, LOTTOTLABOR, LOTTOTEXP , LOTTOTOH FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "-- WHERE INPART = '775345149'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT INMOPART," & vbCrLf
sql = sql & "SUM(IsNull(LOTTOTMATL, 0)), SUM(ISNULL(LOTTOTLABOR,0)) ," & vbCrLf
sql = sql & "Sum (IsNull(LOTTOTEXP, 0)) , SUM(IsNull(BMTOTOH, 0))" & vbCrLf
sql = sql & "From" & vbCrLf
sql = sql & "(SELECT DISTINCT INMOPART,INMORUN,INPART,LOTTOTMATL,LOTTOTLABOR," & vbCrLf
sql = sql & "LOTTOTEXP,LOTTOTOH,SUMTOTMAL,SUMTOTLABOR, SUMTOTEXP, SUMTOTOH,BMTOTOH" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail WHERE SORTKEYLEVEL = @level) as foo" & vbCrLf
sql = sql & "group by INMOPART" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curMODet" & vbCrLf
sql = sql & "FETCH NEXT FROM curMODet INTO @MOPart, @SumTotMat, @SumTotLabor,@SumTotExp,@SumTotOH" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "print 'PartNum : ' + @MOPart" & vbCrLf
sql = sql & "print 'SumTotoh : ' + Convert(varchar(24), @SumTotOH)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET SUMTOTMAL = LOTTOTMATL + @SumTotMat," & vbCrLf
sql = sql & "SUMTOTLABOR = LOTTOTLABOR + @SumTotLabor," & vbCrLf
sql = sql & "SUMTOTEXP = LOTTOTEXP + @SumTotExp, SUMTOTOH = (BMTOTOH + @SumTotOH) * @MOQty ," & vbCrLf
sql = sql & "HASCHILD = 1,PARTSUM = 'TOTAL ' + LTRIM(INPART)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "WHERE INPART = @MOPart AND SORTKEYLEVEL = @level - 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curMODet INTO @MOPart, @SumTotMat, @SumTotLabor,@SumTotExp,@SumTotOH" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curMODet" & vbCrLf
sql = sql & "DEALLOCATE curMODet" & vbCrLf
sql = sql & "SET @level = @level - 1" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// update the Lower level cost detail" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET SUMTOTMAL = LOTTOTMATL, SUMTOTLABOR = LOTTOTLABOR," & vbCrLf
sql = sql & "SUMTOTEXP = LOTTOTEXP, SUMTOTOH = BMTOTOH WHERE HASCHILD IS NULL" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @SumTotMat  = 0" & vbCrLf
sql = sql & "SET @SumTotLabor  = 0" & vbCrLf
sql = sql & "SET @SumTotExp  = 0" & vbCrLf
sql = sql & "SET @SumTotOH  = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// Udpate the Root total" & vbCrLf
sql = sql & "SELECT @SumTotMat = SUM(SUMTOTMAL), @SumTotLabor = SUM(SUMTOTLABOR)," & vbCrLf
sql = sql & "@SumTotExp = SUM(SUMTOTEXP) ,@SumTotOH = SUM(SUMTOTOH)" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail WHERE SORTKEYLEVEL = 0" & vbCrLf
sql = sql & "AND  RTRIM(INMOPART) <> RTRIM(INPART)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET PARTSUM = INPART" & vbCrLf
sql = sql & "WHERE PARTSUM IS NULL" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "----  SELECT * FROM #tempMOPartsDetail WHERE SORTKEYLEVEL = 0" & vbCrLf
sql = sql & "----AND  RTRIM(INMOPART) = RTRIM(INPART)" & vbCrLf
sql = sql & "--// Reverse the partnumbers." & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @level = 0" & vbCrLf
sql = sql & "SET @RowCount = 1" & vbCrLf
sql = sql & "SET @PrevParent = ''" & vbCrLf
sql = sql & "DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
sql = sql & "SELECT INMOPART, SortKey" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "WHERE HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLMASTER ='AA10'--GLTOPMaster = 1 AND" & vbCrLf
sql = sql & "ORDER BY SortKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curAcctStruc" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc INTO @ParentPart, @SortKey" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "if (@PrevParent <> @ParentPart)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET" & vbCrLf
sql = sql & "SortKeyRev = RIGHT ('0000'+ Cast(@RowCount as varchar), 4) + char(36)+ @ParentPart" & vbCrLf
sql = sql & "WHERE INMOPART = @ParentPart AND HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLTOPMaster = 1" & vbCrLf
sql = sql & "SET @RowCount = @RowCount + 1" & vbCrLf
sql = sql & "SET @PrevParent = @ParentPart" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc INTO @ParentPart, @SortKey" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "Close curAcctStruc" & vbCrLf
sql = sql & "DEALLOCATE curAcctStruc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET LOTMOPARTRUNKEY = @MOPartRunKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO EsMOPartsCostDetail (LOTMOPARTRUNKEY, INMOPART,INMORUN,INPART,PARTSUM,LOTNUMBER,LOTUSERLOTID," & vbCrLf
sql = sql & "LOTTOTMATL,SUMTOTMAL, LOTTOTLABOR,SUMTOTLABOR, LOTTOTEXP, SUMTOTEXP, LOTTOTOH,SUMTOTOH,BMTOTOH," & vbCrLf
sql = sql & "LOTDATECOSTED, BMQTYREQD, LOTORGQTY, SORTKEYLEVEL,SortKey,SortKeyRev,HASCHILD, PICKQTY)" & vbCrLf
sql = sql & "SELECT LOTMOPARTRUNKEY,INMOPART,INMORUN,INPART,PARTSUM,LOTNUMBER,LOTUSERLOTID," & vbCrLf
sql = sql & "LOTTOTMATL,SUMTOTMAL, LOTTOTLABOR,SUMTOTLABOR, LOTTOTEXP, SUMTOTEXP, LOTTOTOH,SUMTOTOH,BMTOTOH," & vbCrLf
sql = sql & "LOTDATECOSTED, BMQTYREQD, LOTORGQTY, SORTKEYLEVEL,SortKey,SortKeyRev,HASCHILD, PICKQTY" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail--WHERE SORTKEYLEVEL = 1" & vbCrLf
sql = sql & "order by SortKey--SortKeyRev" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DROP table #tempMOPartsDetail" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript False, sql

sql = "create procedure DropFunctionIfExists" & vbCrLf
sql = sql & "@Function_Name varchar(50)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* drop a function if it exists" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "DropFunctionIfExists 'WCHoursBeforeTime'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "IF EXISTS (" & vbCrLf
sql = sql & "SELECT * FROM sysobjects WHERE id = object_id(@Function_Name)" & vbCrLf
sql = sql & "AND xtype IN (N'FN', N'IF', N'TF')" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "declare @sql varchar(100)" & vbCrLf
sql = sql & "set @sql = 'DROP FUNCTION ' + @Function_Name" & vbCrLf
sql = sql & "execute(@sql)" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql

sql = "DropFunctionIfExists 'WCHoursBeforeTime'" & vbCrLf
ExecuteScript False, sql

sql = "create function dbo.WCHoursBeforeTime" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@shop varchar(12)," & vbCrLf
sql = sql & "@wc varchar(12)," & vbCrLf
sql = sql & "@cutoffDateTime datetime" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "returns decimal(9,2)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "get remaining WC hours for day < the start time of the next operation in MO op scheduling" & vbCrLf
sql = sql & "Author: Terry Lindeman" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "select dbo.WCHoursBeforeTime( 'ST','CPSH','10/27/2017 2:00 PM' )" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "-- if no hours in calendar, return" & vbCrLf
sql = sql & "declare @date datetime, @fullWcHours decimal(9,4), @hours decimal(9,4)" & vbCrLf
sql = sql & "declare @Hours1 decimal(9,4),@Hours2 decimal(9,4),@Hours3 decimal(9,4),@Hours4 decimal(9,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @date = cast(CONVERT(varchar(10),@cutoffDateTime,101) as datetime) -- remove time portion" & vbCrLf
sql = sql & "select @Hours1 = isnull(sum(WCCSHH1),0)," & vbCrLf
sql = sql & "@Hours2 = isnull(sum(WCCSHH2),0)," & vbCrLf
sql = sql & "@Hours3 = isnull(sum(WCCSHH3),0)," & vbCrLf
sql = sql & "@Hours4 = isnull(sum(WCCSHH4),0)" & vbCrLf
sql = sql & "from WcclTable" & vbCrLf
sql = sql & "where WCCDATE = @date and WCCSHOP = @shop and WCCCENTER = @wc" & vbCrLf
sql = sql & "set @hours = @Hours1 + @Hours2 + @Hours3 + @Hours4" & vbCrLf
sql = sql & "if @hours = 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "return @hours" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @day varchar(3)" & vbCrLf
sql = sql & "set @day = UPPER(DATENAME(weekday,@cutoffDateTime))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get shift start times" & vbCrLf
sql = sql & "declare @Start1 varchar(6),@Start2 varchar(6),@Start3 varchar(6),@Start4 varchar(6)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @day = 'SUN'" & vbCrLf
sql = sql & "select @Start1 = WCNSUNSH1, @Start2 = WCNSUNSH2, @Start3 = WCNSUNSH3, @Start4 = WCNSUNSH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else if @day = 'MON'" & vbCrLf
sql = sql & "select @Start1 = WCNMONSH1, @Start2 = WCNMONSH2, @Start3 = WCNMONSH3, @Start4 = WCNMONSH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else if @day = 'TUE'" & vbCrLf
sql = sql & "select @Start1 = WCNTUESH1, @Start2 = WCNTUESH2, @Start3 = WCNTUESH3, @Start4 = WCNTUESH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else if @day = 'WED'" & vbCrLf
sql = sql & "select @Start1 = WCNWEDSH1, @Start2 = WCNWEDSH2, @Start3 = WCNWEDSH3, @Start4 = WCNWEDSH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else if @day = 'THU'" & vbCrLf
sql = sql & "select @Start1 = WCNTHUSH1, @Start2 = WCNTHUSH2, @Start3 = WCNTHUSH3, @Start4 = WCNTHUSH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else if @day = 'FRI'" & vbCrLf
sql = sql & "select @Start1 = WCNFRISH1, @Start2 = WCNFRISH2, @Start3 = WCNFRISH3, @Start4 = WCNFRISH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "select @Start1 = WCNSATSH1, @Start2 = WCNSATSH2, @Start3 = WCNSATSH3, @Start4 = WCNSATSH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @temp table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "ShiftNo int," & vbCrLf
sql = sql & "ShiftHours decimal(9,4)," & vbCrLf
sql = sql & "ShiftStart varchar(6)," & vbCrLf
sql = sql & "StartTime datetime," & vbCrLf
sql = sql & "MinutesBeforeCutoff decimal(9,4)" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @temp (ShiftNo, ShiftHours, ShiftStart) values ( 1, @Hours1, @Start1 )" & vbCrLf
sql = sql & "insert @temp (ShiftNo, ShiftHours, ShiftStart) values ( 1, @Hours2, @Start2 )" & vbCrLf
sql = sql & "insert @temp (ShiftNo, ShiftHours, ShiftStart) values ( 1, @Hours3, @Start3 )" & vbCrLf
sql = sql & "insert @temp (ShiftNo, ShiftHours, ShiftStart) values ( 1, @Hours4, @Start4 )" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @temp set ShiftStart = rtrim(ShiftStart)" & vbCrLf
sql = sql & "delete from @temp where ShiftHours = 0 or ShiftStart = ''" & vbCrLf
sql = sql & "update @temp set StartTime = cast(CONVERT(varchar(10), @date, 101) + ' ' + ShiftStart + 'm' as datetime)" & vbCrLf
sql = sql & "update @temp set MinutesBeforeCutoff = DATEDIFF(minute, StartTime, @cutoffDateTime)" & vbCrLf
sql = sql & "delete from @temp where MinutesBeforeCutoff < 0" & vbCrLf
sql = sql & "update @temp set MinutesBeforeCutoff = 60 * ShiftHours where MinutesBeforeCutoff > 60 * ShiftHours" & vbCrLf
sql = sql & "select @hours = isnull(sum(MinutesBeforeCutoff),0) / 60. from @temp" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "return @hours" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql




''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function

Private Function UpdateDatabase96()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 171
   If ver < newver Then

        clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "DropFunctionIfExists 'WCHoursBeforeTime'" & vbCrLf
ExecuteScript False, sql

sql = "create function dbo.WCHoursBeforeTime" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@shop varchar(12)," & vbCrLf
sql = sql & "@wc varchar(12)," & vbCrLf
sql = sql & "@cutoffDateTime datetime" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "returns decimal(9,2)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "get remaining WC hours for day < the start time of the next operation in MO op scheduling" & vbCrLf
sql = sql & "Author: Terry Lindeman" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "select dbo.WCHoursBeforeTime( 'ST','CPSH','10/27/2017 2:00 PM' ) -- LAFEGG" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "-- if no hours in calendar, return" & vbCrLf
sql = sql & "declare @date datetime, @fullWcHours decimal(9,4), @hours decimal(9,4)" & vbCrLf
sql = sql & "declare @Hours1 decimal(9,4),@Hours2 decimal(9,4),@Hours3 decimal(9,4),@Hours4 decimal(9,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @date = cast(CONVERT(varchar(10),@cutoffDateTime,101) as datetime) -- remove time portion" & vbCrLf
sql = sql & "select @Hours1 = isnull(sum(WCCSHH1),0)," & vbCrLf
sql = sql & "@Hours2 = isnull(sum(WCCSHH2),0)," & vbCrLf
sql = sql & "@Hours3 = isnull(sum(WCCSHH3),0)," & vbCrLf
sql = sql & "@Hours4 = isnull(sum(WCCSHH4),0)" & vbCrLf
sql = sql & "from WcclTable" & vbCrLf
sql = sql & "where WCCDATE = @date and WCCSHOP = @shop and WCCCENTER = @wc" & vbCrLf
sql = sql & "set @hours = @Hours1 + @Hours2 + @Hours3 + @Hours4" & vbCrLf
sql = sql & "if @hours = 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "return @hours" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @day varchar(3)" & vbCrLf
sql = sql & "set @day = UPPER(DATENAME(weekday,@cutoffDateTime))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get shift start times" & vbCrLf
sql = sql & "declare @Start1 varchar(6),@Start2 varchar(6),@Start3 varchar(6),@Start4 varchar(6)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @day = 'SUN'" & vbCrLf
sql = sql & "select @Start1 = WCNSUNSH1, @Start2 = WCNSUNSH2, @Start3 = WCNSUNSH3, @Start4 = WCNSUNSH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else if @day = 'MON'" & vbCrLf
sql = sql & "select @Start1 = WCNMONSH1, @Start2 = WCNMONSH2, @Start3 = WCNMONSH3, @Start4 = WCNMONSH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else if @day = 'TUE'" & vbCrLf
sql = sql & "select @Start1 = WCNTUESH1, @Start2 = WCNTUESH2, @Start3 = WCNTUESH3, @Start4 = WCNTUESH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else if @day = 'WED'" & vbCrLf
sql = sql & "select @Start1 = WCNWEDSH1, @Start2 = WCNWEDSH2, @Start3 = WCNWEDSH3, @Start4 = WCNWEDSH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else if @day = 'THU'" & vbCrLf
sql = sql & "select @Start1 = WCNTHUSH1, @Start2 = WCNTHUSH2, @Start3 = WCNTHUSH3, @Start4 = WCNTHUSH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else if @day = 'FRI'" & vbCrLf
sql = sql & "select @Start1 = WCNFRISH1, @Start2 = WCNFRISH2, @Start3 = WCNFRISH3, @Start4 = WCNFRISH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "select @Start1 = WCNSATSH1, @Start2 = WCNSATSH2, @Start3 = WCNSATSH3, @Start4 = WCNSATSH4" & vbCrLf
sql = sql & "from WcntTable  where WCNSHOP = @shop and WCNREF = @wc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @temp table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "ShiftNo int," & vbCrLf
sql = sql & "ShiftHours decimal(9,4)," & vbCrLf
sql = sql & "ShiftStart varchar(6)," & vbCrLf
sql = sql & "StartTime datetime," & vbCrLf
sql = sql & "MinutesBeforeCutoff decimal(9,4)" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @temp (ShiftNo, ShiftHours, ShiftStart) values ( 1, @Hours1, @Start1 )" & vbCrLf
sql = sql & "insert @temp (ShiftNo, ShiftHours, ShiftStart) values ( 1, @Hours2, @Start2 )" & vbCrLf
sql = sql & "insert @temp (ShiftNo, ShiftHours, ShiftStart) values ( 1, @Hours3, @Start3 )" & vbCrLf
sql = sql & "insert @temp (ShiftNo, ShiftHours, ShiftStart) values ( 1, @Hours4, @Start4 )" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @temp set ShiftStart = rtrim(ShiftStart)" & vbCrLf
sql = sql & "delete from @temp where ShiftHours = 0 or ShiftStart = ''" & vbCrLf
sql = sql & "update @temp set StartTime = cast(CONVERT(varchar(10), @date, 101) + ' ' + ShiftStart + 'm' as datetime)" & vbCrLf
sql = sql & "update @temp set MinutesBeforeCutoff = DATEDIFF(minute, StartTime, @cutoffDateTime)" & vbCrLf
sql = sql & "delete from @temp where MinutesBeforeCutoff < 0" & vbCrLf
sql = sql & "update @temp set MinutesBeforeCutoff = 60 * ShiftHours where MinutesBeforeCutoff > 60 * ShiftHours" & vbCrLf
sql = sql & "select @hours = isnull(sum(MinutesBeforeCutoff),0) / 60. from @temp" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "return @hours" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

        ' update the version
        ExecuteScript False, "Update Version Set Version = " & newver

    End If
End Function

Private Function UpdateDatabase97()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 172     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

      sql = "AddOrUpdateColumn 'SohdTable', 'SOITAREAR', 'bit null default 0'"
      ExecuteScript False, sql
      
      sql = "Update SohdTable set SOITAREAR = 0"
      ExecuteScript False, sql
      
      sql = "AddOrUpdateColumn 'SohdTable', 'SOITAREAR', 'bit not null'"
      ExecuteScript False, sql
      
      sql = "DropStoredProcedureIfExists 'AddEngrTimeCharge'" & vbCrLf
      ExecuteScript False, sql

sql = "create procedure AddEngrTimeCharge" & vbCrLf
sql = sql & "@EmpNo as int," & vbCrLf
sql = sql & "@Date as datetime," & vbCrLf
sql = sql & "@MoPartRef as varchar(30)," & vbCrLf
sql = sql & "@RunNo as int," & vbCrLf
sql = sql & "@OpNo as int," & vbCrLf
sql = sql & "@Hours as decimal(12,2)," & vbCrLf
sql = sql & "@Comment as varchar(1024)," & vbCrLf
sql = sql & "@journalId as varchar(12)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "/* add a time charge, creating a new time card if required" & vbCrLf
sql = sql & "9/15/17 TEL - created for CASGAS Engineering Time Charges" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if time card does not exist for employee, add it" & vbCrLf
sql = sql & "declare @card char(11)" & vbCrLf
sql = sql & "select @card = TMCARD from TchdTable where TMEMP = @EmpNo and TMDAY = @date" & vbCrLf
sql = sql & "if @card is null" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "declare @now as datetime, @nowDays int, @nowMs int, @today datetime, @time datetime" & vbCrLf
sql = sql & "set @now = getdate()" & vbCrLf
sql = sql & "set @today = cast(convert(varchar(10), @now,101) as datetime)" & vbCrLf
sql = sql & "set @time = cast(convert(varchar(50), @now, 114) as datetime)" & vbCrLf
sql = sql & "set @nowDays = DATEDIFF(DAY,'1/1/1900',@today)" & vbCrLf
sql = sql & "set @nowMs = 1000000.0 *cast(DATEDIFF(MILLISECOND,'1/1/1900',@time) as float)/(3600.0*24*1000)" & vbCrLf
sql = sql & "set @card = cast(@nowDays as varchar(5)) + cast(@nowMs as varchar(6))" & vbCrLf
sql = sql & "insert into TchdTable (TMCARD,TMEMP,TMDAY) values (@card, @EmpNo, @Date)" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- calculate fields needed for time charge" & vbCrLf
sql = sql & "declare @startTime smalldatetime, @stopTime smalldatetime" & vbCrLf
sql = sql & "set @startTime = DATEADD(dd, DATEDIFF(dd, 0, @Date), 0)    -- truncate time portion" & vbCrLf
sql = sql & "set @stopTime = DATEADD(MINUTE, @Hours * 60, @startTime)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- construct hh:mma/p version of times" & vbCrLf
sql = sql & "declare @start varchar(6), @stop varchar(6)" & vbCrLf
sql = sql & "set @start = REPLACE(SUBSTRING(CONVERT(VARCHAR(20),@startTime,0),13,6),' ','0')" & vbCrLf
sql = sql & "set @stop = REPLACE(SUBSTRING(CONVERT(VARCHAR(20),@stopTime,0),13,6),' ','0')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- calculate # minutes" & vbCrLf
sql = sql & "set @time = dateadd(minute,datediff(MINUTE,@startTime, @stopTime),'1/1/1900')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get regular timecode" & vbCrLf
sql = sql & "declare @timecode varchar(2)" & vbCrLf
sql = sql & "select @timecode = TYPECODE from TmcdTable where typetype = 'R'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get engineering rate" & vbCrLf
sql = sql & "declare @rate decimal(10,2)" & vbCrLf
sql = sql & "select @rate = EngineeringLaborRate from Preferences" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get employee account number" & vbCrLf
sql = sql & "declare @acct varchar(12)" & vbCrLf
sql = sql & "select @acct = EmplTable.PREMACCTS from EmplTable where PREMNUMBER = @EmpNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get shop and wc for operation" & vbCrLf
sql = sql & "declare @shop varchar(12), @wc varchar(12)" & vbCrLf
sql = sql & "select @shop = OPSHOP, @wc = OPCENTER" & vbCrLf
sql = sql & "from RnopTable where opref = @MoPartRef and oprun = @RunNo and OPNO = @OpNo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get journal id" & vbCrLf
sql = sql & "--declare @journalId varchar(12)" & vbCrLf
sql = sql & "--set @journalId = dbo.fnGetOpenJournalID('TJ', @Date)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now insert the new time charge" & vbCrLf
sql = sql & "INSERT INTO TcitTable (TCCARD,TCEMP,TCSTART,TCSTOP,TCSTARTTIME,TCSTOPTIME," & vbCrLf
sql = sql & "TCHOURS,TCTIME,TCCODE,TCRATE,TCOHRATE,TCRATENO,TCACCT,TCACCOUNT," & vbCrLf
sql = sql & "TCSHOP,TCWC,TCPAYTYPE,TCSURUN,TCYIELD,TCPARTREF,TCRUNNO," & vbCrLf
sql = sql & "TCOPNO,TCSORT,TCOHFIXED,TCGLJOURNAL,TCGLREF,TCSOURCE," & vbCrLf
sql = sql & "TCMULTIJOB,TCACCEPT,TCREJECT,TCSCRAP,TCCOMMENTS)" & vbCrLf
sql = sql & "values( @card,@EmpNo, @start,@stop,@startTime, @stopTime," & vbCrLf
sql = sql & "@Hours,@time,@timecode,@rate,@rate,1,@acct,@acct," & vbCrLf
sql = sql & "@shop,@wc,0,'I',0,@MoPartRef,@RunNo," & vbCrLf
sql = sql & "@OpNo,0,0,@journalId,0,'Engr'," & vbCrLf
sql = sql & "0,0,0,0,@Comment)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now roll up totals for this timecard" & vbCrLf
sql = sql & "EXECUTE UpdateTimeCardTotals @EmpNo, @Date" & vbCrLf
ExecuteScript False, sql

      

''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase98()

   Dim sql As String
   sql = ""

   newver = 173     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''


' ANDELE PO HISTORY BY VENDOR - used by prdpr21.rpt, Purchasing History by Vendor to Date Costing

sql = "dbo.DropStoredProcedureIfExists 'RptPoHistoryByVendor'" & vbCrLf
ExecuteScript False, sql

sql = "create procedure [dbo].[RptPoHistoryByVendor]" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@StartDate varchar(16)," & vbCrLf
sql = sql & "@EndDate varchar(16)," & vbCrLf
sql = sql & "@Vendor varchar(10)," & vbCrLf
sql = sql & "@PartRef varchar(30)," & vbCrLf
sql = sql & "@PartClass varchar(4)," & vbCrLf
sql = sql & "@ProdCode varchar(6)," & vbCrLf
sql = sql & "@UsePoDate integer,       -- if 0, use projected date (PIPDATE), if 1, use PO date (PODATE)" & vbCrLf
sql = sql & "@IncludeOpen14 integer," & vbCrLf
sql = sql & "@IncludeReceived15 integer," & vbCrLf
sql = sql & "@IncludeCanceled16 integer," & vbCrLf
sql = sql & "@IncludeInvoiced17 integer" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* created for ANDELE version of the report" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "RptPoHistoryByVendor '9/1/2017','12/31/2017','','','','',1,1,1,1,1" & vbCrLf
sql = sql & "RptPoHistoryByVendor '9/1/2017','12/31/2017','','85','','',1,1,1,1,1" & vbCrLf
sql = sql & "RptPoHistoryByVendor '9/1/2017','12/4/2017','ACOPIAN','','','',1,1,1,1,1" & vbCrLf
sql = sql & "RptPoHistoryByVendor 'ALL','ALL','ACOPIAN','','','',1,1,1,1,1" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @Start datetime" & vbCrLf
sql = sql & "declare @End datetime" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @StartDate = 'ALL' set @Start = '1/1/1900'" & vbCrLf
sql = sql & "else set  @Start = cast(@StartDate as datetime)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @EndDate = 'ALL' set @End = '12/31/2199'" & vbCrLf
sql = sql & "else set @End = cast(@EndDate as datetime)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "print @Start" & vbCrLf
sql = sql & "print @End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @Vendor = rtrim(@Vendor)" & vbCrLf
sql = sql & "set @PartRef = rtrim(@PartRef)" & vbCrLf
sql = sql & "set @PartClass = rtrim(@PartClass)" & vbCrLf
sql = sql & "set @ProdCode = rtrim(@ProdCode)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @Vendor = 'ALL'set @Vendor = '%' else if len(@Vendor) < 10 set @Vendor = @Vendor + '%'" & vbCrLf
sql = sql & "if @PartRef = 'ALL' set @PartRef = '%' else if len(@PartRef) < 30 set @PartRef = @PartRef + '%'" & vbCrLf
sql = sql & "if @PartClass = 'ALL' set @PartClass = '%' else if len(@PartClass) < 4 set @PartClass = @PartClass + '%'" & vbCrLf
sql = sql & "if @ProdCode = 'ALL' set @ProdCode = '%' else if len(@ProdCode) < 6 set @ProdCode = @ProdCode + '%'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @IncludeOpen14 = 1 set @IncludeOpen14 = 14 else set @IncludeOpen14 = 0" & vbCrLf
sql = sql & "if @IncludeReceived15 = 1 set @IncludeReceived15 = 15  else set @IncludeReceived15 = 0" & vbCrLf
sql = sql & "if @IncludeCanceled16 = 1 set @IncludeCanceled16 = 16  else set @IncludeCanceled16 = 0" & vbCrLf
sql = sql & "if @IncludeInvoiced17 = 1 set @IncludeInvoiced17 = 17  else set @IncludeInvoiced17 = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT distinct ViitTable.VITCOST, InvaTable.INAMT, PoitTable.PIESTUNIT," & vbCrLf
sql = sql & "coalesce(ViitTable.VITCOST, InvaTable.INAMT, PoitTable.PIESTUNIT) as UnitCost," & vbCrLf
sql = sql & "rtrim(VEBCITY) + ', ' + RTRIM(VEBSTATE) + ' ' + RTRIM(VEBZIP) as Address," & vbCrLf
sql = sql & "PoitTable.PIPQTY, PohdTable.POVENDOR, PoitTable.PINUMBER, PoitTable.PIITEM, PoitTable.PIREV," & vbCrLf
sql = sql & "PartTable.PARTNUM, PoitTable.PITYPE, PartTable.PAUNITS, PohdTable.PODATE, PoitTable.PIPDATE," & vbCrLf
sql = sql & "VndrTable.VENUMBER, VndrTable.VENICKNAME, VndrTable.VEBPHONE, VndrTable.VEBNAME, PoitTable.PICOMT," & vbCrLf
sql = sql & "PartTable.PAEXTDESC, PartTable.PADESC, PoitTable.PIAMT, PohdTable.PONUMBER," & vbCrLf
sql = sql & "case when @UsePoDate = 1 then PohdTable.PODATE else PoitTable.PIPDATE end as TranDate," & vbCrLf
sql = sql & "case PITYPE when 14 then 'O' when 15 then 'R' when 16 then 'C' else 'I' end as tp" & vbCrLf
sql = sql & "FROM   ((PartTable PartTable" & vbCrLf
sql = sql & "INNER JOIN PoitTable PoitTable ON PartTable.PARTREF=PoitTable.PIPART)" & vbCrLf
sql = sql & "INNER JOIN PohdTable PohdTable ON PoitTable.PINUMBER=PohdTable.PONUMBER)" & vbCrLf
sql = sql & "INNER JOIN VndrTable VndrTable ON PohdTable.POVENDOR=VndrTable.VEREF" & vbCrLf
sql = sql & "LEFT JOIN InvaTable on InvaTable.INPONUMBER = PoitTable.PINUMBER and InvaTable.INPOITEM = PoitTable.PIITEM and InvaTable.INPOREV = PoitTable.PIREV" & vbCrLf
sql = sql & "LEFT JOIN ViitTable on ViitTable.VITPO = PoitTable.PINUMBER and ViitTable.VITPOITEM = PoitTable.PIITEM and ViitTable.VITPOITEMREV = PoitTable.PIREV" & vbCrLf
sql = sql & "where PIPDATE between @Start and @End" & vbCrLf
sql = sql & "and PITYPE in (@IncludeOpen14, @IncludeReceived15, @IncludeCanceled16, @IncludeInvoiced17)" & vbCrLf
sql = sql & "and VndrTable.VEREF like @Vendor and PartTable.PAPRODCODE like @ProdCode" & vbCrLf
sql = sql & "and PartTable.PACLASS like @PartClass" & vbCrLf
sql = sql & "and PARTREF like @PartRef" & vbCrLf
sql = sql & "ORDER BY PohdTable.POVENDOR, PONUMBER" & vbCrLf
ExecuteScript False, sql

sql = "DropStoredProcedureIfExists 'RptLaborEfficiency'" & vbCrLf
ExecuteScript False, sql

sql = "create procedure dbo.RptLaborEfficiency" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@RunPart varchar(30)," & vbCrLf
sql = sql & "@RunNo int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "Get Labor Efficiency Report Data for Crystal Report, EngRt08.rpt" & vbCrLf
sql = sql & "Created 1/5/2018 TEL" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "exec RptLaborEfficiency '315W1582-1', 67" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @RunRef varchar(30)" & vbCrLf
sql = sql & "set @RunRef = dbo.fnCompress(@RunPart)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select ISNULL(OPNO,TCOPNO) as OP, ISNULL(OPSHOP,TCSHOP) as Shop, ISNULL(OPCENTER,TCWC) as WC," & vbCrLf
sql = sql & "[Rtg SU], [Rtg Run], isnull([TC SU], 0.00) as [TC SU], isnull([TC Run],0.0) as [TC Run]," & vbCrLf
sql = sql & "isnull([Emp No],'') as [Emp No], isnull([Emp Name],'') as [Emp Name], isnull([Date],'') as [Date]," & vbCrLf
sql = sql & "RANK() OVER (PARTITION BY ISNULL(OPNO,TCOPNO) ORDER BY TCSTARTTIME) as Seq, 1 as [Group]" & vbCrLf
sql = sql & "from" & vbCrLf
sql = sql & "(select TCSHOP, TCWC, TCOPNO," & vbCrLf
sql = sql & "cast(TCEMP as varchar(8)) as [Emp No], CONVERT(varchar(10), TCSTARTTIME, 101) as [Date]," & vbCrLf
sql = sql & "case when TCSURUN = 'S' then cast(TCHOURS as decimal(12,2)) else 0.00 end as [TC SU]," & vbCrLf
sql = sql & "case when TCSURUN <> 'S' then cast(TCHOURS as decimal(12,2)) else 0.00 end as [TC Run]," & vbCrLf
sql = sql & "rtrim(PREMFSTNAME) + ' ' + rtrim(PREMLSTNAME) as [Emp Name], TCSTARTTIME" & vbCrLf
sql = sql & "from TcitTable" & vbCrLf
sql = sql & "join EmplTable on PREMNUMBER = TCEMP" & vbCrLf
sql = sql & "where TCPARTREF = @RunRef and TCRUNNO = @RunNo) tm" & vbCrLf
sql = sql & "full outer join (select OPSHOP, OPCENTER, OPNO, Cast(OPSETUP as decimal(12,2)) as [Rtg SU], OPUNIT, RUNQTY, cast(OPUNIT * RUNQTY as decimal(12,2)) as [Rtg Run]" & vbCrLf
sql = sql & "from RtopTable join RunsTable on OPREF = RUNREF and RUNNO = @RunNo where OPREF = @RunRef) op" & vbCrLf
sql = sql & "on tm.TCOPNO = op.OPNO" & vbCrLf
sql = sql & "order by ISNULL(OPNO,TCOPNO), [Date]" & vbCrLf
ExecuteScript False, sql

' script to correctly associate sheet inventory lots and invatable rows
' in a prior update SQL2005 could not handle order by INNUMBER.  It was changed to order by ia.INNUMBER
' first update duplicate INNUMBERS resulting from errors in SheetPick and SheetRestock sp's
sql = "if exists (select 1 from ComnTable where COUSESHEETINVENTORY = 1)" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "declare @dups TABLE (INNUMBER int)" & vbCrLf
sql = sql & "insert @dups" & vbCrLf
sql = sql & "select innumber from invatable" & vbCrLf
sql = sql & "where INPART in (select partref from PartTable where PAPUNITS = 'SH')" & vbCrLf
sql = sql & "group by innumber having count(*) > 1" & vbCrLf
sql = sql & "declare @maxINNUMBER int" & vbCrLf
sql = sql & "select @maxINNUMBER = max(INNUMBER) FROM InvaTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @dups2 table (INNUMBER int, Qty decimal(12,2), newINNUMBER int)" & vbCrLf
sql = sql & "insert @dups2" & vbCrLf
sql = sql & "select ia.INNUMBER, INAQTY, @maxINNUMBER + ROW_NUMBER() OVER (ORDER BY ia.INNUMBER, ia.INAQTY)" & vbCrLf
sql = sql & "from @dups join InvaTable ia on ia.INNUMBER = [@dups].INNUMBER" & vbCrLf
sql = sql & "where INAQTY <> (select min(INAQTY) from @dups join InvaTable ia2 on ia2.INNUMBER = [@dups].INNUMBER where ia2.INNUMBER = ia.INNUMBER)" & vbCrLf
sql = sql & "order by ia.INNUMBER, INAQTY" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ia" & vbCrLf
sql = sql & "set ia.INNUMBER = d.newINNUMBER" & vbCrLf
sql = sql & "from InvaTable ia join @dups2 d on d.INNUMBER = ia.INNUMBER and d.Qty = ia.INAQTY" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now find matching lot item records and point them at the appropriate INNUMBERS" & vbCrLf
sql = sql & "update  loi" & vbCrLf
sql = sql & "set LOIACTIVITY = INNUMBER" & vbCrLf
sql = sql & "from LoitTable loi join InvaTable ia on ia.INLOTNUMBER = loi.LOINUMBER" & vbCrLf
sql = sql & "and ia.INAQTY = loi.LOIQUANTITY" & vbCrLf
sql = sql & "and ia.INADATE = loi.LOIADATE" & vbCrLf
sql = sql & "where LOIPARTREF in (select partref from PartTable where PAPUNITS = 'SH')" & vbCrLf
sql = sql & "and LOIACTIVITY is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "commit tran" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql

sql = "dropstoredprocedureifexists 'InvMrPExcessReport'" & vbCrLf
ExecuteScript False, sql

sql = "create procedure InvMRPExcessReport" & vbCrLf
sql = sql & "@PartClass as Varchar(16), @PartCode as varchar(8), @PartType1 as Integer," & vbCrLf
sql = sql & "@PartType2 as Integer, @PartType3 as Integer, @PartType4 as Integer" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "sp for Excess Inventory Report, InvExcess.rpt" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "InvMRPExcessReport '', '', 1, 1, 1, 1" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@PartClass = 'ALL')" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @PartClass = ''" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartCode = 'ALL')" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @PartCode = ''" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "IF (@PartType1 = 1)" & vbCrLf
sql = sql & "SET @PartType1 = 1" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType1 = 0" & vbCrLf
sql = sql & "IF (@PartType2 = 1)" & vbCrLf
sql = sql & "SET @PartType2 = 2" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType2 = 0" & vbCrLf
sql = sql & "IF (@PartType3 = 1)" & vbCrLf
sql = sql & "SET @PartType3 = 3" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType3 = 0" & vbCrLf
sql = sql & "IF (@PartType4 = 1)" & vbCrLf
sql = sql & "SET @PartType4 = 4" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType4 = 0" & vbCrLf
sql = sql & "CREATE TABLE #tempMrpExRpt" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "PACLASS varchar(4) NULL ," & vbCrLf
sql = sql & "PAPRODCODE varchar(6) NULL ," & vbCrLf
sql = sql & "PALEVEL tinyint NULL ," & vbCrLf
sql = sql & "PARTREF varchar(30) NULL ," & vbCrLf
sql = sql & "PARTNUM varchar(30) NULL ," & vbCrLf
sql = sql & "PADESC varchar(30) NULL ," & vbCrLf
sql = sql & "PAEXTDESC varchar(3072) NULL ," & vbCrLf
sql = sql & "LOTNUMBER varchar(15) NULL," & vbCrLf
sql = sql & "LOTUSERLOTID varchar(40) NULL," & vbCrLf
sql = sql & "MRP_QTYREM int NULL," & vbCrLf
sql = sql & "LOTUNITCOST decimal(12,4) NULL ," & vbCrLf
sql = sql & "PASTDCOST decimal(12,4) NULL ," & vbCrLf
sql = sql & "PAUSEACTUALCOST tinyint NULL ," & vbCrLf
sql = sql & "PALOTTRACK tinyint NULL," & vbCrLf
sql = sql & "MRP_ACTIVITY tinyint NULL," & vbCrLf
sql = sql & "Row int NULL," & vbCrLf
sql = sql & "LOTREMAININGQTY int NULL," & vbCrLf
sql = sql & "PriorLotQty int null," & vbCrLf
sql = sql & "QtyToShow int NULL" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "SELECT mrp_partref, SUM(mrp_partqtyrqd) as rem, cast(0 as bit) as active" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "from MrplTable" & vbCrLf
sql = sql & "WHERE mrp_type NOT IN ('5', '7')" & vbCrLf
sql = sql & "GROUP BY mrp_partref Having Sum(mrp_partqtyrqd) >= 1" & vbCrLf
sql = sql & "order by MRP_PARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #temp set Active = 1 where exists (select 1 from MrplTable x where x.mrp_partref = #temp.MRP_PARTREF and mrp_type NOT IN ('1','17', '5', '6'))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO #tempMrpExRpt (PACLASS, PAPRODCODE, PALEVEL," & vbCrLf
sql = sql & "PARTREF, PARTNUM, PADESC, PAEXTDESC, LOTNUMBER, LOTUSERLOTID," & vbCrLf
sql = sql & "MRP_QTYREM, LOTUNITCOST, PASTDCOST, PAUSEACTUALCOST, PALOTTRACK, MRP_ACTIVITY," & vbCrLf
sql = sql & "Row,LOTREMAININGQTY,PriorLotQty,QtyToShow )" & vbCrLf
sql = sql & "SELECT PACLASS, PAPRODCODE, PALEVEL, PARTREF, PARTNUM, PADESC," & vbCrLf
sql = sql & "PAEXTDESC, LOTNUMBER, LOTUSERLOTID, #temp.rem,LOTUNITCOST," & vbCrLf
sql = sql & "PASTDCOST , PAUSEACTUALCOST, PALOTTRACK, #temp.active," & vbCrLf
sql = sql & "row_number() over (partition by PARTREF order by LOTADATE desc)," & vbCrLf
sql = sql & "LOTREMAININGQTY,0,0" & vbCrLf
sql = sql & "From ViewLohdPartTable join #temp on #temp.MRP_PARTREF = PARTREF" & vbCrLf
sql = sql & "WHERE PACLASS LIKE '%' + @PartClass + '%'" & vbCrLf
sql = sql & "AND PAPRODCODE LIKE '%' + @PartCode + '%'" & vbCrLf
sql = sql & "AND PALEVEL IN (@PartType1, @PartType2, @PartType3, @PartType4)" & vbCrLf
sql = sql & "and LOTREMAININGQTY > 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- calculate running totals for each part" & vbCrLf
sql = sql & "update #tempMrpExRpt set PriorLotQty = isnull((select sum(t2.LOTREMAININGQTY) from #tempMrpExRpt t2" & vbCrLf
sql = sql & "where t2.PARTREF = #tempMrpExRpt.PARTREF and t2.Row < #tempMrpExRpt.Row),0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #tempMrpExRpt set QtyToShow = case when LOTREMAININGQTY > (MRP_QTYREM - PriorLotQty) then (MRP_QTYREM - PriorLotQty) else LOTREMAININGQTY end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "delete from #tempMrpExRpt where QtyToShow <= 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from #tempMrpExRpt order by PARTREF, Row" & vbCrLf
sql = sql & "drop table #tempMrpExRpt" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql



''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function


Private Function UpdateDatabase99()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 174     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "dropstoredprocedureifexists 'InvMRPExcessReport'" & vbCrLf
ExecuteScript False, sql
sql = "create procedure InvMRPExcessReport" & vbCrLf
sql = sql & "@PartClass as Varchar(16), @PartCode as varchar(8), @PartType1 as Integer," & vbCrLf
sql = sql & "@PartType2 as Integer, @PartType3 as Integer, @PartType4 as Integer" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "sp for Excess Inventory Report, InvExcess.rpt" & vbCrLf
sql = sql & "revised 2/5/2018 for MANSER - TEL" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "InvMRPExcessReport '', '', 1, 1, 1, 1" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@PartClass = 'ALL')" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @PartClass = ''" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartCode = 'ALL')" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @PartCode = ''" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "IF (@PartType1 = 1)" & vbCrLf
sql = sql & "SET @PartType1 = 1" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType1 = 0" & vbCrLf
sql = sql & "IF (@PartType2 = 1)" & vbCrLf
sql = sql & "SET @PartType2 = 2" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType2 = 0" & vbCrLf
sql = sql & "IF (@PartType3 = 1)" & vbCrLf
sql = sql & "SET @PartType3 = 3" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType3 = 0" & vbCrLf
sql = sql & "IF (@PartType4 = 1)" & vbCrLf
sql = sql & "SET @PartType4 = 4" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType4 = 0" & vbCrLf
sql = sql & "CREATE TABLE #tempMrpExRpt" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "PACLASS varchar(4) NULL ," & vbCrLf
sql = sql & "PAPRODCODE varchar(6) NULL ," & vbCrLf
sql = sql & "PALEVEL tinyint NULL ," & vbCrLf
sql = sql & "PARTREF varchar(30) NULL ," & vbCrLf
sql = sql & "PARTNUM varchar(30) NULL ," & vbCrLf
sql = sql & "PADESC varchar(30) NULL ," & vbCrLf
sql = sql & "PAEXTDESC varchar(3072) NULL ," & vbCrLf
sql = sql & "MRP_QTYREM int NULL," & vbCrLf
sql = sql & "LOT_QTYREM int NULL," & vbCrLf
sql = sql & "LOT_COST decimal(12,4) NULL ," & vbCrLf
sql = sql & "PASTDCOST decimal(12,4) NULL ," & vbCrLf
sql = sql & "USE_COST decimal(12,4) NULL," & vbCrLf
sql = sql & "TOTAL_COST decimal(12,2) NULL," & vbCrLf
sql = sql & "PAUSEACTUALCOST tinyint NULL ," & vbCrLf
sql = sql & "PALOTTRACK tinyint NULL," & vbCrLf
sql = sql & "MRP_ACTIVITY tinyint NULL" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "SELECT mrp_partref, SUM(mrp_partqtyrqd) as rem, cast(0 as bit) as active" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "from MrplTable" & vbCrLf
sql = sql & "WHERE mrp_type NOT IN ('5', '7')" & vbCrLf
sql = sql & "GROUP BY mrp_partref Having Sum(mrp_partqtyrqd) >= 1" & vbCrLf
sql = sql & "order by MRP_PARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #temp set Active = 1 where exists (select 1 from MrplTable x where x.mrp_partref = #temp.MRP_PARTREF and mrp_type NOT IN ('1','17', '5', '6'))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO #tempMrpExRpt (PACLASS, PAPRODCODE, PALEVEL,PARTREF, PARTNUM, PADESC, PAEXTDESC,MRP_QTYREM," & vbCrLf
sql = sql & "LOT_COST," & vbCrLf
sql = sql & "PASTDCOST, PAUSEACTUALCOST, PALOTTRACK, MRP_ACTIVITY," & vbCrLf
sql = sql & "LOT_QTYREM," & vbCrLf
sql = sql & "USE_COST, TOTAL_COST )" & vbCrLf
sql = sql & "SELECT PACLASS, PAPRODCODE, PALEVEL, PARTREF, PARTNUM, PADESC,PAEXTDESC, #temp.rem," & vbCrLf
sql = sql & "ISNULL((select top 1 LOTUNITCOST from LohdTable lh where LOTPARTREF = pt.PARTREF and LOTREMAININGQTY > 0 order by LOTADATE desc ),0)," & vbCrLf
sql = sql & "PASTDCOST , PAUSEACTUALCOST, PALOTTRACK, #temp.active," & vbCrLf
sql = sql & "ISNULL((select sum(LOTREMAININGQTY) from LohdTable lh2 where lh2.LOTPARTREF = pt.PARTREF and LOTREMAININGQTY > 0 ),0)," & vbCrLf
sql = sql & "0,0" & vbCrLf
sql = sql & "From PartTable pt" & vbCrLf
sql = sql & "join #temp on #temp.MRP_PARTREF = PARTREF" & vbCrLf
sql = sql & "WHERE PACLASS LIKE '%' + @PartClass + '%'" & vbCrLf
sql = sql & "AND PAPRODCODE LIKE '%' + @PartCode + '%'" & vbCrLf
sql = sql & "AND PALEVEL IN (@PartType1, @PartType2, @PartType3, @PartType4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select cost to apply" & vbCrLf
sql = sql & "--1. If PAUSEACTUALCOST = 1 and there is a lot cost, use the lot cost" & vbCrLf
sql = sql & "--2. If PAUSEACTUALCOST = 1 and the lot cost is zero, use standard cost" & vbCrLf
sql = sql & "--3. If PAUSEACTUALCOST = 0 use standard cost" & vbCrLf
sql = sql & "update #tempMrpExRpt set USE_COST = case when PAUSEACTUALCOST = 1 AND LOT_COST > 0 then LOT_COST else PASTDCOST end" & vbCrLf
sql = sql & "update #tempMrpExRpt set TOTAL_COST = MRP_QTYREM * USE_COST" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from #tempMrpExRpt" & vbCrLf
sql = sql & "order by PARTREF" & vbCrLf
sql = sql & "drop table #tempMrpExRpt" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql

' import IMAINC payroll data to GL
sql = "dropstoredprocedureifexists 'InsertPayrollJournal'" & vbCrLf
ExecuteScript False, sql

sql = "create procedure InsertPayrollJournal" & vbCrLf
sql = sql & "@CSV varchar(MAX),     -- (''ACCT1'',AMT1),(''ACCT2'',AMT2)...  (Amount is minus for a credit)" & vbCrLf
sql = sql & "@User varchar(3)," & vbCrLf
sql = sql & "@PayrollDate datetime" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* aggregate IMAINC payroll journal data from an Excel file and create a summary GL Journal" & vbCrLf
sql = sql & "returns blank if successful" & vbCrLf
sql = sql & "returns error message if unsuccessful" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON      -- required to avoid an error in sp with inserts and updates" & vbCrLf
sql = sql & "SET ANSI_WARNINGS OFF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if exists (select 1 from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_PayrollTemp')" & vbCrLf
sql = sql & "drop table _PayrollTemp" & vbCrLf
sql = sql & "if exists (select 1 from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_PayrollTemp2')" & vbCrLf
sql = sql & "drop table _PayrollTemp2" & vbCrLf
sql = sql & "if exists (select 1 from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_PayrollTemp3')" & vbCrLf
sql = sql & "drop table _PayrollTemp3" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create journal name" & vbCrLf
sql = sql & "declare @JournalName varchar(12)" & vbCrLf
sql = sql & "set @JournalName = 'PR-' + cast(year(@PayrollDate) as varchar(4)) + '-'" & vbCrLf
sql = sql & "+ RIGHT('0' + MONTH(@PayrollDate),2) + RIGHT('0' + DAY(@PayrollDate),2)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create table of raw data" & vbCrLf
sql = sql & "create table _PayrollTemp (Account varchar(12), Amount decimal(12,2))" & vbCrLf
sql = sql & "declare @sql varchar(max) = 'insert _PayRollTemp (Account,Amount) values' + char(13) + char(10) + @csv" & vbCrLf
sql = sql & "exec (@sql)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- roll up into accounts" & vbCrLf
sql = sql & "select Account," & vbCrLf
sql = sql & "sum(cast(Amount as decimal(12,2))) as Total" & vbCrLf
sql = sql & "into _PayrollTemp2" & vbCrLf
sql = sql & "from _PayrollTemp" & vbCrLf
sql = sql & "group by [Account]" & vbCrLf
sql = sql & "order by Account" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- construct data to insert" & vbCrLf
sql = sql & "select @JournalName as JINAME, 1 as JITRAN," & vbCrLf
sql = sql & "ROW_NUMBER() over (ORDER BY Account) as JIREF," & vbCrLf
sql = sql & "Account as JIACCOUNT," & vbCrLf
sql = sql & "case when Total < 0 then 0.00 else Total end as JIDEB," & vbCrLf
sql = sql & "case when Total < 0 then -Total else 0.00 end as JICRD" & vbCrLf
sql = sql & "into _PayrollTemp3" & vbCrLf
sql = sql & "from _PayrollTemp2" & vbCrLf
sql = sql & "where Total <> 0" & vbCrLf
sql = sql & "order by Account" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- attempt to create journal" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "if exists (select * from GjhdTable where GJNAME = @JournalName)" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "rollback tran" & vbCrLf
sql = sql & "select 'Journal ' + @JournalName + ' already exists.'" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO dbo.GjhdTable" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "GJNAME" & vbCrLf
sql = sql & ",GJDESC" & vbCrLf
sql = sql & ",GJOPEN" & vbCrLf
sql = sql & ",GJPOST" & vbCrLf
sql = sql & ",GJPOSTED" & vbCrLf
sql = sql & ",GJREVERSE" & vbCrLf
sql = sql & ",GJCLOSE" & vbCrLf
sql = sql & ",GJREVID" & vbCrLf
sql = sql & ",GJREVDATE" & vbCrLf
sql = sql & ",GJEXTDESC" & vbCrLf
sql = sql & ",GJTEMPLATE" & vbCrLf
sql = sql & ",GJYEAREND" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "VALUES" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@JournalName" & vbCrLf
sql = sql & ",''" & vbCrLf
sql = sql & ",CAST(getdate() as date)" & vbCrLf
sql = sql & ",null" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",''" & vbCrLf
sql = sql & ",null" & vbCrLf
sql = sql & ",'PAYROLL JOURNAL FOR PAY DATE '" & vbCrLf
sql = sql & "+ cast(year(@PayrollDate) as varchar(4)) + ' '" & vbCrLf
sql = sql & "+ right('0' + cast(month(@PayrollDate) as varchar(2)),2) + ' '" & vbCrLf
sql = sql & "+ right('0' + cast(DAY(@PayrollDate) as varchar(2)),2)" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now insert the items" & vbCrLf
sql = sql & "declare @now datetime = cast(convert(varchar(19),getdate(),100) as datetime)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO dbo.GjitTable" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "JINAME" & vbCrLf
sql = sql & ",JIDESC" & vbCrLf
sql = sql & ",JITRAN" & vbCrLf
sql = sql & ",JIREF" & vbCrLf
sql = sql & ",JIACCOUNT" & vbCrLf
sql = sql & ",JIDEB" & vbCrLf
sql = sql & ",JICRD" & vbCrLf
sql = sql & ",JIDATE" & vbCrLf
sql = sql & ",JILASTREVBY" & vbCrLf
sql = sql & ",JICLEAR" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "JINAME" & vbCrLf
sql = sql & ",''" & vbCrLf
sql = sql & ",JITRAN" & vbCrLf
sql = sql & ",JIREF" & vbCrLf
sql = sql & ",JIACCOUNT" & vbCrLf
sql = sql & ",JIDEB" & vbCrLf
sql = sql & ",JICRD" & vbCrLf
sql = sql & ",@now" & vbCrLf
sql = sql & ",@User" & vbCrLf
sql = sql & ",null" & vbCrLf
sql = sql & "from _PayrollTemp3" & vbCrLf
sql = sql & "order by JITRAN, JIREF" & vbCrLf
sql = sql & "commit tran" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- show debits and credits" & vbCrLf
sql = sql & "declare @debits decimal(12,2), @credits decimal(12,2)" & vbCrLf
sql = sql & "select @debits = sum(jideb), @credits = sum(jicrd) from GjitTable where jiname = @JournalName" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select 'Payroll Journal ' + @JournalName + ' created.  debits = ' + format(@debits,'N') + '  credits = ' + format(@credits, 'N')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql

'SQL = "AddOrUpdateColumn 'EmplTable', 'PREMEMAIL', 'varchar(60) NULL'"    'do it in 102
'ExecuteScript False, SQL


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase100()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 175     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "exec DropStoredProcedureIfExists 'SheetCancelRestock'" & vbCrLf
ExecuteScript False, sql

sql = "create procedure [dbo].SheetCancelRestock" & vbCrLf
sql = sql & "@UserLotNo varchar(40)as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "cancel restock sheet LOI records.  Make it as if it never happened." & vbCrLf
sql = sql & "exec SheetCancelRestock '030061-2'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @lot varchar(15), @when datetime, @part varchar(30)" & vbCrLf
sql = sql & "select @lot = LOTNUMBER, @part = LOTPARTREF from LohdTable where LOTUSERLOTID = @UserLotNo" & vbCrLf
sql = sql & "select @when = (select top 1 LOIADATE from LoitTable where LOINUMBER = @lot and LOICLOSED is null and LOISHEETACTTYPE = 'RS')" & vbCrLf
sql = sql & "--select @lot, @part, @when" & vbCrLf
sql = sql & "if @lot is null or @part is null or @when is null" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "rollback tran" & vbCrLf
sql = sql & "return" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get total change in quantity" & vbCrLf
sql = sql & "declare @qty decimal(12,4)" & vbCrLf
sql = sql & "select @qty = (select sum(LOIQUANTITY) from LoitTable where LOINUMBER = @lot and LOICLOSED is null and LOISHEETACTTYPE = 'RS')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- re-open previous pick records" & vbCrLf
sql = sql & "update LoitTable set LOICLOSED = NULL where LOINUMBER = @lot and LOISHEETACTTYPE = 'PK'" & vbCrLf
sql = sql & "and LOIRECORD in (select LOIPARENTREC from LoitTable where LOINUMBER = @lot and LOICLOSED is null and LOISHEETACTTYPE = 'RS')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete the inventory activity records" & vbCrLf
sql = sql & "delete from InvaTable" & vbCrLf
sql = sql & "where INLOTNUMBER = @lot and INADATE = @when" & vbCrLf
sql = sql & "and INAQTY in (select LOIQUANTITY FROM LoitTable WHERE LOINUMBER = @lot and LOICLOSED is null and LOISHEETACTTYPE = 'RS')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete the restock records" & vbCrLf
sql = sql & "delete from LoitTable where LOINUMBER = @lot and LOICLOSED is null and LOISHEETACTTYPE = 'RS'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- remove canceled quantity from LohdTable and PartTable" & vbCrLf
sql = sql & "update LohdTable set LOTREMAININGQTY = LOTREMAININGQTY - @qty WHERE LOTNUMBER = @LOT" & vbCrLf
sql = sql & "update PartTable set PAQOH = PAQOH - @qty where partref = @part" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "commit tran" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function


Private Function UpdateDatabase101()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 176     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "DropStoredProcedureIfExists 'UpdateTimeCardTotals'" & vbCrLf
ExecuteScript False, sql

sSql = "create procedure [dbo].[UpdateTimeCardTotals]" & vbCrLf
sSql = sSql & " @EmpNo int," & vbCrLf
sSql = sSql & " @Date datetime" & vbCrLf
sSql = sSql & "as " & vbCrLf
sSql = sSql & "/* test" & vbCrLf
sSql = sSql & "    UpdateTimeCardTotals 52, '9/8/2008'" & vbCrLf
sSql = sSql & "*/" & vbCrLf
sSql = sSql & "update TchdTable " & vbCrLf
sSql = sSql & "set TMSTART = (select top 1 TCSTART From TcitTable " & vbCrLf
sSql = sSql & "join TchdTable on TCCARD = TMCARD" & vbCrLf
sSql = sSql & "where TCEMP = @EmpNo and TMDAY = @Date and TCSTOP <> ''" & vbCrLf
sSql = sSql & "and ISDATE(TCSTART + 'm') = 1 and ISDATE(TCSTOP + 'm') = 1" & vbCrLf
sSql = sSql & "order by tcstarttime)," & vbCrLf
sSql = sSql & "TMSTOP = (select top 1 TCSTOP From TcitTable" & vbCrLf
sSql = sSql & "join TchdTable on TCCARD = TMCARD" & vbCrLf
sSql = sSql & "where TCEMP = @EmpNo and TMDAY = @Date and TCSTOP <> ''" & vbCrLf
sSql = sSql & "and ISDATE(TCSTART + 'm') = 1 and ISDATE(TCSTOP + 'm') = 1" & vbCrLf
sSql = sSql & "order by TCSTOPTIME desc)" & vbCrLf
sSql = sSql & "where TMEMP = @EmpNo and TMDAY = @Date" & vbCrLf
sSql = sSql & "" & vbCrLf
sSql = sSql & "update TchdTable " & vbCrLf
sSql = sSql & "set TMREGHRS = (select isnull(sum(hrs), 0.000) from viewTimeCardHours" & vbCrLf
sSql = sSql & "where EmpNo = @EmpNo" & vbCrLf
sSql = sSql & "and chgDay = @Date" & vbCrLf
sSql = sSql & "and type = 'R')," & vbCrLf
sSql = sSql & "TMOVTHRS = (select isnull(sum(hrs), 0.000) from viewTimeCardHours" & vbCrLf
sSql = sSql & "where EmpNo = @EmpNo" & vbCrLf
sSql = sSql & "and chgDay = @Date" & vbCrLf
sSql = sSql & "and type = 'O')," & vbCrLf
sSql = sSql & "TMDBLHRS = (select isnull(sum(hrs), 0.000) from viewTimeCardHours" & vbCrLf
sSql = sSql & "where EmpNo = @EmpNo" & vbCrLf
sSql = sSql & "and chgDay = @Date" & vbCrLf
sSql = sSql & "and type = 'D')" & vbCrLf
sSql = sSql & "where TMEMP = @EmpNo and TMDAY = @Date"
ExecuteScript False, sSql

' this should have happened in UpdateDatabase27, but that did not happen for all users.
sql = "DropStoredProcedureIfExists 'RptChartOfAccount'" & vbCrLf
ExecuteScript False, sql

sql = "CREATE  PROCEDURE [dbo].[RptChartOfAccount]  " & vbCrLf
sql = sql & "    @InclIncAcct as varchar(1)" & vbCrLf
sql = sql & "AS " & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "  exec RptChartOfAccount 0" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN " & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   declare @glAcctRef as varchar(10) " & vbCrLf
sql = sql & "   declare @glMsAcct as varchar(10) " & vbCrLf
sql = sql & "   declare @TopLevelDesc as varchar(30)" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   declare @level as varchar(12)" & vbCrLf
sql = sql & "   declare @InclInAcct as Integer" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   if (@InclIncAcct = '1')" & vbCrLf
sql = sql & "      SET @InclInAcct = ''" & vbCrLf
sql = sql & "   else" & vbCrLf
sql = sql & "      SET @InclInAcct = '0'" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   CREATE TABLE #tempChartOfAcct(   " & vbCrLf
sql = sql & "   [TOPLEVEL] [varchar](12) NULL,  " & vbCrLf
sql = sql & "   [TOPLEVELDESC] [varchar](30) NULL, " & vbCrLf
sql = sql & "   [GLACCTREF] [varchar](112) NULL,         " & vbCrLf
sql = sql & "   [GLDESCR] [varchar](120) NULL,  " & vbCrLf
sql = sql & "   [GLMASTER] [varchar](12) NULL,   " & vbCrLf
sql = sql & "   [GLFSLEVEL] [INT] NULL," & vbCrLf
sql = sql & "   [GLINACTIVE] [int] NULL," & vbCrLf
sql = sql & "   [SORTKEYLEVEL] [int] NULL,           " & vbCrLf
sql = sql & "   [GLACCSORTKEY] [varchar](512) NULL           " & vbCrLf
sql = sql & ")                             " & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   DECLARE balAcctStruc CURSOR  FOR " & vbCrLf
sql = sql & "      SELECT COASSTACCT, COASSTDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COLIABACCT, COLIABDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COINCMACCT, COINCMDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL " & vbCrLf
sql = sql & "      SELECT COEQTYACCT, COEQTYDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COCOGSACCT, COCOGSDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COEXPNACCT, COEXPNDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COOINCACCT, COOINCDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "      SELECT COFDTXACCT, COFDTXDESC FROM GlmsTable" & vbCrLf
sql = sql & "      UNION ALL" & vbCrLf
sql = sql & "     SELECT COOEXPACCT, COOEXPDESC FROM GlmsTable" & vbCrLf
sql = sql & "      --UNION ALL" & vbCrLf
sql = sql & "      --SELECT COFDTXACCT, COFDTXDESC FROM GlmsTable" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   OPEN balAcctStruc" & vbCrLf
sql = sql & "   FETCH NEXT FROM balAcctStruc INTO @level, @TopLevelDesc" & vbCrLf
sql = sql & "   WHILE (@@FETCH_STATUS <> -1) " & vbCrLf
sql = sql & "   BEGIN " & vbCrLf
sql = sql & "      IF (@@FETCH_STATUS <> -2) " & vbCrLf
sql = sql & "      BEGIN " & vbCrLf
sql = sql & "         " & vbCrLf
sql = sql & "         INSERT INTO #tempChartOfAcct(TOPLEVEL, TOPLEVELDESC, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL,GLINACTIVE, SORTKEYLEVEL,GLACCSORTKEY)" & vbCrLf
sql = sql & "         select @level as TopLevel, @TopLevelDesc as TopLevelDesc, @level as GLACCTREF, " & vbCrLf
sql = sql & "            @TopLevelDesc as GLDESCR, '' as GLMASTER, 0 as GLFSLEVEL, 0,0 as level, @level as SortKey;" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "         with cte" & vbCrLf
sql = sql & "         as" & vbCrLf
sql = sql & "         (select GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL, GLINACTIVE, 0 as level," & vbCrLf
sql = sql & "            cast(cast(@level as varchar(12))+ char(36)+ GLACCTREF as varchar(max)) as SortKey" & vbCrLf
sql = sql & "         from GlacTable" & vbCrLf
sql = sql & "         where GLMASTER = cast(@level as varchar(12)) AND GLINACTIVE LIKE @InclInAcct" & vbCrLf
sql = sql & "         union all" & vbCrLf
sql = sql & "         select a.GLACCTREF, a.GLDESCR, a.GLMASTER, a.GLFSLEVEL, a.GLINACTIVE, level + 1," & vbCrLf
sql = sql & "          cast(COALESCE(SortKey,'') + char(36) + COALESCE(a.GLACCTREF,'') as varchar(max))as SortKey" & vbCrLf
sql = sql & "         from cte" & vbCrLf
sql = sql & "            inner join GlacTable a" & vbCrLf
sql = sql & "               on cte.GLACCTREF = a.GLMASTER" & vbCrLf
sql = sql & "            WHERE a.GLINACTIVE LIKE @InclInAcct" & vbCrLf
sql = sql & "         )" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "         INSERT INTO #tempChartOfAcct(TOPLEVEL, TOPLEVELDESC, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL,GLINACTIVE, SORTKEYLEVEL,GLACCSORTKEY)" & vbCrLf
sql = sql & "         select @level as TopLevel, @TopLevelDesc as TopLevelDesc, " & vbCrLf
sql = sql & "               Replicate('  ', level) + GLACCTREF as GLACCTREF, " & vbCrLf
sql = sql & "               Replicate('  ', level) + GLDESCR as GLDESCR, GLMASTER, " & vbCrLf
sql = sql & "               GLFSLEVEL, GLINACTIVE,level, SortKey" & vbCrLf
sql = sql & "         from cte order by SortKey" & vbCrLf
sql = sql & "         " & vbCrLf
sql = sql & "      END" & vbCrLf
sql = sql & "      FETCH NEXT FROM balAcctStruc INTO @level, @TopLevelDesc" & vbCrLf
sql = sql & "   END         " & vbCrLf
sql = sql & "   CLOSE balAcctStruc" & vbCrLf
sql = sql & "   DEALLOCATE balAcctStruc" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   SELECT TOPLEVEL, TOPLEVELDESC, GLACCTREF, GLDESCR, GLMASTER, " & vbCrLf
sql = sql & "      GLFSLEVEL,GLINACTIVE, SORTKEYLEVEL,GLACCSORTKEY " & vbCrLf
sql = sql & "   FROM #tempChartOfAcct ORDER BY GLACCSORTKEY" & vbCrLf
sql = sql & "                                           " & vbCrLf
sql = sql & "   DROP table #tempChartOfAcct            " & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase102()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 177
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "AddOrUpdateColumn 'EmplTable', 'PREMEMAIL', 'varchar(60) NULL'"
ExecuteScript False, sql

sql = "AddOrUpdateColumn 'EmplTable', 'PREMPREVTERMDT', 'smalldatetime NULL'"
ExecuteScript False, sql

' turn SOTEXT into a computed column so it can never be wrong
sql = "DropFunctionIfExists 'GetDefaultConstraintName'" & vbCrLf
ExecuteScript False, sql

sql = "create function dbo.GetDefaultConstraintName" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@schema varchar(100)," & vbCrLf
sql = sql & "@table varchar(100)," & vbCrLf
sql = sql & "@column varchar(100)" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "returns varchar(100)" & vbCrLf
sql = sql & "/* get default name for a column.  returns blank if none" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "select dbo.GetDefaultConstraintName('dbo','SohdTable','SOTEXT')" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "declare @name varchar(100)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select @name =" & vbCrLf
sql = sql & "default_constraints.name" & vbCrLf
sql = sql & "from sys.all_columns" & vbCrLf
sql = sql & "INNER JOIN sys.tables ON all_columns.object_id = tables.object_id" & vbCrLf
sql = sql & "INNER JOIN sys.schemas ON tables.schema_id = schemas.schema_id" & vbCrLf
sql = sql & "INNER JOIN sys.default_constraints ON all_columns.default_object_id = default_constraints.object_id" & vbCrLf
sql = sql & "where schemas.name = @schema" & vbCrLf
sql = sql & "AND tables.name = @table" & vbCrLf
sql = sql & "AND all_columns.name = @column" & vbCrLf
sql = sql & "return isnull(@name,'')" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql

sql = "declare @constraintname varchar(100)" & vbCrLf
sql = sql & "declare @sql varchar(200)" & vbCrLf
sql = sql & "set @constraintname = dbo.GetDefaultConstraintName('dbo','SohdTable','SOTEXT')" & vbCrLf
sql = sql & "if @constraintname <> ''" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "set @sql = 'alter table SohdTable drop constraint ' + @constraintname" & vbCrLf
sql = sql & "exec (@sql)" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql

sql = "alter table SohdTable drop column SOTEXT" & vbCrLf
ExecuteScript False, sql

sql = "alter table SohdTable add SOTEXT as right('000000' + cast(SONUMBER as varchar(6)),6)" & vbCrLf
ExecuteScript False, sql

'from updatedatabase34 -- not everyone has it
      If StoreProcedureExists("RptAcctBalanceSheet") Then
         sSql = "DROP PROCEDURE RptAcctBalanceSheet"
         ExecuteScript False, sSql
      End If
      
      sSql = "CREATE PROCEDURE [dbo].[RptAcctBalanceSheet]" & vbCrLf
      sSql = sSql & "   @StartDate as varchar(12),@EndDate as varchar(12)," & vbCrLf
      sSql = sSql & "   @InclInAcct as varchar(1)" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "BEGIN " & vbCrLf
      sSql = sSql & "   declare @glAcctRef as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @glMsAcct as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @SumCurBal decimal(15,4)" & vbCrLf
      sSql = sSql & "   declare @SumPrevBal as decimal(15,4)" & vbCrLf
      sSql = sSql & "   declare @level as integer" & vbCrLf
      sSql = sSql & "   declare @TopLevelDesc as varchar(30)" & vbCrLf
      sSql = sSql & "   declare @TopLevAcct as varchar(20)" & vbCrLf
      sSql = sSql & "   declare @PrevMaster as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @RowCount as integer" & vbCrLf
      sSql = sSql & "   declare @GlMasterAcc as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @GlChildAcct as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @ChildKey as varchar(1024)" & vbCrLf
      sSql = sSql & "   declare @GLSortKey as varchar(1024)" & vbCrLf
      sSql = sSql & "   declare @SortKey as varchar(1024)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   DELETE FROM EsReportBalanceSheet" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   if (@InclInAcct = '1')" & vbCrLf
      sSql = sSql & "      SET @InclInAcct = '%'" & vbCrLf
      sSql = sSql & "   Else" & vbCrLf
      sSql = sSql & "      SET @InclInAcct = '0'" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   DECLARE balAcctStruc CURSOR  FOR" & vbCrLf
      sSql = sSql & "      SELECT '1', COASSTACCT, COASSTDESC FROM GlmsTable" & vbCrLf
      sSql = sSql & "      Union All" & vbCrLf
      sSql = sSql & "      SELECT '2', COLIABACCT, COLIABDESC FROM GlmsTable" & vbCrLf
      sSql = sSql & "      Union All" & vbCrLf
      sSql = sSql & "      SELECT '3', COEQTYACCT, COEQTYDESC FROM GlmsTable" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   OPEN balAcctStruc" & vbCrLf
      sSql = sSql & "   FETCH NEXT FROM balAcctStruc INTO @level, @TopLevAcct, @TopLevelDesc" & vbCrLf
      sSql = sSql & "   WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "   BEGIN " & vbCrLf
      sSql = sSql & "      IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "      BEGIN " & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "         ;with cte" & vbCrLf
      sSql = sSql & "         as" & vbCrLf
      sSql = sSql & "         (select GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL, 0 as level," & vbCrLf
      sSql = sSql & "            cast(cast(@level as varchar(4))+ char(36)+ GLACCTREF as varchar(max)) as SortKey" & vbCrLf
      sSql = sSql & "         From GlacTable" & vbCrLf
      sSql = sSql & "         where GLMASTER = cast(@TopLevAcct as varchar(20)) AND GLINACTIVE LIKE @InclInAcct" & vbCrLf
      sSql = sSql & "         Union All" & vbCrLf
      sSql = sSql & "         select a.GLACCTREF, a.GLDESCR, a.GLMASTER, a.GLFSLEVEL, level + 1," & vbCrLf
      sSql = sSql & "          cast(COALESCE(SortKey,'') + char(36) + COALESCE(a.GLACCTREF,'') as varchar(max))as SortKey" & vbCrLf
      sSql = sSql & "         From cte" & vbCrLf
      sSql = sSql & "            inner join GlacTable a" & vbCrLf
      sSql = sSql & "               on cte.GLACCTREF = a.GLMASTER" & vbCrLf
      sSql = sSql & "            WHERE GLINACTIVE LIKE @InclInAcct" & vbCrLf
      sSql = sSql & "         )" & vbCrLf
      sSql = sSql & "         INSERT INTO EsReportBalanceSheet (GLTOPMaster, GLTOPMASTERDESC, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL,SORTKEYLEVEL,GLACCSORTKEY)" & vbCrLf
      sSql = sSql & "         select @level, @TopLevelDesc," & vbCrLf
      sSql = sSql & "         GLACCTREF , GLDESCR, GLMASTER, GLFSLEVEL, Level, SortKey" & vbCrLf
      sSql = sSql & "         from cte order by SortKey" & vbCrLf
      sSql = sSql & "         " & vbCrLf
      sSql = sSql & "      End" & vbCrLf
      sSql = sSql & "      FETCH NEXT FROM balAcctStruc INTO @level, @TopLevAcct, @TopLevelDesc" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & "   Close balAcctStruc" & vbCrLf
      sSql = sSql & "   DEALLOCATE balAcctStruc" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   UPDATE EsReportBalanceSheet SET CurrentBal = foo.Balance" & vbCrLf
      sSql = sSql & "   From" & vbCrLf
      sSql = sSql & "       (SELECT SUM(GjitTable.JIDEB) - SUM(GjitTable.JICRD) as Balance, JIACCOUNT" & vbCrLf
      sSql = sSql & "         FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME" & vbCrLf
      sSql = sSql & "      WHERE GJPOST BETWEEN @StartDate AND @EndDate" & vbCrLf
      sSql = sSql & "         AND GjhdTable.GJPOSTED = 1" & vbCrLf
      sSql = sSql & "      GROUP BY JIACCOUNT) as foo" & vbCrLf
      sSql = sSql & "   Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
      sSql = sSql & "  " & vbCrLf
      sSql = sSql & " UPDATE EsReportBalanceSheet SET PreviousBal = foo.Balance" & vbCrLf
      sSql = sSql & "   From" & vbCrLf
      sSql = sSql & "      (SELECT JIACCOUNT,SUM(GjitTable.JIDEB) - SUM(GjitTable.JICRD) AS Balance" & vbCrLf
      sSql = sSql & "         FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME" & vbCrLf
      sSql = sSql & "      WHERE (GJPOST  <  @StartDate)" & vbCrLf
      sSql = sSql & "         AND GjhdTable.GJPOSTED = 1" & vbCrLf
      sSql = sSql & "      GROUP BY JIACCOUNT) as foo" & vbCrLf
      sSql = sSql & "   Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   set @level = 9" & vbCrLf
      sSql = sSql & "   WHILE (@level >= 1 )" & vbCrLf
      sSql = sSql & "   BEGIN" & vbCrLf
      sSql = sSql & "      DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
      sSql = sSql & "      SELECT GLMASTER, SUM(ISNULL(SUMCURBAL,0) + (ISNULL(CurrentBal,0))) ," & vbCrLf
      sSql = sSql & "         Sum (IsNull(SUMPREVBAL, 0) + (IsNull(PreviousBal, 0)))" & vbCrLf
      sSql = sSql & "      From" & vbCrLf
      sSql = sSql & "         (SELECT DISTINCT GLACCTREF, GLDESCR, GLMASTER," & vbCrLf
      sSql = sSql & "         CurrentBal , PreviousBal, SUMCURBAL, SUMPREVBAL" & vbCrLf
      sSql = sSql & "         FROM EsReportBalanceSheet WHERE GLFSLEVEL = @level) as foo" & vbCrLf
      sSql = sSql & "      group by GLMASTER" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "      OPEN curAcctStruc" & vbCrLf
      sSql = sSql & "      FETCH NEXT FROM curAcctStruc INTO @glMsAcct, @SumCurBal, @SumPrevBal" & vbCrLf
      sSql = sSql & "      WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "      BEGIN" & vbCrLf
      sSql = sSql & "         IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "         BEGIN" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "           UPDATE EsReportBalanceSheet SET SUMCURBAL = (ISNULL(SUMCURBAL, 0) + @SumCurBal)," & vbCrLf
      sSql = sSql & "               SUMPREVBAL = (ISNULL(SUMPREVBAL, 0) + @SumPrevBal), GLDESCR = 'TOTAL '+ LTRIM(GLDESCR)," & vbCrLf
      sSql = sSql & "            HASCHILD = 1" & vbCrLf
      sSql = sSql & "            WHERE GLACCTREF = @glMsAcct" & vbCrLf
      sSql = sSql & "         End" & vbCrLf
      sSql = sSql & "         FETCH NEXT FROM curAcctStruc INTO @glMsAcct, @SumCurBal, @SumPrevBal" & vbCrLf
      sSql = sSql & "      End" & vbCrLf
      sSql = sSql & "            " & vbCrLf
      sSql = sSql & "      Close curAcctStruc" & vbCrLf
      sSql = sSql & "      DEALLOCATE curAcctStruc" & vbCrLf
      sSql = sSql & "      " & vbCrLf
      sSql = sSql & "      SET @level = @level - 1" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   " & vbCrLf
      sSql = sSql & "   UPDATE EsReportBalanceSheet SET SUMCURBAL = CurrentBal WHERE SUMCURBAL IS NULL" & vbCrLf
      sSql = sSql & "   UPDATE EsReportBalanceSheet SET SUMPREVBAL = PreviousBal WHERE SUMPREVBAL IS NULL" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   set @level = 0" & vbCrLf
      sSql = sSql & "   set @RowCount = 1" & vbCrLf
      sSql = sSql & "   SET @PrevMaster = ''" & vbCrLf
      sSql = sSql & "   DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
      sSql = sSql & "      SELECT GLMASTER, GLACCSORTKEY" & vbCrLf
      sSql = sSql & "      From EsReportBalanceSheet" & vbCrLf
      sSql = sSql & "         WHERE HASCHILD IS NULL" & vbCrLf
      sSql = sSql & "      ORDER BY GLACCSORTKEY" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   OPEN curAcctStruc" & vbCrLf
      sSql = sSql & "   FETCH NEXT FROM curAcctStruc INTO @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "   WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "   BEGIN" & vbCrLf
      sSql = sSql & "    IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "    BEGIN" & vbCrLf
      sSql = sSql & "      if (@PrevMaster <> @GlMasterAcc)" & vbCrLf
      sSql = sSql & "      BEGIN" & vbCrLf
      sSql = sSql & "         UPDATE EsReportBalanceSheet SET" & vbCrLf
      sSql = sSql & "            SortKeyRev = RIGHT ('0000'+ Cast(@RowCount as varchar), 4) + char(36)+ @GlMasterAcc" & vbCrLf
      sSql = sSql & "         WHERE GLMASTER = @GlMasterAcc AND HASCHILD IS NULL" & vbCrLf
      sSql = sSql & "         SET @RowCount = @RowCount + 1" & vbCrLf
      sSql = sSql & "         SET @PrevMaster = @GlMasterAcc" & vbCrLf
      sSql = sSql & "      End" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "    End" & vbCrLf
      sSql = sSql & "    FETCH NEXT FROM curAcctStruc INTO @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & "       " & vbCrLf
      sSql = sSql & "   Close curAcctStruc" & vbCrLf
      sSql = sSql & "   DEALLOCATE curAcctStruc" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   set @level = 7" & vbCrLf
      sSql = sSql & "   WHILE (@level >= 1 )" & vbCrLf
      sSql = sSql & "   BEGIN" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      SET @PrevMaster = ''" & vbCrLf
      sSql = sSql & "        DECLARE curAcctStruc1 CURSOR  FOR" & vbCrLf
      sSql = sSql & "         SELECT DISTINCT GLACCTREF, GLMASTER, GLACCSORTKEY" & vbCrLf
      sSql = sSql & "         From EsReportBalanceSheet" & vbCrLf
      sSql = sSql & "            WHERE GLFSLEVEL = @level AND HASCHILD IS NOT NULL" & vbCrLf
      sSql = sSql & "         order by GLACCSORTKEY" & vbCrLf
      sSql = sSql & "    " & vbCrLf
      sSql = sSql & "        OPEN curAcctStruc1" & vbCrLf
      sSql = sSql & "        FETCH NEXT FROM curAcctStruc1 INTO @GlChildAcct, @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "        WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "        BEGIN" & vbCrLf
      sSql = sSql & "          IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "          BEGIN" & vbCrLf
      sSql = sSql & "            if (@PrevMaster <> @GlChildAcct)" & vbCrLf
      sSql = sSql & "            BEGIN" & vbCrLf
      sSql = sSql & "               SELECT TOP 1 @ChildKey = SortKeyRev,@SortKey = GLACCSORTKEY" & vbCrLf
      sSql = sSql & "               From EsReportBalanceSheet" & vbCrLf
      sSql = sSql & "                  WHERE GLFSLEVEL > @level AND GLMASTER = @GlChildAcct" & vbCrLf
      sSql = sSql & "               order by GLACCSORTKEY desc" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "               UPDATE EsReportBalanceSheet SET" & vbCrLf
      sSql = sSql & "                  SortKeyRev = Cast(@ChildKey as varchar(512)) + char(36)+ @GlMasterAcc" & vbCrLf
      sSql = sSql & "               WHERE GLACCTREF = @GlChildAcct AND GLMASTER = @GlMasterAcc" & vbCrLf
      sSql = sSql & "                  AND GLFSLEVEL = @level" & vbCrLf
      sSql = sSql & "               SET @PrevMaster = @GlChildAcct" & vbCrLf
      sSql = sSql & "            End" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "          End" & vbCrLf
      sSql = sSql & "          FETCH NEXT FROM curAcctStruc1 INTO @GlChildAcct, @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "        End" & vbCrLf
      sSql = sSql & "               " & vbCrLf
      sSql = sSql & "        Close curAcctStruc1" & vbCrLf
      sSql = sSql & "        DEALLOCATE curAcctStruc1" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "        SET @level = @level - 1" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   SELECT GLTOPMaster, GLTOPMASTERDESC, GLACCTREF, GLACCTNO, GLDESCR, GLMASTER, GLTYPE,GLINACTIVE, GLFSLEVEL," & vbCrLf
      sSql = sSql & "      SUMCURBAL , CurrentBal, SUMPREVBAL, PreviousBal, SORTKEYLEVEL, GLACCSORTKEY, SortKeyRev,HASCHILD" & vbCrLf
      sSql = sSql & "   FROM EsReportBalanceSheet ORDER BY SortKeyRev" & vbCrLf
      sSql = sSql & "End"
      
      ExecuteScript False, sSql
      
      'not everyone has this
      If StoreProcedureExists("RptAcctTopBalanceSheet") Then
         sSql = "DROP PROCEDURE RptAcctTopBalanceSheet"
         ExecuteScript False, sSql
      End If

      sSql = "CREATE PROCEDURE [dbo].[RptAcctTopBalanceSheet]" & vbCrLf
      sSql = sSql & "   @StartDate as varchar(12),@EndDate as varchar(12)," & vbCrLf
      sSql = sSql & "   @InclInAcct as varchar(1)" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "BEGIN " & vbCrLf
      sSql = sSql & "   declare @glAcctRef as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @glMsAcct as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @SumCurBal decimal(15,4)" & vbCrLf
      sSql = sSql & "   declare @SumPrevBal as decimal(15,4)" & vbCrLf
      sSql = sSql & "   declare @level as integer" & vbCrLf
      sSql = sSql & "   declare @TopLevelDesc as varchar(30)" & vbCrLf
      sSql = sSql & "   declare @TopLevAcct as varchar(20)" & vbCrLf
      sSql = sSql & "   declare @PrevMaster as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @RowCount as integer" & vbCrLf
      sSql = sSql & "   declare @GlMasterAcc as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @GlChildAcct as varchar(10)" & vbCrLf
      sSql = sSql & "   declare @ChildKey as varchar(1024)" & vbCrLf
      sSql = sSql & "   declare @GLSortKey as varchar(1024)" & vbCrLf
      sSql = sSql & "   declare @SortKey as varchar(1024)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   DELETE FROM EsReportBalanceSheet" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   if (@InclInAcct = '1')" & vbCrLf
      sSql = sSql & "      SET @InclInAcct = '%'" & vbCrLf
      sSql = sSql & "   Else" & vbCrLf
      sSql = sSql & "      SET @InclInAcct = '0'" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   DECLARE balAcctStruc CURSOR  FOR" & vbCrLf
      sSql = sSql & "      SELECT '1', COASSTACCT, COASSTDESC FROM GlmsTopTable" & vbCrLf
      sSql = sSql & "      Union All" & vbCrLf
      sSql = sSql & "      SELECT '2', COLIABACCT, COLIABDESC FROM GlmsTopTable" & vbCrLf
      sSql = sSql & "      Union All" & vbCrLf
      sSql = sSql & "      SELECT '3', COEQTYACCT, COEQTYDESC FROM GlmsTopTable" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   OPEN balAcctStruc" & vbCrLf
      sSql = sSql & "   FETCH NEXT FROM balAcctStruc INTO @level, @TopLevAcct, @TopLevelDesc" & vbCrLf
      sSql = sSql & "   WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "   BEGIN " & vbCrLf
      sSql = sSql & "      IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "      BEGIN " & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "         ;with cte" & vbCrLf
      sSql = sSql & "         as" & vbCrLf
      sSql = sSql & "         (select GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL, 0 as level," & vbCrLf
      sSql = sSql & "            cast(cast(@level as varchar(4))+ char(36)+ GLACCTREF as varchar(max)) as SortKey" & vbCrLf
      sSql = sSql & "         From GlacTopTable" & vbCrLf
      sSql = sSql & "         where GLMASTER = cast(@TopLevAcct as varchar(20)) AND GLINACTIVE LIKE @InclInAcct" & vbCrLf
      sSql = sSql & "         Union All" & vbCrLf
      sSql = sSql & "         select a.GLACCTREF, a.GLDESCR, a.GLMASTER, a.GLFSLEVEL, level + 1," & vbCrLf
      sSql = sSql & "          cast(COALESCE(SortKey,'') + char(36) + COALESCE(a.GLACCTREF,'') as varchar(max))as SortKey" & vbCrLf
      sSql = sSql & "         From cte" & vbCrLf
      sSql = sSql & "            inner join GlacTopTable a" & vbCrLf
      sSql = sSql & "               on cte.GLACCTREF = a.GLMASTER" & vbCrLf
      sSql = sSql & "            WHERE GLINACTIVE LIKE @InclInAcct" & vbCrLf
      sSql = sSql & "         )" & vbCrLf
      sSql = sSql & "         INSERT INTO EsReportBalanceSheet (GLTOPMaster, GLTOPMASTERDESC, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL,SORTKEYLEVEL,GLACCSORTKEY)" & vbCrLf
      sSql = sSql & "         select @level, @TopLevelDesc," & vbCrLf
      sSql = sSql & "         GLACCTREF , GLDESCR, GLMASTER, GLFSLEVEL, Level, SortKey" & vbCrLf
      sSql = sSql & "         from cte order by SortKey" & vbCrLf
      sSql = sSql & "         " & vbCrLf
      sSql = sSql & "      End" & vbCrLf
      sSql = sSql & "      FETCH NEXT FROM balAcctStruc INTO @level, @TopLevAcct, @TopLevelDesc" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & "   Close balAcctStruc" & vbCrLf
      sSql = sSql & "   DEALLOCATE balAcctStruc" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   UPDATE EsReportBalanceSheet SET CurrentBal = foo.Balance" & vbCrLf
      sSql = sSql & "   From" & vbCrLf
      sSql = sSql & "       (SELECT SUM(GjitTopTable.JIDEB) - SUM(GjitTopTable.JICRD) as Balance, JIACCOUNT" & vbCrLf
      sSql = sSql & "         FROM GjhdTable INNER JOIN GjitTopTable ON GJNAME = JINAME" & vbCrLf
      sSql = sSql & "      WHERE GJPOST BETWEEN @StartDate AND @EndDate" & vbCrLf
      sSql = sSql & "         AND GjhdTable.GJPOSTED = 1" & vbCrLf
      sSql = sSql & "      GROUP BY JIACCOUNT) as foo" & vbCrLf
      sSql = sSql & "   Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
      sSql = sSql & "  " & vbCrLf
      sSql = sSql & " UPDATE EsReportBalanceSheet SET PreviousBal = foo.Balance" & vbCrLf
      sSql = sSql & "   From" & vbCrLf
      sSql = sSql & "      (SELECT JIACCOUNT,SUM(GjitTopTable.JIDEB) - SUM(GjitTopTable.JICRD) AS Balance" & vbCrLf
      sSql = sSql & "         FROM GjhdTable INNER JOIN GjitTopTable ON GJNAME = JINAME" & vbCrLf
      sSql = sSql & "      WHERE (GJPOST  <  @StartDate)" & vbCrLf
      sSql = sSql & "         AND GjhdTable.GJPOSTED = 1" & vbCrLf
      sSql = sSql & "      GROUP BY JIACCOUNT) as foo" & vbCrLf
      sSql = sSql & "   Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   set @level = 9" & vbCrLf
      sSql = sSql & "   WHILE (@level >= 1 )" & vbCrLf
      sSql = sSql & "   BEGIN" & vbCrLf
      sSql = sSql & "      DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
      sSql = sSql & "      SELECT GLMASTER, SUM(ISNULL(SUMCURBAL,0) + (ISNULL(CurrentBal,0))) ," & vbCrLf
      sSql = sSql & "         Sum (IsNull(SUMPREVBAL, 0) + (IsNull(PreviousBal, 0)))" & vbCrLf
      sSql = sSql & "      From" & vbCrLf
      sSql = sSql & "         (SELECT DISTINCT GLACCTREF, GLDESCR, GLMASTER," & vbCrLf
      sSql = sSql & "         CurrentBal , PreviousBal, SUMCURBAL, SUMPREVBAL" & vbCrLf
      sSql = sSql & "         FROM EsReportBalanceSheet WHERE SORTKEYLEVEL = @level) as foo" & vbCrLf
      sSql = sSql & "      group by GLMASTER" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "      OPEN curAcctStruc" & vbCrLf
      sSql = sSql & "      FETCH NEXT FROM curAcctStruc INTO @glMsAcct, @SumCurBal, @SumPrevBal" & vbCrLf
      sSql = sSql & "      WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "      BEGIN" & vbCrLf
      sSql = sSql & "         IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "         BEGIN" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "           UPDATE EsReportBalanceSheet SET SUMCURBAL = (ISNULL(SUMCURBAL, 0) + @SumCurBal)," & vbCrLf
      sSql = sSql & "               SUMPREVBAL = (ISNULL(SUMPREVBAL, 0) + @SumPrevBal), GLDESCR = 'TOTAL '+ LTRIM(GLDESCR)," & vbCrLf
      sSql = sSql & "            HASCHILD = 1" & vbCrLf
      sSql = sSql & "            WHERE GLACCTREF = @glMsAcct" & vbCrLf
      sSql = sSql & "         End" & vbCrLf
      sSql = sSql & "         FETCH NEXT FROM curAcctStruc INTO @glMsAcct, @SumCurBal, @SumPrevBal" & vbCrLf
      sSql = sSql & "      End" & vbCrLf
      sSql = sSql & "            " & vbCrLf
      sSql = sSql & "      Close curAcctStruc" & vbCrLf
      sSql = sSql & "      DEALLOCATE curAcctStruc" & vbCrLf
      sSql = sSql & "      " & vbCrLf
      sSql = sSql & "      SET @level = @level - 1" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   " & vbCrLf
      sSql = sSql & "   UPDATE EsReportBalanceSheet SET SUMCURBAL = CurrentBal WHERE SUMCURBAL IS NULL" & vbCrLf
      sSql = sSql & "   UPDATE EsReportBalanceSheet SET SUMPREVBAL = PreviousBal WHERE SUMPREVBAL IS NULL" & vbCrLf
      sSql = sSql & " " & vbCrLf
      sSql = sSql & "   set @level = 0" & vbCrLf
      sSql = sSql & "   set @RowCount = 1" & vbCrLf
      sSql = sSql & "   SET @PrevMaster = ''" & vbCrLf
      sSql = sSql & "   DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
      sSql = sSql & "      SELECT GLMASTER, GLACCSORTKEY" & vbCrLf
      sSql = sSql & "      From EsReportBalanceSheet" & vbCrLf
      sSql = sSql & "         WHERE HASCHILD IS NULL" & vbCrLf
      sSql = sSql & "      ORDER BY GLACCSORTKEY" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   OPEN curAcctStruc" & vbCrLf
      sSql = sSql & "   FETCH NEXT FROM curAcctStruc INTO @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "   WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "   BEGIN" & vbCrLf
      sSql = sSql & "    IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "    BEGIN" & vbCrLf
      sSql = sSql & "      if (@PrevMaster <> @GlMasterAcc)" & vbCrLf
      sSql = sSql & "      BEGIN" & vbCrLf
      sSql = sSql & "         UPDATE EsReportBalanceSheet SET" & vbCrLf
      sSql = sSql & "            SortKeyRev = RIGHT ('0000'+ Cast(@RowCount as varchar), 4) + char(36)+ @GlMasterAcc" & vbCrLf
      sSql = sSql & "         WHERE GLMASTER = @GlMasterAcc AND HASCHILD IS NULL" & vbCrLf
      sSql = sSql & "         SET @RowCount = @RowCount + 1" & vbCrLf
      sSql = sSql & "         SET @PrevMaster = @GlMasterAcc" & vbCrLf
      sSql = sSql & "      End" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "    End" & vbCrLf
      sSql = sSql & "    FETCH NEXT FROM curAcctStruc INTO @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & "       " & vbCrLf
      sSql = sSql & "   Close curAcctStruc" & vbCrLf
      sSql = sSql & "   DEALLOCATE curAcctStruc" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   set @level = 8" & vbCrLf
      sSql = sSql & "   WHILE (@level >= 0)" & vbCrLf
      sSql = sSql & "   BEGIN" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      SET @PrevMaster = ''" & vbCrLf
      sSql = sSql & "        DECLARE curAcctStruc1 CURSOR  FOR" & vbCrLf
      sSql = sSql & "         SELECT DISTINCT GLACCTREF, GLMASTER, GLACCSORTKEY" & vbCrLf
      sSql = sSql & "         From EsReportBalanceSheet" & vbCrLf
      sSql = sSql & "            WHERE SORTKEYLEVEL = @level AND HASCHILD IS NOT NULL" & vbCrLf
      sSql = sSql & "         order by GLACCSORTKEY" & vbCrLf
      sSql = sSql & "    " & vbCrLf
      sSql = sSql & "        OPEN curAcctStruc1" & vbCrLf
      sSql = sSql & "        FETCH NEXT FROM curAcctStruc1 INTO @GlChildAcct, @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "        WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "        BEGIN" & vbCrLf
      sSql = sSql & "          IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "          BEGIN" & vbCrLf
      sSql = sSql & "            if (@PrevMaster <> @GlChildAcct)" & vbCrLf
      sSql = sSql & "            BEGIN" & vbCrLf
      sSql = sSql & "               SELECT TOP 1 @ChildKey = SortKeyRev,@SortKey = GLACCSORTKEY" & vbCrLf
      sSql = sSql & "               From EsReportBalanceSheet" & vbCrLf
      sSql = sSql & "                  WHERE SORTKEYLEVEL > @level AND GLMASTER = @GlChildAcct" & vbCrLf
      sSql = sSql & "               order by GLACCSORTKEY desc" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "               UPDATE EsReportBalanceSheet SET" & vbCrLf
      sSql = sSql & "                  SortKeyRev = Cast(@ChildKey as varchar(512)) + char(36)+ @GlMasterAcc" & vbCrLf
      sSql = sSql & "               WHERE GLACCTREF = @GlChildAcct AND GLMASTER = @GlMasterAcc" & vbCrLf
      sSql = sSql & "                  AND SORTKEYLEVEL = @level" & vbCrLf
      sSql = sSql & "               SET @PrevMaster = @GlChildAcct" & vbCrLf
      sSql = sSql & "            End" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "          End" & vbCrLf
      sSql = sSql & "          FETCH NEXT FROM curAcctStruc1 INTO @GlChildAcct, @GlMasterAcc, @GLSortKey" & vbCrLf
      sSql = sSql & "        End" & vbCrLf
      sSql = sSql & "               " & vbCrLf
      sSql = sSql & "        Close curAcctStruc1" & vbCrLf
      sSql = sSql & "        DEALLOCATE curAcctStruc1" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "        SET @level = @level - 1" & vbCrLf
      sSql = sSql & "   End" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   SELECT GLTOPMaster, GLTOPMASTERDESC, GLACCTREF, GLACCTNO, GLDESCR, GLMASTER, GLTYPE,GLINACTIVE, GLFSLEVEL," & vbCrLf
      sSql = sSql & "      SUMCURBAL , CurrentBal, SUMPREVBAL, PreviousBal, SORTKEYLEVEL, GLACCSORTKEY, SortKeyRev,HASCHILD" & vbCrLf
      sSql = sSql & "   FROM EsReportBalanceSheet ORDER BY SortKeyRev" & vbCrLf
      sSql = sSql & "End"
      ExecuteScript False, sSql

sql = "DropStoredProcedureIfExists 'GetMOOverHead'" & vbCrLf
ExecuteScript False, sql

sql = "create PROCEDURE [dbo].[GetMOOverHead]" & vbCrLf
sql = sql & "@InputLotNum as varchar(15),@MOPart as varchar(30),@MORun as int," & vbCrLf
sql = sql & "@MOQty as decimal(15,4), @CalTotOH decimal(15,4) OUTPUT" & vbCrLf
sql = sql & "-- modified 5/7/2018 by TEL -- never worked! semicolon before with CTE, dup set stmt, and comment out prints" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "set nocount on" & vbCrLf
sql = sql & "declare @SumTotMat decimal(15,4)" & vbCrLf
sql = sql & "declare @SumTotLabor decimal(15,4)" & vbCrLf
sql = sql & "declare @SumTotExp decimal(15,4)" & vbCrLf
sql = sql & "declare @SumTotOH decimal(15,4)" & vbCrLf
sql = sql & "declare @LotTotOH decimal (15,4)" & vbCrLf
sql = sql & "declare @level as integer" & vbCrLf
sql = sql & "declare @Part as varchar(30)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @PrevParent  as varchar(30)" & vbCrLf
sql = sql & "declare @RowCount as integer" & vbCrLf
sql = sql & "declare @ChildPart as varchar(30)" & vbCrLf
sql = sql & "declare @ParentPart as varchar(30)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @ChildKey as varchar(1024)" & vbCrLf
sql = sql & "declare @GLSortKey as varchar(1024)" & vbCrLf
sql = sql & "declare @SortKey as varchar(1024)" & vbCrLf
sql = sql & "declare @ParentLotNum as varchar(15)" & vbCrLf
sql = sql & "declare @Maxlevel as int" & vbCrLf
sql = sql & "declare @LotRunNo as int" & vbCrLf
sql = sql & "declare @LotOrgQty as decimal(15,4)" & vbCrLf
sql = sql & "--declare @MOPart as varchar(30)" & vbCrLf
sql = sql & "--declare @MORun as int" & vbCrLf
sql = sql & "--declare @MOQty as decimal(15,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CREATE TABLE #tempMOPartsDetail" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "INMOPART Varchar(30) NULL," & vbCrLf
sql = sql & "INMORUN int NULL ," & vbCrLf
sql = sql & "INPART varchar(30) NULL ," & vbCrLf
sql = sql & "LOTNUMBER varchar(15) NULL," & vbCrLf
sql = sql & "INTOTMATL decimal(12,4) NULL," & vbCrLf
sql = sql & "INTOTLABOR decimal(12,4) NULL," & vbCrLf
sql = sql & "INTOTEXP decimal(12,4) NULL," & vbCrLf
sql = sql & "INTOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTMATL decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTLABOR decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTEXP decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTMAL decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTLABOR decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTEXP decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTDATECOSTED smalldatetime NULL," & vbCrLf
sql = sql & "SortKey varchar(512) NULL," & vbCrLf
sql = sql & "HASCHILD int NULL," & vbCrLf
sql = sql & "SORTKEYLEVEL tinyint NULL," & vbCrLf
sql = sql & "SortKeyRev varchar(512)," & vbCrLf
sql = sql & "PARTSUM varchar(40)," & vbCrLf
sql = sql & "BMQTYREQD decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTORGQTY decimal(12,4) NULL," & vbCrLf
sql = sql & "BMTOTOH decimal(12,4) NULL" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--ALTER TABLE tempMOPartsDetail ADD SortKeyRev varchar(512)" & vbCrLf
sql = sql & "--ALTER TABLE tempMOPartsDetail ADD PARTSUM varchar(40)" & vbCrLf
sql = sql & "--ALTER TABLE tempMOPartsDetail ADD BMQTYREQD decimal(12,4) NULL, LOTORGQTY decimal(12,4) NULL" & vbCrLf
sql = sql & "--ALTER TABLE tempMOPartsDetail ADD BMTOTOH decimal(12,4) NULL" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- DELETE FROM tempMOPartsDetail" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@InputLotNum <> '')" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SELECT @MOPart = LOTMOPARTREF,@MORun =LOTMORUNNO, @MOQty = LOTORIGINALQTY" & vbCrLf
sql = sql & "FROM LohdTable WHERE LOTNUMBER = @InputLotNum" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & ";with cte" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "(select BMASSYPART, BMPARTREF,  BMQTYREQD,0 as level, cast('1' + char(36)+ BMPARTREF as varchar(max)) as SortKey" & vbCrLf
sql = sql & "from BmplTable" & vbCrLf
sql = sql & "where BMASSYPART = @MOPart" & vbCrLf
sql = sql & "union all" & vbCrLf
sql = sql & "select a.BMASSYPART, a.BMPARTREF, a.BMQTYREQD, level + 1," & vbCrLf
sql = sql & "cast(COALESCE(SortKey,'') + char(36) + COALESCE(a.BMPARTREF,'') as varchar(max))as SortKey" & vbCrLf
sql = sql & "from cte" & vbCrLf
sql = sql & "inner join BmplTable a" & vbCrLf
sql = sql & "on cte.BMPARTREF = a.BMASSYPART" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "INSERT INTO #tempMOPartsDetail(INMOPART,INPART,BMQTYREQD,SORTKEYLEVEL,SortKey)" & vbCrLf
sql = sql & "select BMASSYPART, BMPARTREF,BMQTYREQD,level,SortKey" & vbCrLf
sql = sql & "from cte order by SortKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--print 'TopLevel:' + @MOPart + ' RUN:' + convert(varchar(10), @MORun)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// update the top level" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET INMORUN = @MORun, LOTNUMBER = b.INLOTNUMBER," & vbCrLf
sql = sql & "INTOTMATL = b.INTOTMATL, INTOTLABOR = b.INTOTLABOR," & vbCrLf
sql = sql & "INTOTEXP = b.INTOTEXP, INTOTOH = b.INTOTOH," & vbCrLf
sql = sql & "LOTTOTMATL = c.LOTTOTMATL, LOTTOTLABOR = c.LOTTOTLABOR," & vbCrLf
sql = sql & "LOTTOTEXP = c.LOTTOTEXP, LOTTOTOH = c.LOTTOTOH," & vbCrLf
sql = sql & "LOTDATECOSTED = c.LOTDATECOSTED, LOTORGQTY = c.LOTORIGINALQTY," & vbCrLf
sql = sql & "BMTOTOH = (c.LOTTOTOH * BMQTYREQD) / c.LOTORIGINALQTY" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail d, InvaTable b, LohdTable c" & vbCrLf
sql = sql & "WHERE d.INMOPART = b.INMOPART AND d.INPART = b.INPART" & vbCrLf
sql = sql & "AND b.INLOTNUMBER = c.LOTnumber" & vbCrLf
sql = sql & "and c.lotpartref = b.INPART" & vbCrLf
sql = sql & "and b.INMOPART = @MOPart AND b.INMORUN  = @MORun" & vbCrLf
sql = sql & "AND b.INTYPE = 10 AND SORTKEYLEVEL = 0" & vbCrLf
sql = sql & "AND c.LOTORIGINALQTY <> 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// set the totals for" & vbCrLf
sql = sql & "SELECT @Maxlevel =  MAX(SORTKEYLEVEL) FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "SET @level  = 1" & vbCrLf
sql = sql & "WHILE (@level <= @Maxlevel )" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curMORun CURSOR  FOR" & vbCrLf
sql = sql & "SELECT DISTINCT INMOPART,INPART" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail WHERE SORTKEYLEVEL = @level" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curMORun" & vbCrLf
sql = sql & "FETCH NEXT FROM curMORun INTO @MOPart, @Part" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @ParentLotNum = LOTNUMBER FROM #tempMOPartsDetail  WHERE" & vbCrLf
sql = sql & "INPART = @MOPart AND SORTKEYLEVEL = @level - 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @LotRunNo = LOTMORUNNO, @LotOrgQty = LOTORIGINALQTY" & vbCrLf
sql = sql & "FROM lohdTable where LOTNUMBER = @ParentLotNum" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--print 'InLoopLevel:' + @MOPart + ' RUN:' + convert(varchar(10), @LotRunNo)" & vbCrLf
sql = sql & "--// update the top level" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET INMORUN = @LotRunNo, LOTNUMBER = b.INLOTNUMBER," & vbCrLf
sql = sql & "INTOTMATL = b.INTOTMATL, INTOTLABOR = b.INTOTLABOR," & vbCrLf
sql = sql & "INTOTEXP = b.INTOTEXP, INTOTOH = b.INTOTOH," & vbCrLf
sql = sql & "LOTTOTMATL = c.LOTTOTMATL, LOTTOTLABOR = c.LOTTOTLABOR," & vbCrLf
sql = sql & "LOTTOTEXP = c.LOTTOTEXP, LOTTOTOH = c.LOTTOTOH," & vbCrLf
sql = sql & "LOTDATECOSTED = c.LOTDATECOSTED, LOTORGQTY = @LotOrgQty," & vbCrLf
sql = sql & "BMTOTOH = (c.LOTTOTOH * BMQTYREQD) / @LotOrgQty" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail d, InvaTable b, LohdTable c" & vbCrLf
sql = sql & "WHERE d.INMOPART = b.INMOPART AND d.INPART = b.INPART" & vbCrLf
sql = sql & "AND b.INLOTNUMBER = c.LOTnumber" & vbCrLf
sql = sql & "and c.lotpartref = b.INPART" & vbCrLf
sql = sql & "and b.INMOPART = @MOPart AND b.INMORUN  = @LotRunNo" & vbCrLf
sql = sql & "AND b.INTYPE = 10 AND SORTKEYLEVEL = @level" & vbCrLf
sql = sql & "AND c.LOTORIGINALQTY <> 0" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curMORun INTO @MOPart, @Part" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curMORun" & vbCrLf
sql = sql & "DEALLOCATE curMORun" & vbCrLf
sql = sql & "SET @level = @level + 1" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @level =  MAX(SORTKEYLEVEL) FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "WHILE (@level >= 0 )" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curMODet CURSOR  FOR" & vbCrLf
sql = sql & "--SELECT INPART, LOTTOTMATL, LOTTOTLABOR, LOTTOTEXP , LOTTOTOH FROM tempMOPartsDetail" & vbCrLf
sql = sql & "-- WHERE INPART = '775345149'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT INMOPART," & vbCrLf
sql = sql & "SUM(IsNull(LOTTOTMATL, 0)), SUM(ISNULL(LOTTOTLABOR,0)) ," & vbCrLf
sql = sql & "Sum (IsNull(LOTTOTEXP, 0)) , SUM(IsNull(BMTOTOH, 0))" & vbCrLf
sql = sql & "From" & vbCrLf
sql = sql & "(SELECT DISTINCT INMOPART,INMORUN,INPART,LOTTOTMATL,LOTTOTLABOR," & vbCrLf
sql = sql & "LOTTOTEXP,LOTTOTOH,SUMTOTMAL,SUMTOTLABOR, SUMTOTEXP, SUMTOTOH,BMTOTOH" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail WHERE SORTKEYLEVEL = @level) as foo" & vbCrLf
sql = sql & "group by INMOPART" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curMODet" & vbCrLf
sql = sql & "FETCH NEXT FROM curMODet INTO @MOPart, @SumTotMat, @SumTotLabor,@SumTotExp,@SumTotOH" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--print 'PartNum : ' + @MOPart" & vbCrLf
sql = sql & "--print 'SumTotoh : ' + Convert(varchar(24), @SumTotOH)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET SUMTOTMAL = LOTTOTMATL + @SumTotMat," & vbCrLf
sql = sql & "SUMTOTLABOR = LOTTOTLABOR + @SumTotLabor," & vbCrLf
sql = sql & "SUMTOTEXP = LOTTOTEXP + @SumTotExp, SUMTOTOH = (BMTOTOH + @SumTotOH) * @MOQty ," & vbCrLf
sql = sql & "HASCHILD = 1,PARTSUM = 'TOTAL ' + LTRIM(INPART)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "WHERE INPART = @MOPart AND SORTKEYLEVEL = @level - 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curMODet INTO @MOPart, @SumTotMat, @SumTotLabor,@SumTotExp,@SumTotOH" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curMODet" & vbCrLf
sql = sql & "DEALLOCATE curMODet" & vbCrLf
sql = sql & "SET @level = @level - 1" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// update the Lower level cost detail" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET SUMTOTMAL = LOTTOTMATL, SUMTOTLABOR = LOTTOTLABOR," & vbCrLf
sql = sql & "SUMTOTEXP = LOTTOTEXP, SUMTOTOH = BMTOTOH WHERE HASCHILD IS NULL" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @SumTotMat  = 0" & vbCrLf
sql = sql & "SET @SumTotLabor  = 0" & vbCrLf
sql = sql & "SET @SumTotExp  = 0" & vbCrLf
sql = sql & "SET @SumTotOH  = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// Udpate the Root total" & vbCrLf
sql = sql & "SELECT @SumTotMat = SUM(SUMTOTMAL), @SumTotLabor = SUM(SUMTOTLABOR)," & vbCrLf
sql = sql & "@SumTotExp = SUM(SUMTOTEXP) ,@SumTotOH = SUM(SUMTOTOH)" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail WHERE SORTKEYLEVEL = 0" & vbCrLf
sql = sql & "AND  RTRIM(INMOPART) <> RTRIM(INPART)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET PARTSUM = INPART" & vbCrLf
sql = sql & "WHERE PARTSUM IS NULL" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "----  SELECT * FROM tempMOPartsDetail WHERE SORTKEYLEVEL = 0" & vbCrLf
sql = sql & "----AND  RTRIM(INMOPART) = RTRIM(INPART)" & vbCrLf
sql = sql & "--// Reverse the partnumbers." & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @level = 0" & vbCrLf
sql = sql & "SET @RowCount = 1" & vbCrLf
sql = sql & "SET @PrevParent = ''" & vbCrLf
sql = sql & "DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
sql = sql & "SELECT INMOPART, SortKey" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "WHERE HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLMASTER ='AA10'--GLTOPMaster = 1 AND" & vbCrLf
sql = sql & "ORDER BY SortKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curAcctStruc" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc INTO @ParentPart, @SortKey" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "if (@PrevParent <> @ParentPart)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET" & vbCrLf
sql = sql & "SortKeyRev = RIGHT ('0000'+ Cast(@RowCount as varchar), 4) + char(36)+ @ParentPart" & vbCrLf
sql = sql & "WHERE INMOPART = @ParentPart AND HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLTOPMaster = 1" & vbCrLf
sql = sql & "SET @RowCount = @RowCount + 1" & vbCrLf
sql = sql & "SET @PrevParent = @ParentPart" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc INTO @ParentPart, @SortKey" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "Close curAcctStruc" & vbCrLf
sql = sql & "DEALLOCATE curAcctStruc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @level =  MAX(SORTKEYLEVEL) FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "--set @level = 7" & vbCrLf
sql = sql & "WHILE (@level >= 0 )" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @PrevParent = ''" & vbCrLf
sql = sql & "DECLARE curAcctStruc1 CURSOR  FOR" & vbCrLf
sql = sql & "SELECT DISTINCT INPART, INMOPART, SortKey" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "WHERE SORTKEYLEVEL = @level AND HASCHILD IS NOT NULL--GLTOPMaster = 1 AND" & vbCrLf
sql = sql & "order by SortKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curAcctStruc1" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc1 INTO @ChildPart, @ParentPart, @SortKey" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if (@PrevParent <> @ChildPart)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--print 'Record' + @ChildPart + ':' + @ParentPart" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT TOP 1 @ChildKey = SortKeyRev,@SortKey = SortKey" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "WHERE SORTKEYLEVEL > @level AND INMOPART = @ChildPart --GLTOPMaster = 1 AND" & vbCrLf
sql = sql & "order by SortKey desc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET" & vbCrLf
sql = sql & "SortKeyRev = Cast(@ChildKey as varchar(256)) + char(36)+ @ParentPart" & vbCrLf
sql = sql & "WHERE INPART = @ChildPart AND INMOPART = @ParentPart" & vbCrLf
sql = sql & "AND SORTKEYLEVEL = @level --GLTOPMaster = 1 AND" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @PrevParent = @ChildPart" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc1 INTO @ChildPart, @ParentPart, @SortKey" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "Close curAcctStruc1" & vbCrLf
sql = sql & "DEALLOCATE curAcctStruc1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @level = @level - 1" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @LotTotOH = ISNULL(LOTTOTOH, 0) from lohdTable where lotpartref = @MOPart  AND LOTMORUNNO = @MORun" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--print 'LOT OH:' + Convert(varchar(10), @LotTotOH)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @CalTotOH = SUM(ISNULL(BMTOTOH, 0)) + ISNULL(@LotTotOH, 0) FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "WHERE SORTKEYLEVEL = 0" & vbCrLf
sql = sql & "--print 'GetMOOverHead: CalOH:' + convert(varchar(10), @CalTotOH)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DROP table #tempMOPartsDetail" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript False, sql

sql = "DropStoredProcedureIfExists 'RptChartOfAccount'" & vbCrLf
ExecuteScript False, sql

sql = "CREATE PROCEDURE RptChartOfAccount" & vbCrLf
sql = sql & "@InclIncAcct as varchar(1)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "05/09/2018 TEL - changed 'with cte' to ';with cte'" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptChartOfAccount 0" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @glAcctRef as varchar(10)" & vbCrLf
sql = sql & "declare @glMsAcct as varchar(10)" & vbCrLf
sql = sql & "declare @TopLevelDesc as varchar(30)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @level as varchar(12)" & vbCrLf
sql = sql & "declare @InclInAcct as Integer" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if (@InclIncAcct = '1')" & vbCrLf
sql = sql & "SET @InclInAcct = ''" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "SET @InclInAcct = '0'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CREATE TABLE #tempChartOfAcct(" & vbCrLf
sql = sql & "[TOPLEVEL] [varchar](12) NULL," & vbCrLf
sql = sql & "[TOPLEVELDESC] [varchar](30) NULL," & vbCrLf
sql = sql & "[GLACCTREF] [varchar](112) NULL," & vbCrLf
sql = sql & "[GLDESCR] [varchar](120) NULL," & vbCrLf
sql = sql & "[GLMASTER] [varchar](12) NULL," & vbCrLf
sql = sql & "[GLFSLEVEL] [INT] NULL," & vbCrLf
sql = sql & "[GLINACTIVE] [int] NULL," & vbCrLf
sql = sql & "[SORTKEYLEVEL] [int] NULL," & vbCrLf
sql = sql & "[GLACCSORTKEY] [varchar](512) NULL" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE balAcctStruc CURSOR  FOR" & vbCrLf
sql = sql & "SELECT COASSTACCT, COASSTDESC FROM GlmsTable" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT COLIABACCT, COLIABDESC FROM GlmsTable" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT COINCMACCT, COINCMDESC FROM GlmsTable" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT COEQTYACCT, COEQTYDESC FROM GlmsTable" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT COCOGSACCT, COCOGSDESC FROM GlmsTable" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT COEXPNACCT, COEXPNDESC FROM GlmsTable" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT COOINCACCT, COOINCDESC FROM GlmsTable" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT COFDTXACCT, COFDTXDESC FROM GlmsTable" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT COOEXPACCT, COOEXPDESC FROM GlmsTable" & vbCrLf
sql = sql & "--UNION ALL" & vbCrLf
sql = sql & "--SELECT COFDTXACCT, COFDTXDESC FROM GlmsTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN balAcctStruc" & vbCrLf
sql = sql & "FETCH NEXT FROM balAcctStruc INTO @level, @TopLevelDesc" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO #tempChartOfAcct(TOPLEVEL, TOPLEVELDESC, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL,GLINACTIVE, SORTKEYLEVEL,GLACCSORTKEY)" & vbCrLf
sql = sql & "select @level as TopLevel, @TopLevelDesc as TopLevelDesc, @level as GLACCTREF," & vbCrLf
sql = sql & "@TopLevelDesc as GLDESCR, '' as GLMASTER, 0 as GLFSLEVEL, 0,0 as level, @level as SortKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & ";with cte" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "(select GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL, GLINACTIVE, 0 as level," & vbCrLf
sql = sql & "cast(cast(@level as varchar(12))+ char(36)+ GLACCTREF as varchar(max)) as SortKey" & vbCrLf
sql = sql & "from GlacTable" & vbCrLf
sql = sql & "where GLMASTER = cast(@level as varchar(12)) AND GLINACTIVE LIKE @InclInAcct" & vbCrLf
sql = sql & "union all" & vbCrLf
sql = sql & "select a.GLACCTREF, a.GLDESCR, a.GLMASTER, a.GLFSLEVEL, a.GLINACTIVE, level + 1," & vbCrLf
sql = sql & "cast(COALESCE(SortKey,'') + char(36) + COALESCE(a.GLACCTREF,'') as varchar(max))as SortKey" & vbCrLf
sql = sql & "from cte" & vbCrLf
sql = sql & "inner join GlacTable a" & vbCrLf
sql = sql & "on cte.GLACCTREF = a.GLMASTER" & vbCrLf
sql = sql & "WHERE a.GLINACTIVE LIKE @InclInAcct" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO #tempChartOfAcct(TOPLEVEL, TOPLEVELDESC, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL,GLINACTIVE, SORTKEYLEVEL,GLACCSORTKEY)" & vbCrLf
sql = sql & "select @level as TopLevel, @TopLevelDesc as TopLevelDesc," & vbCrLf
sql = sql & "Replicate('  ', level) + GLACCTREF as GLACCTREF," & vbCrLf
sql = sql & "Replicate('  ', level) + GLDESCR as GLDESCR, GLMASTER," & vbCrLf
sql = sql & "GLFSLEVEL, GLINACTIVE,level, SortKey" & vbCrLf
sql = sql & "from cte order by SortKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "FETCH NEXT FROM balAcctStruc INTO @level, @TopLevelDesc" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "CLOSE balAcctStruc" & vbCrLf
sql = sql & "DEALLOCATE balAcctStruc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT TOPLEVEL, TOPLEVELDESC, GLACCTREF, GLDESCR, GLMASTER," & vbCrLf
sql = sql & "GLFSLEVEL,GLINACTIVE, SORTKEYLEVEL,GLACCSORTKEY" & vbCrLf
sql = sql & "FROM #tempChartOfAcct ORDER BY GLACCSORTKEY" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DROP table #tempChartOfAcct" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript False, sql

sql = "DropStoredProcedureIfExists 'RptIncomeStatement'" & vbCrLf
ExecuteScript False, sql

sql = "CREATE PROCEDURE RptIncomeStatement" & vbCrLf
sql = sql & "@StartDate as varchar(12),@EndDate as varchar(12)," & vbCrLf
sql = sql & "@YearBeginDate as varchar(12), @InclIncAcct as varchar(1)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--05/09/2018 TEL - changed 'with cte' to ';with cte'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @glAcctRef as varchar(10)" & vbCrLf
sql = sql & "declare @glMsAcct as varchar(10)" & vbCrLf
sql = sql & "declare @SumCurBal decimal(15,4)" & vbCrLf
sql = sql & "declare @SumYTD decimal(15,4)" & vbCrLf
sql = sql & "declare @SumPrevBal as decimal(15,4)" & vbCrLf
sql = sql & "declare @level as integer" & vbCrLf
sql = sql & "declare @TopLevelDesc as varchar(30)" & vbCrLf
sql = sql & "declare @InclInAcct as Integer" & vbCrLf
sql = sql & "declare @TopLevAcct as varchar(20)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @PrevMaster as varchar(10)" & vbCrLf
sql = sql & "declare @RowCount as integer" & vbCrLf
sql = sql & "declare @GlMasterAcc as varchar(10)" & vbCrLf
sql = sql & "declare @GlChildAcct as varchar(10)" & vbCrLf
sql = sql & "declare @ChildKey as varchar(1024)" & vbCrLf
sql = sql & "declare @GLSortKey as varchar(1024)" & vbCrLf
sql = sql & "declare @SortKey as varchar(1024)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DELETE FROM EsReportIncStatement" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if (@InclIncAcct = '1')" & vbCrLf
sql = sql & "SET @InclInAcct = ''" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @InclInAcct = '0'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE balAcctStruc CURSOR  FOR" & vbCrLf
sql = sql & "SELECT '4', COINCMACCT, COINCMDESC FROM GlmsTable" & vbCrLf
sql = sql & "Union All" & vbCrLf
sql = sql & "SELECT '5', COCOGSACCT, COCOGSDESC FROM GlmsTable" & vbCrLf
sql = sql & "Union All" & vbCrLf
sql = sql & "SELECT '6', COEXPNACCT, COEXPNDESC FROM GlmsTable" & vbCrLf
sql = sql & "Union All" & vbCrLf
sql = sql & "SELECT '7', COOINCACCT, COOINCDESC FROM GlmsTable" & vbCrLf
sql = sql & "Union All" & vbCrLf
sql = sql & "SELECT '8', COOEXPACCT, COOEXPDESC FROM GlmsTable" & vbCrLf
sql = sql & "Union All" & vbCrLf
sql = sql & "SELECT '9', COFDTXACCT, COFDTXDESC FROM GlmsTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN balAcctStruc" & vbCrLf
sql = sql & "FETCH NEXT FROM balAcctStruc INTO @level,@TopLevAcct, @TopLevelDesc" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & ";with cte" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "(select GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL, 1 as level," & vbCrLf
sql = sql & "cast(cast(@level as varchar(4))+ char(36)+ GLACCTREF as varchar(max)) as SortKey" & vbCrLf
sql = sql & "From GlacTable" & vbCrLf
sql = sql & "where GLMASTER = cast(@TopLevAcct as varchar(20)) AND GLINACTIVE LIKE @InclInAcct" & vbCrLf
sql = sql & "Union All" & vbCrLf
sql = sql & "select a.GLACCTREF, a.GLDESCR, a.GLMASTER, a.GLFSLEVEL, level + 1," & vbCrLf
sql = sql & "cast(COALESCE(SortKey,'') + char(36) + COALESCE(a.GLACCTREF,'') as varchar(max))as SortKey" & vbCrLf
sql = sql & "From cte" & vbCrLf
sql = sql & "inner join GlacTable a" & vbCrLf
sql = sql & "on cte.GLACCTREF = a.GLMASTER" & vbCrLf
sql = sql & "WHERE GLINACTIVE LIKE @InclInAcct" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "INSERT INTO EsReportIncStatement(GLTOPMaster, GLTOPMASTERDESC, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL,SORTKEYLEVEL,GLACCSORTKEY)" & vbCrLf
sql = sql & "select @level, @TopLevelDesc, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL, level, SortKey" & vbCrLf
sql = sql & "from cte order by SortKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM balAcctStruc INTO @level,@TopLevAcct, @TopLevelDesc" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close balAcctStruc" & vbCrLf
sql = sql & "DEALLOCATE balAcctStruc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE EsReportIncStatement SET CurrentBal = foo.Balance--, SUMCURBAL = foo.Balance" & vbCrLf
sql = sql & "From" & vbCrLf
sql = sql & "(SELECT SUM(GjitTable.JICRD) - SUM(GjitTable.JIDEB) as Balance, JIACCOUNT" & vbCrLf
sql = sql & "FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME" & vbCrLf
sql = sql & "WHERE GJPOST BETWEEN @StartDate AND @EndDate" & vbCrLf
sql = sql & "AND GjhdTable.GJPOSTED = 1" & vbCrLf
sql = sql & "GROUP BY JIACCOUNT) as foo" & vbCrLf
sql = sql & "Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE EsReportIncStatement SET YTD = foo.Balance--, SUMYTD = foo.Balance" & vbCrLf
sql = sql & "From" & vbCrLf
sql = sql & "(SELECT JIACCOUNT,SUM(GjitTable.JICRD) - SUM(GjitTable.JIDEB) AS Balance" & vbCrLf
sql = sql & "FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME" & vbCrLf
sql = sql & "WHERE (GJPOST BETWEEN @YearBeginDate AND @EndDate)" & vbCrLf
sql = sql & "AND GjhdTable.GJPOSTED = 1" & vbCrLf
sql = sql & "GROUP BY JIACCOUNT) as foo" & vbCrLf
sql = sql & "Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE EsReportIncStatement SET PreviousBal = foo.Balance--, SUMPREVBAL = foo.Balance" & vbCrLf
sql = sql & "From" & vbCrLf
sql = sql & "(SELECT JIACCOUNT,SUM(GjitTable.JICRD) - SUM(GjitTable.JIDEB) AS Balance" & vbCrLf
sql = sql & "FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME" & vbCrLf
sql = sql & "WHERE (GJPOST BETWEEN DATEADD(year, -1, @YearBeginDate) AND DATEADD(year, -1, @EndDate))" & vbCrLf
sql = sql & "AND GjhdTable.GJPOSTED = 1" & vbCrLf
sql = sql & "GROUP BY JIACCOUNT) as foo" & vbCrLf
sql = sql & "Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @level =  MAX(SORTKEYLEVEL) FROM EsReportIncStatement" & vbCrLf
sql = sql & "--set @level = 9" & vbCrLf
sql = sql & "WHILE (@level >= 1 )" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
sql = sql & "SELECT GLMASTER, SUM(ISNULL(SUMCURBAL,0) + (ISNULL(CurrentBal,0))) ," & vbCrLf
sql = sql & "Sum (IsNull(SUMYTD, 0) + (IsNull(YTD, 0))), Sum(IsNull(SUMPREVBAL, 0) + (IsNull(PreviousBal, 0)))" & vbCrLf
sql = sql & "From" & vbCrLf
sql = sql & "(SELECT DISTINCT GLACCTREF, GLDESCR, GLMASTER," & vbCrLf
sql = sql & "CurrentBal , YTD, PreviousBal, SUMCURBAL, SUMYTD, SUMPREVBAL" & vbCrLf
sql = sql & "FROM EsReportIncStatement WHERE SORTKEYLEVEL = @level) as foo" & vbCrLf
sql = sql & "group by GLMASTER" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curAcctStruc" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc INTO @glMsAcct, @SumCurBal, @SumYTD, @SumPrevBal" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "UPDATE EsReportIncStatement SET SUMCURBAL = @SumCurBal, SUMYTD = @SumYTD," & vbCrLf
sql = sql & "SUMPREVBAL = @SumPrevBal, GLDESCR = 'TOTAL ' + LTRIM(GLDESCR)," & vbCrLf
sql = sql & "HASCHILD = 1" & vbCrLf
sql = sql & "WHERE GLACCTREF = @glMsAcct" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc INTO @glMsAcct, @SumCurBal, @SumYTD, @SumPrevBal" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curAcctStruc" & vbCrLf
sql = sql & "DEALLOCATE curAcctStruc" & vbCrLf
sql = sql & "SET @level = @level - 1" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE EsReportIncStatement SET SUMCURBAL = CurrentBal WHERE SUMCURBAL IS NULL" & vbCrLf
sql = sql & "UPDATE EsReportIncStatement SET SUMPREVBAL = PreviousBal WHERE SUMPREVBAL IS NULL" & vbCrLf
sql = sql & "UPDATE EsReportIncStatement SET SUMYTD = YTD  WHERE SUMYTD IS NULL" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @level = 0" & vbCrLf
sql = sql & "set @RowCount = 1" & vbCrLf
sql = sql & "SET @PrevMaster = ''" & vbCrLf
sql = sql & "DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
sql = sql & "SELECT GLMASTER, GLACCSORTKEY" & vbCrLf
sql = sql & "FROM EsReportIncStatement" & vbCrLf
sql = sql & "WHERE HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLMASTER ='AA10'--GLTOPMaster = 1 AND" & vbCrLf
sql = sql & "ORDER BY GLACCSORTKEY" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curAcctStruc" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc INTO @GlMasterAcc, @GLSortKey" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "if (@PrevMaster <> @GlMasterAcc)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "UPDATE EsReportIncStatement SET" & vbCrLf
sql = sql & "SortKeyRev = RIGHT ('0000'+ Cast(@RowCount as varchar), 4) + char(36)+ @GlMasterAcc" & vbCrLf
sql = sql & "WHERE GLMASTER = @GlMasterAcc AND HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLTOPMaster = 1" & vbCrLf
sql = sql & "SET @RowCount = @RowCount + 1" & vbCrLf
sql = sql & "SET @PrevMaster = @GlMasterAcc" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc INTO @GlMasterAcc, @GLSortKey" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "Close curAcctStruc" & vbCrLf
sql = sql & "DEALLOCATE curAcctStruc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @level = MAX(SORTKEYLEVEL) FROM EsReportIncStatement" & vbCrLf
sql = sql & "--set @level = 7" & vbCrLf
sql = sql & "WHILE (@level >= 1 )" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @PrevMaster = ''" & vbCrLf
sql = sql & "DECLARE curAcctStruc1 CURSOR  FOR" & vbCrLf
sql = sql & "SELECT DISTINCT GLACCTREF, GLMASTER, GLACCSORTKEY" & vbCrLf
sql = sql & "FROM EsReportIncStatement" & vbCrLf
sql = sql & "WHERE SORTKEYLEVEL = @level AND HASCHILD IS NOT NULL--GLTOPMaster = 1 AND" & vbCrLf
sql = sql & "order by GLACCSORTKEY" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curAcctStruc1" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc1 INTO @GlChildAcct, @GlMasterAcc, @GLSortKey" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if (@PrevMaster <> @GlChildAcct)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "print 'Record' + @GlChildAcct + ':' + @GlMasterAcc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT TOP 1 @ChildKey = SortKeyRev,@SortKey = GLACCSORTKEY" & vbCrLf
sql = sql & "FROM EsReportIncStatement" & vbCrLf
sql = sql & "WHERE SORTKEYLEVEL > @level AND GLMASTER = @GlChildAcct --GLTOPMaster = 1 AND" & vbCrLf
sql = sql & "order by GLACCSORTKEY desc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE EsReportIncStatement SET" & vbCrLf
sql = sql & "SortKeyRev = Cast(@ChildKey as varchar(512)) + char(36)+ @GlMasterAcc" & vbCrLf
sql = sql & "WHERE GLACCTREF = @GlChildAcct AND GLMASTER = @GlMasterAcc" & vbCrLf
sql = sql & "AND SORTKEYLEVEL = @level --GLTOPMaster = 1 AND" & vbCrLf
sql = sql & "SET @PrevMaster = @GlChildAcct" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc1 INTO @GlChildAcct, @GlMasterAcc, @GLSortKey" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "Close curAcctStruc1" & vbCrLf
sql = sql & "DEALLOCATE curAcctStruc1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @level = @level - 1" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT GLTOPMaster, GLTOPMASTERDESC, GLACCTREF, GLACCTNO, GLDESCR, GLMASTER, GLTYPE,GLINACTIVE, GLFSLEVEL," & vbCrLf
sql = sql & "SUMCURBAL , CurrentBal, SUMYTD, YTD, SUMPREVBAL, PreviousBal, SORTKEYLEVEL, GLACCSORTKEY" & vbCrLf
sql = sql & "FROM EsReportIncStatement ORDER BY SortKeyRev --GLTOPMASTER, GLACCSORTKEY desc, SortKeyLevel" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
ExecuteScript False, sql


sql = "DropStoredProcedureIfExists 'RptMOCostDetail'" & vbCrLf
ExecuteScript False, sql

sql = "CREATE PROCEDURE RptMOCostDetail" & vbCrLf
sql = sql & "@MOPart as varchar(30),@MORun as int, @MOQty as decimal(15,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--05/09/2018 TEL - changed 'with cte' to ';with cte'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @SumTotMat decimal(15,4)" & vbCrLf
sql = sql & "declare @SumTotLabor decimal(15,4)" & vbCrLf
sql = sql & "declare @SumTotExp decimal(15,4)" & vbCrLf
sql = sql & "declare @SumTotOH decimal(15,4)" & vbCrLf
sql = sql & "declare @level as integer" & vbCrLf
sql = sql & "declare @Part as varchar(30)" & vbCrLf
sql = sql & "declare @PrevParent  as varchar(30)" & vbCrLf
sql = sql & "declare @RowCount as integer" & vbCrLf
sql = sql & "declare @ChildPart as varchar(30)" & vbCrLf
sql = sql & "declare @ParentPart as varchar(30)" & vbCrLf
sql = sql & "declare @MOPart1 as varchar(30)" & vbCrLf
sql = sql & "declare @MoRun1 as varchar(20)" & vbCrLf
sql = sql & "declare @Part1 as varchar(30)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @ChildKey as varchar(1024)" & vbCrLf
sql = sql & "declare @GLSortKey as varchar(1024)" & vbCrLf
sql = sql & "declare @SortKey as varchar(1024)" & vbCrLf
sql = sql & "declare @ParentLotNum as varchar(15)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @Maxlevel as int" & vbCrLf
sql = sql & "declare @LotRunNo as int" & vbCrLf
sql = sql & "declare @LotOrgQty as decimal(15,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @LotUSpMat decimal(15,4)" & vbCrLf
sql = sql & "declare @LotUSpLabor decimal(15,4)" & vbCrLf
sql = sql & "declare @LotUSpExp decimal(15,4)" & vbCrLf
sql = sql & "declare @LotUSpOH decimal(15,4)" & vbCrLf
sql = sql & "declare @LotMatl decimal(15,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @MOLotNum as varchar(15)" & vbCrLf
sql = sql & "declare @SplitLot as varchar(15)" & vbCrLf
sql = sql & "declare @cnt  as int" & vbCrLf
sql = sql & "declare @sumQty decimal(15,4)" & vbCrLf
sql = sql & "declare @MOPartRunKey as Varchar(30)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--DROP TABLE #tempMOPartsDetail" & vbCrLf
sql = sql & "-- DELETE FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @MOPartRunKey = RTRIM(@MOPart) + '_' + Convert(varchar(10), @MORun)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CREATE TABLE #tempMOPartsDetail" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "LOTMOPARTRUNKEY varchar(50) NULL," & vbCrLf
sql = sql & "INMOPART Varchar(30) NULL," & vbCrLf
sql = sql & "INMORUN int NULL ," & vbCrLf
sql = sql & "INPART varchar(30) NULL ," & vbCrLf
sql = sql & "LOTNUMBER varchar(15) NULL," & vbCrLf
sql = sql & "LOTUSERLOTID varchar(40) NULL," & vbCrLf
sql = sql & "INTOTMATL decimal(12,4) NULL," & vbCrLf
sql = sql & "INTOTLABOR decimal(12,4) NULL," & vbCrLf
sql = sql & "INTOTEXP decimal(12,4) NULL," & vbCrLf
sql = sql & "INTOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTMATL decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTLABOR decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTEXP decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTTOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTMAL decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTLABOR decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTEXP decimal(12,4) NULL," & vbCrLf
sql = sql & "SUMTOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTDATECOSTED smalldatetime NULL," & vbCrLf
sql = sql & "SortKey varchar(512) NULL," & vbCrLf
sql = sql & "HASCHILD int NULL," & vbCrLf
sql = sql & "SORTKEYLEVEL tinyint NULL," & vbCrLf
sql = sql & "SortKeyRev varchar(512)," & vbCrLf
sql = sql & "PARTSUM varchar(40)," & vbCrLf
sql = sql & "BMQTYREQD decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTORGQTY decimal(12,4) NULL," & vbCrLf
sql = sql & "BMTOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "LOTSPLITFROMSYS varchar(15)," & vbCrLf
sql = sql & "INVNO int NULL," & vbCrLf
sql = sql & "ITPSNUMBER varchar(8) NULL," & vbCrLf
sql = sql & "ITPSITEM smallint NULL," & vbCrLf
sql = sql & "PICKQTY decimal(12,4) NULL" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & ";with cte" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "(select BMASSYPART, BMPARTREF,  BMQTYREQD,0 as level, cast('1' + char(36)+ BMPARTREF as varchar(max)) as SortKey" & vbCrLf
sql = sql & "from BmplTable" & vbCrLf
sql = sql & "where BMASSYPART = @MOPart" & vbCrLf
sql = sql & "union all" & vbCrLf
sql = sql & "select a.BMASSYPART, a.BMPARTREF, a.BMQTYREQD, level + 1," & vbCrLf
sql = sql & "cast(COALESCE(SortKey,'') + char(36) + COALESCE(a.BMPARTREF,'') as varchar(max))as SortKey" & vbCrLf
sql = sql & "from cte" & vbCrLf
sql = sql & "inner join BmplTable a" & vbCrLf
sql = sql & "on cte.BMPARTREF = a.BMASSYPART" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "INSERT INTO #tempMOPartsDetail(INMOPART,INPART,BMQTYREQD,SORTKEYLEVEL,SortKey)" & vbCrLf
sql = sql & "select BMASSYPART, BMPARTREF,BMQTYREQD,level,SortKey" & vbCrLf
sql = sql & "from cte order by SortKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @cnt = 0" & vbCrLf
sql = sql & "print 'Update Start:' + cast(getdate() as char(25))" & vbCrLf
sql = sql & "print 'Count :' + Convert(varchar(10), @cnt)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET INMORUN = @MORun, LOTNUMBER = b.INLOTNUMBER," & vbCrLf
sql = sql & "LOTSPLITFROMSYS = c.LOTSPLITFROMSYS" & vbCrLf
sql = sql & "FROM dbo.LohdTable AS c INNER JOIN" & vbCrLf
sql = sql & "dbo.InvaTable AS b ON c.LOTNUMBER = b.INLOTNUMBER AND c.LOTPARTREF = b.INPART LEFT OUTER JOIN" & vbCrLf
sql = sql & "dbo.#tempMOPartsDetail AS d ON b.INMOPART = d.INMOPART AND b.INPART = d.INPART" & vbCrLf
sql = sql & "WHERE     (b.INMOPART = @MOPart) AND (b.INMORUN = @MORun) AND (b.INTYPE = 10) AND (d.SORTKEYLEVEL = 0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET PICKQTY = sumqty * -1" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail," & vbCrLf
sql = sql & "(SELECT SUM(b.INAQTY) sumqty, d.INMOPART mopart, d.INMORUN morun, d.LOTNUMBER lotnum, d.INPART subpart" & vbCrLf
sql = sql & "FROM dbo.LohdTable AS c INNER JOIN" & vbCrLf
sql = sql & "dbo.InvaTable AS b ON c.LOTNUMBER = b.INLOTNUMBER AND c.LOTPARTREF = b.INPART LEFT OUTER JOIN" & vbCrLf
sql = sql & "dbo.#tempMOPartsDetail AS d ON b.INMOPART = d.INMOPART AND b.INPART = d.INPART" & vbCrLf
sql = sql & "WHERE     b.INMOPART = @MOPart AND b.INMORUN  = @MORun AND (b.INTYPE = 10) AND (d.SORTKEYLEVEL = 0)" & vbCrLf
sql = sql & "GROUP BY d.INMOPART, d.INMORUN, d.LOTNUMBER, d.INPART" & vbCrLf
sql = sql & ") as f" & vbCrLf
sql = sql & "WHERE INMOPART = f.mopart AND INMORUN = f.morun" & vbCrLf
sql = sql & "AND LOTNUMBER = lotnum AND INPART = subpart" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// Update the" & vbCrLf
sql = sql & "--// update the top level" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET INMORUN = @MORun, LOTNUMBER = b.INLOTNUMBER," & vbCrLf
sql = sql & "LOTUSERLOTID = c.LOTUSERLOTID," & vbCrLf
sql = sql & "INTOTMATL = b.INTOTMATL, INTOTLABOR = b.INTOTLABOR," & vbCrLf
sql = sql & "INTOTEXP = b.INTOTEXP, INTOTOH = b.INTOTOH," & vbCrLf
sql = sql & "LOTTOTMATL = (c.LOTTOTMATL * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTLABOR = (c.LOTTOTLABOR * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTEXP = (c.LOTTOTEXP * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTOH = (c.LOTTOTOH * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTDATECOSTED = c.LOTDATECOSTED, LOTORGQTY = c.LOTORIGINALQTY," & vbCrLf
sql = sql & "BMTOTOH = (c.LOTTOTOH * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTSPLITFROMSYS = c.LOTSPLITFROMSYS" & vbCrLf
sql = sql & "FROM dbo.LohdTable AS c INNER JOIN" & vbCrLf
sql = sql & "dbo.InvaTable AS b ON c.LOTNUMBER = b.INLOTNUMBER AND c.LOTPARTREF = b.INPART LEFT OUTER JOIN" & vbCrLf
sql = sql & "dbo.#tempMOPartsDetail AS d ON b.INMOPART = d.INMOPART AND b.INPART = d.INPART" & vbCrLf
sql = sql & "WHERE     (b.INMOPART = @MOPart) AND (b.INMORUN = @MORun) AND (b.INTYPE = 10) AND (d.SORTKEYLEVEL = 0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "print 'Update 2:' + cast(getdate() as char(25))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--FROM #tempMOPartsDetail d, InvaTable b, LohdTable c" & vbCrLf
sql = sql & "--WHERE d.INMOPART = b.INMOPART AND d.INPART = b.INPART" & vbCrLf
sql = sql & "-- AND b.INLOTNUMBER = c.LOTnumber" & vbCrLf
sql = sql & "-- and c.lotpartref = b.INPART" & vbCrLf
sql = sql & "-- and b.INMOPART = @MOPart AND b.INMORUN  = @MORun" & vbCrLf
sql = sql & "-- AND b.INTYPE = 10 AND SORTKEYLEVEL = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// set the totals for" & vbCrLf
sql = sql & "SELECT @Maxlevel =  MAX(SORTKEYLEVEL) FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "SET @level  = 1" & vbCrLf
sql = sql & "WHILE (@level <= @Maxlevel )" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curMORun CURSOR  FOR" & vbCrLf
sql = sql & "SELECT DISTINCT INMOPART,INPART" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail WHERE SORTKEYLEVEL = @level" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curMORun" & vbCrLf
sql = sql & "FETCH NEXT FROM curMORun INTO @MOPart, @Part" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @ParentLotNum = LOTNUMBER FROM #tempMOPartsDetail WHERE" & vbCrLf
sql = sql & "INPART = @MOPart AND SORTKEYLEVEL = @level - 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @LotRunNo = LOTMORUNNO, @LotOrgQty = LOTORIGINALQTY" & vbCrLf
sql = sql & "FROM lohdTable where LOTNUMBER = @ParentLotNum" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET INMORUN = @LotRunNo, LOTNUMBER = b.INLOTNUMBER," & vbCrLf
sql = sql & "LOTSPLITFROMSYS = c.LOTSPLITFROMSYS" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail d, InvaTable b, LohdTable c" & vbCrLf
sql = sql & "WHERE d.INMOPART = b.INMOPART AND d.INPART = b.INPART" & vbCrLf
sql = sql & "AND b.INLOTNUMBER = c.LOTnumber" & vbCrLf
sql = sql & "and c.lotpartref = b.INPART" & vbCrLf
sql = sql & "and b.INMOPART = @MOPart AND b.INMORUN  = @LotRunNo" & vbCrLf
sql = sql & "AND b.INTYPE = 10 AND SORTKEYLEVEL = @level" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET PICKQTY = sumqty * -1" & vbCrLf
sql = sql & "FROM" & vbCrLf
sql = sql & "(SELECT SUM(b.INAQTY) sumqty, d.INMOPART mopart, d.INMORUN morun, d.LOTNUMBER lotnum, d.INPART subpart" & vbCrLf
sql = sql & "FROM dbo.LohdTable AS c INNER JOIN" & vbCrLf
sql = sql & "dbo.InvaTable AS b ON c.LOTNUMBER = b.INLOTNUMBER AND c.LOTPARTREF = b.INPART LEFT OUTER JOIN" & vbCrLf
sql = sql & "dbo.#tempMOPartsDetail AS d ON b.INMOPART = d.INMOPART AND b.INPART = d.INPART" & vbCrLf
sql = sql & "WHERE     b.INMOPART = @MOPart AND b.INMORUN  = @LotRunNo AND (b.INTYPE = 10) AND (d.SORTKEYLEVEL = @level)" & vbCrLf
sql = sql & "GROUP BY d.INMOPART, d.INMORUN, d.LOTNUMBER, d.INPART" & vbCrLf
sql = sql & ") as f" & vbCrLf
sql = sql & "WHERE INMOPART = mopart AND INMORUN = morun" & vbCrLf
sql = sql & "AND LOTNUMBER = lotnum AND INPART = subpart" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// update the top level" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET INMORUN = @LotRunNo, LOTNUMBER = b.INLOTNUMBER," & vbCrLf
sql = sql & "LOTUSERLOTID = c.LOTUSERLOTID," & vbCrLf
sql = sql & "INTOTMATL = b.INTOTMATL, INTOTLABOR = b.INTOTLABOR," & vbCrLf
sql = sql & "INTOTEXP = b.INTOTEXP, INTOTOH = b.INTOTOH," & vbCrLf
sql = sql & "LOTTOTMATL = (c.LOTTOTMATL * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTLABOR = (c.LOTTOTLABOR * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTEXP = (c.LOTTOTEXP * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTTOTOH = (c.LOTTOTOH * PICKQTY) / c.LOTORIGINALQTY," & vbCrLf
sql = sql & "LOTDATECOSTED = c.LOTDATECOSTED, LOTORGQTY = @LotOrgQty," & vbCrLf
sql = sql & "BMTOTOH = (c.LOTTOTOH * PICKQTY) / @LotOrgQty," & vbCrLf
sql = sql & "LOTSPLITFROMSYS = c.LOTSPLITFROMSYS" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail d, InvaTable b, LohdTable c" & vbCrLf
sql = sql & "WHERE d.INMOPART = b.INMOPART AND d.INPART = b.INPART" & vbCrLf
sql = sql & "AND b.INLOTNUMBER = c.LOTnumber" & vbCrLf
sql = sql & "and c.lotpartref = b.INPART" & vbCrLf
sql = sql & "and b.INMOPART = @MOPart AND b.INMORUN  = @LotRunNo" & vbCrLf
sql = sql & "AND b.INTYPE = 10 AND SORTKEYLEVEL = @level" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curMORun INTO @MOPart, @Part" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curMORun" & vbCrLf
sql = sql & "DEALLOCATE curMORun" & vbCrLf
sql = sql & "SET @level = @level + 1" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "print 'Update 2:'+ cast(getdate() as char(25))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curMOSplit CURSOR  FOR" & vbCrLf
sql = sql & "SELECT DISTINCT LOTNUMBER, LOTSPLITFROMSYS, LOTTOTMATL--, LOTTOTLABOR, LOTTOTEXP, LOTTOTOH" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail WHERE LOTSPLITFROMSYS <> ''" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curMOSplit" & vbCrLf
sql = sql & "FETCH NEXT FROM curMOSplit INTO @MOLotNum, @SplitLot, @LotMatl" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "print 'LotSplit LotNum:' + @SplitLot" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF @LotMatl = 0" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SELECT @LotUSpMat = (LOTTOTMATL / LOTORIGINALQTY)," & vbCrLf
sql = sql & "@LotUSpLabor = (LOTTOTLABOR / LOTORIGINALQTY)," & vbCrLf
sql = sql & "@LotUSpExp = (LOTTOTEXP / LOTORIGINALQTY)," & vbCrLf
sql = sql & "@LotUSpOH = (LOTTOTOH / LOTORIGINALQTY)" & vbCrLf
sql = sql & "FROM Lohdtable WHERE LOTNUMBER = @SplitLot" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET LOTTOTMATL = (@LotUSpMat * PICKQTY)," & vbCrLf
sql = sql & "LOTTOTLABOR = (@LotUSpLabor * PICKQTY)," & vbCrLf
sql = sql & "LOTTOTEXP = (@LotUSpExp * PICKQTY)," & vbCrLf
sql = sql & "LOTTOTOH = (@LotUSpOH * PICKQTY)" & vbCrLf
sql = sql & "WHERE LOTNUMBER = @MOLotNum AND LOTSPLITFROMSYS = @SplitLot" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curMOSplit INTO @MOLotNum, @SplitLot, @LotMatl" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curMOSplit" & vbCrLf
sql = sql & "DEALLOCATE curMOSplit" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT @level =  MAX(SORTKEYLEVEL) FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "WHILE (@level >= 0 )" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curMODet CURSOR  FOR" & vbCrLf
sql = sql & "--SELECT INPART, LOTTOTMATL, LOTTOTLABOR, LOTTOTEXP , LOTTOTOH FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "-- WHERE INPART = '775345149'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT INMOPART," & vbCrLf
sql = sql & "SUM(IsNull(LOTTOTMATL, 0)), SUM(ISNULL(LOTTOTLABOR,0)) ," & vbCrLf
sql = sql & "Sum (IsNull(LOTTOTEXP, 0)) , SUM(IsNull(BMTOTOH, 0))" & vbCrLf
sql = sql & "From" & vbCrLf
sql = sql & "(SELECT DISTINCT INMOPART,INMORUN,INPART,LOTTOTMATL,LOTTOTLABOR," & vbCrLf
sql = sql & "LOTTOTEXP,LOTTOTOH,SUMTOTMAL,SUMTOTLABOR, SUMTOTEXP, SUMTOTOH,BMTOTOH" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail WHERE SORTKEYLEVEL = @level) as foo" & vbCrLf
sql = sql & "group by INMOPART" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curMODet" & vbCrLf
sql = sql & "FETCH NEXT FROM curMODet INTO @MOPart, @SumTotMat, @SumTotLabor,@SumTotExp,@SumTotOH" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "print 'PartNum : ' + @MOPart" & vbCrLf
sql = sql & "print 'SumTotoh : ' + Convert(varchar(24), @SumTotOH)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET SUMTOTMAL = LOTTOTMATL + @SumTotMat," & vbCrLf
sql = sql & "SUMTOTLABOR = LOTTOTLABOR + @SumTotLabor," & vbCrLf
sql = sql & "SUMTOTEXP = LOTTOTEXP + @SumTotExp, SUMTOTOH = (BMTOTOH + @SumTotOH) * @MOQty ," & vbCrLf
sql = sql & "HASCHILD = 1,PARTSUM = 'TOTAL ' + LTRIM(INPART)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "WHERE INPART = @MOPart AND SORTKEYLEVEL = @level - 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curMODet INTO @MOPart, @SumTotMat, @SumTotLabor,@SumTotExp,@SumTotOH" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curMODet" & vbCrLf
sql = sql & "DEALLOCATE curMODet" & vbCrLf
sql = sql & "SET @level = @level - 1" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// update the Lower level cost detail" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET SUMTOTMAL = LOTTOTMATL, SUMTOTLABOR = LOTTOTLABOR," & vbCrLf
sql = sql & "SUMTOTEXP = LOTTOTEXP, SUMTOTOH = BMTOTOH WHERE HASCHILD IS NULL" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET @SumTotMat  = 0" & vbCrLf
sql = sql & "SET @SumTotLabor  = 0" & vbCrLf
sql = sql & "SET @SumTotExp  = 0" & vbCrLf
sql = sql & "SET @SumTotOH  = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--// Udpate the Root total" & vbCrLf
sql = sql & "SELECT @SumTotMat = SUM(SUMTOTMAL), @SumTotLabor = SUM(SUMTOTLABOR)," & vbCrLf
sql = sql & "@SumTotExp = SUM(SUMTOTEXP) ,@SumTotOH = SUM(SUMTOTOH)" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail WHERE SORTKEYLEVEL = 0" & vbCrLf
sql = sql & "AND  RTRIM(INMOPART) <> RTRIM(INPART)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET PARTSUM = INPART" & vbCrLf
sql = sql & "WHERE PARTSUM IS NULL" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "----  SELECT * FROM #tempMOPartsDetail WHERE SORTKEYLEVEL = 0" & vbCrLf
sql = sql & "----AND  RTRIM(INMOPART) = RTRIM(INPART)" & vbCrLf
sql = sql & "--// Reverse the partnumbers." & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @level = 0" & vbCrLf
sql = sql & "SET @RowCount = 1" & vbCrLf
sql = sql & "SET @PrevParent = ''" & vbCrLf
sql = sql & "DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
sql = sql & "SELECT INMOPART, SortKey" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail" & vbCrLf
sql = sql & "WHERE HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLMASTER ='AA10'--GLTOPMaster = 1 AND" & vbCrLf
sql = sql & "ORDER BY SortKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curAcctStruc" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc INTO @ParentPart, @SortKey" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "if (@PrevParent <> @ParentPart)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET" & vbCrLf
sql = sql & "SortKeyRev = RIGHT ('0000'+ Cast(@RowCount as varchar), 4) + char(36)+ @ParentPart" & vbCrLf
sql = sql & "WHERE INMOPART = @ParentPart AND HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLTOPMaster = 1" & vbCrLf
sql = sql & "SET @RowCount = @RowCount + 1" & vbCrLf
sql = sql & "SET @PrevParent = @ParentPart" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curAcctStruc INTO @ParentPart, @SortKey" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "Close curAcctStruc" & vbCrLf
sql = sql & "DEALLOCATE curAcctStruc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE #tempMOPartsDetail SET LOTMOPARTRUNKEY = @MOPartRunKey" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO EsMOPartsCostDetail (LOTMOPARTRUNKEY, INMOPART,INMORUN,INPART,PARTSUM,LOTNUMBER,LOTUSERLOTID," & vbCrLf
sql = sql & "LOTTOTMATL,SUMTOTMAL, LOTTOTLABOR,SUMTOTLABOR, LOTTOTEXP, SUMTOTEXP, LOTTOTOH,SUMTOTOH,BMTOTOH," & vbCrLf
sql = sql & "LOTDATECOSTED, BMQTYREQD, LOTORGQTY, SORTKEYLEVEL,SortKey,SortKeyRev,HASCHILD, PICKQTY)" & vbCrLf
sql = sql & "SELECT LOTMOPARTRUNKEY,INMOPART,INMORUN,INPART,PARTSUM,LOTNUMBER,LOTUSERLOTID," & vbCrLf
sql = sql & "LOTTOTMATL,SUMTOTMAL, LOTTOTLABOR,SUMTOTLABOR, LOTTOTEXP, SUMTOTEXP, LOTTOTOH,SUMTOTOH,BMTOTOH," & vbCrLf
sql = sql & "LOTDATECOSTED, BMQTYREQD, LOTORGQTY, SORTKEYLEVEL,SortKey,SortKeyRev,HASCHILD, PICKQTY" & vbCrLf
sql = sql & "FROM #tempMOPartsDetail--WHERE SORTKEYLEVEL = 1" & vbCrLf
sql = sql & "order by SortKey--SortKeyRev" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DROP table #tempMOPartsDetail" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript False, sql

sql = "DropStoredProcedureIfExists 'RptMRPMOQtyShortage'" & vbCrLf
ExecuteScript False, sql

sql = "create PROCEDURE RptMRPMOQtyShortage" & vbCrLf
sql = sql & "@InMOPart as varchar(30), @StartDate as datetime, @EndDate as datetime" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--05/09/2018 TEL - changed 'with cte' to ';with cte'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @ChildKey as varchar(1024)" & vbCrLf
sql = sql & "declare @GLSortKey as varchar(1024)" & vbCrLf
sql = sql & "declare @SortKey as varchar(1024)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @MOPart as varchar(30)" & vbCrLf
sql = sql & "declare @MORun as Integer" & vbCrLf
sql = sql & "declare @MOQtyRqd as decimal(12,4)" & vbCrLf
sql = sql & "declare @MOPartRqDt as datetime" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @Part as varchar(30)" & vbCrLf
sql = sql & "declare @PAQOH as decimal(12,4)" & vbCrLf
sql = sql & "declare @RunTot as decimal(12,4)" & vbCrLf
sql = sql & "declare @AssyPart as varchar(30)" & vbCrLf
sql = sql & "declare @BMQtyReq as decimal(12,4)" & vbCrLf
sql = sql & "declare @RunQtyReq as decimal (12, 4)" & vbCrLf
sql = sql & "declare @PartDateQrd as datetime" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--DROP TABLE #tempMOPartsDetail" & vbCrLf
sql = sql & "DELETE FROM tempMrplPartShort" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@InMOPart = '')" & vbCrLf
sql = sql & "SET @InMOPart = @InMOPart + '%'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curMrpExp CURSOR  FOR" & vbCrLf
sql = sql & "SELECT MRP_PARTREF,0 as RUNNO, MRP_PARTQTYRQD, MRP_ACTIONDATE" & vbCrLf
sql = sql & "FROM MrplTable, PartTable" & vbCrLf
sql = sql & "WHERE MRP_PARTREF = PartRef" & vbCrLf
sql = sql & "AND MrplTable.MRP_PARTREF LIKE @InMOPart" & vbCrLf
sql = sql & "AND MrplTable.MRP_PARTPRODCODE LIKE '%'" & vbCrLf
sql = sql & "AND MrplTable.MRP_PARTCLASS LIKE '%'" & vbCrLf
sql = sql & "AND MrplTable.MRP_POBUYER LIKE '%'" & vbCrLf
sql = sql & "AND MrplTable.MRP_PARTDATERQD BETWEEN @StartDate AND @EndDate" & vbCrLf
sql = sql & "AND MrplTable.MRP_TYPE IN (6, 5)" & vbCrLf
sql = sql & "AND PartTable.PAMAKEBUY ='M'" & vbCrLf
sql = sql & "UNION" & vbCrLf
sql = sql & "SELECT DISTINCT RUNREF, RUNNO, RUNQTY,runpkstart  as MRP_ACTIONDATE FROM RunsTable WHERE" & vbCrLf
sql = sql & "RUNREF LIKE @InMOPart AND RUNSTATUS = 'SC'" & vbCrLf
sql = sql & "AND RUNPKSTART BETWEEN @StartDate  AND @EndDate + ' 23:00' order by MRP_ACTIONDATE" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curMrpExp" & vbCrLf
sql = sql & "FETCH NEXT FROM curMrpExp INTO @MOPart, @MORun, @MOQtyRqd, @MOPartRqDt" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "--print 'MO:' + @MOPart + '; RUN:' + Convert(varchar(10), @MORun) + '; Date:' + Convert(varchar(24), @MOPartRqDt, 101);" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & ";with cte" & vbCrLf
sql = sql & "as (select BMASSYPART,BMPARTREF,BMPARTREV, BMQTYREQD , RTrim(BMUNITS) BMUNITS," & vbCrLf
sql = sql & "BMCONVERSION, BMSEQUENCE, 0 as level," & vbCrLf
sql = sql & "cast(LTRIM(RTrim(BMASSYPART)) + char(36)+ COALESCE(cast(BMSEQUENCE as varchar(4)), '') + LTRIM(RTrim(BMPARTREF)) as varchar(max)) as SortKey" & vbCrLf
sql = sql & "from BmplTable" & vbCrLf
sql = sql & "where BMASSYPART = @MOPart" & vbCrLf
sql = sql & "union all" & vbCrLf
sql = sql & "select a.BMASSYPART,a.BMPARTREF,a.BMPARTREV, a.BMQTYREQD , RTrim(a.BMUNITS) BMUNITS," & vbCrLf
sql = sql & "a.BMCONVERSION, a.BMSEQUENCE, level + 1," & vbCrLf
sql = sql & "cast(COALESCE(SortKey,'') + char(36) + COALESCE(cast(a.BMSEQUENCE as varchar(4)), '') + COALESCE(LTRIM(RTrim(a.BMPARTREF)) ,'') as varchar(max))as SortKey" & vbCrLf
sql = sql & "from cte" & vbCrLf
sql = sql & "inner join BmplTable a" & vbCrLf
sql = sql & "on cte.BMPARTREF = a.BMASSYPART" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "INSERT INTO tempMrplPartShort(BMASSYPART,BMPARTREF,BMQTYREQD," & vbCrLf
sql = sql & "SORTKEYLEVEL,BMSEQUENCE, SortKey, PAQOH, RUNNO,MRP_PARTQTYRQD, MRP_ACTIONDATE)" & vbCrLf
sql = sql & "select BMASSYPART, BMPARTREF,BMQTYREQD,level,BMSEQUENCE, SortKey, PAQOH, @MORun, @MOQtyRqd, @MOPartRqDt" & vbCrLf
sql = sql & "from cte, PartTable WHERE PARTREF = BMPARTREF  AND BMPARTREF <> 'NULL' order by SortKey,BMSEQUENCE" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "FETCH NEXT FROM curMrpExp INTO @MOPart, @MORun, @MOQtyRqd, @MOPartRqDt" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curMrpExp" & vbCrLf
sql = sql & "DEALLOCATE curMrpExp" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curRunTot CURSOR  FOR" & vbCrLf
sql = sql & "select DISTINCT BMPARTREF, PAQOH from tempMrplPartShort order by BMPARTREF" & vbCrLf
sql = sql & "OPEN curRunTot" & vbCrLf
sql = sql & "FETCH NEXT FROM curRunTot INTO @Part, @PAQOH" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @RunTot = 0.0000" & vbCrLf
sql = sql & "SET @RunTot = @PAQOH" & vbCrLf
sql = sql & "DECLARE curRunTot1 CURSOR  FOR" & vbCrLf
sql = sql & "select DISTINCT BMASSYPART, BMQTYREQD, MRP_PARTQTYRQD, MRP_ACTIONDATE from tempMrplPartShort" & vbCrLf
sql = sql & "WHERE BMPARTREF = @Part AND sortkeylevel = 0" & vbCrLf
sql = sql & "order by MRP_ACTIONDATE  -- BMASSYPART," & vbCrLf
sql = sql & "OPEN curRunTot1" & vbCrLf
sql = sql & "FETCH NEXT FROM curRunTot1 INTO @AssyPart, @BMQtyReq, @RunQtyReq, @PartDateQrd" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "--Set @RunTot = ROUND(@RunTot,4)" & vbCrLf
sql = sql & "Set @RunTot = @RunTot -  ( @BMQtyReq * @RunQtyReq)" & vbCrLf
sql = sql & "UPDATE tempMrplPartShort SET PAQRUNTOT = @RunTot WHERE" & vbCrLf
sql = sql & "BMASSYPART = @AssyPart AND BMPARTREF = @Part AND MRP_ACTIONDATE = @PartDateQrd" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "FETCH NEXT FROM curRunTot1 INTO @AssyPart, @BMQtyReq, @RunQtyReq,@PartDateQrd" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curRunTot1" & vbCrLf
sql = sql & "DEALLOCATE curRunTot1" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "FETCH NEXT FROM curRunTot INTO @Part, @PAQOH" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "Close curRunTot" & vbCrLf
sql = sql & "DEALLOCATE curRunTot" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript False, sql


sql = "DropStoredProcedureIfExists 'RptTopIncomeStatement'" & vbCrLf
ExecuteScript False, sql

sql = "CREATE PROCEDURE [dbo].[RptTopIncomeStatement]" & vbCrLf
sql = sql & "   @StartDate as varchar(12),@EndDate as varchar(12)," & vbCrLf
sql = sql & "   @YearBeginDate as varchar(12), @InclIncAcct as varchar(1)" & vbCrLf
sql = sql & "AS " & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   declare @glAcctRef as varchar(10)" & vbCrLf
sql = sql & "   declare @glMsAcct as varchar(10)" & vbCrLf
sql = sql & "   declare @SumCurBal decimal(15,4)" & vbCrLf
sql = sql & "   declare @SumYTD decimal(15,4)" & vbCrLf
sql = sql & "   declare @SumPrevBal as decimal(15,4)" & vbCrLf
sql = sql & "   declare @level as integer" & vbCrLf
sql = sql & "   declare @TopLevelDesc as varchar(30)" & vbCrLf
sql = sql & "   declare @InclInAcct as Integer" & vbCrLf
sql = sql & "   declare @TopLevAcct as varchar(20)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   declare @PrevMaster as varchar(10)" & vbCrLf
sql = sql & "   declare @RowCount as integer" & vbCrLf
sql = sql & "   declare @GlMasterAcc as varchar(10)" & vbCrLf
sql = sql & "   declare @GlChildAcct as varchar(10)" & vbCrLf
sql = sql & "   declare @ChildKey as varchar(1024)" & vbCrLf
sql = sql & "   declare @GLSortKey as varchar(1024)" & vbCrLf
sql = sql & "   declare @SortKey as varchar(1024)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   DELETE FROM EsReportIncStatement" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   if (@InclIncAcct = '1')" & vbCrLf
sql = sql & "      SET @InclInAcct = ''" & vbCrLf
sql = sql & "   Else" & vbCrLf
sql = sql & "      SET @InclInAcct = '0'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   DECLARE balAcctStruc CURSOR  FOR" & vbCrLf
sql = sql & "      SELECT '4', COINCMACCT, COINCMDESC FROM GlmsTopTable" & vbCrLf
sql = sql & "      Union All" & vbCrLf
sql = sql & "      SELECT '5', COCOGSACCT, COCOGSDESC FROM GlmsTopTable" & vbCrLf
sql = sql & "      Union All" & vbCrLf
sql = sql & "      SELECT '6', COEXPNACCT, COEXPNDESC FROM GlmsTopTable" & vbCrLf
sql = sql & "      Union All" & vbCrLf
sql = sql & "      SELECT '7', COOINCACCT, COOINCDESC FROM GlmsTopTable" & vbCrLf
sql = sql & "      Union All" & vbCrLf
sql = sql & "      SELECT '8', COOEXPACCT, COOEXPDESC FROM GlmsTopTable" & vbCrLf
sql = sql & "      Union All" & vbCrLf
sql = sql & "      SELECT '9', COFDTXACCT, COFDTXDESC FROM GlmsTopTable" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   OPEN balAcctStruc" & vbCrLf
sql = sql & "   FETCH NEXT FROM balAcctStruc INTO @level,@TopLevAcct, @TopLevelDesc" & vbCrLf
sql = sql & "   WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "   BEGIN" & vbCrLf
sql = sql & "      IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "      BEGIN" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "        ;with cte" & vbCrLf
sql = sql & "        as " & vbCrLf
sql = sql & "        (select GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL, 1 as level," & vbCrLf
sql = sql & "          cast(cast(@level as varchar(4))+ char(36)+ GLACCTREF as varchar(max)) as SortKey" & vbCrLf
sql = sql & "        From GlacTopTable" & vbCrLf
sql = sql & "        where GLMASTER = cast(@TopLevAcct as varchar(20)) AND GLINACTIVE LIKE @InclInAcct" & vbCrLf
sql = sql & "        Union All" & vbCrLf
sql = sql & "        select a.GLACCTREF, a.GLDESCR, a.GLMASTER, a.GLFSLEVEL, level + 1," & vbCrLf
sql = sql & "         cast(COALESCE(SortKey,'') + char(36) + COALESCE(a.GLACCTREF,'') as varchar(max))as SortKey" & vbCrLf
sql = sql & "        From cte" & vbCrLf
sql = sql & "          inner join GlacTopTable a" & vbCrLf
sql = sql & "            on cte.GLACCTREF = a.GLMASTER" & vbCrLf
sql = sql & "          WHERE GLINACTIVE LIKE @InclInAcct" & vbCrLf
sql = sql & "        )" & vbCrLf
sql = sql & "        INSERT INTO EsReportIncStatement(GLTOPMaster, GLTOPMASTERDESC, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL,SORTKEYLEVEL,GLACCSORTKEY)" & vbCrLf
sql = sql & "        select @level, @TopLevelDesc, GLACCTREF, GLDESCR, GLMASTER, GLFSLEVEL, level, SortKey" & vbCrLf
sql = sql & "        from cte order by SortKey" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "      End" & vbCrLf
sql = sql & "      FETCH NEXT FROM balAcctStruc INTO @level,@TopLevAcct, @TopLevelDesc" & vbCrLf
sql = sql & "   End" & vbCrLf
sql = sql & "   Close balAcctStruc" & vbCrLf
sql = sql & "   DEALLOCATE balAcctStruc" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   UPDATE EsReportIncStatement SET CurrentBal = foo.Balance--, SUMCURBAL = foo.Balance" & vbCrLf
sql = sql & "   From" & vbCrLf
sql = sql & "       (SELECT SUM(GjitTopTable.JICRD) - SUM(GjitTopTable.JIDEB) as Balance, JIACCOUNT" & vbCrLf
sql = sql & "         FROM GjhdTable INNER JOIN GjitTopTable ON GJNAME = JINAME" & vbCrLf
sql = sql & "      WHERE GJPOST BETWEEN @StartDate AND @EndDate" & vbCrLf
sql = sql & "         AND GjhdTable.GJPOSTED = 1" & vbCrLf
sql = sql & "      GROUP BY JIACCOUNT) as foo" & vbCrLf
sql = sql & "   Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   UPDATE EsReportIncStatement SET YTD = foo.Balance--, SUMYTD = foo.Balance" & vbCrLf
sql = sql & "   From" & vbCrLf
sql = sql & "      (SELECT JIACCOUNT,SUM(GjitTopTable.JICRD) - SUM(GjitTopTable.JIDEB) AS Balance" & vbCrLf
sql = sql & "         FROM GjhdTable INNER JOIN GjitTopTable ON GJNAME = JINAME" & vbCrLf
sql = sql & "      WHERE (GJPOST BETWEEN @YearBeginDate AND @EndDate)" & vbCrLf
sql = sql & "         AND GjhdTable.GJPOSTED = 1" & vbCrLf
sql = sql & "      GROUP BY JIACCOUNT) as foo" & vbCrLf
sql = sql & "   Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   UPDATE EsReportIncStatement SET PreviousBal = foo.Balance--, SUMPREVBAL = foo.Balance" & vbCrLf
sql = sql & "   From" & vbCrLf
sql = sql & "      (SELECT JIACCOUNT,SUM(GjitTopTable.JICRD) - SUM(GjitTopTable.JIDEB) AS Balance" & vbCrLf
sql = sql & "         FROM GjhdTable INNER JOIN GjitTopTable ON GJNAME = JINAME" & vbCrLf
sql = sql & "      WHERE (GJPOST BETWEEN DATEADD(year, -1, @YearBeginDate) AND DATEADD(year, -1, @EndDate))" & vbCrLf
sql = sql & "         AND GjhdTable.GJPOSTED = 1" & vbCrLf
sql = sql & "      GROUP BY JIACCOUNT) as foo" & vbCrLf
sql = sql & "   Where foo.JIACCOUNT = GLACCTREF" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   SELECT @level =  MAX(SORTKEYLEVEL) FROM EsReportIncStatement" & vbCrLf
sql = sql & "   --set @level = 9" & vbCrLf
sql = sql & "   WHILE (@level >= 1 )" & vbCrLf
sql = sql & "   BEGIN" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "      DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
sql = sql & "      SELECT GLMASTER, SUM(ISNULL(SUMCURBAL,0) + (ISNULL(CurrentBal,0))) ," & vbCrLf
sql = sql & "         Sum (IsNull(SUMYTD, 0) + (IsNull(YTD, 0))), Sum(IsNull(SUMPREVBAL, 0) + (IsNull(PreviousBal, 0)))" & vbCrLf
sql = sql & "      From" & vbCrLf
sql = sql & "         (SELECT DISTINCT GLACCTREF, GLDESCR, GLMASTER," & vbCrLf
sql = sql & "         CurrentBal , YTD, PreviousBal, SUMCURBAL, SUMYTD, SUMPREVBAL" & vbCrLf
sql = sql & "         FROM EsReportIncStatement WHERE SORTKEYLEVEL = @level) as foo" & vbCrLf
sql = sql & "      group by GLMASTER" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "      OPEN curAcctStruc" & vbCrLf
sql = sql & "      FETCH NEXT FROM curAcctStruc INTO @glMsAcct, @SumCurBal, @SumYTD, @SumPrevBal" & vbCrLf
sql = sql & "      WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "      BEGIN" & vbCrLf
sql = sql & "         IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "         BEGIN" & vbCrLf
sql = sql & "            UPDATE EsReportIncStatement SET SUMCURBAL = @SumCurBal, SUMYTD = @SumYTD," & vbCrLf
sql = sql & "               SUMPREVBAL = @SumPrevBal, GLDESCR = 'TOTAL ' + LTRIM(GLDESCR)," & vbCrLf
sql = sql & "            HASCHILD = 1" & vbCrLf
sql = sql & "            WHERE GLACCTREF = @glMsAcct" & vbCrLf
sql = sql & "         End" & vbCrLf
sql = sql & "         FETCH NEXT FROM curAcctStruc INTO @glMsAcct, @SumCurBal, @SumYTD, @SumPrevBal" & vbCrLf
sql = sql & "      End" & vbCrLf
sql = sql & "      Close curAcctStruc" & vbCrLf
sql = sql & "      DEALLOCATE curAcctStruc" & vbCrLf
sql = sql & "      SET @level = @level - 1" & vbCrLf
sql = sql & "   End" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "   UPDATE EsReportIncStatement SET SUMCURBAL = CurrentBal WHERE SUMCURBAL IS NULL" & vbCrLf
sql = sql & "   UPDATE EsReportIncStatement SET SUMPREVBAL = PreviousBal WHERE SUMPREVBAL IS NULL" & vbCrLf
sql = sql & "   UPDATE EsReportIncStatement SET SUMYTD = YTD  WHERE SUMYTD IS NULL" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   set @level = 0 " & vbCrLf
sql = sql & "   set @RowCount = 1" & vbCrLf
sql = sql & "   SET @PrevMaster = ''" & vbCrLf
sql = sql & "   DECLARE curAcctStruc CURSOR  FOR" & vbCrLf
sql = sql & "      SELECT GLMASTER, GLACCSORTKEY" & vbCrLf
sql = sql & "      FROM EsReportIncStatement " & vbCrLf
sql = sql & "         WHERE HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLMASTER ='AA10'--GLTOPMaster = 1 AND " & vbCrLf
sql = sql & "      ORDER BY GLACCSORTKEY" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   OPEN curAcctStruc" & vbCrLf
sql = sql & "   FETCH NEXT FROM curAcctStruc INTO @GlMasterAcc, @GLSortKey" & vbCrLf
sql = sql & "   WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "   BEGIN" & vbCrLf
sql = sql & "    IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "    BEGIN" & vbCrLf
sql = sql & "      if (@PrevMaster <> @GlMasterAcc)" & vbCrLf
sql = sql & "      BEGIN" & vbCrLf
sql = sql & "         UPDATE EsReportIncStatement SET " & vbCrLf
sql = sql & "            SortKeyRev = RIGHT ('0000'+ Cast(@RowCount as varchar), 4) + char(36)+ @GlMasterAcc" & vbCrLf
sql = sql & "         WHERE GLMASTER = @GlMasterAcc AND HASCHILD IS NULL --GLFSLEVEL = 8 --AND GLTOPMaster = 1 " & vbCrLf
sql = sql & "         SET @RowCount = @RowCount + 1" & vbCrLf
sql = sql & "         SET @PrevMaster = @GlMasterAcc" & vbCrLf
sql = sql & "      END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    End" & vbCrLf
sql = sql & "    FETCH NEXT FROM curAcctStruc INTO @GlMasterAcc, @GLSortKey" & vbCrLf
sql = sql & "   End" & vbCrLf
sql = sql & "       " & vbCrLf
sql = sql & "   Close curAcctStruc" & vbCrLf
sql = sql & "   DEALLOCATE curAcctStruc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "    SELECT @level = MAX(SORTKEYLEVEL) FROM EsReportIncStatement" & vbCrLf
sql = sql & "   --set @level = 7" & vbCrLf
sql = sql & "   WHILE (@level >= 1 )" & vbCrLf
sql = sql & "   BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "      SET @PrevMaster = ''" & vbCrLf
sql = sql & "        DECLARE curAcctStruc1 CURSOR  FOR" & vbCrLf
sql = sql & "         SELECT DISTINCT GLACCTREF, GLMASTER, GLACCSORTKEY" & vbCrLf
sql = sql & "         FROM EsReportIncStatement " & vbCrLf
sql = sql & "            WHERE SORTKEYLEVEL = @level AND HASCHILD IS NOT NULL--GLTOPMaster = 1 AND " & vbCrLf
sql = sql & "         order by GLACCSORTKEY" & vbCrLf
sql = sql & "    " & vbCrLf
sql = sql & "        OPEN curAcctStruc1" & vbCrLf
sql = sql & "        FETCH NEXT FROM curAcctStruc1 INTO @GlChildAcct, @GlMasterAcc, @GLSortKey" & vbCrLf
sql = sql & "        WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "        BEGIN" & vbCrLf
sql = sql & "          IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "          BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "            if (@PrevMaster <> @GlChildAcct)" & vbCrLf
sql = sql & "            BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "   print 'Record' + @GlChildAcct + ':' + @GlMasterAcc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "               SELECT TOP 1 @ChildKey = SortKeyRev,@SortKey = GLACCSORTKEY" & vbCrLf
sql = sql & "               FROM EsReportIncStatement " & vbCrLf
sql = sql & "                  WHERE SORTKEYLEVEL > @level AND GLMASTER = @GlChildAcct --GLTOPMaster = 1 AND " & vbCrLf
sql = sql & "               order by GLACCSORTKEY desc" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "               UPDATE EsReportIncStatement SET " & vbCrLf
sql = sql & "                  SortKeyRev = Cast(@ChildKey as varchar(512)) + char(36)+ @GlMasterAcc" & vbCrLf
sql = sql & "               WHERE GLACCTREF = @GlChildAcct AND GLMASTER = @GlMasterAcc " & vbCrLf
sql = sql & "                  AND SORTKEYLEVEL = @level --GLTOPMaster = 1 AND " & vbCrLf
sql = sql & "               SET @PrevMaster = @GlChildAcct" & vbCrLf
sql = sql & "            END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "          End" & vbCrLf
sql = sql & "          FETCH NEXT FROM curAcctStruc1 INTO @GlChildAcct, @GlMasterAcc, @GLSortKey" & vbCrLf
sql = sql & "        End" & vbCrLf
sql = sql & "               " & vbCrLf
sql = sql & "        Close curAcctStruc1" & vbCrLf
sql = sql & "        DEALLOCATE curAcctStruc1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "        SET @level = @level - 1" & vbCrLf
sql = sql & "   End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "  SELECT GLTOPMaster, GLTOPMASTERDESC, GLACCTREF, GLACCTNO, GLDESCR, GLMASTER, GLTYPE,GLINACTIVE, GLFSLEVEL," & vbCrLf
sql = sql & "      SUMCURBAL , CurrentBal, SUMYTD, YTD, SUMPREVBAL, PreviousBal, SORTKEYLEVEL, GLACCSORTKEY" & vbCrLf
sql = sql & "   FROM EsReportIncStatement ORDER BY SortKeyRev --GLTOPMASTER, GLACCSORTKEY desc, SortKeyLevel" & vbCrLf
sql = sql & " " & vbCrLf
sql = sql & "End"
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase103()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 178     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "DropStoredProcedureIfExists 'RptCashReceipts'" & vbCrLf
ExecuteScript False, sql
sql = "create procedure RptCashReceipts" & vbCrLf
sql = sql & "@CustNickName varchar(10),      -- blank for all" & vbCrLf
sql = sql & "@StartDate as varchar(10),      -- 'mm/dd/yyyy'- blank for all" & vbCrLf
sql = sql & "@ReceiptAmount as varchar(10),  -- blank for all" & vbCrLf
sql = sql & "@CheckNumber as varchar(20), -- blank for all" & vbCrLf
sql = sql & "@ShowUninvoicedItems int     -- = 0 to not show, 1 to show" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* for View Cash Receipts Report" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "exec RptCashReceipts 'BOECO','12/1/2017','','<All>',0" & vbCrLf
sql = sql & "exec RptCashReceipts 'BOECO','12/1/2017','5,375.01','<All>',0" & vbCrLf
sql = sql & "exec RptCashReceipts 'BOEWIN','12/1/2017','','',0" & vbCrLf
sql = sql & "exec RptCashReceipts 'hexstr','12/1/2017','','',0" & vbCrLf
sql = sql & "exec RptCashReceipts 'MISC','12/1/2017','','',1" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if exists (select * from INFORMATION_SCHEMA.tables where table_name = 'RptViewCR')" & vbCrLf
sql = sql & "drop table RptViewCR" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get accounts from ComnTable" & vbCrLf
sql = sql & "declare @CashAccount varchar(12), @ArAccount varchar(12),@TransFeeAccount varchar(12)" & vbCrLf
sql = sql & "select @CashAccount = COCRCASHACCT, @ArAccount = COSJARACCT, @TransFeeAccount = COTRANSFEEACCT from ComnTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @sql varchar(1000)" & vbCrLf
sql = sql & "set @sql =" & vbCrLf
sql = sql & "'SELECT" & vbCrLf
sql = sql & "cash.CACUST, cash.CACHECKNO, cash.CACDATE, cash.CAINVNO, cash.CARCDATE, cash.CACRAMT, cash.CACKAMT, cash.CAENTRY," & vbCrLf
sql = sql & "jr.DCDEBIT, jr.DCCREDIT, jr.DCHEAD, jr.DCTRAN, jr.DCREF," & vbCrLf
sql = sql & "inv.INVDATE, inv.INVPRE, inv.INVTYPE," & vbCrLf
sql = sql & "jr.DCACCTNO," & vbCrLf
sql = sql & "acct.GLDESCR" & vbCrLf
sql = sql & "into RptViewCR" & vbCrLf
sql = sql & "FROM CashTable cash" & vbCrLf
sql = sql & "LEFT OUTER JOIN JritTable jr ON cash.CACUST=jr.DCCUST AND cash.CACHECKNO=jr.DCCHECKNO AND cash.CAINVNO = jr.DCINVNO" & vbCrLf
sql = sql & "INNER JOIN GlacTable acct ON jr.DCACCTNO=acct.GLACCTREF" & vbCrLf
sql = sql & "LEFT OUTER JOIN CihdTable inv ON cash.CAINVNO=inv.INVNO" & vbCrLf
sql = sql & "INNER JOIN CustTable cust ON cash.CACUST=cust.CUREF" & vbCrLf
sql = sql & "'" & vbCrLf
sql = sql & "declare @where varchar(500)" & vbCrLf
sql = sql & "--set @where = 'WHERE DCACCTNO not in (''' + @CashAccount + ''',''' + @ArAccount + ''')" & vbCrLf
sql = sql & "--'" & vbCrLf
sql = sql & "set @where = ''" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @CustNickName <> ''" & vbCrLf
sql = sql & "set @where = @where + 'and CACUST = ''' + @CustNickName + '''" & vbCrLf
sql = sql & "'" & vbCrLf
sql = sql & "if @StartDate <> ''" & vbCrLf
sql = sql & "set @where = @where + 'and CARCDATE >= ''' + @StartDate + '''" & vbCrLf
sql = sql & "'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @ReceiptAmount <> ''" & vbCrLf
sql = sql & "set @where = @where + 'and CACKAMT = ' + replace(@ReceiptAmount,',','') + '" & vbCrLf
sql = sql & "'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @Checknumber <> '' and @Checknumber <> '<All>'" & vbCrLf
sql = sql & "set @where = @where + 'and CACHECKNO = ''' + @Checknumber + '''" & vbCrLf
sql = sql & "'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @ShowUninvoicedItems = 0" & vbCrLf
sql = sql & "set @where = @where + 'and INVDATE is not null" & vbCrLf
sql = sql & "'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @where <> ''" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "set @where = stuff(@where,1,3,'where')" & vbCrLf
sql = sql & "set @sql = @sql + @where" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @sql = @sql + 'order by cash.CACHECKNO,cash.CAINVNO,DCREF'" & vbCrLf
sql = sql & "print @sql" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec (@sql)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- construct a single line for each check/invoice combination" & vbCrLf
sql = sql & "select CACUST as Customer, CACHECKNO as [Check #], convert(varchar(8),CACDATE,1) AS [Check Date],CACKAMT as [Check Amt]," & vbCrLf
sql = sql & "convert(varchar(8),CARCDATE,1) as [Receipt Date]," & vbCrLf
sql = sql & "CACRAMT as [Receipt Amt],CAENTRY as [Entered by]," & vbCrLf
sql = sql & "ISNULL(INVTYPE + ' ' + INVPRE + cast(CAINVNO as varchar(10)),'NONE') as [Invoice #],CAINVNO," & vbCrLf
sql = sql & "isnull(convert(varchar(8),INVDATE,1),'') AS [Inv Date],  DCDEBIT as [Cash], cast(0.00 as decimal(12,2)) as Adj," & vbCrLf
sql = sql & "cast(0.00 as decimal(12,2)) as [Total Applied],cast('' as varchar(50)) as [Adjustment to]" & vbCrLf
sql = sql & "into #temp from RptViewCR" & vbCrLf
sql = sql & "where DCACCTNO = @CashAccount" & vbCrLf
sql = sql & "order by CACUST, CARCDATE, CAINVNO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update tmp" & vbCrLf
sql = sql & "set Adj = DCDEBIT, [Adjustment to] = GLDESCR" & vbCrLf
sql = sql & "from #temp tmp join RptViewCr cr on cr.CACUST = tmp.Customer and cr.CACHECKNO = tmp.[Check #]" & vbCrLf
sql = sql & "and cr.CAINVNO = tmp.CAINVNO and DCDEBIT <> 0 AND DCACCTNO <> @CashAccount" & vbCrLf
sql = sql & "drop table RptViewCR" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #temp set [Total Applied] = [Cash] + Adj" & vbCrLf
sql = sql & "select * from #temp order by Customer, [Receipt Date], [Check #], CAINVNO" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
ExecuteScript False, sql



''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase104()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 179     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "AddOrUpdateColumn 'PshdTable', 'PSCOMMENTS', 'varchar(2040)'"
ExecuteScript False, sql

''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase105()

   Dim sql As String
   sql = ""

   newver = 180     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "exec DropStoredProcedureIfExists 'DropTempTableIfExists'" & vbCrLf
ExecuteScript False, sql

sql = "create procedure DropTempTableIfExists" & vbCrLf
sql = sql & "@Table_Name varchar(50)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "IF OBJECT_ID('tempdb.dbo.' + @Table_Name) IS NOT NULL" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "declare @sql varchar(100)" & vbCrLf
sql = sql & "set @sql = 'DROP TABLE ' + @Table_Name" & vbCrLf
sql = sql & "execute(@sql)" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript False, sql

' delete millions of POM error messages
' do it 3 times because for some reason, one delete does not get them all
sql = "delete from SystemEvents where Event_SQL = 'select getdate() AS ServerTime'" & vbCrLf
ExecuteScript False, sql
sql = "delete from SystemEvents where Event_SQL = 'select getdate() AS ServerTime'" & vbCrLf
ExecuteScript False, sql
sql = "delete from SystemEvents where Event_SQL = 'select getdate() AS ServerTime'" & vbCrLf
ExecuteScript False, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase106()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 184     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

' IMAINC Auto Release
' 097% auto-release performance improvement
sql = "CREATE NONCLUSTERED INDEX [IX_TempMrplPartShort]" & vbCrLf
sql = sql & "ON [dbo].[tempMrplPartShort] ([BMPARTREF],[SORTKEYLEVEL])" & vbCrLf
ExecuteScript True, sql

sql = "exec DropStoredProcedureIfExists 'GetScMOs'" & vbCrLf
ExecuteScript True, sql

' for IMAINC only
' Get a list of all SC status MOs that can be auto-released
sql = "create procedure GetScMOs" & vbCrLf
sql = sql & "@Parts varchar(30),    -- leading characters for MO parts to select" & vbCrLf
sql = sql & "@StartDate date,    -- start pick date" & vbCrLf
sql = sql & "@EndDate date       -- end pick date" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* test" & vbCrLf
sql = sql & "exec GetScMOs '', '7/1/18','7/1/18'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "set nocount on" & vbCrLf
sql = sql & "declare @EndDatePlus1 date = dateadd(day,1,@EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get list of all SC runs" & vbCrLf
sql = sql & "SELECT DISTINCT RUNREF as MRP_PARTREF, RUNNO, ISNULL(RUNOPCUR, 0) as RUNOPCUR, PARTNUM as MRP_PARTNUM," & vbCrLf
sql = sql & "RUNQTY As MRP_PARTQTYRQD, Convert(varchar(10), RUNPKSTART, 101) As MRP_PARTDATERQD," & vbCrLf
sql = sql & "Convert(varchar(10), RUNSCHED, 101) As MRP_ACTIONDATE, RUNSTATUS,RUNPKSTART,PABOMREV,op.OPCENTER," & vbCrLf
sql = sql & "ROW_NUMBER() over (ORDER BY RUNPKSTART) AS [MO#]" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "FROM RunsTable r" & vbCrLf
sql = sql & "join PartTable p on PARTREF = RUNREF" & vbCrLf
sql = sql & "join RnopTable op on op.opref = r.RUNREF and op.OPRUN = r.RUNNO and op.OPNO = r.RUNOPCUR" & vbCrLf
sql = sql & "WHERE RUNREF LIKE @Parts + '%' AND RUNSTATUS = 'SC' and op.OPCENTER = '0120'" & vbCrLf
sql = sql & "AND RUNPKSTART >= @StartDate AND RUNPKSTART < @EndDatePlus1" & vbCrLf
sql = sql & "order by RUNPKSTART" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all pick list requirements" & vbCrLf
sql = sql & "select t.[MO#], BMPARTREF, BMQTYREQD,MRP_PARTQTYRQD," & vbCrLf
sql = sql & "cast(BMQTYREQD * MRP_PARTQTYRQD as decimal(15,4)) as Qty," & vbCrLf
sql = sql & "CAST(-1 AS DECIMAL(15,4)) as Surplus" & vbCrLf
sql = sql & "into #picks" & vbCrLf
sql = sql & "from #temp t" & vbCrLf
sql = sql & "join BmplTable bm on bm.BMASSYPART = t.MRP_PARTREF and bm.BMREV = t.PABOMREV" & vbCrLf
sql = sql & "order by MO#, BMPARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all part quantities on hand less open pick list quantities" & vbCrLf
sql = sql & "select MRP_PARTREF,sum(MRP_PARTQTYRQD) as AVAIL, sum(MRP_PARTQTYRQD) as OrigAvail" & vbCrLf
sql = sql & "into #parts" & vbCrLf
sql = sql & "from" & vbCrLf
sql = sql & "PartTable pt left join MrplTable mrp on mrp.MRP_PARTREF = pt.PARTREF" & vbCrLf
sql = sql & "where MRP_TYPE in (1,12)" & vbCrLf
sql = sql & "group by MRP_PARTREF" & vbCrLf
sql = sql & "order by MRP_PARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert parts not in mpr" & vbCrLf
sql = sql & "insert #parts" & vbCrLf
sql = sql & "select PARTREF,PAQOH, PAQOH" & vbCrLf
sql = sql & "from PartTable pt" & vbCrLf
sql = sql & "left join #parts p on p.MRP_PARTREF = pt.PARTREF" & vbCrLf
sql = sql & "where PALEVEL <= 4 and p.MRP_PARTREF is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- for each MO, determine if pick list quantities are available" & vbCrLf
sql = sql & "declare @MONO int" & vbCrLf
sql = sql & "DECLARE cur CURSOR FOR" & vbCrLf
sql = sql & "SELECT [MO#]" & vbCrLf
sql = sql & "FROM #temp" & vbCrLf
sql = sql & "order by [MO#]" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN cur" & vbCrLf
sql = sql & "FETCH NEXT FROM cur INTO @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "WHILE @@FETCH_STATUS = 0" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "update p set Surplus = pt.AVAIL - p.Qty" & vbCrLf
sql = sql & "from #picks p join #parts pt on p.BMPARTREF = pt.MRP_PARTREF" & vbCrLf
sql = sql & "where p.MO# = @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if any negative quantities for pick list, delete MO" & vbCrLf
sql = sql & "if(select min(Surplus) from #picks where MO# = @MONO) < 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "delete from #temp where MO# = @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- otherwise subtract from quantities available" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "update p set AVAIL = Avail - pk.Qty" & vbCrLf
sql = sql & "from #parts p join #picks pk on pk.BMPARTREF = p.MRP_PARTREF" & vbCrLf
sql = sql & "where pk.MO# = @MONO" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "FETCH NEXT FROM cur INTO @MONO" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CLOSE cur" & vbCrLf
sql = sql & "DEALLOCATE cur" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from #temp order by MO#" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
sql = sql & "drop table #picks" & vbCrLf
sql = sql & "drop table #parts" & vbCrLf
ExecuteScript True, sql

''''''''''''''''''''''''''''''''''''''''''''''

sql = "exec DropStoredProcedureIfExists 'RptPriorityDispatch'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptPriorityDispatch" & vbCrLf
sql = sql & "@Shop varchar(20),  -- show WCs for this shop" & vbCrLf
sql = sql & "@StartDate date, -- starting OPSCHEDDATE to include" & vbCrLf
sql = sql & "@EndDate date    -- ending OPSCHEDDATE to include" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* EBM Priority Dispatch Report 10/29/2018 TEL" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptPriorityDispatch 'IS', '1/1/2018', '9/30/2018'" & vbCrLf
sql = sql & "exec RptPriorityDispatch 'OS', '1/1/2018', '9/30/2018'" & vbCrLf
sql = sql & "exec RptPriorityDispatch '', '1/1/2018', '9/30/2018'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @EndPlus1 date = dateadd(day,1,@EndDate)  -- end date plus 1 to include all times on end datetime" & vbCrLf
sql = sql & "set @Shop = rtrim(@Shop)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select wc.WCNSHOP, wc.WCNREF, wc.WCNDESC, pt.PARTNUM,pt.PADESC,run.RUNNO, op.OPNO," & vbCrLf
sql = sql & "CONVERT(VARCHAR(10), OPSUDATE, 101) + ' ' + CONVERT(VARCHAR(5), OPSUDATE, 108) as [Start]," & vbCrLf
sql = sql & "CONVERT(VARCHAR(10), OPMDATE, 101) + ' ' + CONVERT(VARCHAR(5), OPMDATE, 108) as [End]," & vbCrLf
sql = sql & "RUNQTY, OPSUHRS, OPRUNHRS, RUNSTATUS" & vbCrLf
sql = sql & "from RnopTable op" & vbCrLf
sql = sql & "join RunsTable run on run.RUNREF = op.OPREF and run.RUNNO = op.OPRUN" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = run.RUNREF" & vbCrLf
sql = sql & "join WcntTable wc on wc.WCNREF = op.OPCENTER and wc.WCNSHOP = op.OPSHOP" & vbCrLf
sql = sql & "where (OPSHOP = @Shop or @Shop = '')" & vbCrLf
sql = sql & "and OPSCHEDDATE >= @StartDate and OPSCHEDDATE < @EndPlus1" & vbCrLf
sql = sql & "and OPCOMPDATE IS null" & vbCrLf
sql = sql & "and RUNSTATUS not like 'c%'" & vbCrLf
sql = sql & "order by WCNSHOP, WCNREF, OPSUDATE" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript True, sql

sql = "exec DropStoredProcedureIfExists 'RptEfficiencyByWC'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptEfficiencyByWC" & vbCrLf
sql = sql & "@Shop varchar(20),  -- show WCs for this shop" & vbCrLf
sql = sql & "@StartDate date, -- starting OPCOMPDATE to include" & vbCrLf
sql = sql & "@EndDate date    -- ending OPCOMPDATE to include" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* EBM Efficiency by Workcenter Report 10/29/2018 TEL" & vbCrLf
sql = sql & "-- hours from routing vs hours charged by employee ? individually and entire company: 4 hours" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptEfficiencyByWC 'IS', '1/1/2018', '1/31/2018'" & vbCrLf
sql = sql & "exec RptEfficiencyByWC 'OS', '1/1/2018', '1/31/2018'" & vbCrLf
sql = sql & "exec RptEfficiencyByWC '', '1/1/2018', '1/31/2018'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @DatePlus1 date = dateadd(day,1, @EndDate)" & vbCrLf
sql = sql & "set @Shop = rtrim(@Shop)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get all operations in date range" & vbCrLf
sql = sql & "select WCNSHOP as Shop, OPCENTER as WC, wc.WCNDESC as [WC Desc], pt.PARTNUM as [Part#], OPRUN as [Run#], OPNO as [Op#]," & vbCrLf
sql = sql & "count(*) as [Charges], max(run.RUNQTY) as Qty, CONVERT(varchar(10)," & vbCrLf
sql = sql & "max(OPCOMPDATE),101) AS [Completed], max(OPSUHRS) as Setup, MAX(OPUNITHRS) as Unit," & vbCrLf
sql = sql & "max(cast(OPSUHRS + RUNQTY * OPUNITHRS as decimal(15,4))) AS [Op Hours]," & vbCrLf
sql = sql & "cast(sum(TCHOURS) as decimal(15,4)) as [Emp Hours], cast(0 as decimal(15,0)) as [Eff%]" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "from rnoptable op" & vbCrLf
sql = sql & "join TcitTable tc on tc.TCPARTREF = op.OPREF and tc.TCRUNNO = op.OPRUN and tc.TCOPNO = op.OPNO" & vbCrLf
sql = sql & "and tc.TCWC = op.OPCENTER and tc.TCSHOP = op.OPSHOP" & vbCrLf
sql = sql & "join WcntTable wc on wc.WCNREF = op.OPCENTER and wc.WCNSHOP = op.OPSHOP" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = op.OPREF" & vbCrLf
sql = sql & "join RunsTable run on run.RUNREF = op.OPREF and run.RUNNO = op.OPRUN" & vbCrLf
sql = sql & "where (@Shop = '' or WCNSHOP = @Shop) and OPCOMPLETE = 1" & vbCrLf
sql = sql & "and OPCOMPDATE >= @StartDate and OPCOMPDATE < @DatePlus1" & vbCrLf
sql = sql & "group by WCNSHOP, OPCENTER, WCNDESC, pt.PARTNUM, OPRUN, OPNO" & vbCrLf
sql = sql & "order by WCNSHOP, OPCENTER, max(OPCOMPDATE)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #temp set [Eff%] = 100.00 * [Op Hours] / [Emp Hours]" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from #temp order by Shop, WC, Completed" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
ExecuteScript True, sql

sql = "exec DropStoredProcedureIfExists 'RptEfficiencyByEmployee'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptEfficiencyByEmployee" & vbCrLf
sql = sql & "@Shop varchar(20),  -- show WCs for this shop" & vbCrLf
sql = sql & "@StartDate date, -- starting OPCOMPDATE to include" & vbCrLf
sql = sql & "@EndDate date    -- ending OPCOMPDATE to include" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* EBM Efficiency by Employee Report 10/30/2018 TEL" & vbCrLf
sql = sql & "hours worked vs hours on jobs ? individually and entire company" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptEfficiencyByEmployee 'IS', '1/1/2018', '1/31/2018'" & vbCrLf
sql = sql & "exec RptEfficiencyByEmployee 'OS', '1/1/2018', '1/31/2018'" & vbCrLf
sql = sql & "exec RptEfficiencyByEmployee '', '1/1/2018', '1/31/2018'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get all operations completing in date range" & vbCrLf
sql = sql & "declare @DatePlus1 date = dateadd(day,1, @EndDate)" & vbCrLf
sql = sql & "set @Shop = rtrim(@Shop)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select TCEMP as [Emp#], rtrim(emp.PREMFSTNAME) + ' ' + rtrim(EMP.PREMLSTNAME) as [Employee Name]," & vbCrLf
sql = sql & "CONVERT(VARCHAR(10), TCSTARTTIME, 101) + ' ' + CONVERT(VARCHAR(5), TCSTARTTIME, 108) as [Start]," & vbCrLf
sql = sql & "OPCENTER as WC, pt.PARTNUM as [Part#], OPRUN as [Run#], OPNO as [Op#]," & vbCrLf
sql = sql & "cast(TCHOURS as decimal(15,4)) as [Op Hours]," & vbCrLf
sql = sql & "cast(0 as decimal(15,4)) as [Wk Hours]" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "from rnoptable op" & vbCrLf
sql = sql & "join TcitTable tc on tc.TCPARTREF = op.OPREF and tc.TCRUNNO = OP.OPRUN and tc.TCOPNO = op.OPNO" & vbCrLf
sql = sql & "and tc.TCWC = op.OPCENTER and tc.TCSHOP = op.OPSHOP" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = op.OPREF" & vbCrLf
sql = sql & "join EmplTable emp on emp.PREMNUMBER = tc.TCEMP" & vbCrLf
sql = sql & "where TCSTARTTIME >= @StartDate and TCSTARTTIME < @DatePlus1 and (TCSHOP = @Shop or @Shop = '')" & vbCrLf
sql = sql & "order by TCEMP, TCSTARTTIME" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #temp set [Wk Hours] = (select sum(TMREGHRS + TMOVTHRS + TMDBLHRS)" & vbCrLf
sql = sql & "from TchdTable where TMEMP = Emp# AND TMDAY BETWEEN @StartDate AND @EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from #temp order by Emp#, Start" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
ExecuteScript True, sql


'''''''''''''''''''''''''''''''''''''''

sql = "exec DropStoredProcedureIfExists 'RptApAgingBase'" & vbCrLf
ExecuteScript True, sql


' select invoices to include
' EMB had a conversion issue where invoices prior to 7/1/12 had no debits or credits.
' this date restriction is enforced for all customers

sql = "create procedure RptApAgingBase" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Vendor varchar(20),   -- 'ALL' all" & vbCrLf
sql = sql & "@AgeByPostDate bit  -- = 1 to age by posting date (VIDTRECD), = 0 to post by invoice data (VIDATE)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* test" & vbCrLf
sql = sql & "exec RptApAgingBase '6/30/2018', '', 1" & vbCrLf
sql = sql & "exec RptApAgingBase '6/30/2018', 'ALACOP', 1" & vbCrLf
sql = sql & "exec RptApAgingBase '6/30/2018', 'ALACOP', 0" & vbCrLf
sql = sql & "exec RptApAgingBase '6/30/2018', 'IRSINT', 1" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF OBJECT_ID('tempdb..##TempApAging') IS NOT NULL" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "drop table ##TempApAging" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "PRINT @Vendor" & vbCrLf
sql = sql & "select RTRIM(VIVENDOR) AS VENDOR, rtrim(VEBNAME) AS VENDORNAME, VENUMBER AS [VENDOR#] ," & vbCrLf
sql = sql & "RTRIM(VINO) AS VINO, VIDUE, VIPAY," & vbCrLf
sql = sql & "cast(VIDTRECD as date) as VIDTRECD, cast(VIDATE as date) as VIDATE," & vbCrLf
sql = sql & "cast(VIREVDATE as date) as VIREVDATE, VIPIF," & vbCrLf
sql = sql & "cast(0 as int) as PONUMBER," & vbCrLf
sql = sql & "cast(0 as int) as PORELEASE," & vbCrLf
sql = sql & "cast('' as varchar(15)) as PoRef," & vbCrLf
sql = sql & "case when @AgeByPostDate = 1 then VIDTRECD else VIDATE end as AgeDate," & vbCrLf
sql = sql & "cast(0 as int) as AgeDays," & vbCrLf
sql = sql & "VEDDAYS as DiscDays," & vbCrLf
sql = sql & "VEDISCOUNT as DiscRate," & vbCrLf
sql = sql & "cast(null as datetime) as DiscDate," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as DiscAmt," & vbCrLf
sql = sql & "--cast('' as varchar(100)) as DiscMsg," & vbCrLf
sql = sql & "cast(0 as decimal(12,2)) AS AmtPaid," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as AmtDue," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [0-30 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [31-60 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [61-90 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [91+ Days]" & vbCrLf
sql = sql & "into ##TempApAging" & vbCrLf
sql = sql & "from VihdTable inv" & vbCrLf
sql = sql & "left join VndrTable ven on inv.VIVENDOR = ven.VEREF" & vbCrLf
sql = sql & "where VIDATE >= '7/1/12' and ((@AgeByPostDate = 0 and VIDATE <= @AsOfDate)" & vbCrLf
sql = sql & "or (@AgeByPostDate = 1 and VIDTRECD <= @AsOfDate)) and (@Vendor = 'ALL' or VIVENDOR = @Vendor)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update to use posting date in purchases journal.  Sometimes VIDTRECD is off by a day.  don't know why" & vbCrLf
sql = sql & "if @AgeByPostDate = 1" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "update inv" & vbCrLf
sql = sql & "set INV.AgeDate = dc.DCDATE" & vbCrLf
sql = sql & "from ##TempApAging inv join JritTable dc on dc.DCVENDOR = inv.VENDOR and dc.DCVENDORINV = inv.VINO" & vbCrLf
sql = sql & "where dc.DCHEAD like 'pj%' and DCDATE <> VIDTRECD" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert amount paid for computer checks as of that date (amount less non-voided checks as of that date)" & vbCrLf
sql = sql & "update ##TempApAging set AmtPaid = (select isnull(sum(DCDEBIT),0) - isnull(sum(DCCREDIT),0) from JritTable j" & vbCrLf
sql = sql & "join ChksTable c on c.CHKVENDOR = j.DCVENDOR and c.CHKNUMBER = j.DCCHECKNO" & vbCrLf
sql = sql & "join GlacTable gl on gl.GLACCTNO = j.DCACCTNO" & vbCrLf
sql = sql & "where j.DCVENDOR = VENDOR and j.DCVENDORINV = VINO and isnull(c.CHKVOID,0) = 0" & vbCrLf
sql = sql & "and (DCHEAD like 'CC%' or DCHEAD like 'XC%') and GLTYPE = 2 and DCDATE <= @AsOfDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now add amounts for external checks where the check number may or may not be specified" & vbCrLf
sql = sql & "update ##TempApAging set AmtPaid = AmtPaid + (select isnull(sum(DCDEBIT),0) - isnull(sum(DCCREDIT),0) from JritTable j" & vbCrLf
sql = sql & "left join ChksTable c on c.CHKVENDOR = j.DCVENDOR and c.CHKNUMBER = j.DCCHECKNO" & vbCrLf
sql = sql & "join GlacTable gl on gl.GLACCTNO = j.DCACCTNO" & vbCrLf
sql = sql & "where j.DCVENDOR = ##TempApAging.VENDOR and j.DCVENDORINV = VINO and c.CHKNUMBER is null" & vbCrLf
sql = sql & "and (DCHEAD like 'CC%' or DCHEAD like 'XC%') and GLTYPE = 2 and DCDATE <= @AsOfDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--delete from ##TempApAging where paid amt = invoice amt as of date requested" & vbCrLf
sql = sql & "delete from ##TempApAging where vidue = AmtPaid" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete invoices paid in full where VIDUE < 0 AND VIPAY > 0 (this is a bug)" & vbCrLf
sql = sql & "delete from ##TempApAging where VIDUE = - VIPAY and VIDUE < 0 and VIPIF = 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete where a credit has been overapplied (this is a bug too)" & vbCrLf
sql = sql & "delete from ##TempApAging where VIDUE < 0 and AmtPaid < VIDUE" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if there is a PO, use it's payment terms.  also set PO # and release" & vbCrLf
sql = sql & "update t" & vbCrLf
sql = sql & "set DiscDays = po.PODDAYS, DiscRate = po.PODISCOUNT," & vbCrLf
sql = sql & "PONUMBER = po.PONUMBER, PORELEASE = po.PORELEASE" & vbCrLf
sql = sql & "from ##TempApAging t" & vbCrLf
sql = sql & "join JritTable dc on dc.DCVENDOR = t.VENDOR and dc.DCVENDORINV = t.VINO" & vbCrLf
sql = sql & "join PohdTable po on po.PONUMBER = dc.DCPONUMBER and po.PORELEASE = dc.DCPORELEASE and DCPONUMBER <> 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ##TempApAging set PoRef = cast(PONUMBER as varchar(6)) + '-' + cast(PORELEASE as varchar(6)) where PONUMBER <> 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- determine aging" & vbCrLf
sql = sql & "update ##TempApAging set AmtDue = VIDUE - AmtPaid" & vbCrLf
sql = sql & "delete from ##TempApAging where AmtDue = 0" & vbCrLf
sql = sql & "update ##TempApAging set AgeDays = DATEDIFF(day,AgeDate,@AsOfDate)" & vbCrLf
sql = sql & "update ##TempApAging set [0-30 Days] = case when AgeDays between 0 and 30 then AmtDue else 0 end" & vbCrLf
sql = sql & "update ##TempApAging set [31-60 Days] = case when AgeDays between 31 and 60 then AmtDue else 0 end" & vbCrLf
sql = sql & "update ##TempApAging set [61-90 Days] = case when AgeDays between 61 and 90 then AmtDue else 0 end" & vbCrLf
sql = sql & "update ##TempApAging set [91+ Days] = case when AgeDays >= 90 then AmtDue else 0 end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- calculate discounts available" & vbCrLf
sql = sql & "update ##TempApAging set DiscDate = DATEADD(day,DiscDays,AgeDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ##TempApAging set DiscAmt = AmtDue * DiscRate / 100.00  where DiscDate >= @AsOfDate" & vbCrLf
ExecuteScript True, sql




'----------------------------
'----------------------------

sql = "exec DropStoredProcedureIfExists 'RptApAgingDetail'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptApAgingDetail" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Vendor varchar(20),   -- blank for all" & vbCrLf
sql = sql & "@AgeByPostDate bit  -- = 1 to age by posting date (VIDTRECD), = 0 to post by invoice data (VIDATE)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* test" & vbCrLf
sql = sql & "exec RptApAgingDetail '6/30/2018', 'ALL', 1" & vbCrLf
sql = sql & "exec RptApAgingDetail '6/30/2018', 'ALACOP', 0" & vbCrLf
sql = sql & "exec RptApAgingDetail '6/30/2018', 'ALACOP', 1" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptApAgingBase @AsOfDate, @Vendor, @AgeByPostDate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from ##TempApAging ORDER BY VENDOR,AgeDate,VINO" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript True, sql


'------------------------------------
'------------------------------------


sql = "exec DropStoredProcedureIfExists 'RptApAgingSummary'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptApAgingSummary" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Vendor varchar(20),   -- blank for all" & vbCrLf
sql = sql & "@AgeByPostDate bit  -- = 1 to age by posting date (VIDTRECD), = 0 to post by invoice data (VIDATE)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* test" & vbCrLf
sql = sql & "exec RptApAgingSummary '6/30/2018', 'ALL', 1" & vbCrLf
sql = sql & "exec RptApAgingSummary '6/30/2018', 'IRSINT', 1" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptApAgingBase @AsOfDate, @Vendor, @AgeByPostDate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select VENDOR, VENDORNAME, VENDOR# as [VDR#], count(*) as INVOICES, sum(AmtDue) as Total, sum([0-30 Days]) as [0-30 Days],  sum([31-60 Days]) as [31-60 Days]," & vbCrLf
sql = sql & "sum([61-90 Days]) as [61-90 Days], sum([91+ Days]) as [91+ Days] from ##TempApAging" & vbCrLf
sql = sql & "group by VENDOR, VENDORNAME, VENDOR#" & vbCrLf
sql = sql & "order by VENDOR" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript True, sql

'---------------------------
'---------------------------


' Vendor Statement
sql = "exec DropStoredProcedureIfExists 'RptApStatement'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptApStatement" & vbCrLf
sql = sql & "@Vendor varchar(50),      -- blank for all" & vbCrLf
sql = sql & "@StartDate date," & vbCrLf
sql = sql & "@EndDate date," & vbCrLf
sql = sql & "@IncludePaidInvoices bit  -- = 1 to show paid invoices" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "/* generate vendor statements" & vbCrLf
sql = sql & "Created 11/12/2018 TEL" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "exec RptApStatement 'ALL', '9/1/2017', '11/30/2018', 0" & vbCrLf
sql = sql & "exec RptApStatement 'RSTAHL', '1/1/2017', '11/30/2018', 1  --RSTAHL has 3 checks for in-119255" & vbCrLf
sql = sql & "exec RptApStatement 'GRAYBAR', '10/1/2017', '11/30/2018', 1" & vbCrLf
sql = sql & "exec RptApStatement 'GRAYBAR', '10/1/2017', '10/31/2017', 1" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "IF OBJECT_ID('tempdb..#temp') IS NOT NULL" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get all invoices in date range" & vbCrLf
sql = sql & "SELECT VEREF as Vendor, VEBNAME as [Vendor Name], VIDATE [Inv Date], VINO as [Inv #]," & vbCrLf
sql = sql & "VIDUE as [Inv Amt], cast('' as varchar(12)) as Journal, CHKACCT [Check Acct], CHKNUMBER as [Check #], CHKPOSTDATE as [Check Date]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as Discount," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as Voucher," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as ApApplied" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "FROM VndrTable v" & vbCrLf
sql = sql & "JOIN VihdTable inv on v.VEREF = inv.VIVENDOR" & vbCrLf
sql = sql & "LEFT OUTER JOIN JritTable on inv.VIVENDOR = JritTable.DCVENDOR AND inv.VINO = JritTable.DCVENDORINV" & vbCrLf
sql = sql & "LEFT OUTER JOIN ChksTable on JritTable.DCCHECKNO = ChksTable.CHKNUMBER AND DCCHKACCT = CHKACCT" & vbCrLf
sql = sql & "WHERE VIDATE between @StartDate and @EndDate" & vbCrLf
sql = sql & "and (DCHEAD like 'CC%' or DCHEAD like 'XC%')" & vbCrLf
sql = sql & "and (RTRIM(@Vendor) = 'ALL' or VIVENDOR = @Vendor)" & vbCrLf
sql = sql & "and CHKVOID = 0     -- there should be a voiddate, but there is not" & vbCrLf
sql = sql & "and CHKPOSTDATE <= @EndDate" & vbCrLf
sql = sql & "group by VEREF, VEBNAME, VINO, VIDATE, VIDUE, CHKACCT, CHKNUMBER, CHKPOSTDATE" & vbCrLf
sql = sql & "order by VEREF,VINO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now add discounts taken on or before the end date" & vbCrLf
sql = sql & "update #temp set Discount = isnull((select sum(dccredit - dcdebit)" & vbCrLf
sql = sql & "from JritTable dc where dc.DCVENDOR = #temp.Vendor and dc.DCVENDORINV = #temp.[Inv #]" & vbCrLf
sql = sql & "and dc.DCCHKACCT = #temp.[Check Acct] and dc.DCCHECKNO = #temp.[Check #] and dc.DCREF = 3" & vbCrLf
sql = sql & "and DCDATE <= @EndDate),0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- add voucher amount" & vbCrLf
sql = sql & "update #temp set Voucher = isnull((select sum(dccredit - dcdebit)" & vbCrLf
sql = sql & "from JritTable dc where dc.DCVENDOR = #temp.Vendor and dc.DCVENDORINV = #temp.[Inv #]" & vbCrLf
sql = sql & "and dc.DCCHKACCT = #temp.[Check Acct] and dc.DCCHECKNO = #temp.[Check #] and dc.DCREF = 2" & vbCrLf
sql = sql & "and DCDATE <= @EndDate),0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get journal ID" & vbCrLf
sql = sql & "update #temp set Journal = (select top 1 DCHEAD" & vbCrLf
sql = sql & "from JritTable dc where dc.DCVENDOR = #temp.Vendor and dc.DCVENDORINV = #temp.[Inv #]" & vbCrLf
sql = sql & "and dc.DCCHKACCT = #temp.[Check Acct] and dc.DCCHECKNO = #temp.[Check #] and dc.DCREF = 2" & vbCrLf
sql = sql & "and DCDATE <= @EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete paid in full invoices if not requested" & vbCrLf
sql = sql & "if @IncludePaidInvoices = 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "update #temp set ApApplied = isnull((select sum(Discount + Voucher)" & vbCrLf
sql = sql & "from #temp t2 where t2.Vendor = #temp.Vendor and t2.[Inv #] = #temp.[Inv #]),0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "delete from #temp where ApApplied = [Inv Amt]" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from #temp order by Vendor, [Inv Date], [Inv #]" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
ExecuteScript True, sql


sql = "exec DropStoredProcedureIfExists 'InsertGeneralJournal'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[InsertGeneralJournal]" & vbCrLf
sql = sql & "@JournalName varchar(100)," & vbCrLf
sql = sql & "@JournlDesc varchar(30)," & vbCrLf
sql = sql & "@JournalDate date," & vbCrLf
sql = sql & "@CSV varchar(MAX),     -- ('ACCT1',AMT1,'COMMENT1'),('ACCT2',AMT2,'COMMENT2')...  (Amount is minus for a credit)" & vbCrLf
sql = sql & "@User varchar(3)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* create a GL Journal" & vbCrLf
sql = sql & "returns blank if successful" & vbCrLf
sql = sql & "returns error message if unsuccessful" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON      -- required to avoid an error in sp with inserts and updates" & vbCrLf
sql = sql & "SET ANSI_WARNINGS OFF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if exists (select 1 from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_JournalTemp')" & vbCrLf
sql = sql & "drop table _JournalTemp" & vbCrLf
sql = sql & "if exists (select 1 from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_JournalTemp2')" & vbCrLf
sql = sql & "drop table _JournalTemp2" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create table of raw data" & vbCrLf
sql = sql & "create table _JournalTemp (Account varchar(12), Debit decimal(12,2), Credit decimal(12,2), Comment varchar(1000), ID int identity(1,1))" & vbCrLf
sql = sql & "declare @sql varchar(max) = 'insert _JournalTemp (Account, Debit, Credit, Comment) values' + char(13) + char(10) + @csv" & vbCrLf
sql = sql & "exec (@sql)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- construct data to insert" & vbCrLf
sql = sql & "select @JournalName as JINAME, 1 as JITRAN," & vbCrLf
sql = sql & "ID as JIREF," & vbCrLf
sql = sql & "Account as JIACCOUNT," & vbCrLf
sql = sql & "Debit as JIDEB," & vbCrLf
sql = sql & "Credit as JICRD," & vbCrLf
sql = sql & "Comment as JIDESC" & vbCrLf
sql = sql & "into _JournalTemp2" & vbCrLf
sql = sql & "from _JournalTemp" & vbCrLf
sql = sql & "where DEBIT <> 0 OR CREDIT <> 0" & vbCrLf
sql = sql & "order by ID" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- attempt to create journal" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "if exists (select * from GjhdTable where GJNAME = @JournalName)" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "rollback tran" & vbCrLf
sql = sql & "select 'Journal ' + @JournalName + ' already exists.'" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO dbo.GjhdTable" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "GJNAME" & vbCrLf
sql = sql & ",GJDESC" & vbCrLf
sql = sql & ",GJOPEN" & vbCrLf
sql = sql & ",GJPOST" & vbCrLf
sql = sql & ",GJPOSTED" & vbCrLf
sql = sql & ",GJREVERSE" & vbCrLf
sql = sql & ",GJCLOSE" & vbCrLf
sql = sql & ",GJREVID" & vbCrLf
sql = sql & ",GJREVDATE" & vbCrLf
sql = sql & ",GJEXTDESC" & vbCrLf
sql = sql & ",GJTEMPLATE" & vbCrLf
sql = sql & ",GJYEAREND" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "VALUES" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@JournalName" & vbCrLf
sql = sql & ",case when @JournlDesc = '' then @JournalName else @JournlDesc end" & vbCrLf
sql = sql & ",cast(getdate() as date)" & vbCrLf
sql = sql & ",@JournalDate" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",''" & vbCrLf
sql = sql & ",null" & vbCrLf
sql = sql & ",''" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now insert the debits and credits" & vbCrLf
sql = sql & "declare @now datetime = cast(convert(varchar(19),getdate(),100) as datetime)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO dbo.GjitTable" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "JINAME" & vbCrLf
sql = sql & ",JIDESC" & vbCrLf
sql = sql & ",JITRAN" & vbCrLf
sql = sql & ",JIREF" & vbCrLf
sql = sql & ",JIACCOUNT" & vbCrLf
sql = sql & ",JIDEB" & vbCrLf
sql = sql & ",JICRD" & vbCrLf
sql = sql & ",JIDATE" & vbCrLf
sql = sql & ",JILASTREVBY" & vbCrLf
sql = sql & ",JICLEAR" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "JINAME" & vbCrLf
sql = sql & ",JIDESC" & vbCrLf
sql = sql & ",JITRAN" & vbCrLf
sql = sql & ",JIREF" & vbCrLf
sql = sql & ",JIACCOUNT" & vbCrLf
sql = sql & ",JIDEB" & vbCrLf
sql = sql & ",JICRD" & vbCrLf
sql = sql & ",@now" & vbCrLf
sql = sql & ",@User" & vbCrLf
sql = sql & ",null" & vbCrLf
sql = sql & "from _JournalTemp2" & vbCrLf
sql = sql & "order by JITRAN, JIREF" & vbCrLf
sql = sql & "commit tran" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- show debits and credits" & vbCrLf
sql = sql & "declare @debits decimal(12,2), @credits decimal(12,2)" & vbCrLf
sql = sql & "select @debits = sum(jideb), @credits = sum(jicrd) from GjitTable where jiname = @JournalName" & vbCrLf
sql = sql & "" & vbCrLf
'SQL = SQL & "select 'Journal ' + @JournalName + ' created.  debits = ' + format(@debits,'N') + '  credits = ' + format(@credits, 'N')" & vbCrLf
sql = sql & "select 'Journal ' + @JournalName + ' created.  debits = ' + cast(@debits as varchar(12)) + '  credits = ' + cast(@credits as varchar(12))" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript True, sql



''''''''''''''''''''''''''''''''''''''''''''''''''


      ' update the version
      ExecuteScript True, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase107()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 185     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

' IMAINC 26-week Capacity vs Load Reports
sql = "dropfunctionifexists 'fnt_GetCapacity'" & vbCrLf
ExecuteScript True, sql

sql = "create function fnt_GetCapacity" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@Shop varchar(12),     -- <ALL> for all" & vbCrLf
sql = sql & "@Workcenter varchar(12),  -- <ALL> for all" & vbCrLf
sql = sql & "@StartDate date," & vbCrLf
sql = sql & "@Weeks int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "RETURNS @Capacity TABLE (Shop varchar(12), WC varchar(12), Weekend date, Hours decimal(10,2))" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/* return capacity for a given number of weeks" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "select * from dbo.fnt_GetCapacity('01', '0100', '12/1/2017',2)" & vbCrLf
sql = sql & "select * from dbo.fnt_GetCapacity('01', '<ALL>', '12/11/2018',26)" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "-- determine end date" & vbCrLf
sql = sql & "declare @EndDate date = DATEADD(day,7-datepart(WEEKDAY, @StartDate),@StartDate)" & vbCrLf
sql = sql & "set @EndDate = DATEADD(week, @Weeks - 1,@EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @Capacity" & vbCrLf
sql = sql & "select X.Shop, X.WC, X.Weekend, Sum(X.DayHours) as Hours" & vbCrLf
sql = sql & "from" & vbCrLf
sql = sql & "(select" & vbCrLf
sql = sql & "rtrim(WCCSHOP) as Shop," & vbCrLf
sql = sql & "rtrim(WCCCENTER) as WC," & vbCrLf
sql = sql & "case when WCCSHR1 = 0 then 1 else WCCSHR1 end * WCCSHH1" & vbCrLf
sql = sql & "+ case when WCCSHR2 = 0 then 1 else WCCSHR2 end * WCCSHH2" & vbCrLf
sql = sql & "+ case when WCCSHR3 = 0 then 1 else WCCSHR3 end * WCCSHH3" & vbCrLf
sql = sql & "+ case when WCCSHR4 = 0 then 1 else WCCSHR4 end * WCCSHH4 as DayHours," & vbCrLf
sql = sql & "DATEADD(day,7-datepart(WEEKDAY, WCCDATE),WCCDATE) as WeekEnd" & vbCrLf
sql = sql & "FROM WcclTable WHERE (WCCSHOP=@Shop or @Shop = '<ALL>')" & vbCrLf
sql = sql & "AND (WCCCENTER=@Workcenter or @Workcenter = '<ALL>' or @Shop = '<ALL>')" & vbCrLf
sql = sql & "AND WCCDATE between @StartDate and @EndDate) as X" & vbCrLf
sql = sql & "group by Shop, WC, WeekEnd" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "union" & vbCrLf
sql = sql & "select rtrim(WC2.WCNSHOP) as SHOP, rtrim(WC2.WCNREF) as WC, dateadd(day,-1,@StartDate) as WEEKEND, 0 as Hours" & vbCrLf
sql = sql & "from WcntTable WC2" & vbCrLf
sql = sql & "WHERE (WCNSHOP=@Shop or @Shop = '<ALL>') AND (WCNREF=@Workcenter or @Workcenter = '<ALL>' or @Shop = '<ALL>')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "order by Shop, WC, WeekEnd" & vbCrLf
sql = sql & "RETURN" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript True, sql

'----------------------
'----------------------

sql = "dropfunctionifexists 'fnt_GetLoad'" & vbCrLf
ExecuteScript True, sql

sql = "create function [dbo].[fnt_GetLoad]" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@Shop varchar(12),     -- <ALL> for all" & vbCrLf
sql = sql & "@Workcenter varchar(12),  -- <ALL> for all" & vbCrLf
sql = sql & "@StartDate date," & vbCrLf
sql = sql & "@Weeks int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "RETURNS @Load TABLE (Shop varchar(12), WC varchar(12), Weekend date, Hours decimal(10,2))" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/* return capacity for a given number of weeks" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "select * from dbo.fnt_GetLoad('01', '0600', '12/19/2018',26)" & vbCrLf
sql = sql & "select * from dbo.fnt_GetLoad('<ALL>', '<ALL>', '12/11/2017',26)" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "-- determine end date" & vbCrLf
sql = sql & "declare @EndDate date = DATEADD(day,7-datepart(WEEKDAY, @StartDate),@StartDate)" & vbCrLf
sql = sql & "set @EndDate = DATEADD(week, @Weeks - 1,@EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @Load" & vbCrLf
sql = sql & "select X.Shop, X.WC, X.Weekend, Sum(X.Hours) as Hours" & vbCrLf
sql = sql & "from" & vbCrLf
sql = sql & "(SELECT DISTINCT OPREF,OPRUN,OPNO,rtrim(OPSHOP) as Shop,RTRIM(OPCENTER) as WC,PADESC,RUNREMAININGQTY,RUNSTATUS," & vbCrLf
sql = sql & "cast(OPSUHRS+OPUNITHRS*RUNREMAININGQTY as decimal(10,2)) as Hours," & vbCrLf
sql = sql & "cast(OPSUDATE as date) as OPSUDATE,cast(OPSCHEDDATE as date) as OPSCHEDDATE," & vbCrLf
sql = sql & "cast(case when OPSCHEDDATE < @StartDate then dateadd(day,-1,@StartDate) else" & vbCrLf
sql = sql & "DATEADD(day,7-datepart(WEEKDAY, OPSCHEDDATE),OPSCHEDDATE) end  as Date) as WeekEnd" & vbCrLf
sql = sql & "FROM RnopTable op" & vbCrLf
sql = sql & "join RunsTable run on run.RUNREF = op.OPREF and run.RUNNO = op.OPRUN" & vbCrLf
sql = sql & "join WcntTable wc on wc.WCNREF = op.OPCENTER and wc.WCNSHOP = op.OPSHOP" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = op.OPREF" & vbCrLf
sql = sql & "WHERE (OPREF=RUNREF AND OPRUN=RUNNO AND OPCENTER=WCNREF AND WCNSERVICE=0 AND OPCOMPLETE=0)" & vbCrLf
sql = sql & "AND OPSCHEDDATE <= @EndDate AND (OPSHOP = @Shop or @Shop = '<ALL>')" & vbCrLf
sql = sql & "AND (OPCENTER LIKE @Workcenter or @Workcenter = '<ALL>' or @Shop = '<ALL>') and RUNSTATUS <> 'CA') as X" & vbCrLf
sql = sql & "group by X.Shop, X.WC, X.WeekEnd" & vbCrLf
sql = sql & "order by X.Shop, X.WC, X.WeekEnd" & vbCrLf
sql = sql & "RETURN" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript True, sql



'-----------------------------------
'-----------------------------------

sql = "dropstoredprocedureifexists 'RptCapacityVsLoad'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptCapacityVsLoad" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@Shop varchar(12),        -- <ALL> for all" & vbCrLf
sql = sql & "@Workcenter varchar(12),  -- <ALL> for all" & vbCrLf
sql = sql & "@StartDate date," & vbCrLf
sql = sql & "@Weeks int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "compare capacity and load by week" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "exec RptCapacityVsLoad '<ALL>','<ALL>','12/1/17',26" & vbCrLf
sql = sql & "exec RptCapacityVsLoad '01','0100','12/1/17',2" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select cap.*, isnull(ld.Hours,0) as Load, cap.Hours - isnull(ld.Hours,0) as Available,rtrim(wc.WCNDESC) as WCNDESC" & vbCrLf
sql = sql & "from dbo.fnt_GetCapacity(@Shop,@Workcenter,@StartDate,@Weeks) cap" & vbCrLf
sql = sql & "left join dbo.fnt_GetLoad(@Shop,@Workcenter,@StartDate,@Weeks) ld" & vbCrLf
sql = sql & "on ld.Shop = cap.Shop and ld.WC = cap.WC and ld.Weekend = cap.Weekend" & vbCrLf
sql = sql & "left join WcntTable wc on wc.WCNREF = cap.WC" & vbCrLf
sql = sql & "order by cap.Shop,cap.WC,cap.Weekend" & vbCrLf
ExecuteScript True, sql

'-------------------------------
'-------------------------------

sql = "dropstoredprocedureifexists 'RptLoadDetails'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptLoadDetails" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@Shop varchar(12),        -- <ALL> for all" & vbCrLf
sql = sql & "@Workcenter varchar(12),  -- <ALL> for all" & vbCrLf
sql = sql & "@StartDate date," & vbCrLf
sql = sql & "@Weeks int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "WC Load Details" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "exec RptLoadDetails '01','0100','12/1/17',26" & vbCrLf
sql = sql & "exec RptLoadDetails '01','0100','12/10/18',26" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @EndDate date = DATEADD(day,7-datepart(WEEKDAY, @StartDate),@StartDate)" & vbCrLf
sql = sql & "set @EndDate = DATEADD(week, @Weeks - 1,@EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT DISTINCT RTRIM(OPSHOP) AS OPSHOP, RTRIM(OPCENTER) AS OPCENTER, RTRIM(OPREF) AS OPREF,OPRUN,OPNO,OPSHOP as Shop,OPCENTER as WC," & vbCrLf
sql = sql & "RTRIM(PADESC) AS PADESC,RUNREMAININGQTY,RTRIM(RUNSTATUS) as RUNSTATUS," & vbCrLf
sql = sql & "cast(OPSUHRS+OPUNITHRS*RUNREMAININGQTY as decimal(10,2)) as Hours," & vbCrLf
sql = sql & "cast(OPSUDATE as date) as OPSUDATE,cast(OPSCHEDDATE as date) as OPSCHEDDATE," & vbCrLf
sql = sql & "case when OPSCHEDDATE < @StartDate then ' PRIOR ' else" & vbCrLf
sql = sql & "convert(varchar(8),DATEADD(day,7-datepart(WEEKDAY, OPSCHEDDATE),OPSCHEDDATE),1) end as WeekEnd" & vbCrLf
sql = sql & "FROM RnopTable op" & vbCrLf
sql = sql & "join RunsTable run on run.RUNREF = op.OPREF and run.RUNNO = op.OPRUN" & vbCrLf
sql = sql & "join WcntTable wc on wc.WCNREF = op.OPCENTER and wc.WCNSHOP = op.OPSHOP" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = op.OPREF" & vbCrLf
sql = sql & "WHERE (OPREF=RUNREF AND OPRUN=RUNNO AND OPCENTER=WCNREF AND WCNSERVICE=0 AND OPCOMPLETE=0)" & vbCrLf
sql = sql & "AND OPSCHEDDATE <= @EndDate AND (OPSHOP = @Shop or @Shop = '<ALL>')" & vbCrLf
sql = sql & "AND (OPCENTER LIKE @Workcenter or @Workcenter = '<ALL>' or @Shop = '<ALL>') and RUNSTATUS <> 'CA'" & vbCrLf
sql = sql & "order by Shop, WC, OPSCHEDDATE" & vbCrLf
ExecuteScript True, sql

sql = "dropstoredprocedureifexists 'GetScMOs'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[GetScMOs]" & vbCrLf
sql = sql & "@Parts varchar(30),    -- leading characters for MO parts to select" & vbCrLf
sql = sql & "@StartDate date,    -- start pick date" & vbCrLf
sql = sql & "@EndDate date       -- end pick date" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* test" & vbCrLf
sql = sql & "exec GetScMOs '', '12/14/18','12/31/18'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "set nocount on" & vbCrLf
sql = sql & "declare @EndDatePlus1 date = dateadd(day,1,@EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get list of all SC runs" & vbCrLf
sql = sql & "SELECT DISTINCT RUNREF as MRP_PARTREF, RUNNO, ISNULL(RUNOPCUR, 0) as RUNOPCUR, PARTNUM as MRP_PARTNUM," & vbCrLf
sql = sql & "RUNQTY As MRP_PARTQTYRQD, Convert(varchar(10), RUNPKSTART, 101) As MRP_PARTDATERQD," & vbCrLf
sql = sql & "Convert(varchar(10), RUNSCHED, 101) As MRP_ACTIONDATE, RUNSTATUS,RUNPKSTART,PABOMREV,op.OPCENTER," & vbCrLf
sql = sql & "ROW_NUMBER() over (ORDER BY RUNPKSTART) AS [MO#]" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "FROM RunsTable r" & vbCrLf
sql = sql & "join PartTable p on PARTREF = RUNREF" & vbCrLf
sql = sql & "join RnopTable op on op.opref = r.RUNREF and op.OPRUN = r.RUNNO and op.OPNO = r.RUNOPCUR" & vbCrLf
sql = sql & "WHERE RUNREF LIKE @Parts + '%' AND RUNSTATUS = 'SC' and op.OPCENTER = '0120'" & vbCrLf
sql = sql & "AND RUNPKSTART >= @StartDate AND RUNPKSTART < @EndDatePlus1" & vbCrLf
sql = sql & "order by RUNPKSTART" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all pick list requirements" & vbCrLf
sql = sql & "select t.[MO#], BMPARTREF, BMQTYREQD,MRP_PARTQTYRQD," & vbCrLf
sql = sql & "cast(BMQTYREQD * MRP_PARTQTYRQD as decimal(15,4)) as Qty," & vbCrLf
sql = sql & "CAST(-1 AS DECIMAL(15,4)) as Surplus" & vbCrLf
sql = sql & "into #picks" & vbCrLf
sql = sql & "from #temp t" & vbCrLf
sql = sql & "join BmplTable bm on bm.BMASSYPART = t.MRP_PARTREF and bm.BMREV = t.PABOMREV" & vbCrLf
sql = sql & "order by MO#, BMPARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all part quantities on hand less open pick list quantities" & vbCrLf
sql = sql & "-- use real time quantities rather than MRP" & vbCrLf
sql = sql & "select PARTREF as MRP_PARTREF, Min(PAQOH) as PAQOH," & vbCrLf
sql = sql & "cast(0 as decimal(15,4)) as Unpicked," & vbCrLf
sql = sql & "cast(0 as decimal(15,4)) as Avail," & vbCrLf
sql = sql & "cast(0 as decimal(15,4)) as OrigAvail" & vbCrLf
sql = sql & "into #parts" & vbCrLf
sql = sql & "from" & vbCrLf
sql = sql & "PartTable pt" & vbCrLf
sql = sql & "where PALEVEL <= 4" & vbCrLf
sql = sql & "group by PARTREF" & vbCrLf
sql = sql & "order by PARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #parts set Unpicked = isnull((select sum(PKPQTY) from MopkTable" & vbCrLf
sql = sql & "where PKPARTREF = MRP_PARTREF and PKTYPE = 9),0.0000)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #parts set Avail = PAQOH - Unpicked, OrigAvail = PAQOH - Unpicked" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- for each MO, determine if pick list quantities are available" & vbCrLf
sql = sql & "declare @MONO int" & vbCrLf
sql = sql & "DECLARE cur CURSOR FOR" & vbCrLf
sql = sql & "SELECT [MO#]" & vbCrLf
sql = sql & "FROM #temp" & vbCrLf
sql = sql & "order by [MO#]" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN cur" & vbCrLf
sql = sql & "FETCH NEXT FROM cur INTO @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "WHILE @@FETCH_STATUS = 0" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "update p set Surplus = pt.AVAIL - p.Qty" & vbCrLf
sql = sql & "from #picks p join #parts pt on p.BMPARTREF = pt.MRP_PARTREF" & vbCrLf
sql = sql & "where p.MO# = @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if any negative quantities for pick list, delete MO" & vbCrLf
sql = sql & "if(select min(Surplus) from #picks where MO# = @MONO) < 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "delete from #temp where MO# = @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- otherwise subtract from quantities available" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "update p set AVAIL = Avail - pk.Qty" & vbCrLf
sql = sql & "from #parts p join #picks pk on pk.BMPARTREF = p.MRP_PARTREF" & vbCrLf
sql = sql & "where pk.MO# = @MONO" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "FETCH NEXT FROM cur INTO @MONO" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CLOSE cur" & vbCrLf
sql = sql & "DEALLOCATE cur" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from #temp order by MO#" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
sql = sql & "drop table #picks" & vbCrLf
sql = sql & "drop table #parts" & vbCrLf
ExecuteScript True, sql




''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function


Private Function UpdateDatabase108()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 186     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "dropstoredprocedureifexists 'Qry_FillInspectorsActive'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure Qry_FillInspectorsActive" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "SELECT INSID FROM RinsTable WHERE INSACTIVE=1" & vbCrLf
sql = sql & "ORDER BY INSID" & vbCrLf
ExecuteScript True, sql

''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase109()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 187     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "dropstoredprocedureifexists 'RptApStatement'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptApStatement]" & vbCrLf
sql = sql & "@Vendor varchar(50),      -- blank for all" & vbCrLf
sql = sql & "@StartDate date," & vbCrLf
sql = sql & "@EndDate date," & vbCrLf
sql = sql & "@IncludePaidInvoices bit  -- = 1 to show paid invoices" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "/* generate vendor statements" & vbCrLf
sql = sql & "Created 11/12/2018 TEL" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "exec RptApStatement 'ALL', '9/1/2017', '11/30/2018', 0" & vbCrLf
sql = sql & "exec RptApStatement 'RSTAHL', '1/1/2017', '11/30/2018', 1  --RSTAHL has 3 checks for in-119255" & vbCrLf
sql = sql & "exec RptApStatement 'GRAYBAR', '10/1/2017', '11/30/2018', 1" & vbCrLf
sql = sql & "exec RptApStatement 'GRAYBAR', '6/1/18', '2/1/2019', 1" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "IF OBJECT_ID('tempdb..#temp') IS NOT NULL" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get all invoices in date range" & vbCrLf
sql = sql & "SELECT VEREF as Vendor, VEBNAME as [Vendor Name], VIDATE [Inv Date], VINO as [Inv #]," & vbCrLf
sql = sql & "VIDUE as [Inv Amt], cast('' as varchar(12)) as Journal, CHKACCT [Check Acct], CHKNUMBER as [Check #], CHKPOSTDATE as [Check Date]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as Discount," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as Voucher," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as ApApplied" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "FROM VndrTable v" & vbCrLf
sql = sql & "JOIN VihdTable inv on v.VEREF = inv.VIVENDOR" & vbCrLf
sql = sql & "LEFT OUTER JOIN JritTable on inv.VIVENDOR = JritTable.DCVENDOR AND inv.VINO = JritTable.DCVENDORINV" & vbCrLf
sql = sql & "and (DCHEAD like 'CC%' or DCHEAD like 'XC%')" & vbCrLf
sql = sql & "LEFT OUTER JOIN ChksTable on JritTable.DCCHECKNO = ChksTable.CHKNUMBER AND DCCHKACCT = CHKACCT" & vbCrLf
sql = sql & "WHERE VIDATE between @StartDate and @EndDate" & vbCrLf
sql = sql & "--and (DCHEAD like 'CC%' or DCHEAD like 'XC%')" & vbCrLf
sql = sql & "and (RTRIM(@Vendor) = 'ALL' or VIVENDOR = @Vendor)" & vbCrLf
sql = sql & "and isnull(CHKVOID,0) = 0     -- there should be a voiddate, but there is not" & vbCrLf
sql = sql & "and isnull(CHKPOSTDATE,'1/1/1900') <= @EndDate" & vbCrLf
sql = sql & "group by VEREF, VEBNAME, VINO, VIDATE, VIDUE, CHKACCT, CHKNUMBER, CHKPOSTDATE" & vbCrLf
sql = sql & "order by VEREF,VINO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now add discounts taken on or before the end date" & vbCrLf
sql = sql & "update #temp set Discount = isnull((select sum(dccredit - dcdebit)" & vbCrLf
sql = sql & "from JritTable dc where dc.DCVENDOR = #temp.Vendor and dc.DCVENDORINV = #temp.[Inv #]" & vbCrLf
sql = sql & "and dc.DCCHKACCT = #temp.[Check Acct] and dc.DCCHECKNO = #temp.[Check #] and dc.DCREF = 3" & vbCrLf
sql = sql & "and DCDATE <= @EndDate),0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- add voucher amount" & vbCrLf
sql = sql & "update #temp set Voucher = isnull((select sum(dccredit - dcdebit)" & vbCrLf
sql = sql & "from JritTable dc where dc.DCVENDOR = #temp.Vendor and dc.DCVENDORINV = #temp.[Inv #]" & vbCrLf
sql = sql & "and dc.DCCHKACCT = #temp.[Check Acct] and dc.DCCHECKNO = #temp.[Check #] and dc.DCREF = 2" & vbCrLf
sql = sql & "and DCDATE <= @EndDate),0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get journal ID" & vbCrLf
sql = sql & "update #temp set Journal = (select top 1 DCHEAD" & vbCrLf
sql = sql & "from JritTable dc where dc.DCVENDOR = #temp.Vendor and dc.DCVENDORINV = #temp.[Inv #]" & vbCrLf
sql = sql & "and dc.DCCHKACCT = #temp.[Check Acct] and dc.DCCHECKNO = #temp.[Check #] and dc.DCREF = 2" & vbCrLf
sql = sql & "and DCDATE <= @EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete paid in full invoices if not requested" & vbCrLf
sql = sql & "if @IncludePaidInvoices = 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "update #temp set ApApplied = isnull((select sum(Discount + Voucher)" & vbCrLf
sql = sql & "from #temp t2 where t2.Vendor = #temp.Vendor and t2.[Inv #] = #temp.[Inv #]),0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "delete from #temp where ApApplied = [Inv Amt]" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from #temp order by Vendor, [Inv Date], [Inv #]" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
ExecuteScript True, sql

sql = "dropstoredprocedureifexists 'InsertPayrollJournal'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[InsertPayrollJournal]" & vbCrLf
sql = sql & "@CSV varchar(MAX),     -- (''ACCT1'',AMT1),(''ACCT2'',AMT2)...  (Amount is minus for a credit)" & vbCrLf
sql = sql & "@User varchar(3)," & vbCrLf
sql = sql & "@PayrollDate datetime" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* aggregate IMAINC payroll journal data from an Excel file and create a summary GL Journal" & vbCrLf
sql = sql & "returns blank if successful" & vbCrLf
sql = sql & "returns error message if unsuccessful" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON      -- required to avoid an error in sp with inserts and updates" & vbCrLf
sql = sql & "SET ANSI_WARNINGS OFF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if exists (select 1 from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_PayrollTemp')" & vbCrLf
sql = sql & "drop table _PayrollTemp" & vbCrLf
sql = sql & "if exists (select 1 from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_PayrollTemp2')" & vbCrLf
sql = sql & "drop table _PayrollTemp2" & vbCrLf
sql = sql & "if exists (select 1 from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_PayrollTemp3')" & vbCrLf
sql = sql & "drop table _PayrollTemp3" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create journal name" & vbCrLf
sql = sql & "declare @JournalName varchar(12)" & vbCrLf
sql = sql & "set @JournalName = 'PR-' + cast(year(@PayrollDate) as varchar(4)) + '-'" & vbCrLf
sql = sql & "+ RIGHT('0' + MONTH(@PayrollDate),2) + RIGHT('0' + DAY(@PayrollDate),2)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create table of raw data" & vbCrLf
sql = sql & "-- SQL only allows 1000 rows to be added using insert ... values ... statements, so break this down into 10K chunks" & vbCrLf
sql = sql & "create table _PayrollTemp (Account varchar(12), Amount decimal(12,2))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--declare @sql varchar(max) = 'insert _PayRollTemp (Account,Amount) values' + char(13) + char(10) + @csv" & vbCrLf
sql = sql & "-- exec (@sql)" & vbCrLf
sql = sql & "declare @start int = 1" & vbCrLf
sql = sql & "declare @length int = len(@csv)" & vbCrLf
sql = sql & "declare @comma int = @length" & vbCrLf
sql = sql & "declare @sql varchar(max)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "while @start < @length" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "if @length - @start < 6000 break;" & vbCrLf
sql = sql & "set @comma = charindex('),(', @csv,@start + 5000)" & vbCrLf
sql = sql & "set @sql = 'insert _PayRollTemp (Account,Amount) values' + char(13) + char(10) + substring(@csv,@start,@comma-@start+1)" & vbCrLf
sql = sql & "--print (@sql)" & vbCrLf
sql = sql & "exec (@sql)" & vbCrLf
sql = sql & "set @start = @comma + 2" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @sql = 'insert _PayRollTemp (Account,Amount) values' + char(13) + char(10) + substring(@csv,@start,@length-@start+1)" & vbCrLf
sql = sql & "--print (@sql)" & vbCrLf
sql = sql & "exec (@sql)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- roll up into accounts" & vbCrLf
sql = sql & "select Account," & vbCrLf
sql = sql & "sum(cast(Amount as decimal(12,2))) as Total" & vbCrLf
sql = sql & "into _PayrollTemp2" & vbCrLf
sql = sql & "from _PayrollTemp" & vbCrLf
sql = sql & "group by [Account]" & vbCrLf
sql = sql & "order by Account" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- construct data to insert" & vbCrLf
sql = sql & "select @JournalName as JINAME, 1 as JITRAN," & vbCrLf
sql = sql & "ROW_NUMBER() over (ORDER BY Account) as JIREF," & vbCrLf
sql = sql & "Account as JIACCOUNT," & vbCrLf
sql = sql & "case when Total < 0 then 0.00 else Total end as JIDEB," & vbCrLf
sql = sql & "case when Total < 0 then -Total else 0.00 end as JICRD" & vbCrLf
sql = sql & "into _PayrollTemp3" & vbCrLf
sql = sql & "from _PayrollTemp2" & vbCrLf
sql = sql & "where Total <> 0" & vbCrLf
sql = sql & "order by Account" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- attempt to create journal" & vbCrLf
sql = sql & "begin tran" & vbCrLf
sql = sql & "if exists (select * from GjhdTable where GJNAME = @JournalName)" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "rollback tran" & vbCrLf
sql = sql & "select 'Journal ' + @JournalName + ' already exists.'" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO dbo.GjhdTable" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "GJNAME" & vbCrLf
sql = sql & ",GJDESC" & vbCrLf
sql = sql & ",GJOPEN" & vbCrLf
sql = sql & ",GJPOST" & vbCrLf
sql = sql & ",GJPOSTED" & vbCrLf
sql = sql & ",GJREVERSE" & vbCrLf
sql = sql & ",GJCLOSE" & vbCrLf
sql = sql & ",GJREVID" & vbCrLf
sql = sql & ",GJREVDATE" & vbCrLf
sql = sql & ",GJEXTDESC" & vbCrLf
sql = sql & ",GJTEMPLATE" & vbCrLf
sql = sql & ",GJYEAREND" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "VALUES" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@JournalName" & vbCrLf
sql = sql & ",''" & vbCrLf
sql = sql & ",CAST(getdate() as date)" & vbCrLf
sql = sql & ",null" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",''" & vbCrLf
sql = sql & ",null" & vbCrLf
sql = sql & ",'PAYROLL JOURNAL FOR PAY DATE '" & vbCrLf
sql = sql & "+ cast(year(@PayrollDate) as varchar(4)) + ' '" & vbCrLf
sql = sql & "+ right('0' + cast(month(@PayrollDate) as varchar(2)),2) + ' '" & vbCrLf
sql = sql & "+ right('0' + cast(DAY(@PayrollDate) as varchar(2)),2)" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ",0" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now insert the items" & vbCrLf
sql = sql & "declare @now datetime = cast(convert(varchar(19),getdate(),100) as datetime)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO dbo.GjitTable" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "JINAME" & vbCrLf
sql = sql & ",JIDESC" & vbCrLf
sql = sql & ",JITRAN" & vbCrLf
sql = sql & ",JIREF" & vbCrLf
sql = sql & ",JIACCOUNT" & vbCrLf
sql = sql & ",JIDEB" & vbCrLf
sql = sql & ",JICRD" & vbCrLf
sql = sql & ",JIDATE" & vbCrLf
sql = sql & ",JILASTREVBY" & vbCrLf
sql = sql & ",JICLEAR" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "JINAME" & vbCrLf
sql = sql & ",''" & vbCrLf
sql = sql & ",JITRAN" & vbCrLf
sql = sql & ",JIREF" & vbCrLf
sql = sql & ",JIACCOUNT" & vbCrLf
sql = sql & ",JIDEB" & vbCrLf
sql = sql & ",JICRD" & vbCrLf
sql = sql & ",@now" & vbCrLf
sql = sql & ",@User" & vbCrLf
sql = sql & ",null" & vbCrLf
sql = sql & "from _PayrollTemp3" & vbCrLf
sql = sql & "order by JITRAN, JIREF" & vbCrLf
sql = sql & "commit tran" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- show debits and credits" & vbCrLf
sql = sql & "declare @debits decimal(12,2), @credits decimal(12,2)" & vbCrLf
sql = sql & "select @debits = sum(jideb), @credits = sum(jicrd) from GjitTable where jiname = @JournalName" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select 'Payroll Journal ' + @JournalName + ' created.  debits = ' + format(@debits,'N') + '  credits = ' + format(@credits, 'N')" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript True, sql

sql = "update Preferences set PurchaseAccount = 0 where PreRecord = 1"
ExecuteScript True, sql

sql = "delete from LOLCTABLE where LOTEXLOCATION = ''" & vbCrLf
ExecuteScript True, sql

sql = "DropStoredProcedureIfExists 'UpdateMoPriorities'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[UpdateMoPriorities]" & vbCrLf
sql = sql & "@CSV varchar(MAX),     -- (''''Part'''',Run,Pri),(''''Part'''',Run,Pri)" & vbCrLf
sql = sql & "@User varchar(3)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* Update MO priorities" & vbCrLf
sql = sql & "returns blank if successful" & vbCrLf
sql = sql & "returns error message if unsuccessful" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "exec UpdateMoPriorities" & vbCrLf
sql = sql & "'(''146U7734526'',1,10),(''452T32223'',5,10),(''284W16171'',8,20),(''344T2201118'',10,30),(''256W26361'',4,10)','    '" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON      -- required to avoid an error in sp with inserts and updates" & vbCrLf
sql = sql & "SET ANSI_WARNINGS OFF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if exists (select 1 from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_MoPriorities')" & vbCrLf
sql = sql & "drop table _MoPriorities" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create table of raw data" & vbCrLf
sql = sql & "-- SQL only allows 1000 rows to be added using insert ... values ... statements, so break this down into 5K chunks" & vbCrLf
sql = sql & "create table _MoPriorities (Part varchar(30), Run int, Pri int)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @start int = 1" & vbCrLf
sql = sql & "declare @length int = len(@csv)" & vbCrLf
sql = sql & "declare @comma int = @length" & vbCrLf
sql = sql & "declare @sql varchar(max)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "while @start < @length" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "if @length - @start < 6000 break;" & vbCrLf
sql = sql & "set @comma = charindex('),(', @csv,@start + 5000)" & vbCrLf
sql = sql & "set @sql = 'insert _MoPriorities (Part,Run,Pri) values' + char(13) + char(10) + substring(@csv,@start,@comma-@start+1)" & vbCrLf
sql = sql & "exec (@sql)" & vbCrLf
sql = sql & "set @start = @comma + 2" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert last subset" & vbCrLf
sql = sql & "set @sql = 'insert _MoPriorities (Part,Run,Pri) values' + char(13) + char(10) +  substring(@csv,@start,@length-@start+1)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec (@sql)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- perform updates" & vbCrLf
sql = sql & "declare @ct int" & vbCrLf
sql = sql & "select @ct = count(*) from _MoPriorities" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update r set RUNPRIORITY = p.Pri" & vbCrLf
sql = sql & "from RunsTable r join _MoPriorities p on p.Part = r.RUNREF and p.Run = r.RUNNO" & vbCrLf
sql = sql & "where p.Pri <> r.RUNPRIORITY" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select cast(@@ROWCOUNT as varchar(5)) + ' of ' + cast(@ct as varchar(6)) + ' MO priorities updated' as [Result]" & vbCrLf
ExecuteScript True, sql

sql = "AddOrUpdateColumn 'Preferences', 'GetNextAvailablePoNumber', 'tinyint not null default 0'"
'SQL = "alter table Preferences add GetNextAvailablePoNumber tinyint not null default 0" & vbCrLf
ExecuteScript False, sql

sql = "exec DropStoredProcedureIfExists 'GetScMOs'" & vbCrLf
ExecuteScript True, sql

' for IMAINC only
' Get a list of all SC status MOs that can be auto-released
sql = "create procedure GetScMOs" & vbCrLf
sql = sql & "@Parts varchar(30),    -- leading characters for MO parts to select" & vbCrLf
sql = sql & "@StartDate date,    -- start pick date" & vbCrLf
sql = sql & "@EndDate date       -- end pick date" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* test" & vbCrLf
sql = sql & "exec GetScMOs '', '7/1/18','7/1/18'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "set nocount on" & vbCrLf
sql = sql & "declare @EndDatePlus1 date = dateadd(day,1,@EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get list of all SC runs" & vbCrLf
sql = sql & "SELECT DISTINCT RUNREF as MRP_PARTREF, RUNNO, ISNULL(RUNOPCUR, 0) as RUNOPCUR, PARTNUM as MRP_PARTNUM," & vbCrLf
sql = sql & "RUNQTY As MRP_PARTQTYRQD, Convert(varchar(10), RUNPKSTART, 101) As MRP_PARTDATERQD," & vbCrLf
sql = sql & "Convert(varchar(10), RUNSCHED, 101) As MRP_ACTIONDATE, RUNSTATUS,RUNPKSTART,PABOMREV,op.OPCENTER," & vbCrLf
sql = sql & "ROW_NUMBER() over (ORDER BY RUNPKSTART) AS [MO#]" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "FROM RunsTable r" & vbCrLf
sql = sql & "join PartTable p on PARTREF = RUNREF" & vbCrLf
sql = sql & "join RnopTable op on op.opref = r.RUNREF and op.OPRUN = r.RUNNO and op.OPNO = r.RUNOPCUR" & vbCrLf
sql = sql & "WHERE RUNREF LIKE @Parts + '%' AND RUNSTATUS = 'SC' and op.OPCENTER = '0120'" & vbCrLf
sql = sql & "AND RUNPKSTART >= @StartDate AND RUNPKSTART < @EndDatePlus1" & vbCrLf
sql = sql & "order by RUNPKSTART" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all pick list requirements" & vbCrLf
sql = sql & "select t.[MO#], BMPARTREF, BMQTYREQD,MRP_PARTQTYRQD," & vbCrLf
sql = sql & "cast(BMQTYREQD * MRP_PARTQTYRQD as decimal(15,4)) as Qty," & vbCrLf
sql = sql & "CAST(-1 AS DECIMAL(15,4)) as Surplus" & vbCrLf
sql = sql & "into #picks" & vbCrLf
sql = sql & "from #temp t" & vbCrLf
sql = sql & "join BmplTable bm on bm.BMASSYPART = t.MRP_PARTREF and bm.BMREV = t.PABOMREV" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = BMPARTREF and pt.PALEVEL <= 4" & vbCrLf
sql = sql & "order by MO#, BMPARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all part quantities on hand less open pick list quantities" & vbCrLf
sql = sql & "select MRP_PARTREF,sum(MRP_PARTQTYRQD) as AVAIL, sum(MRP_PARTQTYRQD) as OrigAvail" & vbCrLf
sql = sql & "into #parts" & vbCrLf
sql = sql & "from" & vbCrLf
sql = sql & "PartTable pt left join MrplTable mrp on mrp.MRP_PARTREF = pt.PARTREF" & vbCrLf
sql = sql & "where MRP_TYPE in (1,12)" & vbCrLf
sql = sql & "group by MRP_PARTREF" & vbCrLf
sql = sql & "order by MRP_PARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert parts not in mpr" & vbCrLf
sql = sql & "insert #parts" & vbCrLf
sql = sql & "select PARTREF,PAQOH, PAQOH" & vbCrLf
sql = sql & "from PartTable pt" & vbCrLf
sql = sql & "left join #parts p on p.MRP_PARTREF = pt.PARTREF" & vbCrLf
sql = sql & "where PALEVEL <= 4 and p.MRP_PARTREF is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- for each MO, determine if pick list quantities are available" & vbCrLf
sql = sql & "declare @MONO int" & vbCrLf
sql = sql & "DECLARE cur CURSOR FOR" & vbCrLf
sql = sql & "SELECT [MO#]" & vbCrLf
sql = sql & "FROM #temp" & vbCrLf
sql = sql & "order by [MO#]" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN cur" & vbCrLf
sql = sql & "FETCH NEXT FROM cur INTO @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "WHILE @@FETCH_STATUS = 0" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "update p set Surplus = pt.AVAIL - p.Qty" & vbCrLf
sql = sql & "from #picks p join #parts pt on p.BMPARTREF = pt.MRP_PARTREF" & vbCrLf
sql = sql & "where p.MO# = @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if any negative quantities for pick list, delete MO" & vbCrLf
sql = sql & "if(select min(Surplus) from #picks where MO# = @MONO) < 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "delete from #temp where MO# = @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- otherwise subtract from quantities available" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "update p set AVAIL = Avail - pk.Qty" & vbCrLf
sql = sql & "from #parts p join #picks pk on pk.BMPARTREF = p.MRP_PARTREF" & vbCrLf
sql = sql & "where pk.MO# = @MONO" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "FETCH NEXT FROM cur INTO @MONO" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CLOSE cur" & vbCrLf
sql = sql & "DEALLOCATE cur" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from #temp order by MO#" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
sql = sql & "drop table #picks" & vbCrLf
sql = sql & "drop table #parts" & vbCrLf
ExecuteScript True, sql

''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function




Private Function UpdateDatabase110()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 189     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "Dropstoredprocedureifexists 'RptEfficiencyByWC'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptEfficiencyByWC]" & vbCrLf
sql = sql & "@Shop varchar(20),  -- show WCs for this shop" & vbCrLf
sql = sql & "@StartDate date, -- starting OPCOMPDATE to include" & vbCrLf
sql = sql & "@EndDate date    -- ending OPCOMPDATE to include" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* EBM Efficiency by Workcenter Report 10/29/2018 TEL" & vbCrLf
sql = sql & "-- hours from routing vs hours charged by employee" & vbCrLf
sql = sql & "-- individually and entire company" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptEfficiencyByWC 'IS', '1/1/2018', '1/31/2018'" & vbCrLf
sql = sql & "exec RptEfficiencyByWC 'OS', '1/1/2018', '1/31/2018'" & vbCrLf
sql = sql & "exec RptEfficiencyByWC '', '1/1/2018', '3/15/19'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @DatePlus1 date = dateadd(day,1, @EndDate)" & vbCrLf
sql = sql & "set @Shop = rtrim(@Shop)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get all operations in date range" & vbCrLf
sql = sql & "select WCNSHOP as Shop, OPCENTER as WC, wc.WCNDESC as [WC Desc], pt.PARTNUM as [Part#], OPRUN as [Run#], OPNO as [Op#]," & vbCrLf
sql = sql & "count(*) as [Charges], max(run.RUNQTY) as Qty, CONVERT(varchar(10)," & vbCrLf
sql = sql & "max(OPCOMPDATE),101) AS [Completed], max(OPSUHRS) as Setup, MAX(OPUNITHRS) as Unit," & vbCrLf
sql = sql & "max(cast(OPSUHRS + RUNQTY * OPUNITHRS as decimal(15,4))) AS [Op Hours]," & vbCrLf
sql = sql & "cast(sum(TCHOURS) as decimal(15,4)) as [Emp Hours], cast(0 as decimal(15,0)) as [Eff%]" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "from rnoptable op" & vbCrLf
sql = sql & "join TcitTable tc on tc.TCPARTREF = op.OPREF and tc.TCRUNNO = op.OPRUN and tc.TCOPNO = op.OPNO" & vbCrLf
sql = sql & "and tc.TCWC = op.OPCENTER and tc.TCSHOP = op.OPSHOP" & vbCrLf
sql = sql & "join WcntTable wc on wc.WCNREF = op.OPCENTER and wc.WCNSHOP = op.OPSHOP" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = op.OPREF" & vbCrLf
sql = sql & "join RunsTable run on run.RUNREF = op.OPREF and run.RUNNO = op.OPRUN" & vbCrLf
sql = sql & "where (@Shop = '' or WCNSHOP = @Shop) and OPCOMPLETE = 1" & vbCrLf
sql = sql & "and OPCOMPDATE >= @StartDate and OPCOMPDATE < @DatePlus1" & vbCrLf
sql = sql & "group by WCNSHOP, OPCENTER, WCNDESC, pt.PARTNUM, OPRUN, OPNO" & vbCrLf
sql = sql & "order by WCNSHOP, OPCENTER, max(OPCOMPDATE)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "delete from #temp where [Emp Hours] = 0 and [Op Hours] = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #temp set [Eff%] = 100.00 * [Op Hours] / [Emp Hours] where [Emp Hours] > 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from #temp order by Shop, WC, Completed" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
ExecuteScript True, sql


sql = "Dropstoredprocedureifexists 'GetScMOs'" & vbCrLf
ExecuteScript True, sql
sql = "create procedure [dbo].[GetScMOs]" & vbCrLf
sql = sql & "@Parts varchar(30),    -- leading characters for MO parts to select" & vbCrLf
sql = sql & "@StartDate date,    -- start pick date" & vbCrLf
sql = sql & "@EndDate date       -- end pick date" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "rev 1: 3/20/19 add PKPARTREF and SURPLUS columns to be compatible with GetScMOsBlocked" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec GetScMOs '', '7/1/18','7/1/18'" & vbCrLf
sql = sql & "exec GetScMOs '34', '3/1/19','4/1/19'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "set nocount on" & vbCrLf
sql = sql & "declare @EndDatePlus1 date = dateadd(day,1,@EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get list of all SC runs" & vbCrLf
sql = sql & "--SELECT DISTINCT RUNREF as MRP_PARTREF, RUNNO, ISNULL(RUNOPCUR, 0) as RUNOPCUR, PARTNUM as MRP_PARTNUM," & vbCrLf
sql = sql & "--RUNQTY As MRP_PARTQTYRQD, Convert(varchar(10), RUNPKSTART, 101) As MRP_PARTDATERQD," & vbCrLf
sql = sql & "--Convert(varchar(10), RUNSCHED, 101) As MRP_ACTIONDATE, RUNSTATUS,RUNPKSTART,PABOMREV,op.OPCENTER," & vbCrLf
sql = sql & "--cast('' as varchar(30)) as PKPARTREF, cast(0 as decimal(15,4)) as Surplus," & vbCrLf
sql = sql & "--ROW_NUMBER() over (ORDER BY RUNPKSTART) AS [MO#]" & vbCrLf
sql = sql & "--into #temp" & vbCrLf
sql = sql & "--FROM RunsTable r" & vbCrLf
sql = sql & "--join PartTable p on PARTREF = RUNREF" & vbCrLf
sql = sql & "--join RnopTable op on op.opref = r.RUNREF and op.OPRUN = r.RUNNO and op.OPNO = r.RUNOPCUR" & vbCrLf
sql = sql & "--WHERE RUNREF LIKE @Parts + '%' AND RUNSTATUS = 'SC' and op.OPCENTER = '0120'" & vbCrLf
sql = sql & "--AND RUNPKSTART >= @StartDate AND RUNPKSTART < @EndDatePlus1" & vbCrLf
sql = sql & "--order by RUNPKSTART" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT DISTINCT RUNREF as MRP_PARTREF, RUNNO, ISNULL(RUNOPCUR, 0) as RUNOPCUR, PARTNUM as MRP_PARTNUM," & vbCrLf
sql = sql & "cast(''as varchar(30)) AS PKPARTREF,RUNQTY As MRP_PARTQTYRQD, Convert(varchar(10), RUNPKSTART, 101) As MRP_PARTDATERQD," & vbCrLf
sql = sql & "Convert(varchar(10), RUNSCHED, 101) As MRP_ACTIONDATE, RUNSTATUS,RUNPKSTART,PABOMREV,op.OPCENTER," & vbCrLf
sql = sql & "ROW_NUMBER() over (ORDER BY RUNPKSTART) AS [MO#], cast(0 as decimal(15,4)) as SURPLUS," & vbCrLf
sql = sql & "cast(0 as decimal(15,4)) as PAQOH," & vbCrLf
sql = sql & "cast(0 as decimal(15,4)) as Unpicked" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "FROM RunsTable r" & vbCrLf
sql = sql & "join PartTable p on PARTREF = RUNREF" & vbCrLf
sql = sql & "join RnopTable op on op.opref = r.RUNREF and op.OPRUN = r.RUNNO and op.OPNO = r.RUNOPCUR" & vbCrLf
sql = sql & "WHERE RUNREF LIKE @Parts + '%' AND RUNSTATUS = 'SC' and op.OPCENTER = '0120'" & vbCrLf
sql = sql & "AND RUNPKSTART >= @StartDate AND RUNPKSTART < @EndDatePlus1" & vbCrLf
sql = sql & "order by RUNPKSTART" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all pick list requirements" & vbCrLf
sql = sql & "--select t.[MO#], BMPARTREF, BMQTYREQD,MRP_PARTQTYRQD," & vbCrLf
sql = sql & "--cast(BMQTYREQD * MRP_PARTQTYRQD as decimal(15,4)) as Qty," & vbCrLf
sql = sql & "--CAST(-1 AS DECIMAL(15,4)) as Surplus" & vbCrLf
sql = sql & "--into #picks" & vbCrLf
sql = sql & "--from #temp t" & vbCrLf
sql = sql & "--join BmplTable bm on bm.BMASSYPART = t.MRP_PARTREF and bm.BMREV = t.PABOMREV" & vbCrLf
sql = sql & "--join PartTable pt on pt.PARTREF = BMPARTREF and pt.PALEVEL <= 4" & vbCrLf
sql = sql & "--order by MO#, BMPARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select t.[MO#], BMPARTREF, BMQTYREQD,MRP_PARTQTYRQD," & vbCrLf
sql = sql & "cast(BMQTYREQD * MRP_PARTQTYRQD as decimal(15,4)) as Qty," & vbCrLf
sql = sql & "CAST(-1 AS DECIMAL(15,4)) as Surplus, cast(0 as decimal(15,4)) as OrigAvail" & vbCrLf
sql = sql & "into #picks" & vbCrLf
sql = sql & "from #temp t" & vbCrLf
sql = sql & "join BmplTable bm on bm.BMASSYPART = t.MRP_PARTREF and bm.BMREV = t.PABOMREV" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = BMPARTREF and pt.PALEVEL <= 4" & vbCrLf
sql = sql & "order by MO#, BMPARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all part quantities on hand less open pick list quantities" & vbCrLf
sql = sql & "--select MRP_PARTREF,sum(MRP_PARTQTYRQD) as AVAIL, sum(MRP_PARTQTYRQD) as OrigAvail" & vbCrLf
sql = sql & "--into #parts" & vbCrLf
sql = sql & "--from" & vbCrLf
sql = sql & "--PartTable pt left join MrplTable mrp on mrp.MRP_PARTREF = pt.PARTREF" & vbCrLf
sql = sql & "--where MRP_TYPE in (1,12)" & vbCrLf
sql = sql & "--group by MRP_PARTREF" & vbCrLf
sql = sql & "--order by MRP_PARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all part current quantities on hand" & vbCrLf
sql = sql & "select PARTREF as MRP_PARTREF,PAQOH, cast(0 as decimal(15,4)) as Unpicked," & vbCrLf
sql = sql & "PAQOH as Available , PAQOH as OrigAvail" & vbCrLf
sql = sql & "into #parts" & vbCrLf
sql = sql & "from PartTable where PALEVEL <= 4" & vbCrLf
sql = sql & "order by PARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- add unpicked quantities" & vbCrLf
sql = sql & "update #parts set Unpicked = (select isnull(sum(PKPQTY),0)" & vbCrLf
sql = sql & "from MopkTable where PKPARTREF = MRP_PARTREF AND PKTYPE = 9" & vbCrLf
sql = sql & "and PKADATE is null)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #parts set Available = PAQOH - Unpicked, OrigAvail = PAQOH - Unpicked" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "---- insert parts not in mpr" & vbCrLf
sql = sql & "--insert #parts" & vbCrLf
sql = sql & "--select PARTREF,PAQOH, PAQOH" & vbCrLf
sql = sql & "--from PartTable pt" & vbCrLf
sql = sql & "--left join #parts p on p.MRP_PARTREF = pt.PARTREF" & vbCrLf
sql = sql & "--where PALEVEL <= 4 and p.MRP_PARTREF is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- for each MO, determine if pick list quantities are available" & vbCrLf
sql = sql & "declare @MONO int" & vbCrLf
sql = sql & "DECLARE cur CURSOR FOR" & vbCrLf
sql = sql & "SELECT [MO#]" & vbCrLf
sql = sql & "FROM #temp" & vbCrLf
sql = sql & "order by [MO#]" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN cur" & vbCrLf
sql = sql & "FETCH NEXT FROM cur INTO @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "WHILE @@FETCH_STATUS = 0" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "--update p set Surplus = pt.AVAIL - p.Qty" & vbCrLf
sql = sql & "--from #picks p join #parts pt on p.BMPARTREF = pt.MRP_PARTREF" & vbCrLf
sql = sql & "--where p.MO# = @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update p set Surplus = pt.Available - p.Qty, p.OrigAvail = pt.OrigAvail" & vbCrLf
sql = sql & "from #picks p join #parts pt on p.BMPARTREF = pt.MRP_PARTREF" & vbCrLf
sql = sql & "where p.MO# = @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if any negative quantities for pick list, delete MO" & vbCrLf
sql = sql & "if(select min(Surplus) from #picks where MO# = @MONO) < 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "delete from #temp where MO# = @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- otherwise subtract from quantities available" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "--update p set AVAIL = Avail - pk.Qty" & vbCrLf
sql = sql & "--from #parts p join #picks pk on pk.BMPARTREF = p.MRP_PARTREF" & vbCrLf
sql = sql & "--where pk.MO# = @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update p set Available = Available - pk.Qty" & vbCrLf
sql = sql & "from #parts p join #picks pk on pk.BMPARTREF = p.MRP_PARTREF" & vbCrLf
sql = sql & "where pk.MO# = @MONO" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "FETCH NEXT FROM cur INTO @MONO" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CLOSE cur" & vbCrLf
sql = sql & "DEALLOCATE cur" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from #temp order by MO#,RUNNO,PKPARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
sql = sql & "drop table #picks" & vbCrLf
sql = sql & "drop table #parts" & vbCrLf
ExecuteScript True, sql



sql = "dropstoredprocedureifexists 'GetScMOsBlocked'" & vbCrLf
ExecuteScript True, sql
sql = "create procedure [dbo].[GetScMOsBlocked]" & vbCrLf
sql = sql & "@Parts varchar(30), -- leading characters for MO parts to select" & vbCrLf
sql = sql & "@StartDate date,    -- start pick date" & vbCrLf
sql = sql & "@EndDate date       -- end pick date" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "Get SC MOs that are blocked from released by PAQOH or unpicked quantities" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec GetScMOs '775', '7/1/18','4/1/19'" & vbCrLf
sql = sql & "exec GetScMOsBlocked '775', '7/1/18','4/1/19'" & vbCrLf
sql = sql & "exec GetScMOsBlocked '34', '3/1/19','4/1/19'" & vbCrLf
sql = sql & "select paqoh from parttable where partref = dbo.fncompress('SHA6013KE00.080X48.000X144.00 ')" & vbCrLf
sql = sql & "exec GetScMOs '7753453313', '7/1/18','4/1/19'" & vbCrLf
sql = sql & "exec GetScMOsBlocked '7753453313', '7/1/18','4/1/19'" & vbCrLf
sql = sql & "exec GetScMOs '145T190125', '7/1/18','7/1/18'" & vbCrLf
sql = sql & "exec GetScMOsBlocked '145T190125', '7/1/18','7/1/18'" & vbCrLf
sql = sql & "exec GetScMOs '287W4136204', '7/1/18','7/1/18'" & vbCrLf
sql = sql & "exec GetScMOsBlocked '287W4136-04', '7/1/18','7/1/18'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "declare @Details bit = 1  -- = 1 if negative pick quantities are to be included" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set nocount on" & vbCrLf
sql = sql & "declare @EndDatePlus1 date = dateadd(day,1,@EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get list of all SC runs" & vbCrLf
sql = sql & "SELECT DISTINCT RUNREF as MRP_PARTREF, RUNNO, ISNULL(RUNOPCUR, 0) as RUNOPCUR, PARTNUM as MRP_PARTNUM," & vbCrLf
sql = sql & "cast(''as varchar(30)) AS PKPARTREF,RUNQTY As MRP_PARTQTYRQD, Convert(varchar(10), RUNPKSTART, 101) As MRP_PARTDATERQD," & vbCrLf
sql = sql & "Convert(varchar(10), RUNSCHED, 101) As MRP_ACTIONDATE, RUNSTATUS,RUNPKSTART,PABOMREV,op.OPCENTER," & vbCrLf
sql = sql & "ROW_NUMBER() over (ORDER BY RUNPKSTART) AS [MO#], cast(0 as decimal(15,4)) as SURPLUS," & vbCrLf
sql = sql & "cast(0 as decimal(15,4)) as PAQOH," & vbCrLf
sql = sql & "cast(0 as decimal(15,4)) as Unpicked" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "FROM RunsTable r" & vbCrLf
sql = sql & "join PartTable p on PARTREF = RUNREF" & vbCrLf
sql = sql & "join RnopTable op on op.opref = r.RUNREF and op.OPRUN = r.RUNNO and op.OPNO = r.RUNOPCUR" & vbCrLf
sql = sql & "WHERE RUNREF LIKE @Parts + '%' AND RUNSTATUS = 'SC' and op.OPCENTER = '0120'" & vbCrLf
sql = sql & "AND RUNPKSTART >= @StartDate AND RUNPKSTART < @EndDatePlus1" & vbCrLf
sql = sql & "order by RUNPKSTART" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all pick list requirements" & vbCrLf
sql = sql & "select t.[MO#], BMPARTREF, BMQTYREQD,MRP_PARTQTYRQD," & vbCrLf
sql = sql & "cast(BMQTYREQD * MRP_PARTQTYRQD as decimal(15,4)) as Qty," & vbCrLf
sql = sql & "CAST(-1 AS DECIMAL(15,4)) as Surplus, cast(0 as decimal(15,4)) as OrigAvail" & vbCrLf
sql = sql & "into #picks" & vbCrLf
sql = sql & "from #temp t" & vbCrLf
sql = sql & "join BmplTable bm on bm.BMASSYPART = t.MRP_PARTREF and bm.BMREV = t.PABOMREV" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = BMPARTREF and pt.PALEVEL <= 4" & vbCrLf
sql = sql & "order by MO#, BMPARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "---- get a list of all part quantities on hand less open pick list quantities" & vbCrLf
sql = sql & "--select MRP_PARTREF,sum(MRP_PARTQTYRQD) as AVAIL, sum(MRP_PARTQTYRQD) as OrigAvail" & vbCrLf
sql = sql & "--into #parts" & vbCrLf
sql = sql & "--from" & vbCrLf
sql = sql & "--PartTable pt left join MrplTable mrp on mrp.MRP_PARTREF = pt.PARTREF" & vbCrLf
sql = sql & "--where MRP_TYPE in (1,12)" & vbCrLf
sql = sql & "--group by MRP_PARTREF" & vbCrLf
sql = sql & "--order by MRP_PARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all part current quantities on hand plus open pick list quantities" & vbCrLf
sql = sql & "--select MRP_PARTREF,sum(MRP_PARTQTYRQD) as AVAIL, sum(MRP_PARTQTYRQD) as OrigAvail," & vbCrLf
sql = sql & "--sum(case when MRP_TYPE = 1 then MRP_PARTQTYRQD else 0 end) as MRP_STARTQTY," & vbCrLf
sql = sql & "--sum(case when MRP_TYPE = 12 then MRP_PARTQTYRQD else 0 end) as PICKQTY" & vbCrLf
sql = sql & "--into #parts" & vbCrLf
sql = sql & "--from" & vbCrLf
sql = sql & "--PartTable pt left join MrplTable mrp on mrp.MRP_PARTREF = pt.PARTREF" & vbCrLf
sql = sql & "--where MRP_TYPE in (1,12)" & vbCrLf
sql = sql & "--group by MRP_PARTREF" & vbCrLf
sql = sql & "--order by MRP_PARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get a list of all part current quantities on hand" & vbCrLf
sql = sql & "select PARTREF as MRP_PARTREF,PAQOH, cast(0 as decimal(15,4)) as Unpicked," & vbCrLf
sql = sql & "PAQOH as Available , PAQOH as OrigAvail" & vbCrLf
sql = sql & "into #parts" & vbCrLf
sql = sql & "from PartTable where PALEVEL <= 4" & vbCrLf
sql = sql & "order by PARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- add unpicked quantities" & vbCrLf
sql = sql & "update #parts set Unpicked = (select isnull(sum(PKPQTY),0)" & vbCrLf
sql = sql & "from MopkTable where PKPARTREF = MRP_PARTREF AND PKTYPE = 9" & vbCrLf
sql = sql & "and PKADATE is null)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #parts set Available = PAQOH - Unpicked, OrigAvail = PAQOH - Unpicked" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert parts not in mrp" & vbCrLf
sql = sql & "--insert #parts" & vbCrLf
sql = sql & "--select PARTREF,pt.PAQOH as Avail, pt.PAQOH as OrigAvail, PT.PAQOH, 0" & vbCrLf
sql = sql & "--from PartTable pt" & vbCrLf
sql = sql & "--left join #parts p on p.MRP_PARTREF = pt.PARTREF" & vbCrLf
sql = sql & "--where PALEVEL <= 4 and p.MRP_PARTREF is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- for each MO, determine if pick list quantities are available" & vbCrLf
sql = sql & "declare @MONO int" & vbCrLf
sql = sql & "DECLARE cur CURSOR FOR" & vbCrLf
sql = sql & "SELECT [MO#]" & vbCrLf
sql = sql & "FROM #temp" & vbCrLf
sql = sql & "order by [MO#]" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN cur" & vbCrLf
sql = sql & "FETCH NEXT FROM cur INTO @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "WHILE @@FETCH_STATUS = 0" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "update p set Surplus = pt.Available - p.Qty, p.OrigAvail = pt.OrigAvail" & vbCrLf
sql = sql & "from #picks p join #parts pt on p.BMPARTREF = pt.MRP_PARTREF" & vbCrLf
sql = sql & "where p.MO# = @MONO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if any negative quantities for pick list, keep it in the list as a blocked item" & vbCrLf
sql = sql & "if(select min(Surplus) from #picks where MO# = @MONO) < 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "declare @x int = 0" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "-- this item is pickable" & vbCrLf
sql = sql & "-- subtract from quantities available and then delete the item since it is not blocked" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "update p set Available = Available - pk.Qty" & vbCrLf
sql = sql & "from #parts p join #picks pk on pk.BMPARTREF = p.MRP_PARTREF" & vbCrLf
sql = sql & "where pk.MO# = @MONO" & vbCrLf
sql = sql & "delete from #temp where MO# = @MONO" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "FETCH NEXT FROM cur INTO @MONO" & vbCrLf
sql = sql & "END" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CLOSE cur" & vbCrLf
sql = sql & "DEALLOCATE cur" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if details requested, insert all negative quantities" & vbCrLf
sql = sql & "if @Details = 1" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "insert #temp (MRP_PARTREF,RUNNO,RUNOPCUR,MRP_PARTNUM, PKPARTREF, MRP_PARTQTYRQD," & vbCrLf
sql = sql & "MRP_PARTDATERQD,MRP_ACTIONDATE,RUNSTATUS,RUNPKSTART,PABOMREV,OPCENTER, SURPLUS, PAQOH, Unpicked, MO#)" & vbCrLf
sql = sql & "select t1.MRP_PARTREF,RUNNO,RUNOPCUR,MRP_PARTNUM, pt.PARTNUM as PKPARTREF, pk.qty as MRP_PARTQTYRQD," & vbCrLf
sql = sql & "MRP_PARTDATERQD,'','',RUNPKSTART,'','', pk.Surplus, pt.PAQOH, p.Unpicked, pk.MO#" & vbCrLf
sql = sql & "from #temp t1 join #picks pk on t1.MO# = pk.MO# and pk.Surplus < 0" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = pk.BMPARTREF" & vbCrLf
sql = sql & "join #parts p on p.MRP_PARTREF = pt.PARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--update t set Unpicked = p.Unpicked" & vbCrLf
sql = sql & "--from #temp t join #parts p on p.MRP_PARTREF = t.PKPARTREF" & vbCrLf
sql = sql & "--where t.PKPARTREF <> ''" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "select * from #temp order by MO#,RUNNO,PKPARTREF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
sql = sql & "drop table #picks" & vbCrLf
sql = sql & "drop table #parts" & vbCrLf
ExecuteScript True, sql



''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase111()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 190     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

' add .EsReportUsers.UserID column where it does not exist
sql = "begin tran" & vbCrLf
sql = sql & "delete from EsReportUsers -- delete any rows so can create a non-nullable column" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if not exists(select * from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = 'EsReportUsers' and COLUMN_NAME = 'UserID')" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- drop old primary key to name" & vbCrLf
sql = sql & "if exists (SELECT * FROM information_schema.table_constraints" & vbCrLf
sql = sql & "WHERE constraint_type = 'PRIMARY KEY'" & vbCrLf
sql = sql & "AND table_name = 'EsReportUsers')" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[EsReportUsers] DROP CONSTRAINT [PK_EsReportUsers]" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "alter table EsReportUsers add UserID varchar(30) not null" & vbCrLf
sql = sql & "ALTER TABLE [EsReportUsers] ADD CONSTRAINT [PK_EsReportUsers] PRIMARY KEY CLUSTERED" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "[UserID] ASC" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "commit tran" & vbCrLf
ExecuteScript True, sql

sql = "if exists(select * from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = 'CustTable' and COLUMN_NAME = 'CustTerms')"
sql = sql & "   alter table CustTable drop column CustTerms"
ExecuteScript True, sql

sql = "ALTER TABLE CustTable ADD CustTerms AS" & vbCrLf
sql = sql & "case when CUARDISC = 0 then ''" & vbCrLf
sql = sql & "when ROUND(CUARDISC,0) = CUARDISC then cast(cast(CUARDISC as int) as varchar(10)) + '/' + cast(CUDAYS as varchar(3)) + ' '" & vbCrLf
sql = sql & "when ROUND(CUARDISC,1) = CUARDISC then  cast(cast(CUARDISC as decimal(12,1)) as varchar(10)) + '/' + cast(CUDAYS as varchar(3)) + ' '" & vbCrLf
sql = sql & "when ROUND(CUARDISC,2) = CUARDISC then  cast(cast(CUARDISC as decimal(12,2)) as varchar(10)) + '/' + cast(CUDAYS as varchar(3)) + ' '" & vbCrLf
sql = sql & "else cast(CUARDISC as varchar(10)) + '/' + cast(CUDAYS as varchar(3)) + ' ' end" & vbCrLf
sql = sql & "+ 'NET ' + case when CUNETDAYS = 0 then '30' else cast(CUNETDAYS as varchar(10)) end" & vbCrLf
ExecuteScript True, sql

sql = "if exists (select * from INFORMATION_SCHEMA.VIEWS where TABLE_NAME = 'viewFAIRequired')" & vbCrLf
sql = sql & "drop view viewFAIRequired" & vbCrLf
ExecuteScript True, sql

sql = "create view viewFAIRequired as" & vbCrLf
sql = sql & "select RUNREF as FAIPartRef," & vbCrLf
sql = sql & "max(RUNCOMPLETE) AS [Last]," & vbCrLf
sql = sql & "cast(case when datediff(day,max(RUNCOMPLETE),getdate()) <= 730 then 0 else 1 end as int) as FAIRequired" & vbCrLf
sql = sql & "from RunsTable" & vbCrLf
sql = sql & "group by RUNREF" & vbCrLf
ExecuteScript True, sql

' repeat the below.  There was an error in update 107
sql = "dropfunctionifexists 'fnt_GetLoad'" & vbCrLf
ExecuteScript True, sql

sql = "create function [dbo].[fnt_GetLoad]" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "@Shop varchar(12),     -- <ALL> for all" & vbCrLf
sql = sql & "@Workcenter varchar(12),  -- <ALL> for all" & vbCrLf
sql = sql & "@StartDate date," & vbCrLf
sql = sql & "@Weeks int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "RETURNS @Load TABLE (Shop varchar(12), WC varchar(12), Weekend date, Hours decimal(10,2))" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/* return capacity for a given number of weeks" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "select * from dbo.fnt_GetLoad('01', '0600', '12/19/2018',26)" & vbCrLf
sql = sql & "select * from dbo.fnt_GetLoad('<ALL>', '<ALL>', '12/11/2017',26)" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "-- determine end date" & vbCrLf
sql = sql & "declare @EndDate date = DATEADD(day,7-datepart(WEEKDAY, @StartDate),@StartDate)" & vbCrLf
sql = sql & "set @EndDate = DATEADD(week, @Weeks - 1,@EndDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @Load" & vbCrLf
sql = sql & "select X.Shop, X.WC, X.Weekend, Sum(X.Hours) as Hours" & vbCrLf
sql = sql & "from" & vbCrLf
sql = sql & "(SELECT DISTINCT OPREF,OPRUN,OPNO,rtrim(OPSHOP) as Shop,RTRIM(OPCENTER) as WC,PADESC,RUNREMAININGQTY,RUNSTATUS," & vbCrLf
sql = sql & "cast(OPSUHRS+OPUNITHRS*RUNREMAININGQTY as decimal(10,2)) as Hours," & vbCrLf
sql = sql & "cast(OPSUDATE as date) as OPSUDATE,cast(OPSCHEDDATE as date) as OPSCHEDDATE," & vbCrLf
sql = sql & "cast(case when OPSCHEDDATE < @StartDate then dateadd(day,-1,@StartDate) else" & vbCrLf
sql = sql & "DATEADD(day,7-datepart(WEEKDAY, OPSCHEDDATE),OPSCHEDDATE) end  as Date) as WeekEnd" & vbCrLf
sql = sql & "FROM RnopTable op" & vbCrLf
sql = sql & "join RunsTable run on run.RUNREF = op.OPREF and run.RUNNO = op.OPRUN" & vbCrLf
sql = sql & "join WcntTable wc on wc.WCNREF = op.OPCENTER and wc.WCNSHOP = op.OPSHOP" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = op.OPREF" & vbCrLf
sql = sql & "WHERE (OPREF=RUNREF AND OPRUN=RUNNO AND OPCENTER=WCNREF AND WCNSERVICE=0 AND OPCOMPLETE=0)" & vbCrLf
sql = sql & "AND OPSCHEDDATE <= @EndDate AND (OPSHOP = @Shop or @Shop = '<ALL>')" & vbCrLf
sql = sql & "AND (OPCENTER LIKE @Workcenter or @Workcenter = '<ALL>' or @Shop = '<ALL>') and RUNSTATUS <> 'CA') as X" & vbCrLf
sql = sql & "group by X.Shop, X.WC, X.WeekEnd" & vbCrLf
sql = sql & "order by X.Shop, X.WC, X.WeekEnd" & vbCrLf
sql = sql & "RETURN" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript True, sql

sql = "exec AddOrUpdateColumn 'RunsTable', 'RunParentRunRef', 'varchar(30) not null default '''''"
ExecuteScript True, sql

sql = "exec AddOrUpdateColumn 'RunsTable', 'RunParentRunNo', 'int not null default 0'"
ExecuteScript True, sql

sql = "DropStoredProcedureIfExists 'InsertMultilevelMo'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].InsertMultilevelMo" & vbCrLf
sql = sql & "@ParentPart char(30)," & vbCrLf
sql = sql & "@ParentNumber int," & vbCrLf
sql = sql & "@ChildPart char(30)," & vbCrLf
sql = sql & "@ChildNumber int" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* insert child MO below parent MO" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec InsertMultilevelMo '100T14882', 1, '1003110042D02', 26" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- check that the max of 5 levels is not exceeded." & vbCrLf
sql = sql & "-- also check that there is no circular reference" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- first find all children of the child" & vbCrLf
sql = sql & "declare @Levels table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "Lev int," & vbCrLf
sql = sql & "MoRef char(30) ," & vbCrLf
sql = sql & "MoNo int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & ";WITH CTE (Lev, MoPart, MoRun)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "SELECT 6 as Lev, @ChildPart as MoPart, @ChildNumber as MoRun" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT Lev + 1 as Lev, RUNREF as MoPart, RUNNO as MoRun" & vbCrLf
sql = sql & "FROM RunsTable child" & vbCrLf
sql = sql & "join CTE on child.RunParentRunRef = CTE.MoPart and child.RunParentRunNo = CTE.MoRun" & vbCrLf
sql = sql & "where Lev < 10" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "insert @Levels" & vbCrLf
sql = sql & "SELECT * from CTE" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now find all the parents" & vbCrLf
sql = sql & ";WITH CTE (Lev, MoPart, MoRun)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "SELECT 5 as Lev, @ParentPart as MoPart, @ParentNumber as MoRun" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT Lev - 1, parent.RUNREF, parent.RUNNO" & vbCrLf
sql = sql & "FROM RunsTable parent" & vbCrLf
sql = sql & "join RunsTable child on child.RunParentRunRef = parent.RUNREF and child.RunParentRunNo = parent.RUNNO" & vbCrLf
sql = sql & "join CTE on child.RUNREF = CTE.MoPart and child.RUNNO = CTE.MoRun and CTE.Lev = Lev" & vbCrLf
sql = sql & "where Lev > 1" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "insert @Levels" & vbCrLf
sql = sql & "SELECT * from CTE" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if more than 5 levels do not allow insert" & vbCrLf
sql = sql & "if (select count(distinct Lev) from @Levels) > 5" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "select 'More than 5 levels not allowed'" & vbCrLf
sql = sql & "return" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if same MO more than once do not allow" & vbCrLf
sql = sql & "declare @ct int, @dupPart varchar(30), @dupNo int" & vbCrLf
sql = sql & "select top 1 @ct = ct, @dupPart = MoRef, @dupNo = Mono" & vbCrLf
sql = sql & "from (select count(*) as ct, MoRef,  MoNo from @Levels group by MoRef, Mono having count(*) > 1) x" & vbCrLf
sql = sql & "if @ct is not null" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "select 'MO ' + @DupPart + ' run ' + cast(@DupNo as varchar(10)) + ' cannot appear more than once'" & vbCrLf
sql = sql & "return" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- everything is OK.  Link the MO to its parent" & vbCrLf
sql = sql & "update RunsTable set RunParentRunRef = @ParentPart, RunParentRunNo = @ParentNumber" & vbCrLf
sql = sql & "where RUNREF = @ChildPart and RUNNO = @ChildNumber" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @@ROWCOUNT > 0" & vbCrLf
sql = sql & "select '' as result    -- indicate everything is OK" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "select 'unable to link MOs'" & vbCrLf
ExecuteScript True, sql

sql = "DropStoredProcedureIfExists 'RptMultilevelMoCostAnalysis'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptMultilevelMoCostAnalysis]" & vbCrLf
sql = sql & "@MoPart char(30)," & vbCrLf
sql = sql & "@MoNumber int" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* test" & vbCrLf
sql = sql & "RptMultilevelMoCostAnalysis '174570301',7976" & vbCrLf
sql = sql & "select top 100 * from tcittable" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- make a list of MOs at each level" & vbCrLf
sql = sql & ";WITH CTE (Lev, MoPart, MoRun)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "SELECT 1, @MoPart, @MoNumber" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT Lev + 1, RUNREF, RUNNO" & vbCrLf
sql = sql & "FROM RunsTable child" & vbCrLf
sql = sql & "join CTE on child.RunParentRunRef = CTE.MoPart and child.RunParentRunNo = CTE.MoRun" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "SELECT *" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "from CTE" & vbCrLf
sql = sql & "order by Lev, MoPart, MoRun" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "Lev as [Level]," & vbCrLf
sql = sql & "rtrim(pt.PARTNUM) as [MO Part]," & vbCrLf
sql = sql & "MoRun as [MO Run]," & vbCrLf
sql = sql & "'Labor' as [Type]," & vbCrLf
sql = sql & "cast(sum(TCHOURS) as decimal(12,2)) as [Hours]," & vbCrLf
sql = sql & "cast(sum(TCHOURS*TCRATE) as decimal(12,2)) as Labor," & vbCrLf
sql = sql & "cast(sum(TCHOURS*TCOHRATE) as decimal(12,2)) as OH," & vbCrLf
sql = sql & "null as Part," & vbCrLf
sql = sql & "null as [Part Desc]," & vbCrLf
sql = sql & "null as Qty," & vbCrLf
sql = sql & "null as [Unit Cost]," & vbCrLf
sql = sql & "null as [Ext Matl]," & vbCrLf
sql = sql & "null as [Ext Svc]," & vbCrLf
sql = sql & "'' as Vendor" & vbCrLf
sql = sql & "from TcitTable tc" & vbCrLf
sql = sql & "join #temp tmp on tmp.MoPart = tc.TCPARTREF and tmp.MoRun = tc.TCRUNNO" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = tc.TCPARTREF" & vbCrLf
sql = sql & "group by Lev,pt.PARTNUM,MoRun" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now do parts and services" & vbCrLf
sql = sql & "union all" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "Lev as [Level]," & vbCrLf
sql = sql & "rtrim(mopart.PARTNUM) as [MO Part]," & vbCrLf
sql = sql & "MoRun as [MO Run]," & vbCrLf
sql = sql & "case when pt.PALEVEL = 5 then 'Svc' else 'Matl' end as [Type]," & vbCrLf
sql = sql & "null as [Hours]," & vbCrLf
sql = sql & "null as [Labor]," & vbCrLf
sql = sql & "null as [OH]," & vbCrLf
sql = sql & "rtrim(pt.PARTNUM) as Part," & vbCrLf
sql = sql & "rtrim(pt.PADESC) as [Part Desc]," & vbCrLf
sql = sql & "PKAQTY as Qty," & vbCrLf
sql = sql & "INAMT as [Unit Cost]," & vbCrLf
sql = sql & "--cast(INAMT * PKAQTY as decimal(12,2)) as [Ext Cost]," & vbCrLf
sql = sql & "case when pt.PALEVEL <> 5 then cast(INAMT * PKAQTY as decimal(12,2)) else null end as [Ext Matl]," & vbCrLf
sql = sql & "case when pt.PALEVEL = 5 then cast(INAMT * PKAQTY as decimal(12,2)) else null end as [Ext Svc]," & vbCrLf
sql = sql & "isnull(rtrim(POVENDOR),'') as Vendor" & vbCrLf
sql = sql & "from MopkTable p" & vbCrLf
sql = sql & "join #temp tmp on tmp.MoPart = p.PKMOPART and tmp.MoRun = p.PKMORUN" & vbCrLf
sql = sql & "--join RunsTable r on r.RUNREF = p.PKMOPART and r.RUNNO = p.PKMORUN and r.RUNSTATUS <> 'CA'" & vbCrLf
sql = sql & "join PartTable pt ON PT.PARTREF = p.PKPARTREF" & vbCrLf
sql = sql & "join InvaTable inv on inv.INPART = p.PKPARTREF and inv.INMOPART = p.PKMOPART and inv.INMORUN = p.PKMORUN" & vbCrLf
sql = sql & "left join PoitTable poi on poi.PINUMBER = inv.INPONUMBER and poi.PIITEM = inv.INPOITEM and inv.inporev = poi.PIREV" & vbCrLf
sql = sql & "left join PohdTable ph on ph.PONUMBER = poi.PINUMBER" & vbCrLf
sql = sql & "join PartTable mopart on mopart.PARTREF = tmp.MoPart" & vbCrLf
sql = sql & "where INTYPE = 10" & vbCrLf
sql = sql & "--order by Lev, MoPart, MoRun,[Type]" & vbCrLf
sql = sql & "order by [Level],[MO Part],[MO Run],[Type]" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript True, sql


sql = "dropstoredprocedureifexists 'RptMORunOptDetail'" & vbCrLf
ExecuteScript True, sql

sql = "create PROCEDURE RptMORunOptDetail" & vbCrLf
sql = sql & "@MOPart as varchar(30),@cutoffDt as varchar(10)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "used by Part Info and Status Report" & vbCrLf
sql = sql & "rev 5/16/2019 TEL to correctly show PO info." & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "CREATE TABLE #tempRunOpDet" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "RUNREF Varchar(30) NULL," & vbCrLf
sql = sql & "RUNNO int Null," & vbCrLf
sql = sql & "CUROPNO smallint NULL," & vbCrLf
sql = sql & "NextOPNO smallint NULL," & vbCrLf
sql = sql & "NextOPSHOP varchar(12) NULL," & vbCrLf
sql = sql & "NextOPCENTER varchar(12) NULL," & vbCrLf
sql = sql & "PONUMBER int NULL," & vbCrLf
sql = sql & "POVENDOR varchar(30) NULL" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "INSERT INTO #tempRunOpDet(RUNREF, RUNNO, CUROPNO,NextOPNO, NextOPSHOP, NextOPCENTER)" & vbCrLf
sql = sql & "select distinct runstable.Runref, runstable.runno, runstable.RUNOPCUR, f.NextOPNO, f.NextOPSHOP, f.NextOPCENTER" & vbCrLf
sql = sql & "from runstable,RnopTable," & vbCrLf
sql = sql & "(select a.runref, a.runno, b.OPNO NextOPNO, OPSHOP NextOPSHOP, OPCENTER NextOPCENTER," & vbCrLf
sql = sql & "ROW_NUMBER() OVER (PARTITION BY opref, oprun" & vbCrLf
sql = sql & "ORDER BY opref DESC, oprun) as rn" & vbCrLf
sql = sql & "from runstable a,rnopTable b" & vbCrLf
sql = sql & "where a.runref = b.opref and" & vbCrLf
sql = sql & "a.runno = b.oprun and b.opcompdate is null and b.opno <> a.runopcur" & vbCrLf
sql = sql & "--and b.opno > a.runopcur" & vbCrLf
sql = sql & ") as f" & vbCrLf
sql = sql & "where RunsTable.RUNSCHED <= @cutoffDt AND" & vbCrLf
sql = sql & "RunsTable.RUNSTATUS NOT IN ('CA','CL','CO') and RunsTable.runref =  @MOPart" & vbCrLf
sql = sql & "and RunsTable.runref =  f.runref AND RunsTable.runno = f.runno" & vbCrLf
sql = sql & "and f.rn = 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--UPDATE a SET a.PONUMBER = poitTable.PINUMBER, a.POVENDOR = PohdTable.POVENDOR" & vbCrLf
sql = sql & "--FROM #tempRunOpDet a, ShopTable, poitTable,PohdTable, rnopTable WHERE SHPREF = OPSHOP" & vbCrLf
sql = sql & "--   AND SHPSERVICE = 1 AND PIRUNPART = RUNREF AND PIRUNNO = RUNNO" & vbCrLf
sql = sql & "--   and rnopTable.Opref = RUNREF AND rnopTable.OPRUN = RUNNO" & vbCrLf
sql = sql & "--   AND poitTable.PINUMBER = PohdTable.PONUMBER" & vbCrLf
sql = sql & "--   AND PIRUNOPNO = a.CUROPNO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select *" & vbCrLf
sql = sql & "--FROM #tempRunOpDet a" & vbCrLf
sql = sql & "--join rnopTable on rnopTable.Opref = a.RUNREF AND rnopTable.OPRUN = a.RUNNO and RnopTable.OPNO = a.CUROPNO" & vbCrLf
sql = sql & "--join ShopTable on SHPREF = OPSHOP -- AND SHPSERVICE = 1" & vbCrLf
sql = sql & "--join poitTable on PIRUNPART = RUNREF AND PIRUNNO = RUNNO and PIRUNOPNO = a.CUROPNO" & vbCrLf
sql = sql & "--join PohdTable on poitTable.PINUMBER = PohdTable.PONUMBER" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "UPDATE a SET a.PONUMBER = poitTable.PINUMBER, a.POVENDOR = PohdTable.POVENDOR" & vbCrLf
sql = sql & "FROM #tempRunOpDet a" & vbCrLf
sql = sql & "join rnopTable on rnopTable.Opref = a.RUNREF AND rnopTable.OPRUN = a.RUNNO and RnopTable.OPNO = a.CUROPNO" & vbCrLf
sql = sql & "join ShopTable on SHPREF = OPSHOP -- AND SHPSERVICE = 1 not required" & vbCrLf
sql = sql & "join poitTable on PIRUNPART = RUNREF AND PIRUNNO = RUNNO and PIRUNOPNO = a.CUROPNO" & vbCrLf
sql = sql & "join PohdTable on poitTable.PINUMBER = PohdTable.PONUMBER" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from #tempRunOpDet" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT DISTINCT a.NextOPNO, a.NextOPSHOP, a.NextOPCENTER, a.PONUMBER, RnopTable.OPSHOP as CurOPShop," & vbCrLf
sql = sql & "RnopTable.OPCENTER as CurOPCenter, POVENDOR, RunsTable.*" & vbCrLf
sql = sql & "FROM runstable, RnopTable, #tempRunOpDet a" & vbCrLf
sql = sql & "WHERE a.RunRef = runstable.Runref AND a.RunNo = runstable.RunNO" & vbCrLf
sql = sql & "AND RnopTable.OPREF = RunsTable.runref" & vbCrLf
sql = sql & "AND RnopTable.OPRUN = RunsTable.runno" & vbCrLf
sql = sql & "and RunsTable.runopcur = RnopTable.OPNO" & vbCrLf
sql = sql & "AND a.RunRef = RnopTable.OPREF AND a.RunNo = RnopTable.OPRUN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DROP table #tempRunOpDet" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript True, sql



''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function


Private Function UpdateDatabase112()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 191     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''


sql = "DropStoredProcedureIfExists 'RptArAgingBase'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptArAgingBase]" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Customer varchar(20)   -- blank for all" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* Detail AR Aging Base sp" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptArAgingBase '3/1/2019', 'BCA614'" & vbCrLf
sql = sql & "exec RptArAgingBase '3/1/2019', ''" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF OBJECT_ID('tempdb..##TempArAging') IS NOT NULL" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "drop table ##TempArAging" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get all invoices existing but not fully paid on desired date" & vbCrLf
sql = sql & "select rtrim(cust.CUNAME) as [Customer Name], rtrim(inv.INVCUST) as Nickname, INVNO as [Inv #]," & vbCrLf
sql = sql & "cast(INVDATE as DATE) AS [Inv Date], INVTOTAL as [Inv Total], isnull(x.Debits,0) as [Amt Paid]," & vbCrLf
sql = sql & "isnull(x.ct,0) as ct, case when INVCHECKDATE is null then '' else convert(varchar(10),INVCHECKDATE,101) end as [Ck Date]," & vbCrLf
sql = sql & "cast(0 as int) as [Age Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [Amt Due]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [0-30 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [31-60 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [61-90 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [91-120 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [121+ Days]" & vbCrLf
sql = sql & "into ##TempArAging" & vbCrLf
sql = sql & "from CihdTable as inv" & vbCrLf
sql = sql & "join CustTable cust on cust.CUREF = inv.INVCUST" & vbCrLf
sql = sql & "left join (select DCCUST, DCINVNO, isnull(sum(DCDEBIT),0) AS Debits, COUNT(*) AS ct from JritTable" & vbCrLf
sql = sql & "where DCHEAD like 'cr%'" & vbCrLf
sql = sql & "AND DCDATE <= @AsOfDate" & vbCrLf
sql = sql & "--and DCCUST = inv.INVCUST" & vbCrLf
sql = sql & "and DCINVNO <> 0" & vbCrLf
sql = sql & "and DCDEBIT <> 0" & vbCrLf
sql = sql & "group by DCCUST, DCINVNO) x on x.DCCUST = inv.INVCUST and x.DCINVNO = inv.INVNO" & vbCrLf
sql = sql & "where INVDATE <= @AsOfDate" & vbCrLf
sql = sql & "and INVCUST like @Customer + '%'" & vbCrLf
sql = sql & "and INVTOTAL <> isnull(x.Debits,0)" & vbCrLf
sql = sql & "and (INVPIF = 0" & vbCrLf
sql = sql & "or isnull(INVCHECKDATE, DATEFROMPARTS(2050,1,1)) > @AsOfDate)" & vbCrLf
sql = sql & "ORDER by INVCUST, INVDATE, INVNO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ##TempArAging set [Amt Due] = [Inv Total] - [Amt Paid]" & vbCrLf
sql = sql & "update ##TempArAging set [Age Days] = DATEDIFF(day,[Inv Date],@AsOfDate)" & vbCrLf
sql = sql & "update ##TempArAging set [0-30 Days] = case when [Age Days] between 0 and 30 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [31-60 Days] = case when [Age Days] between 31 and 60 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [61-90 Days] = case when [Age Days] between 61 and 90 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [91-120 Days] = case when [Age Days] between 91 and 120 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [121+ Days] = case when [Age Days] > 120 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from ##TempArAging order by [Customer Name], [Inv #]" & vbCrLf
sql = sql & "--select INVCUST, count(*) as ct, sum([Amt Due]) AS DUE," & vbCrLf
sql = sql & "--SUM([0-30 Days]) as [0-30 Days]," & vbCrLf
sql = sql & "--SUM([31-60 Days]) as [31-60 Days]," & vbCrLf
sql = sql & "--SUM([61-90 Days]) as [61-90 Days]," & vbCrLf
sql = sql & "--SUM([91-120 Days]) as [91-120 Days]," & vbCrLf
sql = sql & "--SUM([121+ Days]) as [121+ Days]" & vbCrLf
sql = sql & "--from ##TempArAging GROUP BY INVCUST ORDER BY INVCUST" & vbCrLf
ExecuteScript True, sql

'----------------------------------------------
'----------------------------------------------

sql = "DropStoredProcedureIfExists 'RptArAgingDetail'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptArAgingDetail]" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Customer varchar(20)   -- blank for all" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* Detail AR Aging" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptArAgingDetail '3/1/2019', 'BCA614'" & vbCrLf
sql = sql & "exec RptArAgingDetail '3/1/2019', ''" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase @AsofDate, @Customer" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from ##TempArAging order by [Customer Name], [Inv #]" & vbCrLf
ExecuteScript True, sql

'----------------------------------------------
'----------------------------------------------


sql = "DropStoredProcedureIfExists 'RptArAgingSummary'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptArAgingSummary]" & vbCrLf
sql = sql & "@AsOfDate date" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* test" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptArAgingSummary '3/1/2019'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase @AsOfDate, ''" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select [Customer Name], Nickname, count(*) as INVOICES, sum([Amt Due]) as Total," & vbCrLf
sql = sql & "sum([0-30 Days]) as [0-30 Days]," & vbCrLf
sql = sql & "sum([31-60 Days]) as [31-60 Days]," & vbCrLf
sql = sql & "sum([61-90 Days]) as [61-90 Days]," & vbCrLf
sql = sql & "sum([91-120 Days]) as [91-120 Days]," & vbCrLf
sql = sql & "sum([121+ Days]) as [121+ Days]" & vbCrLf
sql = sql & "from ##TempArAging" & vbCrLf
sql = sql & "group by [Customer Name],[NickName]" & vbCrLf
sql = sql & "order by [Customer Name]" & vbCrLf
ExecuteScript True, sql

sql = "dropstoredprocedureifexists 'InsertMultilevelMo'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[InsertMultilevelMo]" & vbCrLf
sql = sql & "@ParentPart char(30)," & vbCrLf
sql = sql & "@ParentNumber int," & vbCrLf
sql = sql & "@ChildPart char(30)," & vbCrLf
sql = sql & "@ChildNumber int" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* insert child MO below parent MO" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec InsertMultilevelMo '100T14882', 1, '1003110042D02', 26" & vbCrLf
sql = sql & "exec InsertMultilevelMo '198040001', 9804, '198401301', 9840" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- check that the max of 5 levels is not exceeded." & vbCrLf
sql = sql & "-- also check that there is no circular reference" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- first find all children of the child" & vbCrLf
sql = sql & "SET NOCOUNT ON" & vbCrLf
sql = sql & "declare @Levels table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "Lev int," & vbCrLf
sql = sql & "MoRef char(30) ," & vbCrLf
sql = sql & "MoNo int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & ";WITH CTE (Lev, MoPart, MoRun)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "SELECT 6 as Lev, @ChildPart as MoPart, @ChildNumber as MoRun" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT Lev + 1 as Lev, RUNREF as MoPart, RUNNO as MoRun" & vbCrLf
sql = sql & "FROM RunsTable child" & vbCrLf
sql = sql & "join CTE on child.RunParentRunRef = CTE.MoPart and child.RunParentRunNo = CTE.MoRun" & vbCrLf
sql = sql & "where Lev < 10" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "insert @Levels" & vbCrLf
sql = sql & "SELECT * from CTE" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now find all the parents" & vbCrLf
sql = sql & ";WITH CTE (Lev, MoPart, MoRun)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "SELECT 5 as Lev, @ParentPart as MoPart, @ParentNumber as MoRun" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT Lev - 1, parent.RUNREF, parent.RUNNO" & vbCrLf
sql = sql & "FROM RunsTable parent" & vbCrLf
sql = sql & "join RunsTable child on child.RunParentRunRef = parent.RUNREF and child.RunParentRunNo = parent.RUNNO" & vbCrLf
sql = sql & "join CTE on child.RUNREF = CTE.MoPart and child.RUNNO = CTE.MoRun and CTE.Lev = Lev" & vbCrLf
sql = sql & "where Lev > 1" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "insert @Levels" & vbCrLf
sql = sql & "SELECT * from CTE" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if more than 5 levels do not allow insert" & vbCrLf
sql = sql & "if (select count(distinct Lev) from @Levels) > 5" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "select 'More than 5 levels not allowed'" & vbCrLf
sql = sql & "return" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- if same MO more than once do not allow" & vbCrLf
sql = sql & "declare @ct int, @dupPart varchar(30), @dupNo int" & vbCrLf
sql = sql & "select top 1 @ct = ct, @dupPart = rtrim(MoRef), @dupNo = Mono" & vbCrLf
sql = sql & "from (select count(*) as ct, MoRef,  MoNo from @Levels group by MoRef, Mono having count(*) > 1) x" & vbCrLf
sql = sql & "if @ct is not null" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "select 'MO ' + @DupPart + ' run ' + cast(@DupNo as varchar(10)) + ' cannot appear more than once'" & vbCrLf
sql = sql & "return" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- everything is OK.  Link the MO to its parent" & vbCrLf
sql = sql & "SET NOCOUNT OFF" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update RunsTable set RunParentRunRef = @ParentPart, RunParentRunNo = @ParentNumber" & vbCrLf
sql = sql & "where RUNREF = @ChildPart and RUNNO = @ChildNumber" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @@ROWCOUNT > 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "select '' as result    -- indicate everything is OK" & vbCrLf
sql = sql & "return" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "else" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "select 'unable to link MOs' as result" & vbCrLf
sql = sql & "return" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript True, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function


Private Function UpdateDatabase113()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 192     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "DropStoredProcedureIfExists 'RptArAgingBase'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptArAgingBase]" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Customer varchar(20)   -- blank for all" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* Detail AR Aging Base sp" & vbCrLf
sql = sql & "v2 6/17/2019" & vbCrLf
sql = sql & "take into account canceled invoice indicator" & vbCrLf
sql = sql & "Allow for invoices paid by a different customer" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptArAgingBase '3/1/2019', 'BCA614'" & vbCrLf
sql = sql & "exec RptArAgingBase '3/1/2019', ''" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF OBJECT_ID('tempdb..##TempArAging') IS NOT NULL" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "drop table ##TempArAging" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get all invoices existing but not fully paid on desired date" & vbCrLf
sql = sql & "select rtrim(cust.CUNAME) as [Customer Name], rtrim(inv.INVCUST) as Nickname, INVNO as [Inv #]," & vbCrLf
sql = sql & "cast(INVDATE as DATE) AS [Inv Date], INVTOTAL as [Inv Total], isnull(x.Debits,0) as [Amt Paid]," & vbCrLf
sql = sql & "isnull(x.ct,0) as ct, case when INVCHECKDATE is null then '' else convert(varchar(10),INVCHECKDATE,101) end as [Ck Date]," & vbCrLf
sql = sql & "cast(0 as int) as [Age Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [Amt Due]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [0-30 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [31-60 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [61-90 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [91-120 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [121+ Days]" & vbCrLf
sql = sql & "into ##TempArAging" & vbCrLf
sql = sql & "from CihdTable as inv" & vbCrLf
sql = sql & "join CustTable cust on cust.CUREF = inv.INVCUST" & vbCrLf
sql = sql & "left join (select DCINVNO, isnull(sum(DCDEBIT),0) AS Debits, COUNT(*) AS ct from JritTable" & vbCrLf
sql = sql & "where DCHEAD like 'cr%'" & vbCrLf
sql = sql & "AND DCDATE <= @AsOfDate" & vbCrLf
sql = sql & "and DCINVNO <> 0" & vbCrLf
sql = sql & "and DCDEBIT <> 0" & vbCrLf
sql = sql & "group by DCINVNO) x on x.DCINVNO = inv.INVNO" & vbCrLf
sql = sql & "where INVDATE <= @AsOfDate" & vbCrLf
sql = sql & "and INVCUST like @Customer + '%'" & vbCrLf
sql = sql & "and INVTOTAL <> isnull(x.Debits,0)" & vbCrLf
sql = sql & "and INVCANCELED = 0" & vbCrLf
sql = sql & "and (INVPIF = 0" & vbCrLf
sql = sql & "or isnull(INVCHECKDATE, DATEFROMPARTS(2050,1,1)) > @AsOfDate)" & vbCrLf
sql = sql & "ORDER by INVCUST, INVDATE, INVNO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ##TempArAging set [Amt Due] = [Inv Total] - [Amt Paid]" & vbCrLf
sql = sql & "update ##TempArAging set [Age Days] = DATEDIFF(day,[Inv Date],@AsOfDate)" & vbCrLf
sql = sql & "update ##TempArAging set [0-30 Days] = case when [Age Days] between 0 and 30 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [31-60 Days] = case when [Age Days] between 31 and 60 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [61-90 Days] = case when [Age Days] between 61 and 90 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [91-120 Days] = case when [Age Days] between 91 and 120 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [121+ Days] = case when [Age Days] > 120 then [Amt Due] else 0 end" & vbCrLf
ExecuteScript True, sql

sql = "dropstoredprocedureifexists 'RptArAgingDetail'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptArAgingDetail]" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Customer varchar(20)   -- blank for all" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* Detail AR Aging" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptArAgingDetail '6/17/2019', 'BGSSEA'" & vbCrLf
sql = sql & "exec RptArAgingDetail '3/1/2019', 'BCA614'" & vbCrLf
sql = sql & "exec RptArAgingDetail '3/1/2019', ''" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase @AsofDate, @Customer" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from ##TempArAging order by [Customer Name], [Inv #]" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript True, sql

sql = "DropStoredProcedureIfExists 'RptMultilevelMoCostAnalysis'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptMultilevelMoCostAnalysis]" & vbCrLf
sql = sql & "@MoPart char(30)," & vbCrLf
sql = sql & "@MoNumber int" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* test" & vbCrLf
sql = sql & "RptMultilevelMoCostAnalysis '174570301',7976" & vbCrLf
sql = sql & "select top 100 * from tcittable" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- make a list of MOs at each level" & vbCrLf
sql = sql & ";WITH CTE (Lev, MoPart, MoRun)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "SELECT 1, @MoPart, @MoNumber" & vbCrLf
sql = sql & "UNION ALL" & vbCrLf
sql = sql & "SELECT Lev + 1, RUNREF, RUNNO" & vbCrLf
sql = sql & "FROM RunsTable child" & vbCrLf
sql = sql & "join CTE on child.RunParentRunRef = CTE.MoPart and child.RunParentRunNo = CTE.MoRun" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "SELECT *" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "from CTE" & vbCrLf
sql = sql & "order by Lev, MoPart, MoRun" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "Lev as [Level]," & vbCrLf
sql = sql & "rtrim(pt.PARTNUM) as [MO Part]," & vbCrLf
sql = sql & "MoRun as [MO Run]," & vbCrLf
sql = sql & "'Labor' as [Type]," & vbCrLf
sql = sql & "cast(sum(TCHOURS) as decimal(12,2)) as [Hours]," & vbCrLf
sql = sql & "cast(sum(TCHOURS*TCRATE) as decimal(12,2)) as Labor," & vbCrLf
sql = sql & "cast(sum(TCHOURS*TCOHRATE) as decimal(12,2)) as OH," & vbCrLf
sql = sql & "null as Part," & vbCrLf
sql = sql & "null as [Part Desc]," & vbCrLf
sql = sql & "null as Qty," & vbCrLf
sql = sql & "null as [Unit Cost]," & vbCrLf
sql = sql & "null as [Ext Matl]," & vbCrLf
sql = sql & "null as [Ext Svc]," & vbCrLf
sql = sql & "'' as Vendor" & vbCrLf
sql = sql & "from TcitTable tc" & vbCrLf
sql = sql & "join #temp tmp on tmp.MoPart = tc.TCPARTREF and tmp.MoRun = tc.TCRUNNO" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = tc.TCPARTREF" & vbCrLf
sql = sql & "group by Lev,pt.PARTNUM,MoRun" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now do parts and services" & vbCrLf
sql = sql & "union all" & vbCrLf
sql = sql & "select distinct" & vbCrLf
sql = sql & "Lev as [Level]," & vbCrLf
sql = sql & "rtrim(mopart.PARTNUM) as [MO Part]," & vbCrLf
sql = sql & "MoRun as [MO Run]," & vbCrLf
sql = sql & "case when pt.PALEVEL = 5 then 'Svc' else 'Matl' end as [Type]," & vbCrLf
sql = sql & "null as [Hours]," & vbCrLf
sql = sql & "null as [Labor]," & vbCrLf
sql = sql & "null as [OH]," & vbCrLf
sql = sql & "rtrim(pt.PARTNUM) as Part," & vbCrLf
sql = sql & "rtrim(pt.PADESC) as [Part Desc]," & vbCrLf
sql = sql & "PKAQTY as Qty," & vbCrLf

'SQL = SQL & "INAMT as [Unit Cost]," & vbCrLf
'SQL = SQL & "--cast(INAMT * PKAQTY as decimal(12,2)) as [Ext Cost]," & vbCrLf
'SQL = SQL & "case when pt.PALEVEL <> 5 then cast(INAMT * PKAQTY as decimal(12,2)) else null end as [Ext Matl]," & vbCrLf

sql = sql & "case when INAMT=0 then isnull(poi.PIAMT,0) else INAMT end as [Unit Cost]," & vbCrLf
sql = sql & "case when pt.PALEVEL <> 5 then cast(INAMT * PKAQTY as decimal(12,2)) else null end as [Ext Matl]," & vbCrLf

sql = sql & "case when pt.PALEVEL = 5 then cast(INAMT * PKAQTY as decimal(12,2)) else null end as [Ext Svc]," & vbCrLf
sql = sql & "isnull(rtrim(POVENDOR),'') + ' ' + cast(PINUMBER as varchar(8)) + ' ' + cast(PIITEM as varchar(5)) + PIREV as Vendor" & vbCrLf
sql = sql & "from MopkTable p" & vbCrLf
sql = sql & "join #temp tmp on tmp.MoPart = p.PKMOPART and tmp.MoRun = p.PKMORUN" & vbCrLf
sql = sql & "--join RunsTable r on r.RUNREF = p.PKMOPART and r.RUNNO = p.PKMORUN and r.RUNSTATUS <> 'CA'" & vbCrLf
sql = sql & "join PartTable pt ON PT.PARTREF = p.PKPARTREF" & vbCrLf
sql = sql & "join InvaTable inv on inv.INPART = p.PKPARTREF and inv.INMOPART = p.PKMOPART and inv.INMORUN = p.PKMORUN" & vbCrLf
sql = sql & "   and -inv.INAQTY = p.PKAQTY" & vbCrLf
sql = sql & "left join PoitTable poi on poi.PINUMBER = inv.INPONUMBER and poi.PIITEM = inv.INPOITEM and inv.inporev = poi.PIREV" & vbCrLf
sql = sql & "left join PohdTable ph on ph.PONUMBER = poi.PINUMBER" & vbCrLf
sql = sql & "join PartTable mopart on mopart.PARTREF = tmp.MoPart" & vbCrLf
sql = sql & "where INTYPE = 10" & vbCrLf
sql = sql & "--order by Lev, MoPart, MoRun,[Type]" & vbCrLf
sql = sql & "order by [Level],[MO Part],[MO Run],[Type]" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript True, sql

sql = "DropStoredProcedureIfExists 'RptArAgingBase'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptArAgingBase]" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Customer varchar(20)   -- blank for all" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* Detail AR Aging Base sp" & vbCrLf
sql = sql & "v2 6/17/2019" & vbCrLf
sql = sql & "take into account canceled invoice indicator" & vbCrLf
sql = sql & "Allow for invoices paid by a different customer" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptArAgingBase '6/21/2019', 'ASTK'" & vbCrLf
sql = sql & "SELECT * from ##TempArAging" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase '6/21/2019', ''" & vbCrLf
sql = sql & "SELECT * from ##TempArAging where [amt due] <> [inv total] order by nickname" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase '3/1/2019', ''" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @ArAcct varchar(12)" & vbCrLf
sql = sql & "select @ArAcct = COSJARACCT from ComnTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF OBJECT_ID('tempdb..##TempArAging') IS NOT NULL" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "drop table ##TempArAging" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get all invoices existing but not fully paid on desired date" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "rtrim(cust.CUNAME) as [Customer Name]," & vbCrLf
sql = sql & "rtrim(inv.INVCUST) as Nickname," & vbCrLf
sql = sql & "INVNO as [Inv #]," & vbCrLf
sql = sql & "INVTYPE as TP," & vbCrLf
sql = sql & "cast(INVDATE as DATE) AS [Inv Date]," & vbCrLf
sql = sql & "INVTOTAL as [Inv Total]," & vbCrLf
sql = sql & "isnull(x.Debits,0) as Debits," & vbCrLf
sql = sql & "isnull(x.credits,0) as Credits," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [Amt Paid]," & vbCrLf
sql = sql & "--case when INVTOTAL < 0 THEN INVTOTAL + isnull(x.debits,0) else isnull(x.debits,0) end as [Amt Paid]," & vbCrLf
sql = sql & "isnull(x.ct,0) as ct," & vbCrLf
sql = sql & "case when INVCHECKDATE is null then '' else convert(varchar(10),INVCHECKDATE,101) end as [Ck Date]," & vbCrLf
sql = sql & "cast(0 as int) as [Age Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [Amt Due]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [0-30 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [31-60 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [61-90 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [91-120 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [121+ Days]" & vbCrLf
sql = sql & "into ##TempArAging" & vbCrLf
sql = sql & "from CihdTable as inv" & vbCrLf
sql = sql & "join CustTable cust on cust.CUREF = inv.INVCUST" & vbCrLf
sql = sql & "left join (select DCINVNO," & vbCrLf
sql = sql & "isnull(sum(DCDEBIT),0) AS Debits," & vbCrLf
sql = sql & "isnull(sum(DCCREDIT),0) AS Credits," & vbCrLf
sql = sql & "COUNT(*) AS ct from JritTable" & vbCrLf
sql = sql & "where DCHEAD like 'cr%'" & vbCrLf
sql = sql & "and DCACCTNO = @ArAcct" & vbCrLf
sql = sql & "AND DCDATE <= @AsOfDate" & vbCrLf
sql = sql & "and DCINVNO <> 0" & vbCrLf
sql = sql & "--and DCDEBIT <> 0" & vbCrLf
sql = sql & "group by DCINVNO) x on x.DCINVNO = inv.INVNO" & vbCrLf
sql = sql & "where INVDATE <= @AsOfDate" & vbCrLf
sql = sql & "and INVCUST like @Customer + '%'" & vbCrLf
sql = sql & "and INVTOTAL <> isnull(x.Debits,0)" & vbCrLf
sql = sql & "and INVCANCELED = 0" & vbCrLf
sql = sql & "and (INVPIF = 0" & vbCrLf
sql = sql & "or isnull(INVCHECKDATE, DATEFROMPARTS(2050,1,1)) > @AsOfDate)" & vbCrLf
sql = sql & "ORDER by INVCUST, INVDATE, INVNO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ##TempArAging set [Amt Paid] = sign([Inv Total]) * (Credits -Debits)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ##TempArAging set [Amt Due] = [Inv Total] - [Amt Paid] where [INV TOTAL] >= 0" & vbCrLf
sql = sql & "--update ##TempArAging set [Amt Due] = [Amt Paid] - [Inv Total] where [INV TOTAL] < 0" & vbCrLf
sql = sql & "update ##TempArAging set [Amt Due] = [Amt Paid] where [INV TOTAL] < 0" & vbCrLf
sql = sql & "update ##TempArAging set [Age Days] = DATEDIFF(day,[Inv Date],@AsOfDate)" & vbCrLf
sql = sql & "update ##TempArAging set [0-30 Days] = case when [Age Days] between 0 and 30 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [31-60 Days] = case when [Age Days] between 31 and 60 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [61-90 Days] = case when [Age Days] between 61 and 90 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [91-120 Days] = case when [Age Days] between 91 and 120 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [121+ Days] = case when [Age Days] > 120 then [Amt Due] else 0 end" & vbCrLf
ExecuteScript True, sql

sql = "dropstoredprocedureifexists 'RptArAgingDetail'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptArAgingDetail]" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Customer varchar(20)   -- blank for all" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* Detail AR Aging" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptArAgingDetail '6/17/2019', 'BGSSEA'" & vbCrLf
sql = sql & "exec RptArAgingDetail '3/1/2019', 'BCA614'" & vbCrLf
sql = sql & "exec RptArAgingDetail '3/1/2019', ''" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase @AsofDate, @Customer" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select * from ##TempArAging order by [Customer Name], Nickname, [Inv #]" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript True, sql

sql = "dropstoredprocedureifexists 'RptArAgingSummary'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure [dbo].[RptArAgingSummary]" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Customer varchar(20)   -- blank for all" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* test" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptArAgingSummary '6/21/2019', 'ASTK'" & vbCrLf
sql = sql & "exec RptArAgingSummary '6/21/2019', ''" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase @AsOfDate, @Customer" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select [Customer Name], Nickname, count(*) as INVOICES, sum([Amt Due]) as Total," & vbCrLf
sql = sql & "sum([0-30 Days]) as [0-30 Days]," & vbCrLf
sql = sql & "sum([31-60 Days]) as [31-60 Days]," & vbCrLf
sql = sql & "sum([61-90 Days]) as [61-90 Days]," & vbCrLf
sql = sql & "sum([91-120 Days]) as [91-120 Days]," & vbCrLf
sql = sql & "sum([121+ Days]) as [121+ Days]" & vbCrLf
sql = sql & "from ##TempArAging" & vbCrLf
sql = sql & "group by [Customer Name],[NickName]" & vbCrLf
sql = sql & "order by [Customer Name]" & vbCrLf
ExecuteScript True, sql

sql = "DropStoredProcedureIfExists 'RptApStatement'" & vbCrLf
ExecuteScript True, sql


sql = "create procedure [dbo].[RptApStatement]" & vbCrLf
sql = sql & "@Vendor varchar(50),      -- blank for all" & vbCrLf
sql = sql & "@StartDate date," & vbCrLf
sql = sql & "@EndDate date," & vbCrLf
sql = sql & "@IncludePaidInvoices bit  -- = 1 to show paid invoices" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "/* generate vendor statements" & vbCrLf
sql = sql & "Created 11/12/2018 TEL" & vbCrLf
sql = sql & "Updated 8/1/2019 TEL Show all checks regardless of date per USC request" & vbCrLf
sql = sql & "Also, does not repeat inv amt when there are multiple partial payments" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "exec RptApStatement 'ALL', '1/1/2019', '3/31/2019', 0" & vbCrLf
sql = sql & "exec RptApStatement 'ACTMED', '1/1/2019', '3/31/2019', 1" & vbCrLf
sql = sql & "exec RptApStatement 'ADVTEC', '5/23/2019', '5/23/2019', 0" & vbCrLf
sql = sql & "exec RptApStatement 'RSTAHL', '1/1/2017', '11/30/2018', 1  --RSTAHL has 3 checks for in-119255" & vbCrLf
sql = sql & "exec RptApStatement 'GRAYBAR', '10/1/2017', '11/30/2018', 1" & vbCrLf
sql = sql & "exec RptApStatement 'GRAYBAR', '6/1/18', '2/1/2019', 1" & vbCrLf
sql = sql & "exec RptApStatement 'TEMPLA', '6/12/19', '6/12/2019', 1" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "IF OBJECT_ID('tempdb..#temp') IS NOT NULL" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get ap account" & vbCrLf
sql = sql & "declare @ApAcct varchar(12)" & vbCrLf
sql = sql & "select @ApAcct = COAPACCT from ComnTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get all invoice/check combinations in date range" & vbCrLf
sql = sql & "SELECT VEREF as Vendor, VEBNAME as [Vendor Name], VIDATE [Inv Date], VINO as [Inv #]," & vbCrLf
sql = sql & "VIDUE as [Inv Amt]," & vbCrLf
sql = sql & "isnull(MAX(DCHEAD),'') as Journal," & vbCrLf
sql = sql & "isnull(CHKACCT,'') [Check Acct]," & vbCrLf
sql = sql & "isnull(CHKNUMBER,'') as [Check #]," & vbCrLf
sql = sql & "isnull(CONVERT(varchar(10),CHKPOSTDATE,101),'') as [Check Date]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as Discount," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as Voucher," & vbCrLf
sql = sql & "sum(x.ApApplied) as ApApplied," & vbCrLf
sql = sql & "ROW_NUMBER() over ( partition by veref, VINO order by veref,vidate,vino) as Row" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "FROM VndrTable v" & vbCrLf
sql = sql & "JOIN VihdTable inv on v.VEREF = inv.VIVENDOR" & vbCrLf
sql = sql & "left join" & vbCrLf
sql = sql & "(select (dcdebit-dccredit) as ApApplied, DCVENDOR, DCVENDORINV, CHKACCT, CHKNUMBER, CHKPOSTDATE, DCHEAD" & vbCrLf
sql = sql & "from JritTable jr" & vbCrLf
sql = sql & "join ChksTable ck on ck.CHKVENDOR = jr.DCVENDOR and ck.CHKNUMBER = jr.DCCHECKNO and ck.CHKACCT = jr.DCCHKACCT" & vbCrLf
sql = sql & "and ck.CHKVOID = 0 and DCACCTNO = @ApAcct and (DCHEAD like 'cc%' or DCHEAD like 'xc%')) x" & vbCrLf
sql = sql & "on x.DCVENDOR = inv.VIVENDOR and x.DCVENDORINV = inv.VINO" & vbCrLf
sql = sql & "WHERE VIDATE between @StartDate and @EndDate" & vbCrLf
sql = sql & "and (RTRIM(@Vendor) = 'ALL' or VIVENDOR = @Vendor)" & vbCrLf
sql = sql & "group by VEREF, VEBNAME, VINO, VIDATE, VIDUE, CHKACCT, CHKNUMBER, CHKPOSTDATE" & vbCrLf
sql = sql & "order by VEREF,VINO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- now add discounts taken" & vbCrLf
sql = sql & "update #temp set Discount = isnull((select sum(dccredit - dcdebit)" & vbCrLf
sql = sql & "from JritTable dc where dc.DCVENDOR = #temp.Vendor and dc.DCVENDORINV = #temp.[Inv #]" & vbCrLf
sql = sql & "and dc.DCCHKACCT = #temp.[Check Acct] and dc.DCCHECKNO = #temp.[Check #] and dc.DCREF = 3),0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- add voucher amount" & vbCrLf
sql = sql & "update #temp set Voucher = isnull((select sum(dccredit - dcdebit)" & vbCrLf
sql = sql & "from JritTable dc where dc.DCVENDOR = #temp.Vendor and dc.DCVENDORINV = #temp.[Inv #]" & vbCrLf
sql = sql & "and dc.DCCHKACCT = #temp.[Check Acct] and dc.DCCHECKNO = #temp.[Check #] and dc.DCREF = 2" & vbCrLf
sql = sql & "--and DCDATE <= @EndDate)" & vbCrLf
sql = sql & "),0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get journal ID" & vbCrLf
sql = sql & "--update #temp set Journal = (select top 1 DCHEAD" & vbCrLf
sql = sql & "--from JritTable dc where dc.DCVENDOR = #temp.Vendor and dc.DCVENDORINV = #temp.[Inv #]" & vbCrLf
sql = sql & "--and dc.DCCHKACCT = #temp.[Check Acct] and dc.DCCHECKNO = #temp.[Check #] and dc.DCREF = 2" & vbCrLf
sql = sql & "----and DCDATE <= @EndDate)" & vbCrLf
sql = sql & "--)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- delete paid in full invoices if not requested" & vbCrLf
sql = sql & "if @IncludePaidInvoices = 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "update #temp set ApApplied = isnull((select sum(Discount + Voucher)" & vbCrLf
sql = sql & "from #temp t2 where t2.Vendor = #temp.Vendor and t2.[Inv #] = #temp.[Inv #]),0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "delete from #temp where ApApplied = [Inv Amt]" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select Vendor, [Vendor Name],[Inv Date], [Inv #]," & vbCrLf
sql = sql & "case when Row = 1 then [Inv Amt] else 0 end as [Inv Amt]," & vbCrLf
sql = sql & "Journal, [Check Acct], [Check #], [Check Date], Discount, Voucher" & vbCrLf
sql = sql & "from #temp order by Vendor, [Inv Date], [Inv #]" & vbCrLf
sql = sql & "drop table #temp" & vbCrLf
ExecuteScript True, sql



''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function


Private Function UpdateDatabase193()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 193     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

' option to show names in POM ops rather than OP #
sql = "AddOrUpdateColumn 'ComnTable', 'CoUseNamesInPOM', 'tinyint not null default 0'" & vbCrLf
ExecuteScript True, sql

' VALMAC 12 Month Backlog SSRS
sql = "IF OBJECT_ID('RptBacklog12Month') IS NOT NULL" & vbCrLf
sql = sql & "DROP PROCEDURE RptBacklog12Month" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptBacklog12Month" & vbCrLf
sql = sql & "@StartDate date" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "12 month Backlog by Customer" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "RptBacklog12Month '8/1/2019'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "declare @start date = dateadd(day,-datepart(day,@StartDate) + 1,@StartDate)" & vbCrLf
sql = sql & "declare @end date = dateadd(year,1,@start)" & vbCrLf
sql = sql & "set @end = dateadd(day,-1,@end)" & vbCrLf
sql = sql & "SELECT SOCUST as Customer," & vbCrLf
sql = sql & "case when ITSCHED < @Start then ' Prior' else convert(varchar(7),ITSCHED,102) end as Period," & vbCrLf
sql = sql & "cast(SUM(ITQTY * ITDOLLARS) as decimal(12,2)) as Backlog" & vbCrLf
sql = sql & "FROM SoitTable,SohdTable" & vbCrLf
sql = sql & "WHERE ITSO=SONUMBER AND ITCANCELED=0 AND" & vbCrLf
sql = sql & "ITPSNUMBER='' AND ITPSSHIPPED=0 AND ITINVOICE=0" & vbCrLf
sql = sql & "--AND ITSCHED BETWEEN @start and @end" & vbCrLf
sql = sql & "AND ITSCHED <= @end" & vbCrLf
sql = sql & "group by SOCUST, case when ITSCHED < @Start then ' Prior' else convert(varchar(7),ITSCHED,102) end" & vbCrLf
sql = sql & "order by SOCUST, case when ITSCHED < @Start then ' Prior' else convert(varchar(7),ITSCHED,102) end" & vbCrLf
ExecuteScript True, sql



' VALMAC MinMax SSRS Report
sql = "if exists (select * from sys.procedures where name = 'MinMaxCalculate')" & vbCrLf
sql = sql & "drop procedure MinMaxCalculate" & vbCrLf
ExecuteScript True, sql

sql = "create procedure MinMaxCalculate" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "-- assign row numbers to insure order of items will always be the same" & vbCrLf
sql = sql & ";WITH a AS(" & vbCrLf
sql = sql & "SELECT ROW_NUMBER() OVER(ORDER BY PartNumber, ActivityDate, SortOrder) as rn, Row" & vbCrLf
sql = sql & "FROM MinMax" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "UPDATE a SET Row=rn" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create running totals" & vbCrLf
sql = sql & ";with rt as (" & vbCrLf
sql = sql & "select Row," & vbCrLf
sql = sql & "[Cust Qty], [Our Qty]," & vbCrLf
sql = sql & "SUM ([Cust Activity]) OVER (Partition by PartNumber ORDER BY Row) AS CustQty," & vbCrLf
sql = sql & "SUM ([Our Activity]) OVER (Partition by PartNumber ORDER BY Row) AS OurQty" & vbCrLf
sql = sql & "from MinMax" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "update rt set [Cust Qty] = CustQty, [Our Qty] = OurQty" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript True, sql


sql = "if exists (select * from sys.procedures where name = 'MinMaxActions')" & vbCrLf
sql = sql & "drop procedure MinMaxActions" & vbCrLf
ExecuteScript True, sql

sql = "create procedure MinMaxActions" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if exists (select * from INFORMATION_SCHEMA.tables where TABLE_NAME = 'MinMax')" & vbCrLf
sql = sql & "drop table MinMax" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CREATE TABLE [dbo].[MinMax](" & vbCrLf
sql = sql & "[PartNumber] [varchar](30) NOT NULL," & vbCrLf
sql = sql & "[SortOrder] [int] NOT NULL," & vbCrLf
sql = sql & "[ActivityDate] [date] NOT NULL," & vbCrLf
sql = sql & "[Run #] int," & vbCrLf
sql = sql & "[Cust Activity] [int] NOT NULL," & vbCrLf
sql = sql & "[Our Activity] [int] NOT NULL," & vbCrLf
sql = sql & "[Type] [varchar](20) NOT NULL," & vbCrLf
sql = sql & "[Notes] [varchar](60) NOT NULL," & vbCrLf
sql = sql & "[Min] [int] NOT NULL," & vbCrLf
sql = sql & "[Max] [int] NOT NULL," & vbCrLf
sql = sql & "Row int," & vbCrLf
sql = sql & "[Cust Qty] [int] NULL," & vbCrLf
sql = sql & "[Our Qty] [int] NULL," & vbCrLf
sql = sql & "Action varchar(50)" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Cust Activity]  DEFAULT ((0)) FOR [Cust Activity]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Our Activity]  DEFAULT ((0)) FOR [Our Activity]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Cust Qty]  DEFAULT ((0)) FOR [Cust Qty]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Our Qty]  DEFAULT ((0)) FOR [Our Qty]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Notes]  DEFAULT ('') FOR [Notes]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Min]  DEFAULT ((0)) FOR [Min]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Max]  DEFAULT ((0)) FOR [Max]" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- make sure imported data is proper type" & vbCrLf
sql = sql & "alter table MinMax_Orders alter column [Forecast Date] date" & vbCrLf
sql = sql & "alter table MinMax_Orders alter column Requirement int" & vbCrLf
sql = sql & "alter table MinMax_Orders alter column [On Hand] int" & vbCrLf
sql = sql & "alter table MinMax_Orders alter column Min int" & vbCrLf
sql = sql & "alter table MinMax_Orders alter column Max int" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert Customer Start Quantities - use earliest date in Customer data" & vbCrLf
sql = sql & "declare @startDate date = (select min([Forecast Date]) from MinMax_Orders)" & vbCrLf
sql = sql & "INSERT INTO MinMax (PartNumber,SortOrder,ActivityDate,[Cust Activity],Min,Max,Type)" & vbCrLf
sql = sql & "SELECT distinct [Part Number],10,@startDate,[On Hand],Min,Max,'Start QOH'" & vbCrLf
sql = sql & "FROM dbo.MinMax_Orders" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- add our start qtys" & vbCrLf
sql = sql & "update mm set [Our Activity] = PAQOH" & vbCrLf
sql = sql & "from PartTable p" & vbCrLf
sql = sql & "join MinMax mm on dbo.fnCompress(mm.PartNumber) = p.PARTREF" & vbCrLf
sql = sql & "where mm.SortOrder = 10" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert Customer Forecasts" & vbCrLf
sql = sql & "INSERT INTO MinMax (PartNumber,SortOrder,ActivityDate,[Cust Activity],[Our Activity], Min,Max,Type,Notes)" & vbCrLf
sql = sql & "SELECT [Part Number],50,[Forecast Date],-cast(Requirement as int),0,Min,Max,'Forecast','PO ' + PO" & vbCrLf
sql = sql & "FROM dbo.MinMax_Orders" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert MRP scheduled completions" & vbCrLf
sql = sql & "INSERT INTO MinMax (PartNumber,SortOrder,ActivityDate,[Cust Activity],[Our Activity],Type,[Run #], Notes)" & vbCrLf
sql = sql & "select rtrim(MRP_PARTNUM), 20, MRP_PARTDATERQD, 0, MRP_PARTQTYRQD, 'Current MO',MRP_MORUNNO ," & vbCrLf
sql = sql & "'MO ' + rtrim(MRP_PARTNUM) + ' Run # ' +cast(MRP_MORUNNO as varchar(5))" & vbCrLf
sql = sql & "from MrplTable" & vbCrLf
sql = sql & "where MRP_TYPE in (3,4)" & vbCrLf
sql = sql & "and MRP_PARTREF in (select distinct dbo.fnCompress(PartNumber) from MinMax)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update MinMax set PartNumber = rtrim(PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec MinMaxCalculate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- Insert MinMax shipments until no more are required" & vbCrLf
sql = sql & "declare @loopCount int = 1" & vbCrLf
sql = sql & "while @loopCount <= 50" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "declare @PartsBelowMin table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "PartNumber varchar(30)," & vbCrLf
sql = sql & "Row int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "delete from @PartsBelowMin   -- old insertions remain in loop" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @PartsBelowMin" & vbCrLf
sql = sql & "select PartNumber," & vbCrLf
sql = sql & "min(Row) as Row" & vbCrLf
sql = sql & "from MinMax" & vbCrLf
sql = sql & "where [Cust Qty] < Min" & vbCrLf
sql = sql & "and (@loopCount = 1 or Type <> 'Start QOH')    -- prevents problem if start qoh < min" & vbCrLf
sql = sql & "Group by PartNumber" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if (select count(*) from @PartsBelowMin) <= 0" & vbCrLf
sql = sql & "break;" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO MinMax (PartNumber,SortOrder,ActivityDate,[Cust Activity],[Our Activity],Min,Max,Type,Action)" & vbCrLf
sql = sql & "select mm.PartNumber, 40, ActivityDate, max - [Cust Qty] - 1,-(max - [Cust Qty]-1),Min,Max,'Shipment'," & vbCrLf
sql = sql & "'Ship ' + cast(max - [Cust Qty] - 1 as varchar(10))" & vbCrLf
sql = sql & "from MinMax mm" & vbCrLf
sql = sql & "join @PartsBelowMin blw on blw.PartNumber = mm.PartNumber and blw.Row = mm.Row" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec MinMaxCalculate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from MinMax  where PartNumber = '111N1028-6' order by row" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @loopCount = @loopCount + 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- find first negative [Our Qty] for date and most negative [Our Qty] for each part" & vbCrLf
sql = sql & "declare @Actions table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "PartNumber varchar(30)," & vbCrLf
sql = sql & "ActionDate date," & vbCrLf
sql = sql & "MODate date," & vbCrLf
sql = sql & "MOOrigDate date," & vbCrLf
sql = sql & "MinQty int," & vbCrLf
sql = sql & "FirstMinusRow int," & vbCrLf
sql = sql & "--FirstMinusRowDate date," & vbCrLf
sql = sql & "EndRow int," & vbCrLf
sql = sql & "OurEndQty int," & vbCrLf
sql = sql & "CustEndQty int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "insert @Actions (PartNumber)" & vbCrLf
sql = sql & "select distinct PartNumber from MinMax where [Our Qty] < 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set ActionDate = (select Min(ActivityDate) from MinMax" & vbCrLf
sql = sql & "where PartNumber = [@Actions].PartNumber and [Our Qty] < 0" & vbCrLf
sql = sql & "group by PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions set MOOrigDate = ActionDate" & vbCrLf
sql = sql & "update @Actions set MODate = dateadd(day,-30,ActionDate)" & vbCrLf
sql = sql & "update @Actions set MODate = @startDate where MODate < @startDate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set MinQty = (select Min([Our Qty]) from MinMax" & vbCrLf
sql = sql & "where PartNumber = [@Actions].PartNumber and [Our Qty] < 0" & vbCrLf
sql = sql & "group by PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set EndRow = (select Max(Row) from MinMax" & vbCrLf
sql = sql & "where [@Actions].PartNumber = MinMax.PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set FirstMinusRow = (select Min(Row) from MinMax" & vbCrLf
sql = sql & "where [@Actions].PartNumber = MinMax.PartNumber" & vbCrLf
sql = sql & "and [Our Qty] < 0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set OurEndQty = (select [Our Qty] from MinMax" & vbCrLf
sql = sql & "where [@Actions].EndRow = MinMax.Row)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set CustEndQty = (select [Cust Qty] from MinMax" & vbCrLf
sql = sql & "where [@Actions].EndRow = MinMax.Row)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--update m" & vbCrLf
sql = sql & "--set Action = 'Reschedule MO from ' + convert(varchar(10), [@Actions].MODate,101)" & vbCrLf
sql = sql & "--from MinMax m" & vbCrLf
sql = sql & "--join @Actions on [@Actions].PartNumber = m.PartNumber" & vbCrLf
sql = sql & "--where m.SortOrder = 20" & vbCrLf
sql = sql & "--and m.Row >= [@Actions].FirstMinusRow" & vbCrLf
sql = sql & "--select *" & vbCrLf
sql = sql & "--from MinMax m" & vbCrLf
sql = sql & "--join @Actions on [@Actions].PartNumber = m.PartNumber" & vbCrLf
sql = sql & "--where m.SortOrder = 20" & vbCrLf
sql = sql & "--and m.ActivityDate > MOOrigDate" & vbCrLf
sql = sql & "--and m.Row >= [@Actions].FirstMinusRow" & vbCrLf
sql = sql & "--and m.PartNumber = '112W9721-6' order by row" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update m" & vbCrLf
sql = sql & "set Action = 'Reschedule MO from ' + convert(varchar(10), m.ActivityDate,101)," & vbCrLf
sql = sql & "ActivityDate = MOOrigDate" & vbCrLf
sql = sql & "from MinMax m" & vbCrLf
sql = sql & "join @Actions on [@Actions].PartNumber = m.PartNumber" & vbCrLf
sql = sql & "where m.SortOrder = 20" & vbCrLf
sql = sql & "and m.ActivityDate > MOOrigDate" & vbCrLf
sql = sql & "and m.Row >= [@Actions].FirstMinusRow" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec MinMaxCalculate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- add action items for new MOs" & vbCrLf
sql = sql & "delete from @Actions" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @Actions (PartNumber)" & vbCrLf
sql = sql & "select distinct PartNumber from MinMax where [Our Qty] < 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set ActionDate = (select Min(ActivityDate) from MinMax" & vbCrLf
sql = sql & "where PartNumber = [@Actions].PartNumber and [Our Qty] < 0" & vbCrLf
sql = sql & "group by PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions set MOOrigDate = ActionDate" & vbCrLf
sql = sql & "update @Actions set MODate = dateadd(day,-30,ActionDate)" & vbCrLf
sql = sql & "update @Actions set MODate = @startDate where MODate < @startDate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set MinQty = (select Min([Our Qty]) from MinMax" & vbCrLf
sql = sql & "where PartNumber = [@Actions].PartNumber and [Our Qty] < 0" & vbCrLf
sql = sql & "group by PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set EndRow = (select Max(Row) from MinMax" & vbCrLf
sql = sql & "where [@Actions].PartNumber = MinMax.PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set FirstMinusRow = (select Min(Row) from MinMax" & vbCrLf
sql = sql & "where [@Actions].PartNumber = MinMax.PartNumber" & vbCrLf
sql = sql & "and [Our Qty] < 0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set OurEndQty = (select [Our Qty] from MinMax" & vbCrLf
sql = sql & "where [@Actions].EndRow = MinMax.Row)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set CustEndQty = (select [Cust Qty] from MinMax" & vbCrLf
sql = sql & "where [@Actions].EndRow = MinMax.Row)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO MinMax (PartNumber,SortOrder,ActivityDate,[Cust Activity],[Our Activity],Type,Action)" & vbCrLf
sql = sql & "select PartNumber, 20, MODate, 0, - OurEndQty, 'New MO'," & vbCrLf
sql = sql & "'Schedule MO for qty ' + cast(-OurEndQty as varchar(5)) + ' on ' + convert(varchar(10), [@Actions].MODate,101)" & vbCrLf
sql = sql & "from @Actions" & vbCrLf
sql = sql & "where OurEndQty < 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec MinMaxCalculate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update MinMax set Action = '' where Action is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from MinMax where PartNumber in" & vbCrLf
sql = sql & "--(select PartNumber from MinMax where [Our Qty] < 0)" & vbCrLf
sql = sql & "--order by Row" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select PartNumber as [Part Number]" & vbCrLf
sql = sql & ",ActivityDate as [Activity Date]" & vbCrLf
sql = sql & "--,SortOrder" & vbCrLf
sql = sql & ",[Cust Activity]" & vbCrLf
sql = sql & ",[Our Activity]" & vbCrLf
sql = sql & ",[Type]" & vbCrLf
sql = sql & ",[Min]" & vbCrLf
sql = sql & ",[Max]" & vbCrLf
sql = sql & ",[Cust Qty]" & vbCrLf
sql = sql & ",[Our Qty]" & vbCrLf
sql = sql & ",[Row]" & vbCrLf
sql = sql & ",[Notes]" & vbCrLf
sql = sql & ",[Action]" & vbCrLf
sql = sql & "from MinMax" & vbCrLf
sql = sql & "--where PartNumber = '112W9721-6'" & vbCrLf
sql = sql & "order by Row" & vbCrLf
ExecuteScript True, sql



''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase194()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 194     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

' VALMAC Min/Max v 5
' 9/27/19 two sell price columns added

sql = "if exists (select * from sys.procedures where name = 'MinMaxCalculate')" & vbCrLf
sql = sql & "drop procedure MinMaxCalculate" & vbCrLf
ExecuteScript True, sql

sql = "create procedure MinMaxCalculate" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "-- assign row numbers to insure order of items will always be the same" & vbCrLf
sql = sql & ";WITH a AS(" & vbCrLf
sql = sql & "SELECT ROW_NUMBER() OVER(ORDER BY PartNumber, ActivityDate, SortOrder) as rn, Row" & vbCrLf
sql = sql & "FROM MinMax" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "UPDATE a SET Row=rn" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- create running totals" & vbCrLf
sql = sql & ";with rt as (" & vbCrLf
sql = sql & "select Row," & vbCrLf
sql = sql & "[Cust Qty], [Our Qty]," & vbCrLf
sql = sql & "SUM ([Cust Activity]) OVER (Partition by PartNumber ORDER BY Row) AS CustQty," & vbCrLf
sql = sql & "SUM ([Our Activity]) OVER (Partition by PartNumber ORDER BY Row) AS OurQty" & vbCrLf
sql = sql & "from MinMax" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "update rt set [Cust Qty] = CustQty, [Our Qty] = OurQty" & vbCrLf
sql = sql & "end" & vbCrLf
ExecuteScript True, sql


sql = "if exists (select * from sys.procedures where name = 'MinMaxActions')" & vbCrLf
sql = sql & "drop procedure MinMaxActions" & vbCrLf
ExecuteScript True, sql

sql = "create procedure MinMaxActions" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "v 5 - 11/21/2019" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if exists (select * from INFORMATION_SCHEMA.tables where TABLE_NAME = 'MinMax')" & vbCrLf
sql = sql & "drop table MinMax" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CREATE TABLE [dbo].[MinMax](" & vbCrLf
sql = sql & "[PartNumber] [varchar](30) NOT NULL," & vbCrLf
sql = sql & "[Part Ref] [varchar](30) NOT NULL," & vbCrLf
sql = sql & "[SortOrder] [int] NOT NULL," & vbCrLf
sql = sql & "[ActivityDate] [date] NOT NULL," & vbCrLf
sql = sql & "[Run #] int," & vbCrLf
sql = sql & "[Cust Activity] [int] NOT NULL," & vbCrLf
sql = sql & "[Our Activity] [int] NOT NULL," & vbCrLf
sql = sql & "[Type] [varchar](20) NOT NULL," & vbCrLf
sql = sql & "[Notes] [varchar](60) NOT NULL," & vbCrLf
sql = sql & "[Min] [int] NOT NULL," & vbCrLf
sql = sql & "[Max] [int] NOT NULL," & vbCrLf
sql = sql & "Row int," & vbCrLf
sql = sql & "[Cust Qty] [int] NULL," & vbCrLf
sql = sql & "[Our Qty] [int] NULL," & vbCrLf
sql = sql & "Action varchar(50)," & vbCrLf
sql = sql & "[Selling Price] decimal(12,2) not null default 0," & vbCrLf
sql = sql & "Total decimal(12,2) not null default 0," & vbCrLf
sql = sql & "INDEX IX_MinMax_PartRef NONCLUSTERED ([Part Ref])," & vbCrLf
sql = sql & "INDEX IX_Minmax_PartNum NONCLUSTERED (PartNumber)" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Cust Activity]  DEFAULT ((0)) FOR [Cust Activity]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Our Activity]  DEFAULT ((0)) FOR [Our Activity]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Cust Qty]  DEFAULT ((0)) FOR [Cust Qty]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Our Qty]  DEFAULT ((0)) FOR [Our Qty]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Notes]  DEFAULT ('') FOR [Notes]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Min]  DEFAULT ((0)) FOR [Min]" & vbCrLf
sql = sql & "ALTER TABLE [dbo].[MinMax] ADD  CONSTRAINT [DF_MinMax_Max]  DEFAULT ((0)) FOR [Max]" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- make sure imported data is proper type" & vbCrLf
sql = sql & "alter table MinMax_Orders alter column [Forecast Date] date" & vbCrLf
sql = sql & "alter table MinMax_Orders alter column Requirement int" & vbCrLf
sql = sql & "alter table MinMax_Orders alter column [On Hand] int" & vbCrLf
sql = sql & "alter table MinMax_Orders alter column Min int" & vbCrLf
sql = sql & "alter table MinMax_Orders alter column Max int" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert Customer Start Quantities - use earliest date in Customer data" & vbCrLf
sql = sql & "declare @startDate date = (select min([Forecast Date]) from MinMax_Orders)" & vbCrLf
sql = sql & "INSERT INTO MinMax (PartNumber,[Part Ref],SortOrder,ActivityDate,[Cust Activity],Min,Max,Type)" & vbCrLf
sql = sql & "SELECT distinct [Part Number],dbo.fnCompress([Part Number]),10,@startDate,[On Hand],Min,Max,'Start QOH'" & vbCrLf
sql = sql & "FROM dbo.MinMax_Orders" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- add our start qtys" & vbCrLf
sql = sql & "update mm set [Our Activity] = PAQOH" & vbCrLf
sql = sql & "from PartTable p" & vbCrLf
sql = sql & "join MinMax mm on mm.[Part Ref] = p.PARTREF" & vbCrLf
sql = sql & "where mm.SortOrder = 10" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert Customer Forecasts" & vbCrLf
sql = sql & "INSERT INTO MinMax (PartNumber,[Part Ref], SortOrder,ActivityDate,[Cust Activity],[Our Activity], Min,Max,Type,Notes)" & vbCrLf
sql = sql & "SELECT [Part Number],dbo.fnCompress([Part Number]),50,[Forecast Date],-cast(Requirement as int),0,Min,Max,'Forecast','PO ' + PO" & vbCrLf
sql = sql & "FROM dbo.MinMax_Orders" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert MRP scheduled completions" & vbCrLf
sql = sql & "INSERT INTO MinMax (PartNumber, [Part Ref], SortOrder,ActivityDate,[Cust Activity],[Our Activity],Type,[Run #], Notes)" & vbCrLf
sql = sql & "select rtrim(MRP_PARTNUM), rtrim(MRP_PARTREF), 20, MRP_PARTDATERQD, 0, MRP_PARTQTYRQD, 'Current MO',MRP_MORUNNO ," & vbCrLf
sql = sql & "'MO ' + rtrim(MRP_PARTNUM) + ' Run # ' +cast(MRP_MORUNNO as varchar(5))" & vbCrLf
sql = sql & "from MrplTable" & vbCrLf
sql = sql & "where MRP_TYPE in (3,4)" & vbCrLf
sql = sql & "and MRP_PARTREF in (select distinct [Part Ref] from MinMax)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update MinMax set PartNumber = rtrim(PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec MinMaxCalculate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- Insert MinMax shipments until no more are required" & vbCrLf
sql = sql & "declare @loopCount int = 1" & vbCrLf
sql = sql & "while @loopCount <= 50" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "declare @PartsBelowMin table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "PartNumber varchar(30)," & vbCrLf
sql = sql & "Row int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "delete from @PartsBelowMin   -- old insertions remain in loop" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @PartsBelowMin" & vbCrLf
sql = sql & "select PartNumber," & vbCrLf
sql = sql & "min(Row) as Row" & vbCrLf
sql = sql & "from MinMax" & vbCrLf
sql = sql & "where [Cust Qty] < Min" & vbCrLf
sql = sql & "and (@loopCount = 1 or Type <> 'Start QOH')    -- prevents problem if start qoh < min" & vbCrLf
sql = sql & "Group by PartNumber" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if (select count(*) from @PartsBelowMin) <= 0" & vbCrLf
sql = sql & "break;" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "INSERT INTO MinMax (PartNumber,[Part Ref],SortOrder,ActivityDate,[Cust Activity],[Our Activity],Min,Max,Type,Action,[Selling Price],Total)" & vbCrLf
sql = sql & "select mm.PartNumber, mm.[Part Ref], 40, ActivityDate, max - [Cust Qty] - 1,-(max - [Cust Qty]-1),Min,Max,'Shipment'," & vbCrLf
sql = sql & "'Ship ' + cast(max - [Cust Qty] - 1 as varchar(10)), pt.PAPRICE, pt.paprice * (max - [Cust Qty]-1)" & vbCrLf
sql = sql & "from MinMax mm" & vbCrLf
sql = sql & "join @PartsBelowMin blw on blw.PartNumber = mm.PartNumber and blw.Row = mm.Row" & vbCrLf
sql = sql & "join PartTable pt on pt.PARTREF = mm.[Part Ref]" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec MinMaxCalculate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from MinMax  where PartNumber = '111N1028-6' order by row" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @loopCount = @loopCount + 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- find first negative [Our Qty] for date and most negative [Our Qty] for each part" & vbCrLf
sql = sql & "declare @Actions table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "PartNumber varchar(30)," & vbCrLf
sql = sql & "ActionDate date," & vbCrLf
sql = sql & "MODate date," & vbCrLf
sql = sql & "MOOrigDate date," & vbCrLf
sql = sql & "MinQty int," & vbCrLf
sql = sql & "FirstMinusRow int," & vbCrLf
sql = sql & "--FirstMinusRowDate date," & vbCrLf
sql = sql & "EndRow int," & vbCrLf
sql = sql & "OurEndQty int," & vbCrLf
sql = sql & "CustEndQty int)" & vbCrLf
sql = sql & "insert @Actions (PartNumber)" & vbCrLf
sql = sql & "select distinct PartNumber from MinMax where [Our Qty] < 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set ActionDate = (select Min(ActivityDate) from MinMax" & vbCrLf
sql = sql & "where PartNumber = [@Actions].PartNumber and [Our Qty] < 0" & vbCrLf
sql = sql & "group by PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions set MOOrigDate = ActionDate" & vbCrLf
sql = sql & "update @Actions set MODate = dateadd(day,-30,ActionDate)" & vbCrLf
sql = sql & "update @Actions set MODate = @startDate where MODate < @startDate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set MinQty = (select Min([Our Qty]) from MinMax" & vbCrLf
sql = sql & "where PartNumber = [@Actions].PartNumber and [Our Qty] < 0" & vbCrLf
sql = sql & "group by PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set EndRow = (select Max(Row) from MinMax" & vbCrLf
sql = sql & "where [@Actions].PartNumber = MinMax.PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set FirstMinusRow = (select Min(Row) from MinMax" & vbCrLf
sql = sql & "where [@Actions].PartNumber = MinMax.PartNumber" & vbCrLf
sql = sql & "and [Our Qty] < 0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set OurEndQty = (select [Our Qty] from MinMax" & vbCrLf
sql = sql & "where [@Actions].EndRow = MinMax.Row)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set CustEndQty = (select [Cust Qty] from MinMax" & vbCrLf
sql = sql & "where [@Actions].EndRow = MinMax.Row)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update m" & vbCrLf
sql = sql & "set Action = 'Reschedule MO from ' + convert(varchar(10), m.ActivityDate,101)," & vbCrLf
sql = sql & "ActivityDate = MOOrigDate" & vbCrLf
sql = sql & "from MinMax m" & vbCrLf
sql = sql & "join @Actions on [@Actions].PartNumber = m.PartNumber" & vbCrLf
sql = sql & "where m.SortOrder = 20" & vbCrLf
sql = sql & "and m.ActivityDate > MOOrigDate" & vbCrLf
sql = sql & "and m.Row >= [@Actions].FirstMinusRow" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec MinMaxCalculate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- add action items for new MOs" & vbCrLf
sql = sql & "delete from @Actions" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @Actions (PartNumber)" & vbCrLf
sql = sql & "select distinct PartNumber from MinMax where [Our Qty] < 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set ActionDate = (select Min(ActivityDate) from MinMax" & vbCrLf
sql = sql & "where PartNumber = [@Actions].PartNumber and [Our Qty] < 0" & vbCrLf
sql = sql & "group by PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions set MOOrigDate = ActionDate" & vbCrLf
sql = sql & "update @Actions set MODate = dateadd(day,-30,ActionDate)" & vbCrLf
sql = sql & "update @Actions set MODate = @startDate where MODate < @startDate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set MinQty = (select Min([Our Qty]) from MinMax" & vbCrLf
sql = sql & "where PartNumber = [@Actions].PartNumber and [Our Qty] < 0" & vbCrLf
sql = sql & "group by PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set EndRow = (select Max(Row) from MinMax" & vbCrLf
sql = sql & "where [@Actions].PartNumber = MinMax.PartNumber)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set FirstMinusRow = (select Min(Row) from MinMax" & vbCrLf
sql = sql & "where [@Actions].PartNumber = MinMax.PartNumber" & vbCrLf
sql = sql & "and [Our Qty] < 0)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set OurEndQty = (select [Our Qty] from MinMax" & vbCrLf
sql = sql & "where [@Actions].EndRow = MinMax.Row)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @Actions" & vbCrLf
sql = sql & "set CustEndQty = (select [Cust Qty] from MinMax" & vbCrLf
sql = sql & "where [@Actions].EndRow = MinMax.Row)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "print 'new mo'" & vbCrLf
sql = sql & "INSERT INTO MinMax (PartNumber,[Part Ref],SortOrder,ActivityDate,[Cust Activity],[Our Activity],Type,Action)" & vbCrLf
sql = sql & "select PartNumber, dbo.fnCompress(PartNumber), 20, MODate, 0, - OurEndQty, 'New MO'," & vbCrLf
sql = sql & "'Schedule MO for qty ' + cast(-OurEndQty as varchar(5)) + ' on ' + convert(varchar(10), [@Actions].MODate,101)" & vbCrLf
sql = sql & "from @Actions" & vbCrLf
sql = sql & "where OurEndQty < 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec MinMaxCalculate" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update MinMax set Action = '' where Action is null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from MinMax where PartNumber in" & vbCrLf
sql = sql & "--(select PartNumber from MinMax where [Our Qty] < 0)" & vbCrLf
sql = sql & "--order by Row" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select PartNumber as [Part Number]" & vbCrLf
sql = sql & ",ActivityDate as [Activity Date]" & vbCrLf
sql = sql & "--,SortOrder" & vbCrLf
sql = sql & ",[Cust Activity]" & vbCrLf
sql = sql & ",[Our Activity]" & vbCrLf
sql = sql & ",[Type]" & vbCrLf
sql = sql & ",[Min]" & vbCrLf
sql = sql & ",[Max]" & vbCrLf
sql = sql & ",[Cust Qty]" & vbCrLf
sql = sql & ",[Our Qty]" & vbCrLf
sql = sql & ",[Row]" & vbCrLf
sql = sql & ",[Notes]" & vbCrLf
sql = sql & ",[Action]" & vbCrLf
sql = sql & ",[Selling Price]" & vbCrLf
sql = sql & ",Total" & vbCrLf
sql = sql & "from MinMax" & vbCrLf
sql = sql & "--where PartNumber = '112W9721-6'" & vbCrLf
sql = sql & "order by Row" & vbCrLf
ExecuteScript True, sql

sql = "DropStoredProcedureIfExists 'RptArAgingBase'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptArAgingBase" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Customer varchar(20)   -- blank for all" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* Detail AR Aging Base sp" & vbCrLf
sql = sql & "v3 11/19/2019 Fixes 91-120 Days grand total" & vbCrLf
sql = sql & "fixes credit memos not showing up in aging columns" & vbCrLf
sql = sql & "v2 6/17/2019" & vbCrLf
sql = sql & "take into account canceled invoice indicator" & vbCrLf
sql = sql & "Allow for invoices paid by a different customer" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptArAgingBase '6/21/2019', 'ASTK'" & vbCrLf
sql = sql & "SELECT * from ##TempArAging" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase '6/21/2019', ''" & vbCrLf
sql = sql & "SELECT * from ##TempArAging where [amt due] <> [inv total] order by nickname" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase '3/1/2019', ''" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @ArAcct varchar(12)" & vbCrLf
sql = sql & "select @ArAcct = COSJARACCT from ComnTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF OBJECT_ID('tempdb..##TempArAging') IS NOT NULL" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "drop table ##TempArAging" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get all invoices existing but not fully paid on desired date" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "rtrim(cust.CUNAME) as [Customer Name]," & vbCrLf
sql = sql & "rtrim(inv.INVCUST) as Nickname," & vbCrLf
sql = sql & "INVNO as [Inv #]," & vbCrLf
sql = sql & "INVTYPE as TP," & vbCrLf
sql = sql & "cast(INVDATE as DATE) AS [Inv Date]," & vbCrLf
sql = sql & "INVTOTAL as [Inv Total]," & vbCrLf
sql = sql & "isnull(x.Debits,0) as Debits," & vbCrLf
sql = sql & "isnull(x.credits,0) as Credits," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [Amt Paid]," & vbCrLf
sql = sql & "--case when INVTOTAL < 0 THEN INVTOTAL + isnull(x.debits,0) else isnull(x.debits,0) end as [Amt Paid]," & vbCrLf
sql = sql & "isnull(x.ct,0) as ct," & vbCrLf
sql = sql & "case when INVCHECKDATE is null then '' else convert(varchar(10),INVCHECKDATE,101) end as [Ck Date]," & vbCrLf
sql = sql & "cast(0 as int) as [Age Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [Amt Due]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [0-30 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [31-60 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [61-90 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [91-120 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [121+ Days]" & vbCrLf
sql = sql & "into ##TempArAging" & vbCrLf
sql = sql & "from CihdTable as inv" & vbCrLf
sql = sql & "join CustTable cust on cust.CUREF = inv.INVCUST" & vbCrLf
sql = sql & "left join (select DCINVNO," & vbCrLf
sql = sql & "isnull(sum(DCDEBIT),0) AS Debits," & vbCrLf
sql = sql & "isnull(sum(DCCREDIT),0) AS Credits," & vbCrLf
sql = sql & "COUNT(*) AS ct from JritTable" & vbCrLf
sql = sql & "where DCHEAD like 'cr%'" & vbCrLf
sql = sql & "and DCACCTNO = @ArAcct" & vbCrLf
sql = sql & "AND DCDATE <= @AsOfDate" & vbCrLf
sql = sql & "and DCINVNO <> 0" & vbCrLf
sql = sql & "--and DCDEBIT <> 0" & vbCrLf
sql = sql & "group by DCINVNO) x on x.DCINVNO = inv.INVNO" & vbCrLf
sql = sql & "where INVDATE <= @AsOfDate" & vbCrLf
sql = sql & "and INVCUST like @Customer + '%'" & vbCrLf
sql = sql & "and INVTOTAL <> isnull(x.Debits,0)" & vbCrLf
sql = sql & "and INVCANCELED = 0" & vbCrLf
sql = sql & "and (INVPIF = 0" & vbCrLf
sql = sql & "or isnull(INVCHECKDATE, DATEFROMPARTS(2050,1,1)) > @AsOfDate)" & vbCrLf
sql = sql & "ORDER by INVCUST, INVDATE, INVNO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ##TempArAging set [Amt Paid] = sign([Inv Total]) * (Credits - Debits)" & vbCrLf
sql = sql & "-- v 1" & vbCrLf
sql = sql & "-- update ##TempArAging set [Amt Due] = [Inv Total] - [Amt Paid] -- where [INV TOTAL] >= 0" & vbCrLf
sql = sql & "-- update ##TempArAging set [Amt Due] = [Amt Paid] where [INV TOTAL] < 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- v 2" & vbCrLf
sql = sql & "update ##TempArAging set [Amt Due] = [Inv Total] - [Amt Paid] -- where [INV TOTAL] >= 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ##TempArAging set [Age Days] = DATEDIFF(day,[Inv Date],@AsOfDate)" & vbCrLf
sql = sql & "update ##TempArAging set [0-30 Days] = case when [Age Days] between 0 and 30 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [31-60 Days] = case when [Age Days] between 31 and 60 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [61-90 Days] = case when [Age Days] between 61 and 90 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [91-120 Days] = case when [Age Days] between 91 and 120 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [121+ Days] = case when [Age Days] > 120 then [Amt Due] else 0 end" & vbCrLf
ExecuteScript True, sql

' required to speed up split lot updates
sql = "CREATE NONCLUSTERED INDEX IX_InvaTable_Receive" & vbCrLf
sql = sql & "ON [dbo].[InvaTable] ([INTYPE],[INREF2])" & vbCrLf
sql = sql & "INCLUDE ([INPART],[INAQTY],[INLOTNUMBER],[INLOTTRACK],[INUSEACTUALCOST])" & vbCrLf
ExecuteScript True, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function


Private Function UpdateDatabase195()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 195     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "DropStoredProcedureIfExists 'RptArAgingBase'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptArAgingBase" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Customer varchar(20)   -- blank for all" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/* Detail AR Aging Base sp" & vbCrLf
sql = sql & "v4 11/22/2019 Fixes problem with Cash advance balance not showing in aging columns" & vbCrLf
sql = sql & "v3 11/19/2019 Fixes 91-120 Days grand total" & vbCrLf
sql = sql & "fixes credit memos not showing up in aging columns" & vbCrLf
sql = sql & "v2 6/17/2019" & vbCrLf
sql = sql & "take into account canceled invoice indicator" & vbCrLf
sql = sql & "Allow for invoices paid by a different customer" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "exec RptArAgingBase '6/21/2019', 'ASTK'" & vbCrLf
sql = sql & "SELECT * from ##TempArAging" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase '6/21/2019', ''" & vbCrLf
sql = sql & "SELECT * from ##TempArAging where [amt due] <> [inv total] order by nickname" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase '3/1/2019', ''" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SET NOCOUNT ON" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @ArAcct varchar(12)" & vbCrLf
sql = sql & "select @ArAcct = COSJARACCT from ComnTable" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF OBJECT_ID('tempdb..##TempArAging') IS NOT NULL" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "drop table ##TempArAging" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get all invoices existing but not fully paid on desired date" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "rtrim(cust.CUNAME) as [Customer Name]," & vbCrLf
sql = sql & "rtrim(inv.INVCUST) as Nickname," & vbCrLf
sql = sql & "INVNO as [Inv #]," & vbCrLf
sql = sql & "INVTYPE as TP," & vbCrLf
sql = sql & "cast(INVDATE as DATE) AS [Inv Date]," & vbCrLf
sql = sql & "INVTOTAL as [Inv Total]," & vbCrLf
sql = sql & "isnull(x.Debits,0) as Debits," & vbCrLf
sql = sql & "isnull(x.credits,0) as Credits," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [Amt Paid]," & vbCrLf
sql = sql & "--case when INVTOTAL < 0 THEN INVTOTAL + isnull(x.debits,0) else isnull(x.debits,0) end as [Amt Paid]," & vbCrLf
sql = sql & "isnull(x.ct,0) as ct," & vbCrLf
sql = sql & "case when INVCHECKDATE is null then '' else convert(varchar(10),INVCHECKDATE,101) end as [Ck Date]," & vbCrLf
sql = sql & "cast(0 as int) as [Age Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [Amt Due]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [0-30 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [31-60 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [61-90 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [91-120 Days]," & vbCrLf
sql = sql & "cast(0 as decimal(15,2)) as [121+ Days]" & vbCrLf
sql = sql & "into ##TempArAging" & vbCrLf
sql = sql & "from CihdTable as inv" & vbCrLf
sql = sql & "join CustTable cust on cust.CUREF = inv.INVCUST" & vbCrLf
sql = sql & "left join (select DCINVNO," & vbCrLf
sql = sql & "isnull(sum(DCDEBIT),0) AS Debits," & vbCrLf
sql = sql & "isnull(sum(DCCREDIT),0) AS Credits," & vbCrLf
sql = sql & "COUNT(*) AS ct from JritTable" & vbCrLf
sql = sql & "where DCHEAD like 'cr%'" & vbCrLf
sql = sql & "and DCACCTNO = @ArAcct" & vbCrLf
sql = sql & "AND DCDATE <= @AsOfDate" & vbCrLf
sql = sql & "and DCINVNO <> 0" & vbCrLf
sql = sql & "--and DCDEBIT <> 0" & vbCrLf
sql = sql & "group by DCINVNO) x on x.DCINVNO = inv.INVNO" & vbCrLf
sql = sql & "where INVDATE <= @AsOfDate" & vbCrLf
sql = sql & "and INVCUST like @Customer + '%'" & vbCrLf
sql = sql & "and INVTOTAL <> isnull(x.Debits,0)" & vbCrLf
sql = sql & "and INVCANCELED = 0" & vbCrLf
sql = sql & "and (INVPIF = 0" & vbCrLf
sql = sql & "or isnull(INVCHECKDATE, DATEFROMPARTS(2050,1,1)) > @AsOfDate)" & vbCrLf
sql = sql & "ORDER by INVCUST, INVDATE, INVNO" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ##TempArAging set [Amt Paid] = sign([Inv Total]) * (Credits - Debits)" & vbCrLf
sql = sql & "-- v 1" & vbCrLf
sql = sql & "-- update ##TempArAging set [Amt Due] = [Inv Total] - [Amt Paid] -- where [INV TOTAL] >= 0" & vbCrLf
sql = sql & "-- update ##TempArAging set [Amt Due] = [Amt Paid] where [INV TOTAL] < 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- v 2" & vbCrLf
sql = sql & "update ##TempArAging set [Amt Due] = [Inv Total] - [Amt Paid] -- where [INV TOTAL] >= 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- v 4" & vbCrLf
sql = sql & "update ##TempArAging set [Amt Due] = [Amt Paid] where TP = 'CA'" & vbCrLf
sql = sql & "update ##TempArAging set [Amt Due] = [Inv Total] + [Amt Paid] where TP = 'CM'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update ##TempArAging set [Age Days] = DATEDIFF(day,[Inv Date],@AsOfDate)" & vbCrLf
sql = sql & "update ##TempArAging set [0-30 Days] = case when [Age Days] between 0 and 30 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [31-60 Days] = case when [Age Days] between 31 and 60 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [61-90 Days] = case when [Age Days] between 61 and 90 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [91-120 Days] = case when [Age Days] between 91 and 120 then [Amt Due] else 0 end" & vbCrLf
sql = sql & "update ##TempArAging set [121+ Days] = case when [Age Days] > 120 then [Amt Due] else 0 end" & vbCrLf
ExecuteScript True, sql



''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase196()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 196     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "DropStoredProcedureIfExists 'RptArAgingSummary'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptArAgingSummary" & vbCrLf
sql = sql & "@AsOfDate date," & vbCrLf
sql = sql & "@Customer varchar(30)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "AR Aging Summary" & vbCrLf
sql = sql & "v2 - 12/28/2019 added @Customer parameter" & vbCrLf
sql = sql & "test" & vbCrLf
sql = sql & "tests" & vbCrLf
sql = sql & "exec RptArAgingSummary '3/1/2019', 'ASTK'" & vbCrLf
sql = sql & "exec RptArAgingSummary '3/1/2019', ''" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "exec RptArAgingBase @AsOfDate, ''" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select [Customer Name], Nickname, count(*) as INVOICES, sum([Amt Due]) as Total," & vbCrLf
sql = sql & "sum([0-30 Days]) as [0-30 Days]," & vbCrLf
sql = sql & "sum([31-60 Days]) as [31-60 Days]," & vbCrLf
sql = sql & "sum([61-90 Days]) as [61-90 Days]," & vbCrLf
sql = sql & "sum([91-120 Days]) as [91-120 Days]," & vbCrLf
sql = sql & "sum([121+ Days]) as [121+ Days]" & vbCrLf
sql = sql & "from ##TempArAging" & vbCrLf
sql = sql & "where @Customer = '' or Nickname = @Customer" & vbCrLf
sql = sql & "group by [Customer Name],[NickName]" & vbCrLf
sql = sql & "order by [Customer Name]" & vbCrLf
ExecuteScript True, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function


Private Function UpdateDatabase197()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 197     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "exec DropStoredProcedureIfExists 'RptInvMovFromWIP'" & vbCrLf
ExecuteScript True, sql

sql = "create PROCEDURE RptInvMovFromWIP" & vbCrLf
sql = sql & "@StartDate as varchar(16), @EndDate as Varchar(16)," & vbCrLf
sql = sql & "@PartType1 as Integer, @PartType2 as Integer, @PartType3 as Integer, @PartType4 as Integer" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/* 1/15/2020 TEL - modified to include datetimes for final day of period (time portion made the result greater than the date alone" & vbCrLf
sql = sql & "test:" & vbCrLf
sql = sql & "exec RptInvMovFromWIP '10/1/19','10/31/19', 1, 1, 1, 1  -- 3671" & vbCrLf
sql = sql & "exec RptInvMovFromWIP '11/1/19','11/30/19', 1, 1, 1, 1  -- 2792" & vbCrLf
sql = sql & "exec RptInvMovFromWIP '12/1/19','12/31/19', 1, 1, 1, 1  -- 2211" & vbCrLf
sql = sql & "select * from _temp order by inadate" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @Start date" & vbCrLf
sql = sql & "declare @Cutoff date" & vbCrLf
sql = sql & "set @Start = cast(@StartDate as date)" & vbCrLf
sql = sql & "set @Cutoff = dateadd(day, 1, cast(@EndDate as date))" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartType1 = 1)" & vbCrLf
sql = sql & "SET @PartType1 = 1" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType1 = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartType2 = 1)" & vbCrLf
sql = sql & "SET @PartType2 = 2" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType2 = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartType3 = 1)" & vbCrLf
sql = sql & "SET @PartType3 = 3" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType3 = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartType4 = 1)" & vbCrLf
sql = sql & "SET @PartType4 = 4" & vbCrLf
sql = sql & "Else" & vbCrLf
sql = sql & "SET @PartType4 = 0" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT Inadate, InvaTable.INTYPE,InvaTable.INLOTNUMBER,InvaTable.INPART, InvaTable.INREF2," & vbCrLf
sql = sql & "InvaTable.INAQTY, InvaTable.INAMT, LohdTable.LOTORIGINALQTY, LohdTable.LOTTOTMATL, INTOTMATL," & vbCrLf
sql = sql & "LohdTable.LOTTOTLABOR, INTOTLABOR, LohdTable.LOTTOTEXP, INTOTEXP, LohdTable.LOTTOTOH, INTOTOH," & vbCrLf
sql = sql & "LOTDATECOSTED, INDEBITACCT, INCREDITACCT" & vbCrLf
sql = sql & "FROM" & vbCrLf
sql = sql & "(PartTable PartTable INNER JOIN InvaTable InvaTable ON" & vbCrLf
sql = sql & "PartTable.PARTREF = InvaTable.INPART)" & vbCrLf
sql = sql & "LEFT OUTER JOIN LohdTable LohdTable ON" & vbCrLf
sql = sql & "InvaTable.INLOTNUMBER = LohdTable.LOTNUMBER" & vbCrLf
sql = sql & "WHERE" & vbCrLf
sql = sql & "InvaTable.INTYPE IN (6,12)" & vbCrLf
sql = sql & "and INADATE >= @Start and INADATE < @Cutoff" & vbCrLf
sql = sql & "and PALEVEL IN (@PartType1, @PartType2, @PartType3, @PartType4)" & vbCrLf
sql = sql & "--    AND INPART = 'MX000075'" & vbCrLf
sql = sql & "AND INLOTNUMBER NOT IN (SELECT a.INLOTNUMBER  FROM InvaTable a where" & vbCrLf
sql = sql & "a.INPART = InvaTable.INPART" & vbCrLf
sql = sql & "--a.INPART = 'MX000075'" & vbCrLf
sql = sql & "AND a.INTYPE IN (6,38,12) AND Convert(DateTime, a.Inadate, 101) between @StartDate and @EndDate" & vbCrLf
sql = sql & "GROUP BY INLOTNUMBER" & vbCrLf
sql = sql & "HAVING COUNT(INLOTNUMBER) > 1)" & vbCrLf
sql = sql & "ORDER BY" & vbCrLf
sql = sql & "PartTable.PALEVEL ASC," & vbCrLf
sql = sql & "PartTable.PACLASS ASC" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "END" & vbCrLf
ExecuteScript True, sql

sql = "dropstoredprocedureifexists 'RptVendorDelPerformance'" & vbCrLf
ExecuteScript True, sql

' make sure old standard report users are directed to the report that uses this sp
sql = "update CustomReports set REPORT_CUSTOMREPORT = '' where REPORT_CUSTOMREPORT ='finap12a'" & vbCrLf
ExecuteScript True, sql

sql = "create PROCEDURE [dbo].[RptVendorDelPerformance]" & vbCrLf
sql = sql & "@sVendorRef as VARCHAR(10)," & vbCrLf
sql = sql & "@sBeginDate as VARCHAR(10)," & vbCrLf
sql = sql & "@sEndDate as VARCHAR(10)," & vbCrLf
sql = sql & "@iAllowDaysEarly as INTEGER," & vbCrLf
sql = sql & "@iAllowDaysLate as INTEGER," & vbCrLf
sql = sql & "@iUseOriginalShipDate as INTEGER," & vbCrLf
sql = sql & "@iCalcLateBy as INTEGER" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/* 1/16/2020 - This adds the 7th parameter for users where an older version was in use." & vbCrLf
sql = sql & "Test:" & vbCrLf
sql = sql & "RptVendorDelPerformance '', '2019-11-01', '2019-11-30', 0, 0, 0, 0" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "-- SET NOCOUNT ON added to prevent extra result sets from" & vbCrLf
sql = sql & "-- interfering with SELECT statements." & vbCrLf
sql = sql & "SET NOCOUNT ON;" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (UPPER(@sVendorRef) = 'ALL') Or (UPPER(@sVendorRef) = '<ALL>')" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @sVendorRef = ''" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "IF (@iAllowDaysEarly > 0)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @iAllowDaysEarly = @iAllowDaysEarly * -1" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "-- Insert statements for procedure here" & vbCrLf
sql = sql & "SELECT DISTINCT VEREF, VEBNAME, PINUMBER, PIRELEASE, PIITEM, PIREV, PIPART, PARTNUM, PIADATE, PIPDATE, PIPQTY, PIAQTY, PIONDOCKINSPDATE," & vbCrLf
sql = sql & "PIINSDATE , PIRECEIVED, PIONDOCKINSPECTED, PIODDELDATE, PIPORIGDATE, PIONDOCKQTYACC, PIONDOCKQTYREJ, PIODDELQTY, PIREJECTED, PIWASTE, PIONDOCKQTYWASTE," & vbCrLf
sql = sql & "PIESTUNIT, ISNULL(INPQTY,0.0000) AS INPQTY, ISNULL(INAQTY,0.0000) AS INAQTY, ISNULL(INAMT,0.0000) AS INAMT, ISNULL(INPOITEM,0) AS INPOITEM," & vbCrLf
sql = sql & "ISNULL(INPOREV,'') AS INPOREV, ISNULL(INPONUMBER,0) AS INPONUMBER, ISNULL(INPORELEASE,0) AS INPORELEASE, ISNULL(INTYPE,0) AS INTYPE," & vbCrLf
sql = sql & "CASE WHEN ISNULL(PIODDELDATE, '')>'' THEN PIODDELDATE" & vbCrLf
sql = sql & "WHEN ISNULL(PIONDOCKINSPDATE,'')>'' THEN PIONDOCKINSPDATE" & vbCrLf
sql = sql & "ELSE PIADATE END AS 'DELIVERYDATE'," & vbCrLf
sql = sql & "CASE WHEN (ISNULL(PIRECEIVED,'')='') AND (ISNULL(PIINSDATE,'')='') AND (ISNULL(PIODDELDATE,'')='') THEN 'PO'" & vbCrLf
sql = sql & "WHEN (ISNULL(PIRECEIVED,'')='') AND (ISNULL(PIINSDATE,'')='') THEN 'DEL'" & vbCrLf
sql = sql & "WHEN (ISNULL(PIRECEIVED,'')='') THEN 'DOCK'" & vbCrLf
sql = sql & "WHEN PIRECEIVED > '01/01/1900' THEN 'REC' ELSE '' END AS 'STATUS'," & vbCrLf
sql = sql & "CASE WHEN (@iUseOriginalShipDate=1) AND (ISNULL(PIPORIGDATE,''))='' THEN PIPDATE" & vbCrLf
sql = sql & "WHEN (@iUseOriginalShipDate=1) THEN PIPORIGDATE" & vbCrLf
sql = sql & "ELSE PIPDATE END AS 'DUEDATE'" & vbCrLf
sql = sql & "From VndrTable" & vbCrLf
sql = sql & "INNER JOIN PoitTable ON VEREF=PIVENDOR" & vbCrLf
sql = sql & "LEFT OUTER JOIN InvaTable ON INPONUMBER=PINUMBER AND INPORELEASE=PIRELEASE AND INPOREV=PIREV AND INPOITEM=PIITEM" & vbCrLf
sql = sql & "LEFT OUTER JOIN PartTable ON PARTREF=PIPART" & vbCrLf
sql = sql & "Where" & vbCrLf
sql = sql & "VEREF LIKE @sVendorRef + '%' AND INTYPE = 15" & vbCrLf
sql = sql & "AND" & vbCrLf
sql = sql & "((PIADATE IS NOT NULL) OR (PIODDELDATE IS NOT NULL) OR (PIONDOCKINSPDATE IS NOT NULL))" & vbCrLf
sql = sql & "AND" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "((CASE WHEN @iCalcLateBy=0 THEN PIADATE WHEN @iCalcLateBy=1 Then PIONDOCKINSPDATE ELSE PIODDELDATE END) BETWEEN Cast(@sBeginDate AS DateTime) AND Cast(@sEndDate AS DateTime))" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "End" & vbCrLf
ExecuteScript True, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase198()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 198     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

sql = "EXEC DropStoredProcedureIfExists 'RptInventoryAdjustments'" & vbCrLf
ExecuteScript True, sql

sql = "CREATE PROCEDURE RptInventoryAdjustments" & vbCrLf
sql = sql & "@StartDate as varchar(16), @EndDate as Varchar(16), @PartClass as Varchar(16),@PartCode as varchar(8)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/* Inventory Adjustment Report" & vbCrLf
sql = sql & "2/5/2020 Revised to avoid division by lot original quantity of zero" & vbCrLf
sql = sql & "Test:" & vbCrLf
sql = sql & "EXEC RptInventoryAdjustments '2019-01-22', '2019-01-22', 'ALL', 'ALL'" & vbCrLf
sql = sql & "EXEC RptInventoryAdjustments '2018-01-01', '2019-12-31', 'ALL', 'ALL'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "declare @partRef as varchar(30)" & vbCrLf
sql = sql & "declare @partNum as varchar(30)" & vbCrLf
sql = sql & "declare @partDesc as varchar(30)" & vbCrLf
sql = sql & "declare @partExDesc as varchar(3072)" & vbCrLf
sql = sql & "declare @lotNum as varchar(15)" & vbCrLf
sql = sql & "declare @lotUserID as varchar(40)" & vbCrLf
sql = sql & "declare @Inno as Int" & vbCrLf
sql = sql & "declare @actualDt as varchar(30)" & vbCrLf
sql = sql & "declare @qty as decimal(12, 4)" & vbCrLf
sql = sql & "declare @Orgqty as decimal(12, 4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @invAmt as decimal(12,4)" & vbCrLf
sql = sql & "declare @invTotMatl decimal(12,4)" & vbCrLf
sql = sql & "declare @invTotLabor decimal(12,4)" & vbCrLf
sql = sql & "declare @invTotExp decimal(12,4)" & vbCrLf
sql = sql & "declare @invTotOH decimal(12,4)" & vbCrLf
sql = sql & "declare @creditAcc varchar(12)" & vbCrLf
sql = sql & "declare @debitAcc varchar(12)" & vbCrLf
sql = sql & "declare @lotTotMatl decimal(12,4)" & vbCrLf
sql = sql & "declare @lotTotLabor decimal(12,4)" & vbCrLf
sql = sql & "declare @lotTotExp decimal(12,4)" & vbCrLf
sql = sql & "declare @lotTotOH decimal(12,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @totMatlCost decimal(12,4)" & vbCrLf
sql = sql & "declare @totLaborCost decimal(12,4)" & vbCrLf
sql = sql & "declare @totExpCost decimal(12,4)" & vbCrLf
sql = sql & "declare @totOHCost decimal(12,4)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @fInvMatl int" & vbCrLf
sql = sql & "declare @fInvLabor int" & vbCrLf
sql = sql & "declare @fInvExp int" & vbCrLf
sql = sql & "declare @fInvOH int" & vbCrLf
sql = sql & "declare @invType int" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set NOCOUNT ON" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartClass = 'ALL')" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @PartClass = ''" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "IF (@PartCode = 'ALL')" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @PartCode = ''" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- DELETE FROM #tempINVReport" & vbCrLf
sql = sql & "--ALTER TABLE #tempINVReport  Add INTYPE int null" & vbCrLf
sql = sql & "--ALTER TABLE #tempINVReport  Add LOTORIGINALQTY decimal(12,4) null" & vbCrLf
sql = sql & "CREATE TABLE #tempINVReport" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "PARTNUM Varchar(30) NULL," & vbCrLf
sql = sql & "PADESC varchar(30) NULL ," & vbCrLf
sql = sql & "PAEXTDESC varchar(3072) NULL ," & vbCrLf
sql = sql & "LOTNUMBER varchar(15) NULL," & vbCrLf
sql = sql & "LOTUSERLOTID varchar(40) NULL," & vbCrLf
sql = sql & "LOTORIGINALQTY decimal(12,4) null," & vbCrLf
sql = sql & "INTYPE int null," & vbCrLf
sql = sql & "INNO int NULL," & vbCrLf
sql = sql & "INADATE varchar(30) NULL," & vbCrLf
sql = sql & "INAQTY decimal(12,4) NULL," & vbCrLf
sql = sql & "INAMT decimal (12,4) NULL," & vbCrLf
sql = sql & "TOTMATL decimal(12,4) NULL," & vbCrLf
sql = sql & "TOTLABOR decimal(12,4) NULL," & vbCrLf
sql = sql & "TOTEXP decimal(12,4) NULL," & vbCrLf
sql = sql & "TOTOH decimal(12,4) NULL," & vbCrLf
sql = sql & "CREDITACCT varchar(12) NULL," & vbCrLf
sql = sql & "DEBITACCT varchar(12) NULL," & vbCrLf
sql = sql & "flgMatl int NULL," & vbCrLf
sql = sql & "flgLabor int NULL," & vbCrLf
sql = sql & "flgExp int NULL," & vbCrLf
sql = sql & "flgOH int NULL" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "DECLARE curInvRpt CURSOR   FOR" & vbCrLf
sql = sql & "SELECT INTYPE, INPART, LOTNUMBER, LOTUSERLOTID, INNO, INADATE," & vbCrLf
sql = sql & "INAQTY, LOTORIGINALQTY, INAMT, INTOTMATL," & vbCrLf
sql = sql & "INTOTLABOR, INTOTEXP, INTOTOH," & vbCrLf
sql = sql & "INCREDITACCT, INDEBITACCT, LOTTOTMATL," & vbCrLf
sql = sql & "LOTTOTLABOR, LOTTOTEXP, LOTTOTOH," & vbCrLf
sql = sql & "PartNum , PADESC, PAEXTDESC" & vbCrLf
sql = sql & "From viewRptInventoryAdjustments, PartTable" & vbCrLf
sql = sql & "Where viewRptInventoryAdjustments.INPART = PartTable.PartRef" & vbCrLf
sql = sql & "AND viewRptInventoryAdjustments.INADATE BETWEEN @StartDate AND @EndDate" & vbCrLf
sql = sql & "AND PartTable.PACLASS LIKE '%' + @PartClass + '%'" & vbCrLf
sql = sql & "AND PartTable.PAPRODCODE LIKE '%' + @PartCode + '%'" & vbCrLf
sql = sql & "--         AND INPART = '65B801038'" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "OPEN curInvRpt" & vbCrLf
sql = sql & "FETCH NEXT FROM curInvRpt INTO @invType, @partRef, @lotNum, @lotUserID, @Inno, @actualDt, @qty, @Orgqty," & vbCrLf
sql = sql & "@invAmt, @invTotMatl, @invTotLabor," & vbCrLf
sql = sql & "@invTotExp,@invTotOH, @creditAcc, @debitAcc," & vbCrLf
sql = sql & "@lotTotMatl, @lotTotLabor, @lotTotExp, @lotTotOH," & vbCrLf
sql = sql & "@partNum, @partDesc, @partExDesc" & vbCrLf
sql = sql & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF (@@FETCH_STATUS <> -2)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- Get the costed values from Lothd table" & vbCrLf
sql = sql & "-- if the Inv table does not have the cost for" & vbCrLf
sql = sql & "-- material, expenses, OH and Labour." & vbCrLf
sql = sql & "SET @totMatlCost = @invTotMatl" & vbCrLf
sql = sql & "SET @fInvMatl = 1" & vbCrLf
sql = sql & "IF (@invTotMatl = 0.0000)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF @lotTotMatl IS NOT NULL AND @Orgqty <> 0.0000" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @totMatlCost = (@qty * @lotTotMatl ) / @Orgqty" & vbCrLf
sql = sql & "SET @fInvMatl = 0" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "-- Labor" & vbCrLf
sql = sql & "SET @totLaborCost = @invTotLabor" & vbCrLf
sql = sql & "SET @fInvLabor = 1" & vbCrLf
sql = sql & "IF (@invTotLabor = 0.0000)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF @lotTotLabor IS NOT NULL and  @Orgqty <> 0" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @totLaborCost = (@qty * @lotTotLabor) / @Orgqty" & vbCrLf
sql = sql & "SET @fInvLabor = 0" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "-- Exp" & vbCrLf
sql = sql & "SET @totExpCost = @invTotExp" & vbCrLf
sql = sql & "SET @fInvExp = 1" & vbCrLf
sql = sql & "IF (@invTotExp = 0.0000)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF @lotTotExp IS NOT NULL and @Orgqty <> 0" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "SET @totExpCost = (@qty * @lotTotExp ) / @Orgqty" & vbCrLf
sql = sql & "SET @fInvExp = 0" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "-- OH" & vbCrLf
sql = sql & "SET @totOHCost = @invTotOH" & vbCrLf
sql = sql & "SET @fInvOH = 1" & vbCrLf
sql = sql & "IF (@invTotOH = 0.0000)" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "IF @lotTotOH IS NOT NULL and @OrgQty <> 0" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "SET @totOHCost = (@qty * @lotTotOH ) / @Orgqty" & vbCrLf
sql = sql & "SET @fInvOH = 0" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- Insert to the temp table" & vbCrLf
sql = sql & "INSERT INTO #tempINVReport (PARTNUM, LOTNUMBER, LOTUSERLOTID," & vbCrLf
sql = sql & "INNO, PADESC, PAEXTDESC, INADATE," & vbCrLf
sql = sql & "INAQTY, LOTORIGINALQTY, INAMT, TOTMATL,TOTLABOR,TOTEXP,TOTOH, CREDITACCT, DEBITACCT," & vbCrLf
sql = sql & "flgMatl, flgLabor, flgExp, flgOH, INTYPE)" & vbCrLf
sql = sql & "VALUES (@partNum, @lotNum, @lotUserID, @Inno, @partDesc, @partExDesc, @actualDt,@qty, @Orgqty," & vbCrLf
sql = sql & "@invAmt,@totMatlCost,@totLaborCost,@totExpCost,@totOHCost," & vbCrLf
sql = sql & "@creditAcc,@debitAcc,@fInvMatl,@fInvLabor,@fInvExp,@fInvOH, @invType)" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "FETCH NEXT FROM curInvRpt INTO @invType, @partRef, @lotNum, @lotUserID, @Inno," & vbCrLf
sql = sql & "@actualDt, @qty, @Orgqty, @invAmt, @invTotMatl, @invTotLabor," & vbCrLf
sql = sql & "@invTotExp,@invTotOH, @creditAcc, @debitAcc, @lotTotMatl, @lotTotLabor, @lotTotExp, @lotTotOH," & vbCrLf
sql = sql & "@partNum, @partDesc, @partExDesc" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "CLOSE curInvRpt   --// close the cursor" & vbCrLf
sql = sql & "DEALLOCATE curInvRpt" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- select data for the report" & vbCrLf
sql = sql & "SELECT a.PARTNUM as PARTNUM, LOTNUMBER, LOTUSERLOTID, INNO," & vbCrLf
sql = sql & "a.PADESC as PADESC, a.PAEXTDESC as PAEXTDESC, PALEVEL, INADATE," & vbCrLf
sql = sql & "INAQTY,INAMT, TOTMATL,TOTLABOR,TOTEXP,TOTOH,INTYPE," & vbCrLf
sql = sql & "CREDITACCT, DEBITACCT,flgMatl, flgLabor," & vbCrLf
sql = sql & "flgExp , flgOH" & vbCrLf
sql = sql & "FROM #tempINVReport a, PartTable" & vbCrLf
sql = sql & "WHERE PartTable.PARTNUM = a.PARTNUM" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- drop the temp table" & vbCrLf
sql = sql & "DROP table #tempINVReport" & vbCrLf
sql = sql & "End" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript True, sql

' Early/Late PO update
sql = "dropstoredprocedureifexists 'RptEarlyLateDatesPO'" & vbCrLf
ExecuteScript True, sql

sql = "CREATE PROCEDURE RptEarlyLateDatesPO" & vbCrLf
sql = sql & "@StartDate as varchar(16)," & vbCrLf
sql = sql & "@EndDate as Varchar(16)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "TEL 3/24/2020" & vbCrLf
sql = sql & "replace old flawed triple-cursor based report." & vbCrLf
sql = sql & "this is still very slow, but hopefully a lot more accurate" & vbCrLf
sql = sql & "TEST:" & vbCrLf
sql = sql & "exec RptEarlyLateDatesPO '1/1/2020', '12/31/2020'" & vbCrLf
sql = sql & "exec RptEarlyLateDatesPO '12/6/2019', '1/9/2020'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @table table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "Part varchar(30) NULL," & vbCrLf
sql = sql & "[Req Date] [date] NULL," & vbCrLf
sql = sql & "[Sched Date] [date] NULL," & vbCrLf
sql = sql & "[Days Early] [int] NULL," & vbCrLf
sql = sql & "[Quantity] [decimal](12, 4) NULL," & vbCrLf
sql = sql & "[Balance] [decimal](38, 4) NULL," & vbCrLf
sql = sql & "Comment varchar(80) NULL," & vbCrLf
sql = sql & "MRP_TYPE tinyint NULL," & vbCrLf
sql = sql & "MRP_Row int," & vbCrLf
sql = sql & "Row int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @table" & vbCrLf
sql = sql & "SELECT rtrim(MRP_PARTREF) as Part, cast(null as date) as [Req Date]," & vbCrLf
sql = sql & "case when MRP_TYPE = 1 THEN cast('1/1/1900' as date) else MRP_PARTDATERQD end as [Sched Date]," & vbCrLf
sql = sql & "cast(null as int) as [Days Early], MRP_PARTQTYRQD as Quantity," & vbCrLf
sql = sql & "sum(MRP_PARTQTYRQD) OVER (partition by MRP_PARTREF" & vbCrLf
sql = sql & "order by case when MRP_TYPE = 1 THEN cast('1/1/1900' as date) else MRP_PARTDATERQD end, MRP_ROW) as Balance," & vbCrLf
sql = sql & "MRP_COMMENT as Comment, MRP_TYPE, MRP_ROW," & vbCrLf
sql = sql & "ROW_NUMBER() over (order by MRP_PARTREF, case when MRP_TYPE = 1 THEN cast('1/1/1900' as date) else MRP_PARTDATERQD end, MRP_ROW)" & vbCrLf
sql = sql & "FROM PartTable" & vbCrLf
sql = sql & "JOIN MrplTable ON PartTable.PARTREF=MRP_PARTREF" & vbCrLf
sql = sql & "WHERE MRP_TYPE not in (5,6) and PALEVEL = 4 -- and MRP_PARTREF = '024116' -- and MRP_PARTREF = '141T61232IMA1'" & vbCrLf
sql = sql & "and exists (select * from MrplTable mrp2 where mrp2.MRP_PARTREF = MrplTable.MRP_PARTREF and mrp2.MRP_TYPE = 2" & vbCrLf
sql = sql & "and MRP_PARTDATERQD BETWEEN @StartDate AND  @EndDate)" & vbCrLf
sql = sql & "and MRP_PARTDATERQD BETWEEN @StartDate AND  @EndDate" & vbCrLf
sql = sql & "ORDER BY MRP_PARTREF, case when MRP_TYPE = 1 THEN cast('1/1/1900' as date) else MRP_PARTDATERQD end, MRP_ROW" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-------------------------------------------------" & vbCrLf
sql = sql & "-- determine needed date for parts needed earlier" & vbCrLf
sql = sql & "-------------------------------------------------" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- find first pos preceded by a negative quantity and move po up if possible" & vbCrLf
sql = sql & "declare @candidates table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "Part varchar(30)," & vbCrLf
sql = sql & "FirstMinusRow int," & vbCrLf
sql = sql & "MinusDate date," & vbCrLf
sql = sql & "FirstPoAfter int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @ct int, @s as varchar(30)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set nocount on" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @i int" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @i = 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "while @i < 50" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "delete from @candidates" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @candidates" & vbCrLf
sql = sql & "select Part, Min(Row) as FirstMinusRow, cast(null as date) as [Sched Date], null as FirstPoAfter from @table" & vbCrLf
sql = sql & "where MRP_TYPE <> 2 and Balance < 0" & vbCrLf
sql = sql & "group by PART" & vbCrLf
sql = sql & "order by Part" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @candidates set MinusDate = (select [Sched Date] from @table where Row = [@candidates].FirstMinusRow)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @candidates set FirstPoAfter" & vbCrLf
sql = sql & "= (select min(Row) from @table t where MRP_TYPE = 2 and t.Part = [@candidates].Part and t.[Req Date] is null)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select @ct = count(FirstPoAfter) from @candidates" & vbCrLf
sql = sql & "--set @s = 'minus loop ' + cast(@i as varchar(2)) + ': count ' + cast(@ct as varchar(10))" & vbCrLf
sql = sql & "--print @s" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @ct = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update t set [Req Date] = MinusDate" & vbCrLf
sql = sql & "from @candidates c join @table t on t.Part = c.Part and t.Row = c.FirstPoAfter" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- recalculate and renumber" & vbCrLf
sql = sql & "update x set Row = New_Row" & vbCrLf
sql = sql & "from" & vbCrLf
sql = sql & "(select Row, ROW_NUMBER() over (" & vbCrLf
sql = sql & "order by Part, isnull([Req Date], [Sched Date]), MRP_TYPE, MRP_Row) as New_Row" & vbCrLf
sql = sql & "from @table) x" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update x set Balance = New_Bal from" & vbCrLf
sql = sql & "(select Balance, sum(Quantity) over (partition by Part order by Row) as New_Bal from @table) x" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from @candidates order by Part" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @i = @i + 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-----------------------------------------------------------" & vbCrLf
sql = sql & "-- determine needed date for parts not required until later" & vbCrLf
sql = sql & "-----------------------------------------------------------" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- find last PO for each part and move down if possible" & vbCrLf
sql = sql & "declare @candidates2 table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "Part varchar(30)," & vbCrLf
sql = sql & "LastMoveToRow int," & vbCrLf
sql = sql & "LastMoveToDate date," & vbCrLf
sql = sql & "LastPo int," & vbCrLf
sql = sql & "LastPoQty decimal(12,4)," & vbCrLf
sql = sql & "LastPoDate date" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @i = 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "while @i <= 50" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "delete from @candidates2" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @candidates2 (Part, LastPO)" & vbCrLf
sql = sql & "select Part, Max(Row) from @table" & vbCrLf
sql = sql & "where MRP_TYPE = 2 and [Req Date] is null" & vbCrLf
sql = sql & "group by PART" & vbCrLf
sql = sql & "order by Part" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update c set LastPoQty = Quantity, LastPoDate = [Sched Date]" & vbCrLf
sql = sql & "from @candidates2 c join @table t on c.LastPo = t.Row" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @candidates2 set LastMoveToRow = (select min(Row) from @table t" & vbCrLf
sql = sql & "where t.Part = [@candidates2].part and t.Row > [@candidates2].LastPo and t.Balance < [@candidates2].LastPoQty)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @candidates2 set LastMoveToDate = (select [Sched Date] from @table where Row = [@candidates2].LastMoveToRow)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select @ct = count(*) from @candidates2" & vbCrLf
sql = sql & "--set @s = 'plus loop ' + cast(@i as varchar(2)) + ': count ' + cast(@ct as varchar(10))" & vbCrLf
sql = sql & "--print @s" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @ct = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from @candidates2 order by Part" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--update t set [Req Date] = isnull(LastMoveToDate, DATEADD(day, 9999,LastPoDate))" & vbCrLf
sql = sql & "--from @candidates2 c join @table t on t.Part = c.Part and t.Row = c.LastPo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update t set [Req Date] =" & vbCrLf
sql = sql & "case when LastMoveToDate is not null then LastMoveToDate" & vbCrLf
sql = sql & "when Balance >= Quantity then DATEADD(day, 9999,LastPoDate)" & vbCrLf
sql = sql & "else [Sched Date] end" & vbCrLf
sql = sql & "from @candidates2 c join @table t on t.Part = c.Part and t.Row = c.LastPo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- recalculate and renumber" & vbCrLf
sql = sql & "update x set Row = New_Row" & vbCrLf
sql = sql & "from" & vbCrLf
sql = sql & "(select Row, ROW_NUMBER() over (" & vbCrLf
sql = sql & "order by Part, isnull([Req Date], [Sched Date]), MRP_TYPE, MRP_Row) as New_Row" & vbCrLf
sql = sql & "from @table) x" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update x set Balance = New_Bal from" & vbCrLf
sql = sql & "(select Balance, sum(Quantity) over (partition by Part order by Row) as New_Bal from @table) x" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from @candidates2 order by Part" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @i = @i + 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-------------------------------------------" & vbCrLf
sql = sql & "-- calculate days early/late" & vbCrLf
sql = sql & "-------------------------------------------" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @table set [Days Early] = datediff(d,[Sched Date], [Req Date]) where [Req Date] is not null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- data required for report" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "MRP_PARTREF,MRP_PARTNUM, t.MRP_TYPE, MRP_PONUM, MRP_POITEM, MRP_PARTLEVEL, t.[Req Date] as MRP_HGHPRTDTRQD," & vbCrLf
sql = sql & "t.[Sched Date] as MRP_PARTDATERQD, MRP_PARTQTYRQD, MRP_CATAGORY, MRP_COMMENT," & vbCrLf
sql = sql & "[Days Early] as EARLYLATE_DAYS" & vbCrLf
sql = sql & "from @table t" & vbCrLf
sql = sql & "join MrplTable m on m.MRP_ROW = t.MRP_Row" & vbCrLf
sql = sql & "where t.MRP_TYPE = 2 order by Row" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- test" & vbCrLf
sql = sql & "--select * from @table where Part = 'BACB30VT6K3'" & vbCrLf
sql = sql & "--order by Row" & vbCrLf
ExecuteScript True, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function


Private Function UpdateDatabase199()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 199     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

' COMSOL routing approval
sql = "if not exists (select * from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = 'ComnTable'" & vbCrLf
sql = sql & "and COLUMN_NAME = 'CoRequireApprovedRoutings')" & vbCrLf
sql = sql & "alter table ComnTable add CoRequireApprovedRoutings tinyint default 0 not null" & vbCrLf
ExecuteScript True, sql


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase200()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 200     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

' Early/Late PO update
sql = "dropstoredprocedureifexists 'RptEarlyLateDatesPO'" & vbCrLf
ExecuteScript True, sql

sql = "CREATE PROCEDURE RptEarlyLateDatesPO" & vbCrLf
sql = sql & "@StartDate as varchar(16)," & vbCrLf
sql = sql & "@EndDate as Varchar(16)" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "v 4 7/15/2020 - Fixed a problem where POs showed that they could be moved to the last requirement" & vbCrLf
sql = sql & "when that was not the case." & vbCrLf
sql = sql & "v 3 5/6/2020 - Add Safety Stock to calculations" & vbCrLf
sql = sql & "v 2 3/24/2020" & vbCrLf
sql = sql & "replace old flawed triple-cursor based report." & vbCrLf
sql = sql & "this is still very slow, but hopefully a lot more accurate" & vbCrLf
sql = sql & "TEST:" & vbCrLf
sql = sql & "exec RptEarlyLateDatesPO '1/1/2020', '12/31/2020'" & vbCrLf
sql = sql & "exec RptEarlyLateDatesPO '12/6/2019', '1/9/2020'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- get MRP creation date" & vbCrLf
sql = sql & "declare @MrpDate date" & vbCrLf
sql = sql & "SELECT @MrpDate = MRP_CREATEDATE FROM MrpdTable WHERE MRP_ROW=1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @table table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "Part varchar(30) NULL," & vbCrLf
sql = sql & "[Req Date] [date] NULL," & vbCrLf
sql = sql & "[Sched Date] [date] NULL," & vbCrLf
sql = sql & "[Days Early] [int] NULL," & vbCrLf
sql = sql & "[Quantity] [decimal](12, 4) NULL," & vbCrLf
sql = sql & "[Balance] [decimal](38, 4) NULL," & vbCrLf
sql = sql & "Comment varchar(80) NULL," & vbCrLf
sql = sql & "MRP_TYPE tinyint NULL," & vbCrLf
sql = sql & "MRP_Row int," & vbCrLf
sql = sql & "Row int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @table" & vbCrLf
sql = sql & "SELECT rtrim(MRP_PARTREF) as Part, cast(null as date) as [Req Date]," & vbCrLf
sql = sql & "case when MRP_TYPE = 1 THEN cast('1/1/1900' as date) -- force PAQOH to top" & vbCrLf
sql = sql & "when MRP_TYPE = 17 THEN @MrpDate               -- use MRP date for safety stock reqt" & vbCrLf
sql = sql & "else MRP_PARTDATERQD end as [Sched Date],         -- otherwise use required date" & vbCrLf
sql = sql & "cast(null as int) as [Days Early], MRP_PARTQTYRQD as Quantity," & vbCrLf
sql = sql & "sum(MRP_PARTQTYRQD) OVER (partition by MRP_PARTREF" & vbCrLf
sql = sql & "order by case when MRP_TYPE = 1 THEN cast('1/1/1900' as date) -- force PAQOH to top" & vbCrLf
sql = sql & "when MRP_TYPE = 17 THEN @MrpDate               -- use MRP date for safety stock reqt" & vbCrLf
sql = sql & "else MRP_PARTDATERQD end, MRP_ROW) as Balance," & vbCrLf
sql = sql & "case when MRP_TYPE = 17 then 'Safety Stock' else MRP_COMMENT end as Comment, MRP_TYPE, MRP_ROW," & vbCrLf
sql = sql & "ROW_NUMBER() over (order by MRP_PARTREF," & vbCrLf
sql = sql & "case when MRP_TYPE = 1 THEN cast('1/1/1900' as date) -- force PAQOH to top" & vbCrLf
sql = sql & "when MRP_TYPE = 17 THEN @MrpDate               -- use MRP date for safety stock reqt" & vbCrLf
sql = sql & "else MRP_PARTDATERQD end,       -- otherwise use required date" & vbCrLf
sql = sql & "MRP_ROW)" & vbCrLf
sql = sql & "FROM PartTable" & vbCrLf
sql = sql & "JOIN MrplTable ON PartTable.PARTREF=MRP_PARTREF" & vbCrLf
sql = sql & "WHERE MRP_TYPE not in (5,6) and PALEVEL = 4 -- and MRP_PARTREF = dbo.fnCompress('15-5R 2.250 DIA AMS5659')" & vbCrLf
sql = sql & "and exists (select * from MrplTable mrp2 where mrp2.MRP_PARTREF = MrplTable.MRP_PARTREF and mrp2.MRP_TYPE = 2" & vbCrLf
sql = sql & "and MRP_PARTDATERQD BETWEEN @StartDate AND  @EndDate)" & vbCrLf
sql = sql & "and (MRP_PARTDATERQD BETWEEN @StartDate AND  @EndDate or MRP_TYPE in (1,17))" & vbCrLf
sql = sql & "ORDER BY MRP_PARTREF, case when MRP_TYPE = 1 THEN cast('1/1/1900' as date)" & vbCrLf
sql = sql & "when MRP_TYPE = 17 THEN @MrpDate" & vbCrLf
sql = sql & "else MRP_PARTDATERQD end," & vbCrLf
sql = sql & "MRP_ROW" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from @table ORDER BY [Sched Date], MRP_ROW" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-------------------------------------------------" & vbCrLf
sql = sql & "-- determine needed date for parts needed earlier" & vbCrLf
sql = sql & "-------------------------------------------------" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- find first pos preceded by a negative quantity and move po up if possible" & vbCrLf
sql = sql & "declare @candidates table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "Part varchar(30)," & vbCrLf
sql = sql & "FirstMinusRow int," & vbCrLf
sql = sql & "MinusDate date," & vbCrLf
sql = sql & "FirstPoAfter int" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @ct int, @s as varchar(30)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set nocount on" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "declare @i int" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @i = 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "while @i < 50" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "delete from @candidates" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @candidates" & vbCrLf
sql = sql & "select Part, Min(Row) as FirstMinusRow, cast(null as date) as [Sched Date], null as FirstPoAfter from @table" & vbCrLf
sql = sql & "where MRP_TYPE <> 2 and Balance < 0" & vbCrLf
sql = sql & "group by PART" & vbCrLf
sql = sql & "order by Part" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @candidates set MinusDate = (select [Sched Date] from @table where Row = [@candidates].FirstMinusRow)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @candidates set FirstPoAfter" & vbCrLf
sql = sql & "= (select min(Row) from @table t where MRP_TYPE = 2 and t.Part = [@candidates].Part and t.[Req Date] is null" & vbCrLf
sql = sql & "and t.[Sched Date] > MinusDate)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select @ct = count(FirstPoAfter) from @candidates" & vbCrLf
sql = sql & "--set @s = 'minus loop ' + cast(@i as varchar(2)) + ': count ' + cast(@ct as varchar(10))" & vbCrLf
sql = sql & "--print @s" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @ct = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from @candidates" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update t set [Req Date] = MinusDate" & vbCrLf
sql = sql & "from @candidates c join @table t on t.Part = c.Part and t.Row = c.FirstPoAfter" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- recalculate and renumber" & vbCrLf
sql = sql & "update x set Row = New_Row" & vbCrLf
sql = sql & "from" & vbCrLf
sql = sql & "(select Row, ROW_NUMBER() over (" & vbCrLf
sql = sql & "order by Part, isnull([Req Date], [Sched Date]), MRP_TYPE, MRP_Row) as New_Row" & vbCrLf
sql = sql & "from @table) x" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update x set Balance = New_Bal from" & vbCrLf
sql = sql & "(select Balance, sum(Quantity) over (partition by Part order by Row) as New_Bal from @table) x" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from @candidates order by Part" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @i = @i + 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select 'earlier', * from @table ORDER BY [Sched Date], MRP_ROW" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-----------------------------------------------------------" & vbCrLf
sql = sql & "-- determine needed date for parts not required until later" & vbCrLf
sql = sql & "-----------------------------------------------------------" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- find last PO for each part and move down if possible" & vbCrLf
sql = sql & "declare @candidates2 table" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "Part varchar(30)," & vbCrLf
sql = sql & "LastMoveToRow int," & vbCrLf
sql = sql & "LastMoveToDate date," & vbCrLf
sql = sql & "LastPo int," & vbCrLf
sql = sql & "LastPoQty decimal(12,4)," & vbCrLf
sql = sql & "LastPoDate date" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @i = 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "while @i <= 50" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "delete from @candidates2" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert @candidates2 (Part, LastPO)" & vbCrLf
sql = sql & "select Part, Max(Row) from @table" & vbCrLf
sql = sql & "where MRP_TYPE = 2 and [Req Date] is null" & vbCrLf
sql = sql & "group by PART" & vbCrLf
sql = sql & "order by Part" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update c set LastPoQty = Quantity, LastPoDate = [Sched Date]" & vbCrLf
sql = sql & "from @candidates2 c join @table t on c.LastPo = t.Row" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @candidates2 set LastMoveToRow = (select min(Row) from @table t" & vbCrLf
sql = sql & "where t.Part = [@candidates2].part and t.Row > [@candidates2].LastPo and t.Balance < [@candidates2].LastPoQty)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @candidates2 set LastMoveToDate = (select [Sched Date] from @table where Row = [@candidates2].LastMoveToRow)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select @ct = count(*) from @candidates2" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if @ct = 0 break" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update t set [Req Date] =" & vbCrLf
sql = sql & "case when LastMoveToDate is not null then LastMoveToDate" & vbCrLf
sql = sql & "when Balance >= Quantity then DATEADD(day, 9999,LastPoDate)" & vbCrLf
sql = sql & "else [Sched Date] end" & vbCrLf
sql = sql & "from @candidates2 c join @table t on t.Part = c.Part and t.Row = c.LastPo" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- recalculate and renumber" & vbCrLf
sql = sql & "update x set Row = New_Row" & vbCrLf
sql = sql & "from" & vbCrLf
sql = sql & "(select Row, ROW_NUMBER() over (" & vbCrLf
sql = sql & "order by Part, isnull([Req Date], [Sched Date]), MRP_TYPE, MRP_Row) as New_Row" & vbCrLf
sql = sql & "from @table) x" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update x set Balance = New_Bal from" & vbCrLf
sql = sql & "(select Balance, sum(Quantity) over (partition by Part order by Row) as New_Bal from @table) x" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select * from @candidates2 order by Part" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @i = @i + 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--select 'later', * from @table ORDER BY [Sched Date], MRP_ROW" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-------------------------------------------" & vbCrLf
sql = sql & "-- calculate days early/late" & vbCrLf
sql = sql & "-------------------------------------------" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update @table set [Days Early] = datediff(d,[Sched Date], [Req Date]) where [Req Date] is not null" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- data required for report" & vbCrLf
sql = sql & "select" & vbCrLf
sql = sql & "MRP_PARTREF,MRP_PARTNUM, t.MRP_TYPE, MRP_PONUM, MRP_POITEM, MRP_PARTLEVEL, t.[Req Date] as MRP_HGHPRTDTRQD," & vbCrLf
sql = sql & "t.[Sched Date] as MRP_PARTDATERQD, MRP_PARTQTYRQD, MRP_CATAGORY, MRP_COMMENT," & vbCrLf
sql = sql & "[Days Early] as EARLYLATE_DAYS" & vbCrLf
sql = sql & "from @table t" & vbCrLf
sql = sql & "join MrplTable m on m.MRP_ROW = t.MRP_Row" & vbCrLf
sql = sql & "where t.MRP_TYPE = 2" & vbCrLf
sql = sql & "order by Row" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript True, sql



''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function


Private Function UpdateDatabase201()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 201     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

' MANSER Costed BOM Report
sql = "exec DropStoredProcedureIfExists 'RptCostedBOM'" & vbCrLf
ExecuteScript True, sql

sql = "create procedure RptCostedBOM" & vbCrLf
sql = sql & "@Assembly VARCHAR(30)," & vbCrLf
sql = sql & "@Rev varchar(10)" & vbCrLf
sql = sql & "as" & vbCrLf
sql = sql & "/*" & vbCrLf
sql = sql & "MANSER Costed BOM Report - engbm08a.rpt" & vbCrLf
sql = sql & "Test:" & vbCrLf
sql = sql & "exec RptCostedBOM 'ACC-30806-TD-001'" & vbCrLf
sql = sql & "*/" & vbCrLf
sql = sql & "set @Assembly = dbo.fnCompress(@Assembly)" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "if exists(select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_RptCostedBOM')" & vbCrLf
sql = sql & "drop table _RptCostedBOM" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "create table _RptCostedBOM" & vbCrLf
sql = sql & "(" & vbCrLf
sql = sql & "SortKey varchar(300)," & vbCrLf
sql = sql & "Level int," & vbCrLf
sql = sql & "PartRef varchar(30)," & vbCrLf
sql = sql & "[Part Number] varchar(30)," & vbCrLf
sql = sql & "Description varchar(30)," & vbCrLf
sql = sql & "[Ext Description] varchar(3072)," & vbCrLf
sql = sql & "[Qty Required] decimal(12,4)," & vbCrLf
sql = sql & "[Set Qty] decimal(12,4)," & vbCrLf
sql = sql & "UoM varchar(2)," & vbCrLf
sql = sql & "[Qty on Hand] decimal(12,4)," & vbCrLf
sql = sql & "[Net on Hand] decimal(12,4)," & vbCrLf
sql = sql & "[Last Vendor] varchar(10)," & vbCrLf
sql = sql & "[Last Cost] decimal(12,4)" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert the base level" & vbCrLf
sql = sql & "insert _RptCostedBOM" & vbCrLf
sql = sql & "select '', 0, rtrim(PARTREF), RTRIM(PARTNUM),RTRIM(PADESC), RTRIM(PAEXTDESC), 1, 0, PAUNITS, PAQOH, 0, '', 0 from PartTable p" & vbCrLf
sql = sql & "join BmhdTable bh on bh.BMHREF = p.PARTREF" & vbCrLf
sql = sql & "where bh.BMHREF = @Assembly and bh.bmhrev = @Rev" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- insert the explosion levels" & vbCrLf
sql = sql & "declare @level int = 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "while @level < 10" & vbCrLf
sql = sql & "begin" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "insert _RptCostedBOM" & vbCrLf
sql = sql & "select bom.SortKey + p.PARTREF, @level, RTRIM(p.PARTREF), RTRIM(p.PARTNUM), RTRIM(p.PADESC), RTRIM(p.PAEXTDESC)," & vbCrLf
sql = sql & "b.BMQTYREQD, b.BMSETUP, PAUNITS, PAQOH, 0, '', 0" & vbCrLf
sql = sql & "from _RptCostedBOM bom" & vbCrLf
sql = sql & "join BmplTable b on b.BMASSYPART = bom.PartRef" & vbCrLf
sql = sql & "join PartTable p on p.PARTREF = b.BMPARTREF" & vbCrLf
sql = sql & "where bom.Level = @level - 1" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "set @level = @level + 1" & vbCrLf
sql = sql & "end" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- calculate MRP net balance for all items" & vbCrLf
sql = sql & "update _RptCostedBOM set [Net on Hand]" & vbCrLf
sql = sql & "= isnull((select sum(MRP_PARTQTYRQD) from MrplTable where MRP_PARTREF = _RptCostedBOM.PartRef),0)" & vbCrLf
sql = sql & "-- BAT00501900 -- 115" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "-- update last vendor and last cost" & vbCrLf
sql = sql & "-- this is kludgy but it is impossible to use select top 1 in a subquery" & vbCrLf
sql = sql & "IF OBJECT_ID('tempdb..#temp') IS NOT NULL" & vbCrLf
sql = sql & "DROP TABLE #temp" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "select PIPART, max(PIPDATE) as PIPDATE, cast('' as varchar(30)) as [Last Vendor], cast(0 as decimal(12,4)) as [Last Cost]" & vbCrLf
sql = sql & "into #temp" & vbCrLf
sql = sql & "from PoitTable where PIPART in (select PartRef from _RptCostedBOM) group by PIPART" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update #temp set [Last Vendor] = POVENDOR, [Last Cost] = PIESTUNIT" & vbCrLf
sql = sql & "from #temp" & vbCrLf
sql = sql & "join PoitTable poi on POI.PIPART = #TEMP.PIPART" & vbCrLf
sql = sql & "join PohdTable po on po.PONUMBER = poi.PINUMBER" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "update r" & vbCrLf
sql = sql & "set r.[Last Vendor] = t.[Last Vendor], r.[Last Cost] = t.[Last Cost]" & vbCrLf
sql = sql & "from _RptCostedBOM r" & vbCrLf
sql = sql & "join #temp t on t.PIPART = r.PartRef" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "--update _RptCostedBOM" & vbCrLf
sql = sql & "--   set [Last Vendor] = POVENDOR, [Last Cost] = Cost" & vbCrLf
sql = sql & "--   from (select top 1 POVENDOR, case when PIAMT = 0 then PIESTUNIT else PIAMT end as Cost" & vbCrLf
sql = sql & "--      from PohdTable po" & vbCrLf
sql = sql & "--      join PoitTable poi on po.PONUMBER = poi.PINUMBER" & vbCrLf
sql = sql & "--      where poi.PIPART = _RptCostedBOM.PartRef" & vbCrLf
sql = sql & "--      order by isnull(PIADATE,PIPDATE) desc) x" & vbCrLf
sql = sql & "" & vbCrLf
sql = sql & "SELECT [Level]" & vbCrLf
sql = sql & ",[PartRef]" & vbCrLf
sql = sql & ",[Part Number]" & vbCrLf
sql = sql & ",[Description]" & vbCrLf
sql = sql & ",[Ext Description]" & vbCrLf
sql = sql & ",[Qty Required]" & vbCrLf
sql = sql & ",[Set Qty]" & vbCrLf
sql = sql & ",[UoM]" & vbCrLf
sql = sql & ",[Qty on Hand]" & vbCrLf
sql = sql & ",[Net on Hand]" & vbCrLf
sql = sql & ",[Last Vendor]" & vbCrLf
sql = sql & ",[Last Cost]" & vbCrLf
sql = sql & "FROM _RptCostedBOM" & vbCrLf
sql = sql & "order by SortKey" & vbCrLf
sql = sql & "" & vbCrLf
ExecuteScript True, sql



''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function

Private Function UpdateDatabase202()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 202     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''

' option to show abbreviated lot numbers for COMSOL
sql = "AddOrUpdateColumn 'ComnTable', 'CoUseAbbreviatedLotNumbers', 'tinyint not null default 0'" & vbCrLf
ExecuteScript True, sql



''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function


Private Function UpdateDatabaseXXX()

'update database version template
'set version at top of this file
'set version below
'add SQL updates


   Dim sql As String
   sql = ""

   newver = 999     ' set actual version
   If ver < newver Then

      clsADOCon.ADOErrNum = 0

''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''

      ' update the version
      ExecuteScript False, "Update Version Set Version = " & newver

   End If
End Function










''''''''''''''''''''''''''''''''''

   
Private Function IsUpdateRequired(OldDbVersion As Integer, NewDbVersion As Integer) As Boolean
   'terminates if cannot proceed.
   'returns false if no update required
   'returns true if update required and authorized by and admin

   Err.Clear
   'Dim strFulVer As String
   clsADOCon.ADOErrNum = 0
   sSql = "select * from Updates" & vbCrLf _
      & "where UpdateID = (select max(UpdateID) from Updates)"
   On Error Resume Next
   Dim rdo As ADODB.Recordset
   If clsADOCon.GetDataSet(sSql, rdo) Then
      If clsADOCon.ADOErrNum = 0 Then
         oldRelease = rdo!newRelease
         If oldRelease > 9999 Then
            oldRelease = oldRelease / 100
            rdo.Close
            sSql = "alter table updates add V19OldRelease int"
            clsADOCon.ExecuteSql (sSql)
            sSql = "alter table updates add V19NewRelease int"
            clsADOCon.ExecuteSql (sSql)
            sSql = "update Updates set V19OldRelease = OldRelease, V19NewRelease = NewRelease"
            clsADOCon.ExecuteSql (sSql)
            sSql = "update Updates set OldRelease = OldRelease /100 where OldRelease > 9999"
            clsADOCon.ExecuteSql (sSql)
            sSql = "update Updates set NewRelease = NewRelease /100 where NewRelease > 9999"
            clsADOCon.ExecuteSql (sSql)
         End If
      End If
   Else
      oldRelease = 0
   End If
   
   If IsTestDatabase Then
      oldType = "Test"
   Else
      oldType = "Live"
   End If
   
'   If App.Minor < 10 Then
'    strFulVer = CStr(App.Major) & "0" & CStr(App.Minor)
'   Else
'    strFulVer = CStr(App.Major) & CStr(App.Minor)
'   End If
'
'   newRelease = CInt(strFulVer)
   
   ' 1/2/2020 - no longer consider revision of an exe as a release
   'newRelease = 10000& * App.Major + 100& * App.Minor + App.Revision
   newRelease = 100& * App.Major + App.Minor
   
   If InTestMode() Then
      NewType = "Test"
   Else
      NewType = "Live"
   End If
   
   If oldType <> NewType And NewType = "Test" Then
        MsgBox "You cannot update a Live database with a Test Application", vbCritical
        End
   End If
    
   Dim msg As String
   If oldRelease < newRelease Then
      msg = "Old release (" & oldRelease & ") < New release (" & newRelease & ")" & vbCrLf
   ElseIf oldRelease > newRelease Then
      MsgBox "Old release (" & oldRelease & ") > New release (" & newRelease & ")" & vbCrLf _
         & "You cannot proceed.", vbCritical
      End
   End If
   
   If OldDbVersion < NewDbVersion Then
      msg = msg & "Old db version (" & OldDbVersion & ") < New db version (" & NewDbVersion & ")" & vbCrLf
   ElseIf OldDbVersion > NewDbVersion Then
      MsgBox "Old db version (" & OldDbVersion & ") > New db version (" & NewDbVersion & ")" & vbCrLf _
         & "You cannot proceed.", vbCritical
      End
   End If
   
   If oldType <> NewType Then
      msg = msg & "Db type (" & oldType & ") <> Application type (" & NewType & ")" & vbCrLf
   End If
   
   If msg = "" Then
      IsUpdateRequired = False
      Exit Function
   
   Else
      If Secure.UserAdmn = 0 Then
         msg = msg & "An administrator must perform an update before you can proceed."
         MsgBox msg, vbCritical
         End
      Else
         msg = msg & "Do you wish to update the database now?"
         Select Case MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2)
         Case vbYes
            IsUpdateRequired = True
         Case Else
            End
         End Select
      End If
   End If
   
End Function



Private Function GetConstraintNme(sTableNme As String, sPartialConstraintNme As String) As String
    Dim RdoCon As ADODB.Recordset
    
    On Error GoTo GCNErr1
    sSql = "select O.name as ConstraintName from sysobjects AS O  left join sysobjects AS T on O.parent_obj = T.id " & vbCrLf _
           & "where isnull(objectproperty(O.id,'IsMSShipped'),1) = 0 and O.name not like '%dtproper%' and O.name not like 'dt[_]%'" & vbCrLf _
           & "and T.name = '" & sTableNme & "' and O.name like '" & sPartialConstraintNme & "%'"
  
   GetConstraintNme = ""
   
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoCon, ES_FORWARD)
    If bSqlRows Then
        GetConstraintNme = "" & RdoCon!ConstraintName
    End If
    Set RdoCon = Nothing
    Exit Function
   
GCNErr1:
    On Error GoTo 0
    GetConstraintNme = ""

End Function

Private Function GetConstrinNameforColumn(sTableNme As String, sColName As String) As String
    Dim RdoCon As ADODB.Recordset
    
    On Error GoTo GCNErr1
  
    sSql = "SELECT c_obj.name" & vbCrLf _
           & "   FROM sysobjects t_obj, sysobjects c_obj," & vbCrLf _
           & "   syscolumns Cols" & vbCrLf _
           & "Where Cols.id = t_obj.id And c_obj.id = Cols.cdefault" & vbCrLf _
           & "   AND t_obj.name = '" & sTableNme & "' AND cols.name = '" & sColName & "'"
   
   GetConstrinNameforColumn = ""
   
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoCon, ES_FORWARD)
    If bSqlRows Then
        GetConstrinNameforColumn = "" & RdoCon!Name
    End If
    Set RdoCon = Nothing
    Exit Function
   
GCNErr1:
    On Error GoTo 0
    GetConstrinNameforColumn = ""

End Function

Private Function NormalizeConstraintNme(sTableNme As String, sOldObjName As String, sNewObjName As String) As Boolean
    Dim RdoCon As ADODB.Recordset
    Dim sOldNme As String
    
    On Error GoTo NSONErr
    If GetConstraintNme(sTableNme, sNewObjName) = sNewObjName Then
        NormalizeConstraintNme = True
        Exit Function
    End If
    
    
    sOldNme = GetConstraintNme(sTableNme, sOldObjName)
    
    ExecuteScript True, "exec sp_rename '" & sOldNme & "','" & sNewObjName & "',OBJECT"
    NormalizeConstraintNme = True
    Exit Function
    
NSONErr:
    On Error GoTo 0
    NormalizeConstraintNme = False

End Function

Private Function CheckSeedDataExits(sSeedSql As String)
    
   Dim RdoCon As ADODB.Recordset
   
   On Error GoTo ERR1
      
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCon, ES_FORWARD)
   If bSqlRows Then
       CheckSeedDataExits = True
   Else
       CheckSeedDataExits = False
   End If
   Set RdoCon = Nothing
   Exit Function
   
ERR1:
    CheckSeedDataExits = False

End Function



Private Function StoredProcedureExists(sProcNme As String) As Boolean
   Dim RdoProc As ADODB.Recordset
   Dim iProcedureID As Double
   
   
   On Error GoTo modErr1
   
   StoredProcedureExists = False
   sSql = "SELECT OBJECT_ID('" & sProcNme & "') AS StoredProcID"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoProc, ES_FORWARD)
   If bSqlRows Then
    iProcedureID = Val("" & RdoProc!StoredProcID)
    If iProcedureID > 0 Then StoredProcedureExists = True
   End If
   Set RdoProc = Nothing
   Exit Function
   
modErr1:
   On Error GoTo 0
   StoredProcedureExists = False
End Function


