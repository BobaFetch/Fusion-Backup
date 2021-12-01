Attribute VB_Name = "Databases"
Option Explicit
' 2/20/2009 Changed Database version from 61 to 62
'Public Const DB_VERSION = 62

' MM 7/24/2009 Changed Database version from 62 to 63
' MM 8/2/2009 Changed Database version from 63 to 64
' MM 11/15/2009 Changed Database version from 64 to 65
' MM 1/24/2010 Changed Database version from 65 to 66
' MM 3/28/2010 Changed Database version from 66 to 67
' MM 5/6/2010 Changed Database version from 67 to 68
' MM 6/27/2010 Changed Database version from 68 to 69
' BBS 7/1/2010 Changed Database version from 69 to 70
Public Const DB_VERSION = 70

Public sSaAdmin As String
Public sSaPassword As String
Public sserver As String
Public sDataBase As String
Public sSysCaption As String
Public sProgName As String
Public sDsn As String

'version info
Private oldType As String     'live or test
Private NewType As String
Private oldRelease As Integer
Private newRelease As Integer
Private OldDbVersion As Integer
Private directory As String
Private newver As Integer

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
Private ver As Integer

Function OpenSqlServer(Optional bReStart As Boolean) As Boolean
   'return = True if successful
   
   Dim RdoCheck As rdoResultset
   Dim b As Byte
   Dim CloseExit As Long
   Dim iTimeOut As Integer
   Dim sWindows As String
   
   On Error GoTo PrjOs1
   MouseCursor ccHourglass
   sWindows = GetWindowsDir()
   'sserver = UCase$(GetSetting("Esi2000", "System", "ServerId", sserver))
   sserver = UCase(GetUserSetting(USERSETTING_ServerName))
   sSaAdmin = Trim(GetSysLogon(True))
   sSaPassword = Trim(GetSysLogon(False))
   sSysCaption = GetSystemCaption()
   GetCurrentDatabase
   
   Set RdoEnv = rdoEnvironments(0)
   RdoEnv.CursorDriver = rdUseIfNeeded
   
   Set RdoCon = New rdoConnection
   RdoCon.QueryTimeout = 60
   iTimeOut = SQLSetConnectOption(RdoCon.hdbc, SQL_PRESERVE_CURSORS, SQL_PC_ON)
   RdoCon.Connect = "UID=" & sSaAdmin & ";PWD=" & sSaPassword & ";DRIVER={SQL Server};" _
                    & "SERVER=" & sserver & ";DATABASE=" & sDataBase & ";"
   RdoCon.EstablishConnection rdDriverNoPrompt
   
   'RdoCon.Execute "alter table Version add TestDatabase tinyint not null default 0"
   'Err.Clear
   SaveSetting "Esi2000", "System", "CloseSection", ""
   iTimeOut = RdoCon.QueryTimeout
   If iTimeOut < 60 Then RdoCon.QueryTimeout = 60
   
   'if test version of the executables, allow only a test vesion of the database
'   On Error Resume Next
'   sSql = "select TestDatabase from Version"
'   Dim rdo As rdoResultset
'   Dim TestDatabase As Boolean
'   If GetDataSet(rdo) Then
'      If Err = 0 Then
'         TestDatabase = IIf(rdo.rdoColumns(0) = 1, True, False)
'      Else
'         RdoCon.Execute "alter table Version add TestDatabase tinyint not null default 0"
'      End If
'   End If
'   On Error GoTo PrjOs1
   
   Dim TestDatabase As Boolean
   TestDatabase = IsTestDatabase()
   
'do this in IsUpdateRequired
'   if not Runninginide Then
'      If InTestMode() Then
'         If Not TestDatabase Then
'            MouseCursor ccDefault
'            MsgBox "You cannot run a test application against a live database."
'            Exit Function
'         End If
'      Else
'         If TestDatabase Then
'            MouseCursor ccDefault
'            MsgBox "You cannot run a live application against a test database."
'            Exit Function
'         End If
'      End If
'   End If
   
   If bReStart = 0 Then
      'Get Count of Parts to see how combo's are to be handled
      sSql = "SELECT COUNT(PARTREF) FROM PartTable"
      bSqlRows = GetDataSet(RdoCheck, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoCheck.rdoColumns(0)) Then _
                       ES_PARTCOUNT = RdoCheck.rdoColumns(0) Else _
                       ES_PARTCOUNT = 0
      End If
      'GetDataBases
      UpdateTables
' Mohan Commented
      b = CheckSecuritySettings()
'      'If b = 0 Then GetSectionPermissions
      GetCompany 1
      On Error GoTo ModErr1
      GetCustomerPermissions
      'MdiSect.Caption = MdiSect.Caption & " - " & sDataBase
      MDISect.Caption = GetSystemCaption
   Else
      'MouseCursor ccArrow
   End If
   
   MouseCursor ccDefault
   
   'For Local Testing of Customs
   'If sServer = "SomeLocalServer" Then ES_CUSTOM = "PROPLA"
   'If sServer = "SomeOtherServer" Then ES_CUSTOM = "WATERJET"
   MDISect.SystemMsg = ""
   OpenSqlServer = True
   Exit Function
   
ModErr1:
   Resume modErr2
modErr2:
   'Couldn't open msdb
   For b = 1 To 8
      bCustomerGroups(b) = 1
   Next
   Exit Function
   
PrjOs1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume PrjOs2
PrjOs2:
   MouseCursor ccArrow
   On Error GoTo 0
   MsgBox LTrim(Str(CurrError.Number)) & vbCrLf & CurrError.Description & vbCrLf _
                & "Unable To Make SQL Server Connection.", 48, "Fusion ERP"
End

End Function

'''''''''''''''''''''''''

Function GetTemporarySqlConnection() As rdoConnection

   'creates a seaprate connection and executes the sql
   'this function is designed to store error information in the SystemEvents table
   'be sure to set the returned connection = nothing after use
  
   Dim rdoTemp As rdoConnection
   Set rdoTemp = New rdoConnection
   rdoTemp.QueryTimeout = 60
   rdoTemp.Connect = "UID=" & sSaAdmin & ";PWD=" & sSaPassword & ";DRIVER={SQL Server};" _
                    & "SERVER=" & sserver & ";DATABASE=" & sDataBase & ";"
   rdoTemp.EstablishConnection rdDriverNoPrompt
   Set GetTemporarySqlConnection = rdoTemp
End Function



'''''''''''''''''''''''''''
Public Sub UpdateDatabase()

   
   'Update the database to the current version
   
   'note the following differences between SQL2000 and SQL2005.  This code must work with both
   '1.  Index drops are different - use both
   '     DROP INDEX DdocTable.DoClass     -- SQL2000 - deprecated in SQL2005
   '     DROP INDEX DoClass ON DdocTable  -- SQL2005 - this will be the syntax going forward
   '     ... or use Sub DropIndex, which attemps both
   
   '2. SQL2000 doesn't allow NOT NULL columns without a default
   '     ALTER TABLE TempPsLots ADD PartRef varchar(30) NOT NULL              -- fails in SQL2000
   '     ALTER TABLE TempPsLots ADD PartRef varchar(30) NOT NULL default ''   -- works in both
   
   '3. TOP 100 PERCENT AND MORE GENERALLY, DON'T PUT PARENS AROUND NUMBERS
   '     SQL2000 doesn't lke "SELECT TOP (100) PERCENT"  -- GENERATED BY Larry's SQL tool
   '     USE: "SELECT TOP 100 PERCENT"
   
   
   Dim rdo As rdoResultset
'   Dim newver As Integer

   
   'if no version table, create it and set it to version 0
   'sSql = "select max(Version) as Version from Version"
   sSql = "select Version from Version"
   On Error Resume Next 'need to attempt all steps even if they fail
   bSqlRows = GetDataSet(rdo)
   If Err.Number <> 0 Then
      bSqlRows = False
      Err.Clear
   End If
   If bSqlRows Then
      OldDbVersion = rdo!Version
      ver = rdo!Version
   Else
      RdoCon.Execute ("create table Version ( Version  int )")
      RdoCon.Execute ("insert Version ( Version ) values ( 0 )")
      ver = 0
   End If
Debug.Print "Initial database version " & ver
   Set rdo = Nothing
   
   'Continue if:
   '1. Update required and user is an admin
   '2. User is a developer running in VB IDE (to debug db updates)
   Dim UpdateReqd As Boolean
   If IsUpdateRequired(OldDbVersion, DB_VERSION) Then
      UpdateReqd = True
   Else
      If Not RunningInIDE Then
         Exit Sub
      End If
   End If
   
   Err.Clear
   On Error Resume Next
   
   MouseCursor ccHourglass
   SysMsgBox.tmr1.Enabled = False
   SysMsgBox.msg = "Updating database."
   SysMsgBox.Show
   
   'allow really long timeouts (normal limit is 60 sec)
   Dim timeout As Integer
   timeout = RdoCon.QueryTimeout
   RdoCon.QueryTimeout = 1200    '20 minutes
   
   If ver < 1 Then
      'create procedure to drop default values so columns can be altered
      sSql = _
             "create procedure DropColumnDefault" & vbCrLf _
             & "@TableName varchar(50)," & vbCrLf _
             & "@ColumnName varchar(50)" & vbCrLf _
             & "as" & vbCrLf _
             & "declare @constraint_name sysname" & vbCrLf _
             & "While 1 = 1" & vbCrLf _
             & "begin" & vbCrLf _
             & "    set @constraint_name = ( select top 1 c_obj.name" & vbCrLf _
             & "    from sysobjects  t_obj" & vbCrLf _
             & "    inner join sysobjects c_obj on t_obj.id = c_obj.parent_obj" & vbCrLf _
             & "    inner join syscolumns  cols on cols.colid = c_obj.info" & vbCrLf _
             & "    and cols.id = c_obj.parent_obj" & vbCrLf _
             & "    where t_obj.name = @TableName" & vbCrLf _
             & "    and c_obj.xtype = 'D'" & vbCrLf _
             & "    and cols.[name] = @ColumnName )" & vbCrLf _
             & "    if @constraint_name is null break" & vbCrLf _
             & "    exec ('alter table ' + @TableName + ' drop constraint ' + @constraint_name)" & vbCrLf _
             & "End"
      RdoCon.Execute sSql
      
      '       SystemEvents table is now in local database
      '        'allow large error descriptions including things like SQL and Crystal Reports formulas
      '        RdoCon.Execute "use msdb"
      '        RdoCon.Execute "ALTER TABLE dbo.SystemEvents ALTER COLUMN Event_Text varchar(4096)"
      '        RdoCon.Execute "use " & sDataBase
      
      'convert PsitTable.PISELLPRICE FROM real to decimal(12,4).  Remove default first
      RdoCon.Execute "exec DropColumnDefault 'PsitTable', 'PISELLPRICE'"
      RdoCon.Execute "alter table PsitTable alter column PISELLPRICE decimal(12,4) null"
      RdoCon.Execute "alter table PsitTable add constraint PISELLPRICE_Default default ( 0.00 ) for PISELLPRICE"
      
      'basic security is no longer used
      RdoCon.Execute "EXEC sp_rename 'UsscTable', 'UsscTable_Obsolete'"
      
      'set version = 1
      RdoCon.Execute "update Version set Version = 1"
   End If
   
   If ver < 2 Then
      'create SystemEvents table in local database
      'it wasn't always getting created in msdb
      sSql = "CREATE TABLE SystemEvents(" & vbCrLf _
             & "Event_ID int IDENTITY (1, 1) NOT NULL," & vbCrLf _
             & "Event_Date char(18) NULL," & vbCrLf _
             & "Event_Section char(10) NULL," & vbCrLf _
             & "Event_Form char(20)," & vbCrLf _
             & "Event_User char(30)," & vbCrLf _
             & "Event_Event int NULL," & vbCrLf _
             & "Event_Warning tinyint NULL," & vbCrLf _
             & "Event_Procedure char(20) NULL," & vbCrLf _
             & "Event_Text varchar(4096) NULL )"
      RdoCon.Execute sSql
      
      'set version
      RdoCon.Execute "update Version set Version = 2"
   End If
   
   If ver < 3 Then
      'this constraint prevented temporary invoices from being generated
      sSql = "ALTER TABLE CihdTable DROP CONSTRAINT FK_CihdTable_CustTable"
      RdoCon.Execute sSql
      
      'set version
      RdoCon.Execute "update Version set Version = 3"
   End If
   
   If ver < 4 Then
      'fix primary key and index problems at Jevco China and other places
      'RdoCon.Execute "drop index Cihdtable.Cihd_Idx"
      DropIndex "Cihdtable", "Cihd_Idx"
      RdoCon.Execute "alter table Cihdtable drop constraint PK_Cihdtable_CHECKNUMBER"
      RdoCon.Execute "alter table Cihdtable WITH NOCHECK ADD CONSTRAINT PK_CihdTable PRIMARY KEY CLUSTERED(INVNO)"
      RdoCon.Execute "create index InvCust ON Cihdtable(INVCUST)"
      RdoCon.Execute "create index InvPref ON Cihdtable(INVPRE)"
      RdoCon.Execute "create index InvSon ON Cihdtable(INVSO)"
      
      'set version
      RdoCon.Execute "update Version set Version = 4"
   End If
   
   If ver < 5 Then
      'add columns for costing and problem tracking.  Ticket 2A1-0E6162E2-0859
      Execute False, "alter table InvaTable add INLOTTRACK bit NULL"
      Execute False, "alter table InvaTable add INUSEACTUALCOST bit NULL"
      
      'set version
      Execute False, "update Version set Version = 5"
   End If
   
   If ver < 6 Then
      sSql = _
             "alter procedure DropColumnDefault" & vbCrLf _
             & "@TableName varchar(50)," & vbCrLf _
             & "@ColumnName varchar(50)" & vbCrLf _
             & "as" & vbCrLf _
             & "-- if constraint created with Alter Table. use sp_unbinddefault for others" & vbCrLf _
             & "declare @constraint_name sysname, @sql varchar(1000)" & vbCrLf _
             & "set @constraint_name = " & vbCrLf _
             & "    (select top 1 def.name from syscolumns cols join sysobjects tbl on cols.id = tbl.id and tbl.xtype = 'U'" & vbCrLf _
             & "    join sysobjects def on cols.cdefault = def.id and def.xtype = 'D'" & vbCrLf _
             & "    where tbl.name = @TableName and cols.Name = @ColumnName)" & vbCrLf _
             & "if @constraint_name is null return" & vbCrLf _
             & "set @sql = 'alter table ' + @TableName + ' drop constraint ' + @constraint_name" & vbCrLf _
             & "exec (@sql)" & vbCrLf
      Execute False, sSql
      
      'add column for costing and problem tracking.  Ticket 2A1-0E6162E2-0859
      Execute False, "alter table InvaTable add INCOSTEDBY char(4) NULL"
      
      'alter other columns that are too small
      AlterNumericColumn "ShopTable", "SHPESTRATE", "decimal(12,4)"
      AlterNumericColumn "RunsTable", "RUNMATL", "decimal(12,4)"
      AlterNumericColumn "RunsTable", "RUNLABOR", "decimal(12,4)"
      AlterNumericColumn "RunsTable", "RUNSTDCOST", "decimal(12,4)"
      AlterNumericColumn "RunsTable", "RUNEXP", "decimal(12,4)"
      
      'set version
      Execute False, "update Version set Version = 6"
   End If
   
   If ver < 7 Then
      sSql = "create view Vw_slebl08" & vbCrLf _
             & "as" & vbCrLf _
             & "/* test" & vbCrLf _
             & "Wrong:" & vbCrLf _
             & "select * from Vw_slebl08 where ITCUSTREQ <= '2007-05-24 23:59' -- doesn't constrain subquery" & vbCrLf _
             & "Right:" & vbCrLf _
             & "select * from Vw_slebl08 where ITCUSTREQ <= '2007-05-24 23:59' and OldestDate <= '2007-05-24 23:59'" & vbCrLf _
             & "*/" & vbCrLf _
             & "SELECT TOP 100 PERCENT x.OldestDate, PARTNUM, ITCUSTREQ, PADESC, PAEXTDESC, PAQOH, " & vbCrLf _
             & "ITSO, ITNUMBER, ITREV, ITPART, ITQTY, ITSCHED, ITSCHEDDEL, " & vbCrLf _
             & "ITBOOKDATE, ITCANCELED, ITINVOICE, ITPSSHIPPED, SOTYPE, SOCUST, " & vbCrLf _
             & "SOPO, SODIVISION, SOREGION, SOPREFIX, CUNICKNAME " & vbCrLf
      
      sSql = sSql & "FROM  PartTable part" & vbCrLf _
             & "join SoitTable ON PARTREF = ITPART " & vbCrLf _
             & "join SohdTable ON ITSO = SONUMBER " & vbCrLf _
             & "join CustTable ON SOCUST = CUREF " & vbCrLf _
             & "join (SELECT MIN(ITCUSTREQ) AS OldestDate, PARTREF" & vbCrLf _
             & "    FROM PartTable " & vbCrLf _
             & " join SoitTable ON PARTREF = ITPART " & vbCrLf _
             & " join SohdTable ON ITSO = SONUMBER " & vbCrLf _
             & " join CustTable ON SOCUST = CUREF" & vbCrLf _
             & "    WHERE ITCANCELED <> 1 AND ITINVOICE = 0" & vbCrLf _
             & " AND ITPSSHIPPED <> 1 " & vbCrLf _
             & "    GROUP BY PARTREF) AS x ON x.PARTREF = part.PARTREF " & vbCrLf _
             & "WHERE ITCANCELED <> 1 AND ITINVOICE = 0 AND ITPSSHIPPED <> 1 " & vbCrLf _
             & "ORDER BY x.OldestDate, ITPART, ITCUSTREQ" & vbCrLf
      Execute False, sSql
      
      sSql = "ALTER TABLE PoitTable DROP CONSTRAINT FK_PoitTable_VndrTable"
      Execute False, sSql
      
      AlterStringColumn "RndlTable", "RUNDLSDOCREF", "varchar(30)"
      AlterStringColumn "RndlTable", "RUNDLSDOCREFLONG", "varchar(30)"
      
      'set version
      Execute False, "update Version set Version = 7"
      
   End If
   
   If ver < 8 Then
      sSql = "CREATE PROCEDURE Qry_GetProductCode" & vbCrLf _
             & "(@productcode char(6)) " & vbCrLf _
             & "as " & vbCrLf _
             & "SELECT PCCODE,PCDESC FROM PcodTable " & vbCrLf _
             & "WHERE PCREF=@productcode"
      Execute False, sSql
      
      'ALTER TABLE PoitTable DROP CONSTRAINT FK_PoitTable_PohdTable
      'DROP INDEX PohdTable.PohdRef
      'ALTER TABLE PohdTable DROP CONSTRAINT PK_PohdTable
      'ALTER TABLE PohdTable alter column PONUMBER int not null
      'ALTER TABLE PohdTable alter column PORELEASE smallint not null
      'ALTER TABLE PohdTable ADD CONSTRAINT PK_PohdTable PRIMARY KEY CLUSTERED (PONUMBER,PORELEASE)
      'ALTER TABLE PoitTable ADD CONSTRAINT FK_PoitTable_PohdTable FOREIGN KEY (PINUMBER,PIRELEASE) REFERENCES PohdTable (PONUMBER,PORELEASE)
      
      Execute False, "ALTER TABLE PoitTable DROP CONSTRAINT FK_PoitTable_PohdTable"
      DropIndex "PohdTable", "PohdRef"
      Execute False, "ALTER TABLE PohdTable DROP CONSTRAINT PK_PohdTable"
      AlterNumericColumn "PohdTable", "PONUMBER", "int not null"
      AlterNumericColumn "PohdTable", "PORELEASE", "smallint not null"
      Execute False, "ALTER TABLE PohdTable ADD CONSTRAINT PK_PohdTable PRIMARY KEY CLUSTERED (PONUMBER,PORELEASE)"
      Execute False, "ALTER TABLE PoitTable ADD CONSTRAINT FK_PoitTable_PohdTable FOREIGN KEY (PINUMBER,PIRELEASE) REFERENCES PohdTable (PONUMBER,PORELEASE)"
      
      AlterNumericColumn "MopkTable", "PKTOTMATL", "decimal(15,4)"
      AlterNumericColumn "MopkTable", "PKTOTLABOR", "decimal(15,4)"
      AlterNumericColumn "MopkTable", "PKTOTEXP", "decimal(15,4)"
      AlterNumericColumn "MopkTable", "PKTOTOH", "decimal(15,4)"
      AlterNumericColumn "MopkTable", "PKTOTHRS", "decimal(10,3)"
      
      'set version
      Execute False, "update Version set Version = 8"
   End If
   
   If ver < 9 Then
      AlterNumericColumn "MopkTable", "PKPQTY", "decimal(12,3)"
      AlterNumericColumn "MopkTable", "PKAQTY", "decimal(12,3)"
      AlterNumericColumn "MopkTable", "PKAMT", "decimal(15,4)"
      AlterNumericColumn "MopkTable", "PKOHPCT", "decimal(8,3)"
      AlterNumericColumn "MopkTable", "PKINADDERS", "decimal(15,4)"
      AlterNumericColumn "MopkTable", "PKBOMQTY", "decimal(12,3)"
      
      'set version
      Execute False, "update Version set Version = 9"
   End If
   
   If ver < 10 Then
      sSql = "CREATE view [dbo].[viewCompletedMOs]" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select " & vbCrLf
      sSql = sSql & "RUNREF as MoPartRef, RUNNO as MoRunNo,  " & vbCrLf
      sSql = sSql & "(select count(*) from PoitTable " & vbCrLf
      sSql = sSql & "where PIRUNPART = RUNREF and PIRUNNO = RUNNO and PITYPE = 14 ) as OpenPoItems," & vbCrLf
      sSql = sSql & "(select count(*) from PoitTable " & vbCrLf
      sSql = sSql & "where PIRUNPART = RUNREF and PIRUNNO = RUNNO and PITYPE = 15 ) as RecPoItems," & vbCrLf
      sSql = sSql & "(select count(*) from InvaTable" & vbCrLf
      sSql = sSql & "where INMOPART = RUNREF and INMORUN = RUNNO and ( INTYPE = 9 OR INTYPE = 23 ) ) as OpenPickItems," & vbCrLf
      sSql = sSql & "(select count(*) from InvaTable" & vbCrLf
      sSql = sSql & "join PartTable on INPART = PARTREF and PALEVEL <> 5" & vbCrLf
      sSql = sSql & "where INMOPART = RUNREF and INMORUN = RUNNO and ( INTYPE = 9 OR INTYPE = 23 )" & vbCrLf
      sSql = sSql & "and PALEVEL <> 5 ) as OpenNonType5PickItems, " & vbCrLf
      sSql = sSql & "(select count(*) " & vbCrLf
      sSql = sSql & "from InvaTable  " & vbCrLf
      sSql = sSql & "join PartTable on INPART = PARTREF and PALOTTRACK = 1" & vbCrLf
      sSql = sSql & "join LoitTable on LOIMOPARTREF = INMOPART and LOIMORUNNO = INMORUN and LOIPARTREF = INPART" & vbCrLf
      sSql = sSql & "join LohdTable on LOINUMBER = LOTNUMBER and LOTUNITCOST = 0" & vbCrLf
      sSql = sSql & "where INMOPART = RUNREF and INMORUN = RUNNO and INTYPE = 10) as UncostedLots" & vbCrLf
      sSql = sSql & "from RunsTable " & vbCrLf
      sSql = sSql & "where RunStatus = 'CO'" & vbCrLf
      Execute False, sSql
   
      sSql = "CREATE view [dbo].[viewLotCostsByMoDetails]" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select rtrim(LOIMOPARTREF) as MoPart, LOIMORUNNO as MoRun," & vbCrLf
      sSql = sSql & "rtrim(LOTPARTREF) as Part," & vbCrLf
      sSql = sSql & "-LOIQUANTITY as Quantity, LOTUNITCOST as UnitCost," & vbCrLf
      sSql = sSql & "cast(-LOTUNITCOST * LOIQUANTITY as decimal(15,4)) as TotalCost," & vbCrLf
      sSql = sSql & "LOTNUMBER as LotNumber, rtrim(LOTUSERLOTID) as LotUserID" & vbCrLf
      sSql = sSql & "FROM LohdTable join LoitTable on LOTNUMBER = LOINUMBER and LOITYPE = 10" & vbCrLf
      sSql = sSql & "join PartTable on LOTPARTREF = PARTREF AND PALEVEL < 5 and PALOTTRACK = 1" & vbCrLf
      Execute False, sSql
   
      sSql = "CREATE view [dbo].[viewLotCostsByMoSummary]" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select MoPart, MoRun, sum(TotalCost) as TotalCost" & vbCrLf
      sSql = sSql & "from viewLotCostsByMoDetails" & vbCrLf
      sSql = sSql & "group by MoPart, MoRun" & vbCrLf
      Execute False, sSql
   
      sSql = "CREATE view [dbo].[viewNonLotCostsByMoDetails]" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select rtrim(INMOPART) as MoPart, INMORUN as MoRun, rtrim(INPART) as Part," & vbCrLf
      sSql = sSql & "PALEVEL as PartType, -INAQTY as Quantity, INAMT as UnitCost, " & vbCrLf
      sSql = sSql & "cast( -INAQTY*INAMT as decimal(15,4)) as TotalCost" & vbCrLf
      sSql = sSql & "from InvaTable" & vbCrLf
      sSql = sSql & "join PartTable on PARTREF=INPART" & vbCrLf
      sSql = sSql & "where INTYPE=10 and PALOTTRACK = 0" & vbCrLf
      Execute False, sSql
   
      sSql = "CREATE view [dbo].[viewNonLotCostsByMoSummary]" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select MoPart, MoRun, PartType, sum(TotalCost) as TotalCost" & vbCrLf
      sSql = sSql & "from viewNonLotCostsByMoDetails" & vbCrLf
      sSql = sSql & "group by MoPart, MoRun, PartType" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 10"
   End If

'   'For database version x and up:
'   'execute updates in UpdateEsiDatabase.sql in same directory as executable
'   'warning: the file must be ansi, not unicode!
'   'don't care if script errs out after the first time.
'   Err.Clear
'   Dim s As String
'   Dim hfile As Long
'   hfile = FreeFile
'   Open App.Path + "EsiDatabaseUpdates.sql" For Input As #1
'   If Err = 0 Then
'      Do While Not EOF(1)
'         Line Input #1, s
'      Loop
'      Close #1
'   End If
   
   If ver < 11 Then
      'add TCPRORATE column to TcitTable
      sSql = "ALTER TABLE TcitTable ADD TCPRORATE decimal(8,3) NULL DEFAULT(0)"
      RdoCon.Execute sSql, rdExecDirect
      sSql = "Update TcitTable set TCPRORATE = 0 where TCPRORATE is null"
      RdoCon.Execute sSql, rdExecDirect
      
      'set version
      Execute False, "update Version set Version = 11"
   End If
   
   If ver < 12 Then
      sSql = "alter view viewNonLotCostsByMoDetails" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select rtrim(INMOPART) as MoPart, INMORUN as MoRun, rtrim(INPART) as Part," & vbCrLf
      sSql = sSql & "PALEVEL as PartType, -INAQTY as Quantity, INAMT as UnitCost, " & vbCrLf
      sSql = sSql & "cast( -INAQTY*INAMT as decimal(15,4)) as TotalCost" & vbCrLf
      sSql = sSql & "from InvaTable" & vbCrLf
      sSql = sSql & "join PartTable on PARTREF=INPART" & vbCrLf
      sSql = sSql & "where INTYPE=10 and ( PALOTTRACK = 0 or INLOTNUMBER = '' )" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 12"
   End If
   
   'original MRP tables & procedures
   If ver < 13 Then
   
      'from MrplMRf01a.zMakeParaTable
      Dim RdoQoh As rdoResultset
      Dim sBegDate As String
      Dim sEnddate As String
      sSql = "CREATE TABLE MrpdTable (" _
             & "MRP_ROW INT NULL DEFAULT(1)," _
             & "MRP_CREATEDATE smalldatetime NULL," _
             & "MRP_THROUGHDATE smalldatetime NULL," _
             & "MRP_CREATEDBY CHAR(4) NULL DEFAULT(''))"
      Execute False, sSql
         
      sSql = "CREATE UNIQUE CLUSTERED INDEX MrpDateIdx ON " _
             & "MrpdTable(MRP_ROW) WITH  FILLFACTOR = 80"
      Execute False, sSql
      
      sSql = "SELECT DISTINCT MRP_ROW,MRP_PARTDATERQD,MRP_USER " _
             & "FROM MrplTable WHERE MRP_ROW=1"
      bSqlRows = GetDataSet(RdoQoh, ES_FORWARD)
      If bSqlRows Then
         With RdoQoh
            If Not IsNull(!mrp_partDateRQD) Then
               sBegDate = Format(!mrp_partDateRQD, "mm/dd/yy")
            Else
               sBegDate = Format(GetServerDateTime(), "mm/dd/yy")
            End If
            .Cancel
         End With
      Else
         sBegDate = Format(GetServerDateTime(), "mm/dd/yy")
      End If
      
      sSql = "SELECT MAX(MRP_PARTDATERQD) AS RegDate FROM MrplTable"
      bSqlRows = GetDataSet(RdoQoh, ES_FORWARD)
      If bSqlRows Then
         With RdoQoh
            If Not IsNull(!RegDate) Then
               sEnddate = Format(!RegDate, "mm/dd/yy")
            Else
               sEnddate = Format(GetServerDateTime(), "mm/dd/yy")
            End If
            .Cancel
         End With
      Else
         sEnddate = Format(GetServerDateTime(), "mm/dd/yy")
      End If
      
      sSql = "INSERT INTO MrpdTable (MRP_ROW,MRP_CREATEDATE,MRP_THROUGHDATE) " _
             & "VALUES(1,'" & sBegDate & "','" & sEnddate & "')"
      RdoCon.Execute sSql, rdExecDirect
      ClearResultSet RdoQoh
      On Error GoTo 0
      
      'from MrplMRf01a.zCheckTableColumns
      AlterNumericColumn "MrplTable", "MRP_PARTQTYRQD", "decimal(12,4)"
      AlterNumericColumn "MrplTable", "MRP_PARTUNITCOST", "decimal(12,4)"
      
      'changes for mrp
      DropColumnDefault "MrplTable", "MRP_ROW"
      Execute False, "truncate table MrplTable"
      Execute False, "alter table MrplTable drop constraint PK_MrplTable_MRPREF"
      'Execute False, "drop index MrplTable.MrpRow"
      DropIndex "MrplTable", "MrpRow"
      Execute False, "alter table MrplTable drop column MRP_ROW"
      Execute False, "alter table MrplTable add MRP_ROW int identity not null"
      Execute False, " alter table MrplTable add constraint PK_MrplTable_MRPREF PRIMARY KEY CLUSTERED (MRP_ROW Asc) "
      
      Execute False, "alter table MrplTable add MRP_PARENTROW int null"
      Execute False, "alter table MrplTable add MRP_ActionDate datetime null"
      
      'from MrplMRf01a.CheckLogTable
      On Error Resume Next
      sSql = "SELECT LOG_NUMBER FROM EsReportMrpLog"
      RdoCon.Execute sSql, rdExecDirect
      If Err > 0 Then
         Err.Clear
         sSql = "CREATE TABLE EsReportMrpLog (" _
                & "LOG_NUMBER SMALLINT NULL DEFAULT(0)," _
                & "LOG_TEXT VARCHAR(60) NULL DEFAULT(''))"
         RdoCon.Execute sSql, rdExecDirect
         If Err = 0 Then
            sSql = "CREATE UNIQUE CLUSTERED INDEX LogIndex ON " _
                   & "EsReportMrpLog(LOG_NUMBER) WITH FILLFACTOR = 80"
            RdoCon.Execute sSql, rdExecDirect
         End If
      End If
      On Error GoTo 0
      
      'from MrplMRf01a.zCheckStoredProcs
      sSql = "CREATE PROCEDURE Qry_UpdateMRppTablePOS" & vbCrLf _
             & "(@mrpquantity dec(12,4), " & vbCrLf _
             & "@mrppartnumber char(30)) " & vbCrLf _
             & "as " & vbCrLf _
             & "UPDATE MrppTable SET MRP_PARTQOH=MRP_PARTQOH+@mrpquantity " & vbCrLf _
             & "WHERE MRP_PARTREF=@mrppartnumber"
      Execute False, sSql
      
      sSql = "CREATE PROCEDURE Qry_UpdateMRppTableNEG" & vbCrLf _
             & "(@mrpquantity dec(12,4), " & vbCrLf _
             & "@mrppartnumber char(30)) " & vbCrLf _
             & "as " & vbCrLf _
             & "UPDATE MrppTable SET MRP_PARTQOH=MRP_PARTQOH-@mrpquantity " & vbCrLf _
             & "WHERE MRP_PARTREF=@mrppartnumber"
      Execute False, sSql
      
      sSql = "CREATE PROCEDURE Qry_UpdateMRppTableZERO" & vbCrLf _
             & "(@mrppartnumber char(30)) " & vbCrLf _
             & "as " & vbCrLf _
             & "UPDATE MrppTable SET MRP_PARTQOH=0 WHERE " & vbCrLf _
             & "(MRP_PARTQOH<0 AND MRP_PARTREF=@mrppartnumber)"
      Execute False, sSql

      'Qry_GetMRPGetNextBillLevel renamed to exclude second Get
      sSql = "drop procedure Qry_GetMRPGetNextBillLevel"
      Execute False, sSql
      
      sSql = "CREATE PROCEDURE Qry_GetMRPNextBillLevel" & vbCrLf _
             & "(@usedonpart char(30)) " & vbCrLf _
             & "as " & vbCrLf _
             & "SELECT BMASSYPART,BMPARTREF,BMREV,BMQTYREQD," & vbCrLf _
             & "BMCONVERSION,BMADDER,BMSETUP,PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
             & "PACLASS,PAPRODCODE,PABOMREV,PAFLOWTIME,PALEADTIME,PAMAKEBUY " & vbCrLf _
             & "FROM BmplTable,PartTable WHERE (BMPARTREF=PARTREF AND BMASSYPART=" & vbCrLf _
             & "@usedonpart AND PALEVEL<5) "
      Execute False, sSql
      
      'use Qry_GetMRPGetNextBillLevel -- it's identical
'      sSql = "CREATE PROCEDURE Qry_GetMRPGetNextPickLevel" & vbCrLf _
'             & "(@usedonpart char(30)) " & vbCrLf _
'             & "as " & vbCrLf _
'             & "SELECT BMASSYPART,BMPARTREF,BMREV,BMQTYREQD," & vbCrLf _
'             & "BMCONVERSION,PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
'             & "PACLASS,PAPRODCODE,PABOMREV,PAFLOWTIME,PALEADTIME " & vbCrLf _
'             & "PAMAKEBUY FROM BmplTable,PartTable WHERE " & vbCrLf _
'             & "(BMPARTREF=PARTREF AND BMASSYPART=@usedonpart AND PALEVEL<5) "
      sSql = "drop procedure Qry_GetMRPGetNextPickLevel"
      Execute False, sSql
      
      'use Qry_GetMRPGetNextBillLevel -- it's identical
'      sSql = "CREATE PROCEDURE Qry_GetMRPGetNextSOLevel" & vbCrLf _
'             & "(@usedonpart char(30)) " & vbCrLf _
'             & "as " & vbCrLf _
'             & "SELECT BMASSYPART,BMPARTREF,BMREV,BMQTYREQD," & vbCrLf _
'             & "BMCONVERSION,PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
'             & "PACLASS,PAPRODCODE,PABOMREV,PAFLOWTIME,PALEADTIME,PAMAKEBUY " & vbCrLf _
'             & "FROM BmplTable,PartTable WHERE BMPARTREF=PARTREF AND " & vbCrLf _
'             & "(BMASSYPART=@UsedOnPart AND PALEVEL<5)"
'      Execute false, ssql
      sSql = "drop procedure Qry_GetMRPGetNextSOLevel"
      Execute False, sSql
      
      sSql = "CREATE PROCEDURE Qry_GetMRPPartQoh" & vbCrLf _
             & "(@usedonpart char(30)) " & vbCrLf _
             & "as " & vbCrLf _
             & "SELECT MRP_PARTREF,MRP_PARTQOH FROM MrppTable WHERE " & vbCrLf _
             & "MRP_PARTREF=@UsedOnPart"
      Execute False, sSql
      
'      sSql = "CREATE PROCEDURE Qry_GetMRPMOrders" & vbCrLf _
'             & "(@enddate smalldatetime) " & vbCrLf _
'             & "as " & vbCrLf _
'             & "SELECT RUNREF,RUNNO,RUNSCHED,RUNQTY,RUNSTATUS," & vbCrLf _
'             & "PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
'             & "PACLASS,PAPRODCODE FROM RunsTable,PartTable WHERE " & vbCrLf _
'             & "(RUNREF=PARTREF AND RUNSTATUS NOT LIKE 'C%') " & vbCrLf _
'             & "AND RUNSCHED<=@enddate"
'      Execute false, ssql
'
      sSql = "CREATE PROCEDURE Qry_GetMRPMOPicks" & vbCrLf _
             & "(@enddate smalldatetime) " & vbCrLf _
             & "as " & vbCrLf _
             & "SELECT PKPARTREF,PKMOPART,PKMORUN,PKTYPE,PKPDATE," & vbCrLf _
             & "PKPQTY,PKAQTY,PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
             & "PACLASS,PAPRODCODE FROM MopkTable,PartTable WHERE " & vbCrLf _
             & "(PKPARTREF=PARTREF AND PKAQTY=0 AND PKTYPE<>12) " & vbCrLf _
             & "AND PKPDATE<=@enddate"
      Execute False, sSql
      
      sSql = "CREATE PROCEDURE Qry_GetMRPVendor" & vbCrLf _
             & "(@ponumber integer) " & vbCrLf _
             & "as " & vbCrLf _
             & "SELECT PONUMBER,POVENDOR,POBUYER,VEREF,VENICKNAME FROM " & vbCrLf _
             & "PohdTable,VndrTable WHERE (POVENDOR=VEREF AND PONUMBER=@ponumber)"
      Execute False, sSql
      
      sSql = "CREATE PROCEDURE Qry_GetMRPCustomer" & vbCrLf _
             & "(@sonumber integer) " & vbCrLf _
             & "as " & vbCrLf _
             & "SELECT SONUMBER,SOTYPE,SOCUST,CUREF,CUNICKNAME FROM " & vbCrLf _
             & "SohdTable,CustTable WHERE (SOCUST=CUREF AND SONUMBER=@sonumber )"
      Execute False, sSql
   
      sSql = "create view viewMRPSortOrder" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select mrp1.MRP_ROW,  " & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "select count(*) from MrplTable mrp2 " & vbCrLf
      sSql = sSql & "where mrp2.MRP_PARTNUM < mrp1.MRP_PARTNUM" & vbCrLf
      sSql = sSql & "or (mrp2.MRP_PARTNUM = mrp1.MRP_PARTNUM " & vbCrLf
      sSql = sSql & "and (mrp2.MRP_PARTSORTDATE < mrp1.MRP_PARTSORTDATE " & vbCrLf
      sSql = sSql & "or (mrp2.MRP_PARTSORTDATE = mrp1.MRP_PARTSORTDATE and mrp2.MRP_TYPE < mrp1.MRP_TYPE)" & vbCrLf
      sSql = sSql & "or (mrp2.MRP_PARTSORTDATE = mrp1.MRP_PARTSORTDATE and mrp2.MRP_TYPE = mrp1.MRP_TYPE" & vbCrLf
      sSql = sSql & "and mrp2.MRP_ROW <= mrp1.MRP_ROW)))" & vbCrLf
      sSql = sSql & ") as SortOrder" & vbCrLf
      sSql = sSql & "from MrplTable mrp1" & vbCrLf
      Execute False, sSql
      
      sSql = "create view viewMRPBalances" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select mrp1.MRP_ROW, mrp1.MRP_PARTNUM, MRP_PARTSORTDATE, MRP_PARTQTYRQD, MRP_PARTDATERQD," & vbCrLf
      sSql = sSql & "( select sum(MRP_PARTQTYRQD) " & vbCrLf
      sSql = sSql & "   from MrplTable mrp2 join viewMRPSortOrder vs2 on vs2.MRP_ROW = mrp2.MRP_ROW" & vbCrLf
      sSql = sSql & "   where mrp2.MRP_PARTNUM = mrp1.MRP_PARTNUM" & vbCrLf
      sSql = sSql & "   and vs2.SortOrder <= vs1.SortOrder" & vbCrLf
      sSql = sSql & ") as Balance," & vbCrLf
      sSql = sSql & "vs1.SortOrder" & vbCrLf
      sSql = sSql & "from MrplTable mrp1" & vbCrLf
      sSql = sSql & "join viewMRPSortOrder vs1 on vs1.MRP_ROW = mrp1.MRP_ROW" & vbCrLf
      Execute False, sSql
      
      Execute False, "drop procedure Qry_GetMRPMOrders"
      sSql = "CREATE PROCEDURE Qry_GetMRPMOrders" & vbCrLf _
             & "(@enddate smalldatetime) " & vbCrLf _
             & "as " & vbCrLf _
             & "SELECT RUNREF,RUNNO,RUNPKSTART,RUNSTART,RUNSCHED,RUNQTY,RUNSTATUS," & vbCrLf _
             & "coalesce(RUNPKSTART,RUNSTART,RUNSCHED) as ActionDate," & vbCrLf _
             & "PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
             & "PACLASS,PAPRODCODE FROM RunsTable,PartTable WHERE " & vbCrLf _
             & "(RUNREF=PARTREF AND RUNSTATUS NOT LIKE 'C%') " & vbCrLf _
             & "AND RUNSCHED<=@enddate"
      Execute False, sSql
      
'      sSql = "create view viewMRPBalances" & vbCrLf
'      sSql = sSql & "as" & vbCrLf
'      sSql = sSql & "select vb1.*, " & vbCrLf
'      sSql = sSql & "case when vb1.Balance < 0 " & vbCrLf
'      sSql = sSql & "then datediff(d, " & vbCrLf
'      sSql = sSql & "   isnull((select max(MRP_PARTSORTDATE) " & vbCrLf
'      sSql = sSql & "   from viewMRPBalances2 vb2 where vb2.MRP_PARTNUM = vb1.MRP_PARTNUM" & vbCrLf
'      sSql = sSql & "   and vb2.MRP_PARTSORTDATE < vb1.MRP_PARTSORTDATE" & vbCrLf
'      sSql = sSql & "   and vb2.Balance >= 0)," & vbCrLf
'      sSql = sSql & "   (select min(MRP_PARTSORTDATE) " & vbCrLf
'      sSql = sSql & "   from viewMRPBalances2 vb3 where vb3.MRP_PARTNUM = vb1.MRP_PARTNUM))," & vbCrLf
'      sSql = sSql & "   vb1.MRP_PARTSORTDATE)" & vbCrLf
'      sSql = sSql & "else" & vbCrLf
'      sSql = sSql & "   0" & vbCrLf
'      sSql = sSql & "end  as DaysNegative," & vbCrLf
'      sSql = sSql & "case when exists (select SortOrder from viewMRPBalances2 vb4" & vbCrLf
'      sSql = sSql & "   where vb4.MRP_PARTNUM = vb1.MRP_PARTNUM" & vbCrLf
'      sSql = sSql & "   and vb4.SortOrder > vb1.SortOrder ) then 0" & vbCrLf
'      sSql = sSql & "else" & vbCrLf
'      sSql = sSql & "   1" & vbCrLf
'      sSql = sSql & "end As LastItem" & vbCrLf
'      sSql = sSql & "from viewMRPBalances2 vb1" & vbCrLf
'      Execute false, ssql

      'set version
      Execute False, "update Version set Version = 13"
   End If
   
   
   If ver < 14 Then
      DropColumnDefault "MrplTable", "MRP_ROW"
      Execute False, "truncate table MrplTable"
      Execute False, "alter table MrplTable drop constraint PK_MrplTable_MRPREF"
      'Execute False, "drop index MrplTable.MrpRow"
      DropIndex "MrplTable", "MrpRow"
      Execute False, "alter table MrplTable drop column MRP_ROW"
      Execute False, "alter table MrplTable add MRP_ROW int identity not null"
      Execute False, " alter table MrplTable add constraint PK_MrplTable_MRPREF PRIMARY KEY CLUSTERED (MRP_ROW Asc) "

      Execute False, "drop procedure Qry_GetMRPMOrders"
      sSql = "CREATE PROCEDURE Qry_GetMRPMOrders" & vbCrLf _
             & "(@enddate smalldatetime) " & vbCrLf _
             & "as " & vbCrLf _
             & "SELECT RUNREF,RUNNO,RUNPKSTART,RUNSTART,RUNSCHED,RUNQTY,RUNSTATUS," & vbCrLf _
             & "coalesce(RUNPKSTART,RUNSTART,RUNSCHED) as ActionDate," & vbCrLf _
             & "PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
             & "PACLASS,PAPRODCODE FROM RunsTable,PartTable WHERE " & vbCrLf _
             & "(RUNREF=PARTREF AND RUNSTATUS NOT LIKE 'C%') " & vbCrLf _
             & "AND RUNSCHED<=@enddate"
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 14"
   End If
   
   
   If ver < 15 Then
      sSql = "create index IX_MrplTable_PART_SORT_TYPE_ROW" & vbCrLf _
         & "on MrplTable ( MRP_PARTNUM, MRP_PARTSORTDATE, MRP_TYPE, MRP_ROW )"
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 15"
   End If
   
   If ver < 16 Then
      'Execute False, "drop index InvaTable.INVCLUSTER"
      DropIndex "InvaTable", "INVCLUSTER"
      Execute False, "alter table InvaTable add INNO int identity not null"
      Execute False, "CREATE INDEX IX_InvaTable_PartTypeNum ON InvaTable (INPART,INTYPE,INNUMBER)"
      Execute False, "alter table InvaTable ADD CONSTRAINT PK_InvaTable PRIMARY KEY CLUSTERED(INNO)"
      
      sSql = sSql & "create trigger TR_InvaTable_Insert" & vbCrLf
      sSql = sSql & "on InvaTable " & vbCrLf
      sSql = sSql & "for insert" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "declare cur cursor for " & vbCrLf
      sSql = sSql & "select inno from inserted" & vbCrLf
      sSql = sSql & "where INNUMBER = 0 or INNUMBER is null" & vbCrLf
      sSql = sSql & "open cur" & vbCrLf
      sSql = sSql & "declare @INNO int" & vbCrLf
      sSql = sSql & "fetch next from cur into @INNO " & vbCrLf
      sSql = sSql & "while @@FETCH_STATUS = 0" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   update InvaTable set INNUMBER = (select isnull(max(INNUMBER),0) + 1 from InvaTable)" & vbCrLf
      sSql = sSql & "   where INNO = @INNO" & vbCrLf
      sSql = sSql & "   fetch next from cur into @INNO " & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "close cur" & vbCrLf
      sSql = sSql & "deallocate cur" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 16"
   End If
   
   If ver < 17 Then
      sSql = "ALTER PROCEDURE Qry_GetMRPMOrders" & vbCrLf _
           & "(@enddate smalldatetime) " & vbCrLf _
           & "as " & vbCrLf _
           & "SELECT RUNREF, RUNNO, RUNPKSTART, RUNSTART, RUNSCHED, RUNQTY," & vbCrLf _
           & "RUNQTY - RUNYIELD AS QuantityLeft, RUNSTATUS," & vbCrLf _
           & "coalesce(RUNPKSTART,RUNSTART,RUNSCHED) as ActionDate," & vbCrLf _
           & "PARTREF, PARTNUM, PADESC, PALEVEL, PAUNITS," & vbCrLf _
           & "PACLASS, PAPRODCODE" & vbCrLf _
           & "FROM RunsTable,PartTable " & vbCrLf _
           & "WHERE (RUNREF=PARTREF AND RUNSTATUS NOT LIKE 'C%') " & vbCrLf _
           & "AND RUNSCHED <= @enddate" & vbCrLf _
           & "AND RUNQTY > RUNYIELD" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 17"
   End If
   
   If ver < 18 Then
      sSql = "ALTER view [dbo].[viewNonLotCostsByMoDetails]" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select rtrim(INMOPART) as MoPart, INMORUN as MoRun, rtrim(INPART) as Part," & vbCrLf
      sSql = sSql & "PALEVEL as PartType, -INAQTY as Quantity, INAMT as UnitCost, " & vbCrLf
      sSql = sSql & "cast( -INAQTY*INAMT as decimal(15,4)) as TotalCost" & vbCrLf
      sSql = sSql & "from InvaTable" & vbCrLf
      sSql = sSql & "join PartTable on PARTREF=INPART" & vbCrLf
'      sSql = sSql & "where INTYPE=10 and PALOTTRACK = 0" & vbCrLf
      sSql = sSql & "Where INTYPE = 10 And PALEVEL < 5 And PALOTTRACK = 0" & vbCrLf
      Execute False, sSql
   
      'set version
      Execute False, "update Version set Version = 18"
   End If
   
   If ver < 19 Then
      sSql = "CREATE TABLE TempPsLots(" & vbCrLf
      sSql = sSql & "PsNumber char(8) NOT NULL," & vbCrLf
      sSql = sSql & "PsItem smallint NOT NULL," & vbCrLf
      sSql = sSql & "LotID char(15) NOT NULL," & vbCrLf
      sSql = sSql & "LotQty decimal(12, 4) NOT NULL," & vbCrLf
      sSql = sSql & "WhenCreated datetime NOT NULL CONSTRAINT DF_TempPsLots_WhenCreated DEFAULT (getdate())"
      sSql = sSql & ")"
      Execute False, sSql
   
      'set version
      Execute False, "update Version set Version = 19"
   End If
   
   If ver < 20 Then
      sSql = "CREATE FUNCTION fnGetNextLotItemNumber(@LotNumber char(15))" & vbCrLf
      sSql = sSql & "RETURNS INT" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "BEGIN" & vbCrLf
      sSql = sSql & " DECLARE @i INT" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & " SELECT @i = MAX( LOIRECORD ) FROM LoitTable WHERE LOINUMBER = @LotNumber" & vbCrLf
      sSql = sSql & " IF @i IS NULL SET @i = 1" & vbCrLf
      sSql = sSql & " ELSE SET @i = @i + 1" & vbCrLf
      sSql = sSql & " RETURN (@i)" & vbCrLf
      sSql = sSql & "END" & vbCrLf
      Execute False, sSql
      
      sSql = "DELETE FROM TempPsLots"
      Execute False, sSql
      
      sSql = "ALTER TABLE TempPsLots ADD PartRef varchar(30) NOT NULL"
      Execute False, sSql
   
      'set version
      Execute False, "update Version set Version = 20"
   End If
   
   If ver < 21 Then
      Execute False, "ALTER TABLE PartTable ADD PALOTSEXPIRE TINYINT NULL"
      Execute False, "UPDATE PartTable SET PALOTSEXPIRE = 0 WHERE PALOTSEXPIRE IS NULL"
      Execute False, "ALTER TABLE PartTable ALTER COLUMN PALOTSEXPIRE TINYINT NOT NULL"
      Execute False, "ALTER TABLE PartTable ADD CONSTRAINT DF_PartTable_PALOTSEXPIRE DEFAULT 0 FOR PALOTSEXPIRE"
      
      Execute False, "ALTER TABLE LohdTable ADD LOTEXPIRESON DATETIME NULL"
   
      'set version
      Execute False, "update Version set Version = 21"
   End If
   
   If ver < 22 Then
      'optimize very slow packing slip costing query in ClassInventoryActivity
      sSql = "CREATE NONCLUSTERED INDEX IX_InvaTable_INPSNUMBER ON InvaTable" & vbCrLf _
         & "(" & vbCrLf _
         & "   INPSNUMBER ASC," & vbCrLf _
         & "   INPSITEM Asc" & vbCrLf _
         & ")" & vbCrLf
      Execute False, sSql
      
      sSql = "CREATE NONCLUSTERED INDEX IX_LoitTable_LOIPSNUMBER ON LoitTable" & vbCrLf _
         & "(" & vbCrLf _
         & "   LOIPSNUMBER ASC," & vbCrLf _
         & "   LOIPSITEM Asc" & vbCrLf _
         & ")" & vbCrLf
      Execute False, sSql
      
      Execute False, "update Version set Version = 22"
   End If
   
   If ver < 23 Then
      
      'didn't get added to Exotic Tools
      sSql = "DELETE FROM TempPsLots"
      Execute False, sSql
      
      sSql = "ALTER TABLE TempPsLots ADD PartRef varchar(30) NOT NULL"
      Execute False, sSql
      'Execute false, ssql
      
      Execute False, "update Version set Version = 23"
   End If

   If ver < 24 Then
   
      'whoops - need to delete all rows to add a not null row
      sSql = "DELETE FROM TempPsLots"
      Execute False, sSql
      
      sSql = "ALTER TABLE TempPsLots ADD PartRef varchar(30) NOT NULL"
      Execute False, sSql
   
      'set version
      Execute False, "update Version set Version = 24"
   End If
   
   If ver < 25 Then
   
      AddNonNullColumnWithDefault "PartTable", "PAINACTIVE", "TINYINT", "0"
      AddNonNullColumnWithDefault "PartTable", "PAOBSOLETE", "TINYINT", "0"
      
      sSql = "CREATE NONCLUSTERED INDEX IX_InvaTable_INSONUMBER_ETC ON InvaTable" & vbCrLf _
         & "(" & vbCrLf _
         & "   INSONUMBER," & vbCrLf _
         & "   INSOITEM," & vbCrLf _
         & "   INSOREV" & vbCrLf _
         & ")" & vbCrLf
      Execute False, sSql
      
      sSql = "CREATE NONCLUSTERED INDEX IX_SoitTable_ITINVOICE ON SoitTable" & vbCrLf _
         & "(" & vbCrLf _
         & "   ITINVOICE" & vbCrLf _
         & ")" & vbCrLf
      Execute False, sSql
      
'supposed to help Revise Request Date For Unshipped PS Items, but doesn't
'      sSql = "CREATE NONCLUSTERED INDEX IX_SoitTable_ITPSSHIPPED ON SoitTable" & vbCrLf _
'         & "(" & vbCrLf _
'         & "   ITPSSHIPPED," & vbCrLf _
'         & "   ITPSNUMBER," & vbCrLf _
'         & "   ITPSITEM" & vbCrLf _
'         & ")" & vbCrLf
'      Execute false, ssql
      
      sSql = "CREATE FUNCTION fnGetPartCgsAccount(@PartRef varchar(30))" & vbCrLf
      sSql = sSql & "RETURNS varchar(12)" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "BEGIN" & vbCrLf
      sSql = sSql & " DECLARE @expAcct varchar(12), @returnAcct varchar(12)" & vbCrLf
      sSql = sSql & " DECLARE @prodcode varchar(6)" & vbCrLf
      sSql = sSql & " declare @level tinyint" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & " -- first look for account in part record" & vbCrLf
      sSql = sSql & " select @expAcct = rtrim(PACGSEXPACCT)," & vbCrLf
      sSql = sSql & "    @returnAcct = rtrim(PACGSMATACCT)," & vbCrLf
      sSql = sSql & "    @ProdCode = rtrim(PAPRODCODE)," & vbCrLf
      sSql = sSql & "    @level = PALEVEL" & vbCrLf
      sSql = sSql & "    from PartTable" & vbCrLf
      sSql = sSql & "    where PARTREF = @PartRef" & vbCrLf
      sSql = sSql & "        " & vbCrLf
      sSql = sSql & " if @level = 6 or @level = 7" & vbCrLf
      sSql = sSql & "        set @returnAcct = @expAcct" & vbCrLf
      sSql = sSql & "        " & vbCrLf
      sSql = sSql & " -- if no account in part, look in product code" & vbCrLf
      sSql = sSql & " if @returnAcct = ''" & vbCrLf
      sSql = sSql & " begin" & vbCrLf
      sSql = sSql & "    select @expAcct = rtrim(PCCGSEXPACCT)," & vbCrLf
      sSql = sSql & "       @returnAcct = rtrim(PCCGSMATACCT)" & vbCrLf
      sSql = sSql & "       from PcodTable" & vbCrLf
      sSql = sSql & "       where PCCODE = @ProdCode" & vbCrLf
      sSql = sSql & "         " & vbCrLf
      sSql = sSql & "    if @level = 6 or @level = 7" & vbCrLf
      sSql = sSql & "       set @returnAcct = @expAcct" & vbCrLf
      sSql = sSql & "         " & vbCrLf
      sSql = sSql & "    -- if no account in product code, look in common      " & vbCrLf
      sSql = sSql & "    if @returnAcct = ''" & vbCrLf
      sSql = sSql & "    begin" & vbCrLf
      sSql = sSql & "       if @level = 1 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COCGSEXPACCT1), @returnAcct = rtrim(COCGSMATACCT1) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 2 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COCGSEXPACCT2), @returnAcct = rtrim(COCGSMATACCT2) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 3 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COCGSEXPACCT3), @returnAcct = rtrim(COCGSMATACCT3) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 4 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COCGSEXPACCT4), @returnAcct = rtrim(COCGSMATACCT4) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 5 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COCGSEXPACCT5), @returnAcct = rtrim(COCGSMATACCT5) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 6 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COCGSEXPACCT6), @returnAcct = rtrim(COCGSMATACCT6) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 7 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COCGSEXPACCT7), @returnAcct = rtrim(COCGSMATACCT7) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "       if @level = 6 or @level = 7" & vbCrLf
      sSql = sSql & "          set @returnAcct = @expAcct" & vbCrLf
      sSql = sSql & "    end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & " end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & " return @returnAcct" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      Execute False, sSql
      
      sSql = "CREATE FUNCTION fnGetPartInvAccount(@PartRef varchar(30))" & vbCrLf
      sSql = sSql & "RETURNS varchar(12)" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "BEGIN" & vbCrLf
      sSql = sSql & " DECLARE @expAcct varchar(12), @returnAcct varchar(12)" & vbCrLf
      sSql = sSql & " DECLARE @prodcode varchar(6)" & vbCrLf
      sSql = sSql & " declare @level tinyint" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & " -- first look for account in part record" & vbCrLf
      sSql = sSql & " select @expAcct = rtrim(PAINVEXPACCT)," & vbCrLf
      sSql = sSql & "    @returnAcct = rtrim(PAINVMATACCT)," & vbCrLf
      sSql = sSql & "    @ProdCode = rtrim(PAPRODCODE)," & vbCrLf
      sSql = sSql & "    @level = PALEVEL" & vbCrLf
      sSql = sSql & "    from PartTable" & vbCrLf
      sSql = sSql & "    where PARTREF = @PartRef" & vbCrLf
      sSql = sSql & "        " & vbCrLf
      sSql = sSql & " if @level = 6 or @level = 7" & vbCrLf
      sSql = sSql & "        set @returnAcct = @expAcct" & vbCrLf
      sSql = sSql & "        " & vbCrLf
      sSql = sSql & " -- if no account in part, look in product code" & vbCrLf
      sSql = sSql & " if @returnAcct = ''" & vbCrLf
      sSql = sSql & " begin" & vbCrLf
      sSql = sSql & "    select @expAcct = rtrim(PCINVEXPACCT)," & vbCrLf
      sSql = sSql & "       @returnAcct = rtrim(PCINVMATACCT)" & vbCrLf
      sSql = sSql & "       from PcodTable" & vbCrLf
      sSql = sSql & "       where PCCODE = @ProdCode" & vbCrLf
      sSql = sSql & "         " & vbCrLf
      sSql = sSql & "    if @level = 6 or @level = 7" & vbCrLf
      sSql = sSql & "       set @returnAcct = @expAcct" & vbCrLf
      sSql = sSql & "         " & vbCrLf
      sSql = sSql & "    -- if no account in product code, look in common      " & vbCrLf
      sSql = sSql & "    if @returnAcct = ''" & vbCrLf
      sSql = sSql & "    begin" & vbCrLf
      sSql = sSql & "       if @level = 1 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COINVEXPACCT1), @returnAcct = rtrim(COINVMATACCT1) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 2 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COINVEXPACCT2), @returnAcct = rtrim(COINVMATACCT2) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 3 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COINVEXPACCT3), @returnAcct = rtrim(COINVMATACCT3) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 4 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COINVEXPACCT4), @returnAcct = rtrim(COINVMATACCT4) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 5 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COINVEXPACCT5), @returnAcct = rtrim(COINVMATACCT5) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 6 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COINVEXPACCT6), @returnAcct = rtrim(COINVMATACCT6) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "       if @level = 7 " & vbCrLf
      sSql = sSql & "          select @expAcct = rtrim(COINVEXPACCT7), @returnAcct = rtrim(COINVMATACCT7) from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "       if @level = 6 or @level = 7" & vbCrLf
      sSql = sSql & "          set @returnAcct = @expAcct" & vbCrLf
      sSql = sSql & "    end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & " end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & " return @returnAcct" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      Execute False, sSql
      
      sSql = "create procedure UpdatePackingSlipCosts" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & " @PackingSlip varchar(8)" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "update InvaTable" & vbCrLf
      sSql = sSql & "set INAMT = LOTUNITCOST," & vbCrLf
      sSql = sSql & "INTOTMATL = cast ( abs( INAQTY ) * LOTTOTMATL / LOTORIGINALQTY as decimal(12,4))," & vbCrLf
      sSql = sSql & "INTOTLABOR = cast ( abs( INAQTY ) * LOTTOTLABOR / LOTORIGINALQTY as decimal(12,4))," & vbCrLf
      sSql = sSql & "INTOTEXP = cast ( abs( INAQTY ) * LOTTOTEXP / LOTORIGINALQTY as decimal(12,4)), " & vbCrLf
      sSql = sSql & "INTOTOH = cast ( abs( INAQTY ) * LOTTOTOH / LOTORIGINALQTY as decimal(12,4))," & vbCrLf
      sSql = sSql & "INTOTHRS = cast ( abs( INAQTY ) * LOTTOTHRS / LOTORIGINALQTY as decimal(12,4))" & vbCrLf
      sSql = sSql & "from LoitTable " & vbCrLf
      sSql = sSql & "join LohdTable ON LOINUMBER = LOTNUMBER" & vbCrLf
      sSql = sSql & "join InvaTable ia2 ON INNUMBER = LOIACTIVITY" & vbCrLf
      sSql = sSql & "where ia2.INPSNUMBER = @PackingSlip" & vbCrLf
      sSql = sSql & "and LOTORIGINALQTY <> 0" & vbCrLf
      Execute False, sSql
      
      ' this didn't work in version 16 -- try it again
      sSql = "create trigger TR_InvaTable_Insert" & vbCrLf
      sSql = sSql & "on InvaTable " & vbCrLf
      sSql = sSql & "for insert" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "declare cur cursor for " & vbCrLf
      sSql = sSql & "select inno from inserted" & vbCrLf
      sSql = sSql & "where INNUMBER = 0 or INNUMBER is null" & vbCrLf
      sSql = sSql & "open cur" & vbCrLf
      sSql = sSql & "declare @INNO int" & vbCrLf
      sSql = sSql & "fetch next from cur into @INNO " & vbCrLf
      sSql = sSql & "while @@FETCH_STATUS = 0" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   update InvaTable set INNUMBER = (select isnull(max(INNUMBER),0) + 1 from InvaTable)" & vbCrLf
      sSql = sSql & "   where INNO = @INNO" & vbCrLf
      sSql = sSql & "   fetch next from cur into @INNO " & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "close cur" & vbCrLf
      sSql = sSql & "deallocate cur" & vbCrLf
      Execute False, sSql
      
      sSql = "create procedure UpdateIaWipAccounts" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & " @MinInNumber int     -- min ia.INNUMBER to update" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "/* test" & vbCrLf
      sSql = sSql & " exec UpdateIaWipAccounts 504740" & vbCrLf
      sSql = sSql & "*/" & vbCrLf
      sSql = sSql & "UPDATE InvaTable " & vbCrLf
      sSql = sSql & "SET INCRLABACCT = WIPCRLABACCT," & vbCrLf
      sSql = sSql & "INCRMATACCT = WIPCRMATACCT," & vbCrLf
      sSql = sSql & "INCREXPACCT = WIPCREXPACCT," & vbCrLf
      sSql = sSql & "INCROHDACCT = WIPCROHDACCT," & vbCrLf
      sSql = sSql & "INDRLABACCT = WIPDRLABACCT," & vbCrLf
      sSql = sSql & "INDRMATACCT = WIPDRMATACCT," & vbCrLf
      sSql = sSql & "INDREXPACCT = WIPDREXPACCT," & vbCrLf
      sSql = sSql & "INDROHDACCT = WIPDROHDACCT" & vbCrLf
      sSql = sSql & "FROM Invatable ia" & vbCrLf
      sSql = sSql & "JOIN ComnTable ON COREF = 1 AND ia.INNUMBER >= @MinInNumber" & vbCrLf
      Execute False, sSql
      
      'delete redundant inno invatable row at Intercoastal so primary key constraint may be applied
      Execute False, "delete from invatable where innumber = 843 and inno = 5790 and inpdate = '1/15/1997'"
      
      'this may have timed out or failed (as described above) in db version 16 update
      Execute False, "alter table InvaTable ADD CONSTRAINT PK_InvaTable PRIMARY KEY CLUSTERED(INNO)"
      
      'set version
      Execute False, "update Version set Version = 25"
   End If
   
   If ver < 26 Then
   
      AddNonNullColumnWithDefault "LohdTable", "LOTMAINTCOSTED", "INT", "0"
      AddNonNullColumnWithDefault "InvaTable", "INMAINTCOSTED", "INT", "0"
      AddNonNullColumnWithDefault "RunsTable", "RUNMAINTCOSTED", "INT", "0"
      
      sSql = "alter view viewLotCostsByMoDetails" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select LOITYPE, rtrim(LOIMOPARTREF) as MoPart, LOIMORUNNO as MoRun," & vbCrLf
      sSql = sSql & "rtrim(LOTPARTREF) as Part," & vbCrLf
      sSql = sSql & "-LOIQUANTITY as Quantity, LOTUNITCOST as UnitCost," & vbCrLf
      sSql = sSql & "cast(-LOTUNITCOST * LOIQUANTITY as decimal(15,4)) as TotalCost," & vbCrLf
      sSql = sSql & "LOTNUMBER as LotNumber, rtrim(LOTUSERLOTID) as LotUserID, LOTMAINTCOSTED" & vbCrLf
      sSql = sSql & "FROM LohdTable join LoitTable on LOTNUMBER = LOINUMBER and LOITYPE = 10" & vbCrLf
      sSql = sSql & "join PartTable on LOTPARTREF = PARTREF AND PALEVEL < 5 and PALOTTRACK = 1" & vbCrLf
      Execute False, sSql
      
      sSql = "alter view viewNonLotCostsByMoDetails" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select rtrim(INMOPART) as MoPart, INMORUN as MoRun, rtrim(INPART) as Part," & vbCrLf
      sSql = sSql & "PALEVEL as PartType, -INAQTY as Quantity, INAMT as UnitCost, " & vbCrLf
      sSql = sSql & "cast( -INAQTY*INAMT as decimal(15,4)) as TotalCost, INMAINTCOSTED" & vbCrLf
      sSql = sSql & "from InvaTable" & vbCrLf
      sSql = sSql & "join PartTable on PARTREF=INPART" & vbCrLf
      sSql = sSql & "Where INTYPE = 10 And PALEVEL < 5 And PALOTTRACK = 0" & vbCrLf
      Execute False, sSql
      
      sSql = "alter view viewLotCostsByMoSummary" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select MoPart, MoRun, sum(TotalCost) as TotalCost," & vbCrLf
      sSql = sSql & "isnull(min(cast(LOTMAINTCOSTED as tinyint)), 1) as FullyCosted" & vbCrLf
      sSql = sSql & "from viewLotCostsByMoDetails" & vbCrLf
      sSql = sSql & "group by MoPart, MoRun" & vbCrLf
      Execute False, sSql
      
      sSql = "alter view viewNonLotCostsByMoSummary" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select MoPart, MoRun, PartType, sum(TotalCost) as TotalCost," & vbCrLf
      sSql = sSql & "isnull(min(cast(INMAINTCOSTED as tinyint)), 1) as FullyCosted" & vbCrLf
      sSql = sSql & "from viewNonLotCostsByMoDetails" & vbCrLf
      sSql = sSql & "group by MoPart, MoRun, PartType" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 26"
   End If
   
   If ver < 27 Then
   
      sSql = "create view viewExpensesByMoDetails" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select OPREF, OPRUN, OPNO, cast( min( OPYIELD * OPSVCUNIT ) as decimal(12,4) ) as OpSvcCost, " & vbCrLf
      sSql = sSql & "min( mopart.PAUSEACTUALCOST ) as PAUSEACTUALCOST," & vbCrLf
      sSql = sSql & "min( svcpart.PASTDCOST) as SvcStdCost, " & vbCrLf
      sSql = sSql & "cast( sum( isnull(PIAQTY,0) * isnull( PIAMT,0 ) ) as decimal(12,4) ) as PoSvcCost" & vbCrLf
      sSql = sSql & "from RnopTable" & vbCrLf
      sSql = sSql & "join PartTable mopart on mopart.PARTREF = OPREF" & vbCrLf
      sSql = sSql & "join PartTable svcpart on OPSERVPART = svcpart.PARTREF" & vbCrLf
      sSql = sSql & "left join PoitTable on OPREF = PIRUNPART and OPRUN = PIRUNNO and OPNO = PIRUNOPNO" & vbCrLf
      sSql = sSql & "and PITYPE = 17" & vbCrLf
      sSql = sSql & "where svcpart.PALEVEL >= 6" & vbCrLf
      sSql = sSql & "group by OPREF, OPRUN, OPNO" & vbCrLf
      Execute False, sSql
      
      'views for Larry's reports
      sSql = "CREATE VIEW [dbo].[Vw_Sales]" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "SELECT DISTINCT TOP 100 PERCENT dbo.CihdTable.INVTYPE, dbo.CihdTable.INVPRE, dbo.CihdTable.INVNO, dbo.LoitTable.LOICUSTINVNO, dbo.CihdTable.INVCUST, " & vbCrLf
      sSql = sSql & "               dbo.CustTable.CUNUMBER, dbo.CustTable.CUNAME, dbo.CihdTable.INVDATE, dbo.CihdTable.INVPIF, dbo.PshdTable.PSCANCELED, " & vbCrLf
      sSql = sSql & "               dbo.PshdTable.PSTYPE, dbo.PsitTable.PIPACKSLIP, dbo.PsitTable.PIITNO, dbo.InvaTable.INPSNUMBER, dbo.InvaTable.INPSITEM, " & vbCrLf
      sSql = sSql & "               dbo.LoitTable.LOIPSNUMBER, dbo.LoitTable.LOIPSITEM, dbo.SoitTable.ITPSSHIPPED, dbo.SohdTable.SOSALESMAN AS PSSoSalesman, " & vbCrLf
      sSql = sSql & "               SohdTable_1.SOSALESMAN AS SOSoSlsmn, dbo.SohdTable.SOPO AS PSSoPo, SohdTable_1.SOPO AS SOSoPo, " & vbCrLf
      sSql = sSql & "               dbo.SohdTable.SODIVISION AS PSSoDiv, SohdTable_1.SODIVISION AS SOSoDiv, dbo.SohdTable.SOREGION AS PSSoReg, " & vbCrLf
      sSql = sSql & "               SohdTable_1.SOREGION AS SOSoReg, dbo.SohdTable.SOBUSUNIT AS PSSoBu, SohdTable_1.SOBUSUNIT AS SOSoBu, " & vbCrLf
      sSql = sSql & "               dbo.SprsTable.SPLAST AS PSSlsmnLast, SprsTable_1.SPLAST AS SOSlsmnLast, dbo.SprsTable.SPFIRST AS PSSlsmnFirst, " & vbCrLf
      sSql = sSql & "               SprsTable_1.SPFIRST AS SOSlsmnFirst, dbo.SprsTable.SPMIDD AS PSSlsmnInit, SprsTable_1.SPMIDD AS SOSlsmnInit, " & vbCrLf
      sSql = sSql & "               dbo.SohdTable.SOTYPE AS SOSoType, SohdTable_1.SOTYPE AS PSSoType, dbo.SoitTable.ITSO, dbo.SoitTable.ITNUMBER, dbo.SoitTable.ITREV, " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITPART, dbo.PartTable.PARTNUM, dbo.PartTable.PADESC, dbo.PartTable.PALEVEL, dbo.PartTable.PALOTTRACK, " & vbCrLf
      sSql = sSql & "               dbo.PartTable.PAUSEACTUALCOST, dbo.PartTable.PAUNITS, dbo.PartTable.PAMAKEBUY, dbo.PartTable.PAFAMILY, dbo.PartTable.PAPRODCODE, " & vbCrLf
      sSql = sSql & "               dbo.PcodTable.PCDESC, dbo.PartTable.PACLASS, dbo.PclsTable.CCDESC, dbo.SoitTable.ITQTY, dbo.SoitTable.ITDOLLARS, dbo.SoitTable.ITADJUST, " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITDISCAMOUNT, dbo.SoitTable.ITCOMMISSION, dbo.SoitTable.ITBOOKDATE, dbo.SoitTable.ITSCHED, dbo.SoitTable.ITACTUAL, " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITCANCELDATE, dbo.SoitTable.ITCANCELED, dbo.SoitTable.ITREVACCT, dbo.GlacTable.GLACCTNO AS RevAcct, " & vbCrLf
      sSql = sSql & "               dbo.GlacTable.GLDESCR AS RevAcctDesc, dbo.SoitTable.ITCGSACCT, dbo.SoitTable.ITDISACCT, dbo.SoitTable.ITSTATE, dbo.SoitTable.ITTAXCODE, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INTYPE, dbo.InvaTable.INAQTY, dbo.InvaTable.INAMT, dbo.InvaTable.INTOTLABOR, dbo.InvaTable.INTOTMATL, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INTOTEXP, dbo.InvaTable.INTOTOH, dbo.PartTable.PASTDCOST, dbo.PartTable.PATOTCOST, dbo.PartTable.PATOTLABOR, " & vbCrLf
      sSql = sSql & "               dbo.PartTable.PATOTMATL, dbo.PartTable.PATOTEXP, dbo.PartTable.PATOTOH, dbo.InvaTable.INDRLABACCT, dbo.InvaTable.INDRMATACCT, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INDREXPACCT, dbo.InvaTable.INDROHDACCT, dbo.InvaTable.INCRLABACCT, dbo.InvaTable.INCRMATACCT, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INCREXPACCT, dbo.InvaTable.INCROHDACCT, dbo.LoitTable.LOITYPE, dbo.LoitTable.LOINUMBER, dbo.LoitTable.LOIRECORD, " & vbCrLf
      sSql = sSql & "               dbo.LohdTable.LOTNUMBER, dbo.LohdTable.LOTUSERLOTID, dbo.LohdTable.LOTPARTREF, dbo.LohdTable.LOTDATECOSTED, " & vbCrLf
      sSql = sSql & "          dbo.LohdTable.LOTTOTLABOR, dbo.LohdTable.LOTTOTMATL, dbo.LohdTable.LOTTOTEXP, dbo.LohdTable.LOTTOTOH" & vbCrLf
      sSql = sSql & "FROM  dbo.GlacTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SprsTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SoitTable LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SohdTable ON dbo.SoitTable.ITSO = dbo.SohdTable.SONUMBER ON " & vbCrLf
      sSql = sSql & "               dbo.SprsTable.SPNUMBER = dbo.SohdTable.SOSALESMAN RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.CustTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.CihdTable ON dbo.CustTable.CUREF = dbo.CihdTable.INVCUST ON dbo.SoitTable.ITINVOICE = dbo.CihdTable.INVNO LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SprsTable AS SprsTable_1 RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SohdTable AS SohdTable_1 ON SprsTable_1.SPNUMBER = SohdTable_1.SOSALESMAN ON " & vbCrLf
      sSql = sSql & "               dbo.CihdTable.INVSO = SohdTable_1.SONUMBER LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PcodTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PclsTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PartTable ON dbo.PclsTable.CCREF = dbo.PartTable.PACLASS ON dbo.PcodTable.PCREF = dbo.PartTable.PAPRODCODE ON " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITPART = dbo.PartTable.PARTREF ON dbo.GlacTable.GLACCTREF = dbo.SoitTable.ITREVACCT LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.LoitTable INNER JOIN" & vbCrLf
      sSql = sSql & "               dbo.LohdTable ON dbo.LoitTable.LOINUMBER = dbo.LohdTable.LOTNUMBER RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.InvaTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PshdTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PsitTable ON dbo.PshdTable.PSNUMBER = dbo.PsitTable.PIPACKSLIP ON dbo.InvaTable.INPSNUMBER = dbo.PsitTable.PIPACKSLIP AND " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INPSITEM = dbo.PsitTable.PIITNO ON dbo.LoitTable.LOIPSNUMBER = dbo.InvaTable.INPSNUMBER AND " & vbCrLf
      sSql = sSql & "               dbo.LoitTable.LOIPSITEM = dbo.InvaTable.INPSITEM ON dbo.SoitTable.ITSO = dbo.PsitTable.PISONUMBER AND " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITNUMBER = dbo.PsitTable.PISOITEM AND dbo.SoitTable.ITREV = dbo.PsitTable.PISOREV" & vbCrLf
      sSql = sSql & "ORDER BY dbo.CihdTable.INVNO, dbo.SoitTable.ITSO, dbo.SoitTable.ITNUMBER, dbo.SoitTable.ITREV" & vbCrLf
      Execute False, sSql
      
      sSql = "CREATE VIEW GlActivityView AS" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "SELECT     TOP 100 PERCENT dbo.GlacTable.GLACCTNO AS Account, dbo.GlacTable.GLDESCR AS AcctDesc, dbo.JritTable.DCDATE AS XtnDate, " & vbCrLf
      sSql = sSql & "                      dbo.JrhdTable.MJDESCRIPTION AS XtnDesc, dbo.JritTable.DCDESC AS XtnItmDesc, dbo.JritTable.DCDEBIT AS Debit, " & vbCrLf
      sSql = sSql & "                      dbo.JritTable.DCCREDIT AS Credit, dbo.GlacTable.GLTYPE AS AcctType, dbo.GlacTable.GLINACTIVE AS Inactive, dbo.JritTable.DCHEAD AS XtnID, " & vbCrLf
      sSql = sSql & "                      dbo.JritTable.DCTRAN AS XtnNo, dbo.JritTable.DCREF AS XtnRef" & vbCrLf
      sSql = sSql & "FROM         dbo.GlacTable INNER JOIN" & vbCrLf
      sSql = sSql & "                      dbo.JritTable ON dbo.GlacTable.GLACCTREF = dbo.JritTable.DCACCTNO INNER JOIN" & vbCrLf
      sSql = sSql & "                      dbo.JrhdTable ON dbo.JritTable.DCHEAD = dbo.JrhdTable.MJGLJRNL" & vbCrLf
      sSql = sSql & "WHERE     (NOT (dbo.JritTable.DCHEAD LIKE 'TJ%'))" & vbCrLf
      sSql = sSql & "ORDER BY dbo.GlacTable.GLACCTNO, dbo.JritTable.DCDATE, dbo.JritTable.DCHEAD, dbo.JritTable.DCTRAN, dbo.JritTable.DCREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "Union" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "SELECT     TOP 100 PERCENT dbo.GlacTable.GLACCTNO AS Account, dbo.GlacTable.GLDESCR AS AcctDesc, dbo.GjhdTable.GJPOST AS XtnDate, " & vbCrLf
      sSql = sSql & "                      dbo.GjhdTable.GJDESC AS XtnDesc, dbo.GjitTable.JIDESC AS XtnItmDesc, dbo.GjitTable.JIDEB AS Debit, dbo.GjitTable.JICRD AS Credit, " & vbCrLf
      sSql = sSql & "                      dbo.GlacTable.GLTYPE AS AcctType, dbo.GlacTable.GLINACTIVE AS Inactive, dbo.GjhdTable.GJNAME AS XtnID, dbo.GjitTable.JITRAN AS XtnNo, " & vbCrLf
      sSql = sSql & "                      dbo.GjitTable.JIREF AS XtnRef" & vbCrLf
      sSql = sSql & "FROM         dbo.GlacTable INNER JOIN" & vbCrLf
      sSql = sSql & "                      dbo.GjitTable ON dbo.GlacTable.GLACCTREF = dbo.GjitTable.JIACCOUNT INNER JOIN" & vbCrLf
      sSql = sSql & "                      dbo.GjhdTable ON dbo.GjitTable.JINAME = dbo.GjhdTable.GJNAME" & vbCrLf
      sSql = sSql & "WHERE     (NOT (dbo.GjhdTable.GJNAME LIKE 'CC%')) AND (NOT (dbo.GjhdTable.GJNAME LIKE 'XC%')) AND (NOT (dbo.GjhdTable.GJNAME LIKE 'CR%')) AND " & vbCrLf
      sSql = sSql & "                      (NOT (dbo.GjhdTable.GJNAME LIKE 'PJ%')) AND (NOT (dbo.GjhdTable.GJNAME LIKE 'SJ%')) AND (NOT (dbo.GjhdTable.GJNAME LIKE 'IJ%')) AND " & vbCrLf
      sSql = sSql & "                      (NOT (dbo.GjhdTable.GJNAME LIKE 'TJ%'))" & vbCrLf
      sSql = sSql & "ORDER BY dbo.GlacTable.GLACCTNO, dbo.GjhdTable.GJPOST, dbo.GjhdTable.GJNAME, dbo.GjitTable.JITRAN, dbo.GjitTable.JIREF" & vbCrLf
      Execute False, sSql
      
      sSql = "CREATE VIEW dbo.GLView" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "SELECT     TOP 100 PERCENT dbo.GlacTable.GLTYPE AS Type, dbo.GjitTable.JIACCOUNT AS GLAcctRef, dbo.GlacTable.GLMASTER AS GLMstrAcctNo, " & vbCrLf
      sSql = sSql & "                      dbo.GlacTable.GLINACTIVE AS InactiveFlag, dbo.GlacTable.GLACCTNO, dbo.GlacTable.GLDESCR AS GLDescription, " & vbCrLf
      sSql = sSql & "                      dbo.GjhdTable.GJPOST AS GLPostDate, SUM(dbo.GjitTable.JIDEB) AS Debits, SUM(dbo.GjitTable.JICRD) AS Credits" & vbCrLf
      sSql = sSql & "FROM         dbo.GjitTable INNER JOIN" & vbCrLf
      sSql = sSql & "                      dbo.GlacTable ON dbo.GjitTable.JIACCOUNT = dbo.GlacTable.GLACCTREF LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "                      dbo.GjhdTable ON dbo.GjitTable.JINAME = dbo.GjhdTable.GJNAME" & vbCrLf
      sSql = sSql & "WHERE     (dbo.GjhdTable.GJPOSTED = 1)" & vbCrLf
      sSql = sSql & "GROUP BY dbo.GjitTable.JIACCOUNT, dbo.GjhdTable.GJPOST, dbo.GlacTable.GLDESCR, dbo.GlacTable.GLTYPE, dbo.GlacTable.GLINACTIVE, " & vbCrLf
      sSql = sSql & "                      dbo.GlacTable.GLMASTER, dbo.GlacTable.GLACCTNO" & vbCrLf
      sSql = sSql & "ORDER BY dbo.GlacTable.GLTYPE, dbo.GjitTable.JIACCOUNT, dbo.GjhdTable.GJPOST" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 27"
   End If
   
   If ver < 28 Then
      sSql = "CREATE TABLE TempPickLots(" & vbCrLf
      sSql = sSql & "MoPartRef varchar(30) NOT NULL," & vbCrLf
      sSql = sSql & "MoRunNo int NOT NULL," & vbCrLf
      sSql = sSql & "PickPartRef varchar(30) NOT NULL," & vbCrLf
      sSql = sSql & "LotID char(15) NOT NULL," & vbCrLf
      sSql = sSql & "LotQty decimal(12, 4) NOT NULL," & vbCrLf
      sSql = sSql & "WhenCreated datetime NOT NULL CONSTRAINT DF_TempPickLots_WhenCreated DEFAULT (getdate())"
      sSql = sSql & ")"
      Execute False, sSql
   
      'change from string to date to allow selection of recent errors
      Execute False, "alter table SystemEvents alter column Event_Date datetime"
      
      sSql = "CREATE NONCLUSTERED INDEX IX_TcitTable_TCPARTREF ON TcitTable" & vbCrLf _
         & "(" & vbCrLf _
         & "   TCPARTREF," & vbCrLf _
         & "   TCRUNNO" & vbCrLf _
         & ")"
      Execute False, sSql
      
      sSql = "CREATE NONCLUSTERED INDEX IX_LoitTable_LOIMOPARTREF ON LoitTable " & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "  LOIMOPARTREF," & vbCrLf
      sSql = sSql & "  LOIMORUNNO," & vbCrLf
      sSql = sSql & "  LOITYPE" & vbCrLf
      sSql = sSql & ")"
      Execute False, sSql
      
      sSql = "CREATE NONCLUSTERED INDEX IX_LohdTable_LOTMOPARTREF ON LohdTable " & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "  LOTMOPARTREF," & vbCrLf
      sSql = sSql & "  LOTMORUNNO" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 28"
   End If
   
   If ver < 29 Then
      'in prior versions, canceling a packing slip invoice, sets ITACTUAL = null.
      'this has been fixed in the cancel invoice code.
      'the following sql fixes the erroneous null ITACTUAL's
      sSql = "update SoitTable" & vbCrLf _
         & "set ITACTUAL = PSPRINTED" & vbCrLf _
         & "from SoitTable" & vbCrLf _
         & "join PshdTable on ITPSNUMBER = PSNUMBER" & vbCrLf _
         & "WHERE ITACTUAL IS NULL AND ( ITINVOICE > 0 OR ITPSITEM > 0 )" & vbCrLf _
         & "AND PSPRINTED is not null"
      Execute False, sSql
      
      Execute False, "alter table InvaTable add constraint DF__InvaTable__INPDATE DEFAULT getdate() for INPDATE"
      
      'set version
      Execute False, "update Version set Version = 29"
   End If
   
   If ver < 30 Then
      Execute False, "ALTER TABLE DdocTable DROP CONSTRAINT PK_DdocTable_DOCREF"
      
      'need to temporarily drop index so that column may be revised to not null
      'Execute False, "DROP INDEX DoClass ON DdocTable"
      DropIndex "DdocTable", "DoClass"
      
      Execute False, "ALTER TABLE DdocTable ALTER COLUMN DOCLASS Char(16) NOT NULL"
      
      sSql = "CREATE NONCLUSTERED INDEX DoClass ON DdocTable " & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   [DOCLASS] ASC" & vbCrLf
      sSql = sSql & ")"
      Execute False, sSql
      
      sSql = "ALTER TABLE DdocTable ADD  CONSTRAINT PK_DdocTable_DOCLASS PRIMARY KEY NONCLUSTERED" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   DOCLASS," & vbCrLf
      sSql = sSql & "   DOREF," & vbCrLf
      sSql = sSql & "   DOREV," & vbCrLf
      sSql = sSql & "   DOSHEET" & vbCrLf
      sSql = sSql & ")"
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 30"
   End If
   
   
   If ver < 31 Then
      'don't need a clustered index and this wasn't unique anyway
      Execute False, "ALTER TABLE DlstTable DROP CONSTRAINT PK_DlstTable_DLSREF"
      
      sSql = "CREATE NONCLUSTERED INDEX IX_DlstTable_DLSREF ON DlstTable " & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   DLSREF," & vbCrLf
      sSql = sSql & "   DLSREV," & vbCrLf
      sSql = sSql & "   DLSTYPE" & vbCrLf
      sSql = sSql & ")"
      Execute False, sSql
      
      'delete document list items with no corresponding part
      Execute False, "delete from dlsttable where dlsref not in (select partref from parttable)"

      'to support recosting of PO items
      sSql = "CREATE VIEW viewUnitFreightByInvoice" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "SELECT VITNO AS InvoiceNo," & vbCrLf
      sSql = sSql & "SUM(VITQTY) AS InvoiceQty, MAX(VIFREIGHT) AS Freight," & vbCrLf
      sSql = sSql & "CAST( MAX(VIFREIGHT) / SUM(VITQTY) AS DECIMAL(12,4) ) AS UnitFreight" & vbCrLf
      sSql = sSql & "FROM VihdTable" & vbCrLf
      sSql = sSql & "JOIN ViitTable ON VITNO = VINO" & vbCrLf
      sSql = sSql & "JOIN InvaTable ON INPONUMBER = VITPO AND INPORELEASE = VITPORELEASE" & vbCrLf
      sSql = sSql & "AND INPOITEM = VITPOITEM AND INPOREV = VITPOITEMREV" & vbCrLf
      sSql = sSql & "AND INTYPE IN (15,17)" & vbCrLf
      sSql = sSql & "WHERE VIFREIGHT <> 0 AND VITPO <> 0 AND VITQTY <> 0" & vbCrLf
      sSql = sSql & "GROUP BY VITNO" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 31"
   End If
   
   
   If ver < 32 Then
      'need to repeat ver 30 stuff due to change in drop index syntax
      Execute False, "ALTER TABLE DdocTable DROP CONSTRAINT PK_DdocTable_DOCREF"
      
      'need to temporarily drop index so that column may be revised to not null
      'Execute False, "DROP INDEX DoClass ON DdocTable"    'SQL2005
      'Execute False, "DROP INDEX DdocTable.DoClass"       'SQL2000
      DropIndex "DdocTable", "DoClass"

      Execute False, "ALTER TABLE DdocTable ALTER COLUMN DOCLASS Char(16) NOT NULL"
      
      sSql = "CREATE NONCLUSTERED INDEX DoClass ON DdocTable " & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   [DOCLASS] ASC" & vbCrLf
      sSql = sSql & ")"
      Execute False, sSql
      
      sSql = "ALTER TABLE DdocTable ADD  CONSTRAINT PK_DdocTable_DOCLASS PRIMARY KEY NONCLUSTERED" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   DOCLASS," & vbCrLf
      sSql = sSql & "   DOREF," & vbCrLf
      sSql = sSql & "   DOREV," & vbCrLf
      sSql = sSql & "   DOSHEET" & vbCrLf
      sSql = sSql & ")"
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 32"
   End If
   
   If ver < 33 Then
      sSql = "CREATE NONCLUSTERED INDEX IX_LoitTable_LOIPONUMBER ON LoitTable" & vbCrLf _
         & "( LOIPONUMBER, LOINUMBER, LOITYPE, LOIADATE, LOIQUANTITY, LOIPOITEM, LOIPOREV )"
      Execute False, sSql
      
      'allow backlog report table to be recreated with an int id (previously smallint)
      Execute False, "drop table EsReportSale06"
      
      'set version
      Execute False, "update Version set Version = 33"
   End If
   
   If ver < 34 Then
      ' to prevent null columns below
      Execute False, "DELETE FROM TempPsLots"
      Execute False, "ALTER TABLE TempPsLots ADD PartRef varchar(30) NOT NULL default ''"
      
      'document class table reorg
      Execute False, "ALTER TABLE DclsTableORIG DROP CONSTRAINT PK_DCLSTable_DCLREF"    'need PK name used in conversion
      sSql = "ALTER TABLE DclsTable ADD  CONSTRAINT PK_DclsTable_DCLREF PRIMARY KEY NONCLUSTERED" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   DCLREF" & vbCrLf
      sSql = sSql & ")"
      Execute False, sSql
      Execute False, "alter table DclsTable add constraint DF__DclsTable__DCLNAME DEFAULT '' for DCLNAME"
      Execute False, "alter table DclsTable add constraint DF__DclsTable__DCLDESC DEFAULT '' for DCLDESC"
      Execute False, "alter table DclsTable add constraint DF__DclsTable__DCLNOTES DEFAULT '' for DCLNOTES"
      Execute False, "alter table DclsTable add constraint DF__DclsTable__DCLSHEETS DEFAULT 0 for DCLSHEETS"
      Execute False, "alter table DclsTable add constraint DF__DclsTable__DCLADCN DEFAULT 0 for DCLADCN"
      
      'views for Larry's reports from db ver 27 -- SQL2000 didn't like select top (100) percent
      sSql = "CREATE VIEW [dbo].[Vw_Sales]" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "SELECT DISTINCT TOP 100 PERCENT dbo.CihdTable.INVTYPE, dbo.CihdTable.INVPRE, dbo.CihdTable.INVNO, dbo.LoitTable.LOICUSTINVNO, dbo.CihdTable.INVCUST, " & vbCrLf
      sSql = sSql & "               dbo.CustTable.CUNUMBER, dbo.CustTable.CUNAME, dbo.CihdTable.INVDATE, dbo.CihdTable.INVPIF, dbo.PshdTable.PSCANCELED, " & vbCrLf
      sSql = sSql & "               dbo.PshdTable.PSTYPE, dbo.PsitTable.PIPACKSLIP, dbo.PsitTable.PIITNO, dbo.InvaTable.INPSNUMBER, dbo.InvaTable.INPSITEM, " & vbCrLf
      sSql = sSql & "               dbo.LoitTable.LOIPSNUMBER, dbo.LoitTable.LOIPSITEM, dbo.SoitTable.ITPSSHIPPED, dbo.SohdTable.SOSALESMAN AS PSSoSalesman, " & vbCrLf
      sSql = sSql & "               SohdTable_1.SOSALESMAN AS SOSoSlsmn, dbo.SohdTable.SOPO AS PSSoPo, SohdTable_1.SOPO AS SOSoPo, " & vbCrLf
      sSql = sSql & "               dbo.SohdTable.SODIVISION AS PSSoDiv, SohdTable_1.SODIVISION AS SOSoDiv, dbo.SohdTable.SOREGION AS PSSoReg, " & vbCrLf
      sSql = sSql & "               SohdTable_1.SOREGION AS SOSoReg, dbo.SohdTable.SOBUSUNIT AS PSSoBu, SohdTable_1.SOBUSUNIT AS SOSoBu, " & vbCrLf
      sSql = sSql & "               dbo.SprsTable.SPLAST AS PSSlsmnLast, SprsTable_1.SPLAST AS SOSlsmnLast, dbo.SprsTable.SPFIRST AS PSSlsmnFirst, " & vbCrLf
      sSql = sSql & "               SprsTable_1.SPFIRST AS SOSlsmnFirst, dbo.SprsTable.SPMIDD AS PSSlsmnInit, SprsTable_1.SPMIDD AS SOSlsmnInit, " & vbCrLf
      sSql = sSql & "               dbo.SohdTable.SOTYPE AS SOSoType, SohdTable_1.SOTYPE AS PSSoType, dbo.SoitTable.ITSO, dbo.SoitTable.ITNUMBER, dbo.SoitTable.ITREV, " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITPART, dbo.PartTable.PARTNUM, dbo.PartTable.PADESC, dbo.PartTable.PALEVEL, dbo.PartTable.PALOTTRACK, " & vbCrLf
      sSql = sSql & "               dbo.PartTable.PAUSEACTUALCOST, dbo.PartTable.PAUNITS, dbo.PartTable.PAMAKEBUY, dbo.PartTable.PAFAMILY, dbo.PartTable.PAPRODCODE, " & vbCrLf
      sSql = sSql & "               dbo.PcodTable.PCDESC, dbo.PartTable.PACLASS, dbo.PclsTable.CCDESC, dbo.SoitTable.ITQTY, dbo.SoitTable.ITDOLLARS, dbo.SoitTable.ITADJUST, " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITDISCAMOUNT, dbo.SoitTable.ITCOMMISSION, dbo.SoitTable.ITBOOKDATE, dbo.SoitTable.ITSCHED, dbo.SoitTable.ITACTUAL, " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITCANCELDATE, dbo.SoitTable.ITCANCELED, dbo.SoitTable.ITREVACCT, dbo.GlacTable.GLACCTNO AS RevAcct, " & vbCrLf
      sSql = sSql & "               dbo.GlacTable.GLDESCR AS RevAcctDesc, dbo.SoitTable.ITCGSACCT, dbo.SoitTable.ITDISACCT, dbo.SoitTable.ITSTATE, dbo.SoitTable.ITTAXCODE, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INTYPE, dbo.InvaTable.INAQTY, dbo.InvaTable.INAMT, dbo.InvaTable.INTOTLABOR, dbo.InvaTable.INTOTMATL, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INTOTEXP, dbo.InvaTable.INTOTOH, dbo.PartTable.PASTDCOST, dbo.PartTable.PATOTCOST, dbo.PartTable.PATOTLABOR, " & vbCrLf
      sSql = sSql & "               dbo.PartTable.PATOTMATL, dbo.PartTable.PATOTEXP, dbo.PartTable.PATOTOH, dbo.InvaTable.INDRLABACCT, dbo.InvaTable.INDRMATACCT, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INDREXPACCT, dbo.InvaTable.INDROHDACCT, dbo.InvaTable.INCRLABACCT, dbo.InvaTable.INCRMATACCT, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INCREXPACCT, dbo.InvaTable.INCROHDACCT, dbo.LoitTable.LOITYPE, dbo.LoitTable.LOINUMBER, dbo.LoitTable.LOIRECORD, " & vbCrLf
      sSql = sSql & "               dbo.LohdTable.LOTNUMBER, dbo.LohdTable.LOTUSERLOTID, dbo.LohdTable.LOTPARTREF, dbo.LohdTable.LOTDATECOSTED, " & vbCrLf
      sSql = sSql & "          dbo.LohdTable.LOTTOTLABOR, dbo.LohdTable.LOTTOTMATL, dbo.LohdTable.LOTTOTEXP, dbo.LohdTable.LOTTOTOH" & vbCrLf
      sSql = sSql & "FROM  dbo.GlacTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SprsTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SoitTable LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SohdTable ON dbo.SoitTable.ITSO = dbo.SohdTable.SONUMBER ON " & vbCrLf
      sSql = sSql & "               dbo.SprsTable.SPNUMBER = dbo.SohdTable.SOSALESMAN RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.CustTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.CihdTable ON dbo.CustTable.CUREF = dbo.CihdTable.INVCUST ON dbo.SoitTable.ITINVOICE = dbo.CihdTable.INVNO LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SprsTable AS SprsTable_1 RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SohdTable AS SohdTable_1 ON SprsTable_1.SPNUMBER = SohdTable_1.SOSALESMAN ON " & vbCrLf
      sSql = sSql & "               dbo.CihdTable.INVSO = SohdTable_1.SONUMBER LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PcodTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PclsTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PartTable ON dbo.PclsTable.CCREF = dbo.PartTable.PACLASS ON dbo.PcodTable.PCREF = dbo.PartTable.PAPRODCODE ON " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITPART = dbo.PartTable.PARTREF ON dbo.GlacTable.GLACCTREF = dbo.SoitTable.ITREVACCT LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.LoitTable INNER JOIN" & vbCrLf
      sSql = sSql & "               dbo.LohdTable ON dbo.LoitTable.LOINUMBER = dbo.LohdTable.LOTNUMBER RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.InvaTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PshdTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PsitTable ON dbo.PshdTable.PSNUMBER = dbo.PsitTable.PIPACKSLIP ON dbo.InvaTable.INPSNUMBER = dbo.PsitTable.PIPACKSLIP AND " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INPSITEM = dbo.PsitTable.PIITNO ON dbo.LoitTable.LOIPSNUMBER = dbo.InvaTable.INPSNUMBER AND " & vbCrLf
      sSql = sSql & "               dbo.LoitTable.LOIPSITEM = dbo.InvaTable.INPSITEM ON dbo.SoitTable.ITSO = dbo.PsitTable.PISONUMBER AND " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITNUMBER = dbo.PsitTable.PISOITEM AND dbo.SoitTable.ITREV = dbo.PsitTable.PISOREV" & vbCrLf
      sSql = sSql & "ORDER BY dbo.CihdTable.INVNO, dbo.SoitTable.ITSO, dbo.SoitTable.ITNUMBER, dbo.SoitTable.ITREV" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 34"
   End If

   If ver < 35 Then
      Execute False, "ALTER TABLE PoitTable ADD PIONDOCKQTYWASTE decimal(12,4) NULL default 0"
      Execute False, "UPDATE PoitTable SET PIONDOCKQTYWASTE = 0.0000 WHERE PIONDOCKQTYWASTE IS NULL"
      
      'set version
      Execute False, "update Version set Version = 35"
   End If

   If ver < 36 Then
      Execute False, "ALTER TABLE CustTable ADD CUCREDITLIMIT decimal(12,2) NULL default 0"
      Execute False, "UPDATE CustTable SET CUCREDITLIMIT = 0.00 WHERE CUCREDITLIMIT IS NULL"
      
      Execute False, "ALTER TABLE ComnTable ADD COPODEFAULTTOLASTPRICE bit NULL default 0"
      Execute False, "UPDATE ComnTable SET COPODEFAULTTOLASTPRICE = 0 WHERE COPODEFAULTTOLASTPRICE IS NULL"
      
      Execute False, "ALTER TABLE ComnTable ADD COLASTINVOICENUMBER int NULL default 0"
      Execute False, "UPDATE ComnTable SET COLASTINVOICENUMBER = (SELECT MAX(INVNO) from CihdTable)" & vbCrLf _
         & "WHERE COLASTINVOICENUMBER IS NULL"
      
      'set version
      Execute False, "update Version set Version = 36"
   End If

   If ver < 37 Then
      Execute False, "ALTER TABLE ComnTable ADD COPSPREFIX varchar(2) NULL default 'PS'"
      Execute False, "UPDATE ComnTable SET COPSPREFIX = 'PS'" & vbCrLf _
         & "WHERE COPSPREFIX IS NULL"
      
      Execute False, "ALTER TABLE ComnTable ADD COLASTPSNUMBER INT NULL default 0"
      Execute False, "UPDATE ComnTable SET COLASTPSNUMBER = ISNULL(CAST(CURPSNUMBER AS int), 0)" & vbCrLf _
         & "WHERE COLASTPSNUMBER IS NULL"
      
      'set version
      Execute False, "update Version set Version = 37"
   End If

   If ver < 38 Then
      'views for Larry's reports didn't include LOTUNITCOST
      Execute False, "drop view Vw_Sales"
      
      sSql = "CREATE VIEW [dbo].[Vw_Sales]" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "SELECT DISTINCT TOP 100 PERCENT dbo.CihdTable.INVTYPE, dbo.CihdTable.INVPRE, dbo.CihdTable.INVNO, dbo.LoitTable.LOICUSTINVNO, dbo.CihdTable.INVCUST, " & vbCrLf
      sSql = sSql & "               dbo.CustTable.CUNUMBER, dbo.CustTable.CUNAME, dbo.CihdTable.INVDATE, dbo.CihdTable.INVPIF, dbo.PshdTable.PSCANCELED, " & vbCrLf
      sSql = sSql & "               dbo.PshdTable.PSTYPE, dbo.PsitTable.PIPACKSLIP, dbo.PsitTable.PIITNO, dbo.InvaTable.INPSNUMBER, dbo.InvaTable.INPSITEM, " & vbCrLf
      sSql = sSql & "               dbo.LoitTable.LOIPSNUMBER, dbo.LoitTable.LOIPSITEM, dbo.SoitTable.ITPSSHIPPED, dbo.SohdTable.SOSALESMAN AS PSSoSalesman, " & vbCrLf
      sSql = sSql & "               SohdTable_1.SOSALESMAN AS SOSoSlsmn, dbo.SohdTable.SOPO AS PSSoPo, SohdTable_1.SOPO AS SOSoPo, " & vbCrLf
      sSql = sSql & "               dbo.SohdTable.SODIVISION AS PSSoDiv, SohdTable_1.SODIVISION AS SOSoDiv, dbo.SohdTable.SOREGION AS PSSoReg, " & vbCrLf
      sSql = sSql & "               SohdTable_1.SOREGION AS SOSoReg, dbo.SohdTable.SOBUSUNIT AS PSSoBu, SohdTable_1.SOBUSUNIT AS SOSoBu, " & vbCrLf
      sSql = sSql & "               dbo.SprsTable.SPLAST AS PSSlsmnLast, SprsTable_1.SPLAST AS SOSlsmnLast, dbo.SprsTable.SPFIRST AS PSSlsmnFirst, " & vbCrLf
      sSql = sSql & "               SprsTable_1.SPFIRST AS SOSlsmnFirst, dbo.SprsTable.SPMIDD AS PSSlsmnInit, SprsTable_1.SPMIDD AS SOSlsmnInit, " & vbCrLf
      sSql = sSql & "               dbo.SohdTable.SOTYPE AS SOSoType, SohdTable_1.SOTYPE AS PSSoType, dbo.SoitTable.ITSO, dbo.SoitTable.ITNUMBER, dbo.SoitTable.ITREV, " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITPART, dbo.PartTable.PARTNUM, dbo.PartTable.PADESC, dbo.PartTable.PALEVEL, dbo.PartTable.PALOTTRACK, " & vbCrLf
      sSql = sSql & "               dbo.PartTable.PAUSEACTUALCOST, dbo.PartTable.PAUNITS, dbo.PartTable.PAMAKEBUY, dbo.PartTable.PAFAMILY, dbo.PartTable.PAPRODCODE, " & vbCrLf
      sSql = sSql & "               dbo.PcodTable.PCDESC, dbo.PartTable.PACLASS, dbo.PclsTable.CCDESC, dbo.SoitTable.ITQTY, dbo.SoitTable.ITDOLLARS, dbo.SoitTable.ITADJUST, " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITDISCAMOUNT, dbo.SoitTable.ITCOMMISSION, dbo.SoitTable.ITBOOKDATE, dbo.SoitTable.ITSCHED, dbo.SoitTable.ITACTUAL, " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITCANCELDATE, dbo.SoitTable.ITCANCELED, dbo.SoitTable.ITREVACCT, dbo.GlacTable.GLACCTNO AS RevAcct, " & vbCrLf
      sSql = sSql & "               dbo.GlacTable.GLDESCR AS RevAcctDesc, dbo.SoitTable.ITCGSACCT, dbo.SoitTable.ITDISACCT, dbo.SoitTable.ITSTATE, dbo.SoitTable.ITTAXCODE, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INTYPE, dbo.InvaTable.INAQTY, dbo.InvaTable.INAMT, dbo.InvaTable.INTOTLABOR, dbo.InvaTable.INTOTMATL, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INTOTEXP, dbo.InvaTable.INTOTOH, dbo.PartTable.PASTDCOST, dbo.PartTable.PATOTCOST, dbo.PartTable.PATOTLABOR, " & vbCrLf
      sSql = sSql & "               dbo.PartTable.PATOTMATL, dbo.PartTable.PATOTEXP, dbo.PartTable.PATOTOH, dbo.InvaTable.INDRLABACCT, dbo.InvaTable.INDRMATACCT, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INDREXPACCT, dbo.InvaTable.INDROHDACCT, dbo.InvaTable.INCRLABACCT, dbo.InvaTable.INCRMATACCT, " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INCREXPACCT, dbo.InvaTable.INCROHDACCT, dbo.LoitTable.LOITYPE, dbo.LoitTable.LOINUMBER, dbo.LoitTable.LOIRECORD, " & vbCrLf
      sSql = sSql & "               dbo.LohdTable.LOTNUMBER, dbo.LohdTable.LOTUSERLOTID, dbo.LohdTable.LOTPARTREF, dbo.LohdTable.LOTDATECOSTED, " & vbCrLf
      sSql = sSql & "          dbo.LohdTable.LOTTOTLABOR, dbo.LohdTable.LOTTOTMATL, dbo.LohdTable.LOTTOTEXP, dbo.LohdTable.LOTTOTOH, dbo.LohdTable.LOTUNITCOST" & vbCrLf
      sSql = sSql & "FROM  dbo.GlacTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SprsTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SoitTable LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SohdTable ON dbo.SoitTable.ITSO = dbo.SohdTable.SONUMBER ON " & vbCrLf
      sSql = sSql & "               dbo.SprsTable.SPNUMBER = dbo.SohdTable.SOSALESMAN RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.CustTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.CihdTable ON dbo.CustTable.CUREF = dbo.CihdTable.INVCUST ON dbo.SoitTable.ITINVOICE = dbo.CihdTable.INVNO LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SprsTable AS SprsTable_1 RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.SohdTable AS SohdTable_1 ON SprsTable_1.SPNUMBER = SohdTable_1.SOSALESMAN ON " & vbCrLf
      sSql = sSql & "               dbo.CihdTable.INVSO = SohdTable_1.SONUMBER LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PcodTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PclsTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PartTable ON dbo.PclsTable.CCREF = dbo.PartTable.PACLASS ON dbo.PcodTable.PCREF = dbo.PartTable.PAPRODCODE ON " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITPART = dbo.PartTable.PARTREF ON dbo.GlacTable.GLACCTREF = dbo.SoitTable.ITREVACCT LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.LoitTable INNER JOIN" & vbCrLf
      sSql = sSql & "               dbo.LohdTable ON dbo.LoitTable.LOINUMBER = dbo.LohdTable.LOTNUMBER RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.InvaTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PshdTable RIGHT OUTER JOIN" & vbCrLf
      sSql = sSql & "               dbo.PsitTable ON dbo.PshdTable.PSNUMBER = dbo.PsitTable.PIPACKSLIP ON dbo.InvaTable.INPSNUMBER = dbo.PsitTable.PIPACKSLIP AND " & vbCrLf
      sSql = sSql & "               dbo.InvaTable.INPSITEM = dbo.PsitTable.PIITNO ON dbo.LoitTable.LOIPSNUMBER = dbo.InvaTable.INPSNUMBER AND " & vbCrLf
      sSql = sSql & "               dbo.LoitTable.LOIPSITEM = dbo.InvaTable.INPSITEM ON dbo.SoitTable.ITSO = dbo.PsitTable.PISONUMBER AND " & vbCrLf
      sSql = sSql & "               dbo.SoitTable.ITNUMBER = dbo.PsitTable.PISOITEM AND dbo.SoitTable.ITREV = dbo.PsitTable.PISOREV" & vbCrLf
      sSql = sSql & "ORDER BY dbo.CihdTable.INVNO, dbo.SoitTable.ITSO, dbo.SoitTable.ITNUMBER, dbo.SoitTable.ITREV" & vbCrLf
      Execute False, sSql
      
      sSql = "create view viewOpenAPTerms" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select VIVENDOR, VENICKNAME, VEBNAME, VINO, VIDATE, VIDUE as InvTotal, VIDUEDATE, " & vbCrLf
      sSql = sSql & "VIFREIGHT, VITAX, VIPAY, " & vbCrLf
      sSql = sSql & "isnull(VITPO,0) as VITPO, isnull(VITPORELEASE,0) as VITPORELEASE," & vbCrLf
      sSql = sSql & "CASE WHEN PONETDAYS <> 0 THEN PODISCOUNT ELSE VEDISCOUNT END as DiscRate," & vbCrLf
      sSql = sSql & "CASE WHEN PONETDAYS <> 0 THEN PODDAYS ELSE VEDDAYS END as DiscDays," & vbCrLf
      sSql = sSql & "CASE WHEN PONETDAYS <> 0 THEN PONETDAYS ELSE VENETDAYS END as NetDays," & vbCrLf
      sSql = sSql & "VIDUE - VIPAY as AmountDue," & vbCrLf
      sSql = sSql & "VIFREIGHT + VITAX + (SELECT CAST(SUM(ROUND(cast(VITQTY as decimal(12,3)) " & vbCrLf
      sSql = sSql & "* cast(VITCOST as decimal(12,4)) " & vbCrLf
      sSql = sSql & "+ cast(VITADDERS as decimal(12,2)), 2)) as decimal(12,2))" & vbCrLf
      sSql = sSql & "FROM ViitTable WHERE VITVENDOR=VIVENDOR AND VITNO=VINO) as CalcTotal" & vbCrLf
      sSql = sSql & "from VihdTable" & vbCrLf
      sSql = sSql & "join VndrTable on VIVENDOR = VEREF" & vbCrLf
      sSql = sSql & "left join ViitTable on VITNO = VINO" & vbCrLf
      sSql = sSql & "and VITVENDOR = VIVENDOR" & vbCrLf
      sSql = sSql & "and VITITEM = ( select min(VITITEM) from ViitTable" & vbCrLf
      sSql = sSql & "where VITVENDOR = VEREF and VITNO = VINO" & vbCrLf
      sSql = sSql & "and VITPO is not null and VITPO <> 0)" & vbCrLf
      sSql = sSql & "left join PohdTable on VITPO = PONUMBER" & vbCrLf
      sSql = sSql & "and VITPORELEASE = PORELEASE" & vbCrLf
      sSql = sSql & "where VIPIF <> 1" & vbCrLf
      Execute False, sSql
      
      sSql = "create view viewOpenAP" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select *," & vbCrLf
      sSql = sSql & "cast(AmountDue * DiscRate / 100 as decimal(12,2)) as DiscAmount," & vbCrLf
      sSql = sSql & "CASE WHEN DiscRate > 0 " & vbCrLf
      sSql = sSql & "THEN DATEADD(day,DiscDays,VIDATE)" & vbCrLf
      sSql = sSql & "ELSE null END as DiscCutoff" & vbCrLf
      sSql = sSql & "from viewOpenAPTerms" & vbCrLf
      Execute False, sSql
      
      sSql = "create function fnGetOpenAP" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   @MinDiscountDate datetime, -- show discounts available on or after this date" & vbCrLf
      sSql = sSql & "   @MaxDueDate datetime    -- show invoices due on or before this date" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "returns table" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "return" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "select *," & vbCrLf
      sSql = sSql & "cast(case when DiscCutoff is null then null" & vbCrLf
      sSql = sSql & "when datediff(d, DiscCutoff, @MinDiscountDate) > 0 then null" & vbCrLf
      sSql = sSql & "else DiscCutoff end as datetime) as TakeDiscByDate," & vbCrLf
      sSql = sSql & "case when DiscCutoff is null then 0.00" & vbCrLf
      sSql = sSql & "when datediff(d, DiscCutoff, @MinDiscountDate) > 0 then 0.00" & vbCrLf
      sSql = sSql & "else DiscAmount end as DiscAvail" & vbCrLf
      sSql = sSql & "from viewOpenAP" & vbCrLf
      sSql = sSql & "where VIDUEDATE <= @MaxDueDate " & vbCrLf
      sSql = sSql & "or DiscCutoff >= @MinDiscountDate" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 38"
   End If

   If ver < 39 Then
      sSql = "CREATE TABLE [dbo].[EsReportUsers]" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   [UserName] [varchar](40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf
      sSql = sSql & "   [Initials] [varchar](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & vbCrLf
      sSql = sSql & "   [Nickname] [varchar](20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & vbCrLf
      sSql = sSql & "   [Active] [bit] NULL," & vbCrLf
      sSql = sSql & "   [Created] [datetime] NULL," & vbCrLf
      sSql = sSql & "   [Level] [int] NULL," & vbCrLf
      sSql = sSql & "   [Admin] bit NULL," & vbCrLf
      sSql = sSql & "   CONSTRAINT [PK_EsReportUsers] PRIMARY KEY CLUSTERED " & vbCrLf
      sSql = sSql & "   (" & vbCrLf
      sSql = sSql & "      [UserName] ASC" & vbCrLf
      sSql = sSql & "   )" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      Execute False, sSql

      sSql = "CREATE TABLE [dbo].[EsReportUserPermissions]" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   [UserName] [varchar](40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf
      sSql = sSql & "   [ModuleName] [varchar](40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf
      sSql = sSql & "   [SectionName] [varchar](40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf
      sSql = sSql & "   [GroupPermission] [bit] NOT NULL," & vbCrLf
      sSql = sSql & "   [EditPermission] [bit] NOT NULL," & vbCrLf
      sSql = sSql & "   [ViewPermission] [bit] NOT NULL," & vbCrLf
      sSql = sSql & "   [FunctionPermission] [bit] NOT NULL" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      Execute False, sSql
      
      sSql = "ALTER TABLE [dbo].[EsReportUserPermissions]  WITH CHECK ADD  CONSTRAINT [FK_EsReportUserPermissions_EsReportUsers] FOREIGN KEY([UserName])" & vbCrLf
      sSql = sSql & "REFERENCES [dbo].[EsReportUsers] ([UserName])" & vbCrLf
      Execute False, sSql
      
      sSql = "ALTER TABLE [dbo].[EsReportUserPermissions] CHECK CONSTRAINT [FK_EsReportUserPermissions_EsReportUsers]" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 39"
   End If

   If ver < 40 Then
   
      sSql = "alter view viewLotCostsByMoDetails" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select LOITYPE, rtrim(LOIMOPARTREF) as MoPart, LOIMORUNNO as MoRun," & vbCrLf
      sSql = sSql & "rtrim(LOTPARTREF) as Part," & vbCrLf
      sSql = sSql & "-LOIQUANTITY as Quantity, LOTUNITCOST as UnitCost," & vbCrLf
      sSql = sSql & "cast(-LOTUNITCOST * LOIQUANTITY as decimal(15,4)) as TotalCost," & vbCrLf
      sSql = sSql & "LOTNUMBER as LotNumber, rtrim(LOTUSERLOTID) as LotUserID, LOTMAINTCOSTED,LOTDATECOSTED" & vbCrLf
      sSql = sSql & "FROM LohdTable join LoitTable on LOTNUMBER = LOINUMBER and LOITYPE = 10" & vbCrLf
      sSql = sSql & "join PartTable on LOTPARTREF = PARTREF AND PALEVEL < 5 and PALOTTRACK = 1" & vbCrLf
      Execute False, sSql
      
      sSql = "CREATE INDEX IX_InvaTable_INMOPART_INMORUN_INTYPE ON InvaTable" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   INMOPART," & vbCrLf
      sSql = sSql & "   INMORUN," & vbCrLf
      sSql = sSql & "   INTYPE" & vbCrLf
      sSql = sSql & ")"
      Execute False, sSql
      
      sSql = "CREATE INDEX IX_InvaTable_INAQTY_INNUMBER ON InvaTable" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   INAQTY," & vbCrLf
      sSql = sSql & "   INNUMBER" & vbCrLf
      sSql = sSql & ")"
      Execute False, sSql
      
      sSql = "CREATE INDEX IX_LoitTable_LOINUMBER_LOIACTIVITY ON LoitTable" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   LOINUMBER," & vbCrLf
      sSql = sSql & "   LOIACTIVITY" & vbCrLf
      sSql = sSql & ")"
      Execute False, sSql
   
      'set version
      Execute False, "update Version set Version = 40"
   End If

   If ver < 41 Then
   
      sSql = "create view viewMaintLotRemQtyVsLoiQty" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select 'LOTREMAININGQTY' as QtyCol, 'sum(LOIQUANTITY)' as SumCol, LOTPARTREF as PartRef, LotNumber," & vbCrLf
      sSql = sSql & "LOTREMAININGQTY as Qty, sum(LOIQUANTITY) as SumQty" & vbCrLf
      sSql = sSql & "from LohdTable" & vbCrLf
      sSql = sSql & "join LoitTable on LOINUMBER = LOTNUMBER" & vbCrLf
      sSql = sSql & "group by LOTPARTREF, LOTNUMBER, LOTREMAININGQTY having LOTREMAININGQTY <> sum(LOIQUANTITY)"
      Execute False, sSql

      sSql = "create view viewMaintPaQohVsLotRemQty" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select 'PAQOH' as QtyCol, 'sum(LOTREMAININGQTY)' as SumCol, PartRef, '' as LotNumber," & vbCrLf
      sSql = sSql & "PAQOH as Qty, sum(LOTREMAININGQTY) as SumQty" & vbCrLf
      sSql = sSql & "from LohdTable" & vbCrLf
      sSql = sSql & "join PartTable on LOTPARTREF = PARTREF" & vbCrLf
      sSql = sSql & "group by PARTREF, PAQOH having PAQOH <> sum(LOTREMAININGQTY)"
      Execute False, sSql

      sSql = "create view viewMaintPaQohVsInaQty" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "select 'PAQOH' as QtyCol, 'sum(INAQTY)' as SumCol, PartRef, '' as LotNumber," & vbCrLf
      sSql = sSql & "PAQOH as Qty, sum(INAQTY) as SumQty" & vbCrLf
      sSql = sSql & "from PartTable" & vbCrLf
      sSql = sSql & "join InvaTable on INPART = PARTREF" & vbCrLf
      sSql = sSql & "group by PARTREF, PAQOH having PAQOH <> sum(INAQTY)"
      Execute False, sSql
   
      'set version
      Execute False, "update Version set Version = 41"
   End If

   If ver < 42 Then
      Execute False, "ALTER TABLE CcltTable ADD CLRECONCILED tinyint NULL default 0"
      Execute False, "ALTER TABLE CcltTable ADD CLRECONCILEDDATE datetime NULL"
      
      'set version
      Execute False, "update Version set Version = 42"
   End If

   If ver < 43 Then
   
      sSql = "CREATE TABLE CCLog" & vbCrLf _
         & "(" & vbCrLf _
         & "   CCREF varchar(20) NOT NULL default ''," & vbCrLf _
         & "   PARTREF varchar(30) NOT NULL default ''," & vbCrLf _
         & "   LOTNUMBER varchar(15) NULL," & vbCrLf _
         & "   LOGTEXT varchar(80) NOT NULL default ''" & vbCrLf _
         & ")"
      Execute False, sSql
      
      sSql = "CREATE TABLE CCLotAlloc" & vbCrLf _
         & "(" & vbCrLf _
         & "   CCREF varchar(20) NOT NULL default ''," & vbCrLf _
         & "   PARTREF varchar(30) NOT NULL default ''," & vbCrLf _
         & "   LOTNUMBER varchar(15) NOT NULL default ''," & vbCrLf _
         & "   OLDQTY decimal(15,3) NOT NULL default 0," & vbCrLf _
         & "   NEWQTY decimal(15,3) NOT NULL default 0" & vbCrLf _
         & ")"
      Execute False, sSql
      
      Execute False, "Exec sp_rename 'CchdTable.CCRECONCILED', 'CCUPDATED', 'COLUMN'"
      Execute False, "Exec sp_rename 'CchdTable.CCINVRECDATE', 'CCUPDATEDDATE', 'COLUMN'"
      DropColumnDefault "CchdTable", "CCCOUNTERS"
      Execute False, "Alter Table CchdTable drop column CCCOUNTERS"
      DropColumnDefault "CchdTable", "CCACTUALDATE"
      Execute False, "Alter Table CchdTable drop column CCACTUALDATE"
      
      DropColumnDefault "CcitTable", "CIRECONCILED"
      Execute False, "Alter Table CcitTable drop column CIRECONCILED"
      DropColumnDefault "CcitTable", "CIRECONCILEDDATE"
      Execute False, "Alter Table CcitTable drop column CIRECONCILEDDATE"
      
      Execute False, "Exec sp_rename 'CcltTable.CLRECONCILED', 'CLENTERED', 'COLUMN'"
      Execute False, "Exec sp_rename 'CcltTable.CLRECONCILEDDATE', 'CLENTEREDDATE', 'COLUMN'"
      
      sSql = "ALTER PROCEDURE [dbo].[Qry_GetCycleCount]" & vbCrLf _
         & "(@cyclecount char(20))" & vbCrLf _
         & "as" & vbCrLf _
         & "SELECT CCREF,CCDESC,CCPLANDATE,CCABCCODE,CCCOUNTSAVED" & vbCrLf _
         & "FROM CchdTable WHERE (CCREF=@cyclecount AND CCCOUNTLOCKED=1" & vbCrLf _
         & "AND CCUPDATED=0)"
     Execute False, sSql
      
      sSql = "create procedure AllocateCycleCountLots" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   @CCID varchar(20)    -- cycle count ID" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "/* test" & vbCrLf
      sSql = sSql & "   AllocateCycleCountLots '20080726-A'" & vbCrLf
      sSql = sSql & "*/" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "delete from CCLog where CCREF = @CCID" & vbCrLf
      sSql = sSql & "delete from CCLoTaLLOC where CCREF = @CCID" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- insert info for lot-tracked parts" & vbCrLf
      sSql = sSql & "insert into CCLOTALLOC" & vbCrLf
      sSql = sSql & "select @CCID, rtrim(CLPARTREF), rtrim(CLLOTNUMBER), CLLOTREMAININGQTY, CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER <> '' --and CLLOTREMAININGQTY <> CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "order by CLPARTREF, CLLOTNUMBER" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "---------------------------------------" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- find all non-lot-tracked parts with increasing quantity" & vbCrLf
      sSql = sSql & "-- apply increases to newest lots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare @partRef varchar(30)" & vbCrLf
      sSql = sSql & "declare @initialQty decimal(15,4)" & vbCrLf
      sSql = sSql & "declare @countQty decimal(15,4)" & vbCrLf
      sSql = sSql & "declare @adjustQty decimal(15,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare more cursor for" & vbCrLf
      sSql = sSql & "select rtrim(CLPARTREF), CLLOTREMAININGQTY, CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER = '' and CLLOTREMAININGQTY < CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "order by CLPARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare @lotNumber varchar(15)" & vbCrLf
      sSql = sSql & "declare @lotRemaining decimal(15,4)" & vbCrLf
      sSql = sSql & "declare @availQty decimal(15,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "open more" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "while (1 = 1)" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   fetch next from more into @partRef, @initialQty, @countQty" & vbCrLf
      sSql = sSql & "   if (@@FETCH_STATUS <> 0)break" & vbCrLf
      sSql = sSql & "   set @adjustQty = @countQty - @initialQty" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- find available lots to add to" & vbCrLf
      sSql = sSql & "   declare lifoLots cursor for" & vbCrLf
      sSql = sSql & "   select LOTNUMBER, LOTREMAININGQTY, LOTORIGINALQTY - LOTREMAININGQTY as Available" & vbCrLf
      sSql = sSql & "   from LohdTable" & vbCrLf
      sSql = sSql & "   where LOTPARTREF = @partRef" & vbCrLf
      sSql = sSql & "   and LOTORIGINALQTY > LOTREMAININGQTY" & vbCrLf
      sSql = sSql & "   order by LOTADATE DESC" & vbCrLf
      sSql = sSql & "   open lifoLots" & vbCrLf
      sSql = sSql & "   while(@adjustQty > 0)" & vbCrLf
      sSql = sSql & "   begin" & vbCrLf
      sSql = sSql & "      fetch next from lifoLots into @lotNumber, @lotRemaining, @availQty" & vbCrLf
      sSql = sSql & "      if (@@FETCH_STATUS <> 0)break" & vbCrLf
      sSql = sSql & "      declare @add decimal(15,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      if @availQty <= @adjustQty" & vbCrLf
      sSql = sSql & "         set @add = @availQty" & vbCrLf
      sSql = sSql & "      else" & vbCrLf
      sSql = sSql & "         set @add = @adjustQty" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      set @adjustQty = @adjustQty - @add" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      insert into CCLotAlloc ( CCREF, PARTREF, LOTNUMBER, OLDQTY, NEWQTY )" & vbCrLf
      sSql = sSql & "         values( @CCID, @partRef, @lotNumber, @lotRemaining, @lotRemaining + @add)" & vbCrLf
      sSql = sSql & "   end" & vbCrLf
      sSql = sSql & "   close lifoLots" & vbCrLf
      sSql = sSql & "   deallocate lifoLots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- if insufficient lots, say so in log" & vbCrLf
      sSql = sSql & "   if @adjustQty > 0" & vbCrLf
      sSql = sSql & "      insert CCLog ( CCREF, PARTREF, LOGTEXT )" & vbCrLf
      sSql = sSql & "         values ( @CCID, @partRef, 'no room in lots for adjustment of +' " & vbCrLf
      sSql = sSql & "         + cast(@adjustQty as varchar(12)))" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "close more" & vbCrLf
      sSql = sSql & "deallocate more" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "------------------------------------" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- find all non-lot-tracked parts with decreasing quantity" & vbCrLf
      sSql = sSql & "-- apply decreases to oldest lots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare less cursor for" & vbCrLf
      sSql = sSql & "select rtrim(CLPARTREF), CLLOTREMAININGQTY, CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER = '' and CLLOTREMAININGQTY > 0" & vbCrLf
      sSql = sSql & "order by CLPARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "open less" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "while (1 = 1)" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   fetch next from less into @partRef, @initialQty, @countQty" & vbCrLf
      sSql = sSql & "   if (@@FETCH_STATUS <> 0)break" & vbCrLf
      sSql = sSql & "   set @adjustQty = @initialQty - @countQty" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- find available lots to add to" & vbCrLf
      sSql = sSql & "   declare fifoLots cursor for" & vbCrLf
      sSql = sSql & "   select LOTNUMBER, LOTREMAININGQTY, LOTREMAININGQTY" & vbCrLf
      sSql = sSql & "   from LohdTable" & vbCrLf
      sSql = sSql & "   where LOTPARTREF = @partRef" & vbCrLf
      sSql = sSql & "   and LOTREMAININGQTY > 0" & vbCrLf
      sSql = sSql & "   order by LOTADATE ASC" & vbCrLf
      sSql = sSql & "   open fifoLots" & vbCrLf
      sSql = sSql & "   while(@adjustQty > 0)" & vbCrLf
      sSql = sSql & "   begin" & vbCrLf
      sSql = sSql & "      fetch next from fifoLots into @lotNumber, @lotRemaining, @availQty" & vbCrLf
      sSql = sSql & "      if (@@FETCH_STATUS <> 0)break" & vbCrLf
      sSql = sSql & "      declare @sub decimal(15,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      if @availQty <= @adjustQty" & vbCrLf
      sSql = sSql & "         set @sub = @availQty" & vbCrLf
      sSql = sSql & "      else" & vbCrLf
      sSql = sSql & "         set @sub = @adjustQty" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      set @adjustQty = @adjustQty - @sub" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      insert into CCLotAlloc ( CCREF, PARTREF, LOTNUMBER, OLDQTY, NEWQTY )" & vbCrLf
      sSql = sSql & "         values( @CCID, @partRef, @lotNumber, @lotRemaining, @lotRemaining - @sub)" & vbCrLf
      sSql = sSql & "   end" & vbCrLf
      sSql = sSql & "   close fifoLots" & vbCrLf
      sSql = sSql & "   deallocate fifoLots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- if insufficient lots, say so in log" & vbCrLf
      sSql = sSql & "   if @adjustQty > 0" & vbCrLf
      sSql = sSql & "      insert CCLog ( CCREF, PARTREF, LOGTEXT )" & vbCrLf
      sSql = sSql & "         values ( @CCID, @partRef, 'no room in lots for adjustment of -' " & vbCrLf
      sSql = sSql & "         + cast(@adjustQty as varchar(12)))" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "close less" & vbCrLf
      sSql = sSql & "deallocate less" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 43"
   End If

   If ver < 44 Then
      sSql = "create procedure UpdateCycleCount" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   @CCID varchar(20)," & vbCrLf
      sSql = sSql & "   @user varchar(10)" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- make sure all adjustments can be accomodated by lots" & vbCrLf
      sSql = sSql & "exec AllocateCycleCountLots @CCID" & vbCrLf
      sSql = sSql & "declare @problemCount int" & vbCrLf
      sSql = sSql & "select @problemCount = count(*) from CCLog where CCREF = @CCID" & vbCrLf
      sSql = sSql & "if @problemCount <> 0" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   print 'UpdateCycleCount ' + @CCID + ' cannot proceed.  See CCLog table.'" & vbCrLf
      sSql = sSql & "   return" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "begin transaction" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare @InvAdjAcct varchar(12)" & vbCrLf
      sSql = sSql & "declare @CountDate datetime" & vbCrLf
      sSql = sSql & "declare @PlanDate datetime" & vbCrLf
      sSql = sSql & "declare @NextCountDate datetime" & vbCrLf
      sSql = sSql & "declare @ABCCode varchar(2)" & vbCrLf
      sSql = sSql & "declare @ABCFrequency int" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "SET NOCOUNT ON" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "select @InvAdjAcct = isnull(COADJACCT,'?') from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "select @CountDate = CCCOUNTLOCKEDDATE," & vbCrLf
      sSql = sSql & "@PlanDate = CCPLANDATE," & vbCrLf
      sSql = sSql & "@ABCCode = CCABCCODE " & vbCrLf
      sSql = sSql & "from CchdTable where CCREF = @CCID" & vbCrLf
      sSql = sSql & "select @ABCFrequency = isnull(COABCFREQUENCY,90) from CabcTable where COABCCODE = @ABCCode" & vbCrLf
      sSql = sSql & "select @NextCountDate = dateadd(d, @ABCFrequency, @PlanDate)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- create inventory activities" & vbCrLf
      sSql = sSql & "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INLOTNUMBER,INPDATE,INADATE," & vbCrLf
      sSql = sSql & "INPQTY,INAQTY,INAMT,INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS," & vbCrLf
      sSql = sSql & "INDEBITACCT,INCREDITACCT,INUSER)" & vbCrLf
      sSql = sSql & "select 30, cl.PARTREF, 'ABC Cycle Count', @CCID, cl.LOTNUMBER, @CountDate, @CountDate," & vbCrLf
      sSql = sSql & "cl.NEWQTY - cl.OLDQTY, cl.NEWQTY - cl.OLDQTY, lh.LOTUNITCOST," & vbCrLf
      sSql = sSql & "abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTMATL / lh.LOTORIGINALQTY," & vbCrLf
      sSql = sSql & "abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTLABOR / lh.LOTORIGINALQTY," & vbCrLf
      sSql = sSql & "abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTEXP / lh.LOTORIGINALQTY," & vbCrLf
      sSql = sSql & "abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTOH / lh.LOTORIGINALQTY," & vbCrLf
      sSql = sSql & "abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTHRS / lh.LOTORIGINALQTY," & vbCrLf
      sSql = sSql & "case when cl.NEWQTY > cl.OLDQTY then dbo.fnGetPartInvAccount(cl.PARTREF) else @InvAdjAcct end," & vbCrLf
      sSql = sSql & "case when cl.NEWQTY > cl.OLDQTY then @InvAdjAcct else dbo.fnGetPartInvAccount(cl.PARTREF) end," & vbCrLf
      sSql = sSql & "@user" & vbCrLf
      sSql = sSql & "from CCLotAlloc cl" & vbCrLf
      sSql = sSql & "join LohdTable lh on cl.LOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "where cl.CCREF = @CCID" & vbCrLf
      sSql = sSql & "order by cl.PARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- create corresponding lot item records" & vbCrLf
      sSql = sSql & "insert into LoitTable" & vbCrLf
      sSql = sSql & "(LOINUMBER,LOIRECORD,LOITYPE,LOIPARTREF," & vbCrLf
      sSql = sSql & "LOIPDATE,LOIADATE,LOIQUANTITY,LOIACTIVITY,LOICOMMENT)" & vbCrLf
      sSql = sSql & "select lh.LOTNUMBER,(select isnull(max(LOIRECORD),0) + 1 from LoitTable" & vbCrLf
      sSql = sSql & "where LOINUMBER = lh.LOTNUMBER),30,cl.PARTREF," & vbCrLf
      sSql = sSql & "@CountDate, @CountDate, cl.NEWQTY - cl.OLDQTY, ia.INNUMBER, @CCID" & vbCrLf
      sSql = sSql & "from CCLotAlloc cl" & vbCrLf
      sSql = sSql & "join LohdTable lh on cl.LOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "join InvaTable ia on cl.PARTREF = ia.INPART and ia.INLOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "and ia.INTYPE = 30 and ia.INREF2 = @CCID" & vbCrLf
      sSql = sSql & "where cl.CCREF = @CCID" & vbCrLf
      sSql = sSql & "order by cl.PARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- update lot header remaining quantity" & vbCrLf
      sSql = sSql & "update LohdTable" & vbCrLf
      sSql = sSql & "set LOTAVAILABLE = case when (lh.LOTREMAININGQTY + cl.NEWQTY - cl.OLDQTY) = 0" & vbCrLf
      sSql = sSql & "then 0 else 1 end," & vbCrLf
      sSql = sSql & "LOTREMAININGQTY = lh.LOTREMAININGQTY +  cl.NEWQTY - cl.OLDQTY" & vbCrLf
      sSql = sSql & "from LohdTable lh" & vbCrLf
      sSql = sSql & "join CCLotAlloc cl on cl.LOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "where cl.CCREF = @CCID" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- update part table" & vbCrLf
      sSql = sSql & "update PartTable" & vbCrLf
      sSql = sSql & "set PAQOH = x.LotSum, PALOTQTYREMAINING = x.LotSum," & vbCrLf
      sSql = sSql & "PANEXTCYCLEDATE = @NextCountDate" & vbCrLf
      sSql = sSql & "from PartTable pt" & vbCrLf
      sSql = sSql & "join (select LOTPARTREF, sum(LOTREMAININGQTY) as LotSum from LohdTable lh" & vbCrLf
      sSql = sSql & "where LOTPARTREF in (select PARTREF from CCLotAlloc where CCREF = @CCID)" & vbCrLf
      sSql = sSql & "group by LOTPARTREF) x on pt.PARTREF = x.LOTPARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- set status of cycle count = Updated" & vbCrLf
      sSql = sSql & "update CchdTable" & vbCrLf
      sSql = sSql & "set CCUPDATED = 1," & vbCrLf
      sSql = sSql & "CCUPDATEDDATE = getdate()" & vbCrLf
      sSql = sSql & "where CCREF = @CCID" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "commit transaction" & vbCrLf
      Execute False, sSql
      
      'set version
      Execute False, "update Version set Version = 44"
   End If
   
   If ver < 45 Then
      ver = 45
      Execute False, "drop view viewMaintPartsWithoutLots"
      sSql = "create view viewMaintPartsWithoutLots" & vbCrLf _
         & "as" & vbCrLf _
         & "select 'PARTREF' as QtyCol, 'count(LOTPARTREF)' as SumCol, PARTREF as PartRef," & vbCrLf _
         & "'' as LotNumber, 1 as Qty, 0 as SumQty" & vbCrLf _
         & "from PartTable" & vbCrLf _
         & "where PARTREF not in (select LOTPARTREF from LohdTable)" & vbCrLf
      Execute True, sSql
      
      Execute False, "drop table CCLog"
      sSql = "CREATE TABLE CCLog" & vbCrLf _
         & "(" & vbCrLf _
         & "   CCREF varchar(20) NOT NULL default ''," & vbCrLf _
         & "   PARTREF varchar(30) NOT NULL default ''," & vbCrLf _
         & "   LOTNUMBER varchar(15) NULL," & vbCrLf _
         & "   LOGTEXT varchar(80) NOT NULL default ''" & vbCrLf _
         & ")"
      Execute True, sSql
      
      Execute False, "drop view viewTimeCardHours"
      sSql = "create view viewTimeCardHours" & vbCrLf _
         & "as" & vbCrLf _
         & "select TCEMP as empno, TMDAY as chgday, TYPETYPE as type," & vbCrLf _
         & "sum(cast(TCHOURS as decimal(6,3))) as hrs" & vbCrLf _
         & "from TcitTable" & vbCrLf _
         & "join TchdTable on TCCARD = TMCARD" & vbCrLf _
         & "join TmcdTable on TCCODE = TYPECODE" & vbCrLf _
         & "group by TCEMP, TMDAY, TYPETYPE"
      Execute True, sSql
   
      Execute False, "drop procedure UpdateTimeCardTotals"
      sSql = "create procedure UpdateTimeCardTotals" & vbCrLf
      sSql = sSql & " @EmpNo int," & vbCrLf
      sSql = sSql & " @Date datetime" & vbCrLf
      sSql = sSql & "as update TchdTable " & vbCrLf
      sSql = sSql & " set TMSTART = (select top 1 TCSTART From TcitTable " & vbCrLf
      sSql = sSql & " join TchdTable on TCCARD = TMCARD" & vbCrLf
      sSql = sSql & " where TCEMP = @EmpNo and TMDAY = @Date" & vbCrLf
      sSql = sSql & " order by substring(TCSTART,6,1), TCSTART)," & vbCrLf
      sSql = sSql & " TMSTOP = (select top 1 TCSTOP From TcitTable" & vbCrLf
      sSql = sSql & " join TchdTable on TCCARD = TMCARD" & vbCrLf
      sSql = sSql & " where TCEMP = @EmpNo and TMDAY = @Date" & vbCrLf
      sSql = sSql & " order by substring(TCSTOP,6,1) desc, TCSTOP desc)" & vbCrLf
      sSql = sSql & "where TMEMP = @EmpNo and TMDAY = @Date" & vbCrLf
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
      sSql = sSql & "where TMEMP = @EmpNo and TMDAY = @Date" & vbCrLf
      Execute True, sSql
      
      Execute False, "drop procedure UpdateAllTimeCardTotals"
      sSql = "create procedure UpdateAllTimeCardTotals" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "declare @emp int" & vbCrLf
      sSql = sSql & "declare @date datetime" & vbCrLf
      sSql = sSql & "declare @msg varchar(100)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare timecards cursor" & vbCrLf
      sSql = sSql & "for" & vbCrLf
      sSql = sSql & "select distinct TMEMP, TMDAY" & vbCrLf
      sSql = sSql & "from TchdTable" & vbCrLf
      sSql = sSql & "order by TMEMP, TMDAY" & vbCrLf
      sSql = sSql & "open timecards" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "fetch next from timecards into @emp, @date" & vbCrLf
      sSql = sSql & "while @@fetch_status = 0" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   exec UpdateTimeCardTotals @emp, @date" & vbCrLf
      sSql = sSql & "   fetch next from timecards into @emp, @date" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "close timecards" & vbCrLf
      sSql = sSql & "deallocate timecards" & vbCrLf
      Execute True, sSql
      'Execute True, "UpdateAllTimeCardTotals" -- done in a later revision
     
      Execute False, "drop procedure AllocateCycleCountLots"
      sSql = "create procedure AllocateCycleCountLots" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   @CCID varchar(20)    -- cycle count ID" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "/* test" & vbCrLf
      sSql = sSql & "   AllocateCycleCountLots '20080808-A'" & vbCrLf
      sSql = sSql & "   Select * from CCLog where CCREF = '20080808-A'" & vbCrLf
      sSql = sSql & "*/" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "delete from CCLog where CCREF = @CCID" & vbCrLf
      sSql = sSql & "delete from CCLoTaLLOC where CCREF = @CCID" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- make sure all counts have been entered" & vbCrLf
      sSql = sSql & "declare @CountsRequired int, @TotalItems int, @CountsEntered int, @NoLots int" & vbCrLf
      sSql = sSql & "select @TotalItems = count(*) from CcltTable where CLREF = @CCID" & vbCrLf
      sSql = sSql & "select @CountsEntered = count(*) from CcltTable where CLREF = @CCID" & vbCrLf
      sSql = sSql & "and CLENTERED = 1" & vbCrLf
      sSql = sSql & "select @NoLots = count(*) from CcltTable" & vbCrLf
      sSql = sSql & "join CcitTable on CIREF = CLREF and CIPARTREF = CLPARTREF" & vbCrLf
      sSql = sSql & "and CILOTTRACK = 1 and rtrim(CLLOTNUMBER) = ''" & vbCrLf
      sSql = sSql & "where CLREF = @CCID" & vbCrLf
      sSql = sSql & "set @CountsRequired = @TotalItems - @CountsEntered - @NoLots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "if @CountsRequired <> 0" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   insert into CCLog ( CCREF, PARTREF, LOTNUMBER, LOGTEXT )" & vbCrLf
      sSql = sSql & "   values( @CCID, '', '', cast(@CountsRequired as varchar(10)) + ' counts need to be entered.')" & vbCrLf
      sSql = sSql & "   return" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- insert info for lot-tracked parts" & vbCrLf
      sSql = sSql & "insert into CCLOTALLOC" & vbCrLf
      sSql = sSql & "select @CCID, rtrim(CLPARTREF), rtrim(CLLOTNUMBER), CLLOTREMAININGQTY, CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER <> '' and CLLOTREMAININGQTY <> CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "order by CLPARTREF, CLLOTNUMBER" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "---------------------------------------" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- find all non-lot-tracked parts with increasing quantity" & vbCrLf
      sSql = sSql & "-- apply increases to newest lots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare @partRef varchar(30)" & vbCrLf
      sSql = sSql & "declare @initialQty decimal(15,4)" & vbCrLf
      sSql = sSql & "declare @countQty decimal(15,4)" & vbCrLf
      sSql = sSql & "declare @adjustQty decimal(15,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare more cursor for" & vbCrLf
      sSql = sSql & "select rtrim(CLPARTREF), CLLOTREMAININGQTY, CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER = '' and CLLOTREMAININGQTY < CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "order by CLPARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare @lotNumber varchar(15)" & vbCrLf
      sSql = sSql & "declare @lotRemaining decimal(15,4)" & vbCrLf
      sSql = sSql & "declare @availQty decimal(15,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "open more" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "while (1 = 1)" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   fetch next from more into @partRef, @initialQty, @countQty" & vbCrLf
      sSql = sSql & "   if (@@FETCH_STATUS <> 0)break" & vbCrLf
      sSql = sSql & "   set @adjustQty = @countQty - @initialQty" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- find available lots to add to" & vbCrLf
      sSql = sSql & "   declare lifoLots cursor for" & vbCrLf
      sSql = sSql & "   select LOTNUMBER, LOTREMAININGQTY, LOTORIGINALQTY - LOTREMAININGQTY as Available" & vbCrLf
      sSql = sSql & "   from LohdTable" & vbCrLf
      sSql = sSql & "   where LOTPARTREF = @partRef" & vbCrLf
      sSql = sSql & "   and LOTORIGINALQTY > LOTREMAININGQTY" & vbCrLf
      sSql = sSql & "   order by LOTADATE DESC" & vbCrLf
      sSql = sSql & "   open lifoLots" & vbCrLf
      sSql = sSql & "   while(@adjustQty > 0)" & vbCrLf
      sSql = sSql & "   begin" & vbCrLf
      sSql = sSql & "      fetch next from lifoLots into @lotNumber, @lotRemaining, @availQty" & vbCrLf
      sSql = sSql & "      if (@@FETCH_STATUS <> 0)break" & vbCrLf
      sSql = sSql & "      declare @add decimal(15,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      if @availQty <= @adjustQty" & vbCrLf
      sSql = sSql & "         set @add = @availQty" & vbCrLf
      sSql = sSql & "      else" & vbCrLf
      sSql = sSql & "         set @add = @adjustQty" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      set @adjustQty = @adjustQty - @add" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      insert into CCLotAlloc ( CCREF, PARTREF, LOTNUMBER, OLDQTY, NEWQTY )" & vbCrLf
      sSql = sSql & "         values( @CCID, @partRef, @lotNumber, @lotRemaining, @lotRemaining + @add)" & vbCrLf
      sSql = sSql & "   end" & vbCrLf
      sSql = sSql & "   close lifoLots" & vbCrLf
      sSql = sSql & "   deallocate lifoLots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- if insufficient lots, say so in log" & vbCrLf
      sSql = sSql & "   if @adjustQty > 0" & vbCrLf
      sSql = sSql & "      insert CCLog ( CCREF, PARTREF, LOTNUMBER, LOGTEXT )" & vbCrLf
      sSql = sSql & "         values ( @CCID, @partRef, '', 'no room in hidden lots for adjustment of +' " & vbCrLf
      sSql = sSql & "         + cast(@adjustQty as varchar(15)) + ' of ' + cast(@countQty - @initialQty as varchar(15)) )" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "close more" & vbCrLf
      sSql = sSql & "deallocate more" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "------------------------------------" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- find all non-lot-tracked parts with decreasing quantity" & vbCrLf
      sSql = sSql & "-- apply decreases to oldest lots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare less cursor for" & vbCrLf
      sSql = sSql & "select rtrim(CLPARTREF), CLLOTREMAININGQTY, CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER = '' and CLLOTREMAININGQTY > 0" & vbCrLf
      sSql = sSql & "order by CLPARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "open less" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "while (1 = 1)" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   fetch next from less into @partRef, @initialQty, @countQty" & vbCrLf
      sSql = sSql & "   if (@@FETCH_STATUS <> 0)break" & vbCrLf
      sSql = sSql & "   set @adjustQty = @initialQty - @countQty" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- find available lots to add to" & vbCrLf
      sSql = sSql & "   declare fifoLots cursor for" & vbCrLf
      sSql = sSql & "   select LOTNUMBER, LOTREMAININGQTY, LOTREMAININGQTY" & vbCrLf
      sSql = sSql & "   from LohdTable" & vbCrLf
      sSql = sSql & "   where LOTPARTREF = @partRef" & vbCrLf
      sSql = sSql & "   and LOTREMAININGQTY > 0" & vbCrLf
      sSql = sSql & "   order by LOTADATE ASC" & vbCrLf
      sSql = sSql & "   open fifoLots" & vbCrLf
      sSql = sSql & "   while(@adjustQty > 0)" & vbCrLf
      sSql = sSql & "   begin" & vbCrLf
      sSql = sSql & "      fetch next from fifoLots into @lotNumber, @lotRemaining, @availQty" & vbCrLf
      sSql = sSql & "      if (@@FETCH_STATUS <> 0)break" & vbCrLf
      sSql = sSql & "      declare @sub decimal(15,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      if @availQty <= @adjustQty" & vbCrLf
      sSql = sSql & "         set @sub = @availQty" & vbCrLf
      sSql = sSql & "      else" & vbCrLf
      sSql = sSql & "         set @sub = @adjustQty" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      set @adjustQty = @adjustQty - @sub" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      insert into CCLotAlloc ( CCREF, PARTREF, LOTNUMBER, OLDQTY, NEWQTY )" & vbCrLf
      sSql = sSql & "         values( @CCID, @partRef, @lotNumber, @lotRemaining, @lotRemaining - @sub)" & vbCrLf
      sSql = sSql & "   end" & vbCrLf
      sSql = sSql & "   close fifoLots" & vbCrLf
      sSql = sSql & "   deallocate fifoLots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- if insufficient lots, say so in log" & vbCrLf
      sSql = sSql & "   if @adjustQty > 0" & vbCrLf
      sSql = sSql & "      insert CCLog ( CCREF, PARTREF, LOTNUMBER, LOGTEXT )" & vbCrLf
      sSql = sSql & "         values ( @CCID, @partRef, '', 'no room in hidden lots for adjustment of -' " & vbCrLf
      sSql = sSql & "         + cast(@adjustQty as varchar(15)) + ' of ' + cast(@countQty - @initialQty as varchar(15)) )" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "close less" & vbCrLf
      sSql = sSql & "deallocate less" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- now find problems with positive adjustments to lot-tracked parts" & vbCrLf
      sSql = sSql & "insert into CCLog ( CCREF, PARTREF, LOTNUMBER, LOGTEXT )" & vbCrLf
      sSql = sSql & "select @CCID, rtrim(CLPARTREF), rtrim(LOTNUMBER), 'insufficient room in lot for adjustment of +' " & vbCrLf
      sSql = sSql & " + cast(CLLOTADJUSTQTY - CLLOTREMAININGQTY as varchar(15))" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "join LohdTable on CLLOTNUMBER = LOTNUMBER" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER <> '' and CLLOTREMAININGQTY < CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "and LOTORIGINALQTY - LOTREMAININGQTY < CLLOTADJUSTQTY - CLLOTREMAININGQTY" & vbCrLf
      sSql = sSql & "order by CLPARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- now find problems with negative adjustments to lot-tracked parts" & vbCrLf
      sSql = sSql & "insert into CCLog ( CCREF, PARTREF, LOTNUMBER, LOGTEXT )" & vbCrLf
      sSql = sSql & "select @CCID, rtrim(CLPARTREF), rtrim(LOTNUMBER), 'insufficient qty in lot for adjustment of ' " & vbCrLf
      sSql = sSql & " + cast(CLLOTADJUSTQTY - CLLOTREMAININGQTY as varchar(15))" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "join LohdTable on CLLOTNUMBER = LOTNUMBER" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER <> ''" & vbCrLf
      sSql = sSql & "and LOTREMAININGQTY < CLLOTREMAININGQTY - CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "order by CLPARTREF" & vbCrLf
      Execute True, sSql
      
      'set version
      Execute False, "update Version set Version = 45"
   End If
      
   newver = 46
   If ver < newver Then
      ver = newver
      Execute False, "DROP TABLE EsReportWIP"
      sSql = "CREATE TABLE EsReportWIP (" & vbCrLf _
         & "WIPRUNREF CHAR(30) NULL DEFAULT('')," & vbCrLf _
         & "WIPRUNNO INT NULL DEFAULT(0)," & vbCrLf _
         & "WIPRUNSTATUS CHAR(2) NULL DEFAULT('')," & vbCrLf _
         & "WIPCOSTTYPE CHAR(3) NULL DEFAULT('')," & vbCrLf _
         & "WIPLABOR REAL NULL DEFAULT(0)," & vbCrLf _
         & "WIPMISSTIME TINYINT NULL DEFAULT(0)," & vbCrLf _
         & "WIPMATL REAL NULL DEFAULT(0)," & vbCrLf _
         & "WIPMISSMATL TINYINT NULL DEFAULT(0)," & vbCrLf _
         & "WIPOH REAL NULL DEFAULT(0)," & vbCrLf _
         & "WIPEXP REAL NULL DEFAULT(0)," & vbCrLf _
         & "WIPMISSEXP TINYINT NULL DEFAULT(0)," & vbCrLf _
         & "WIPFREIGHT REAL NULL DEFAULT(0)," & vbCrLf _
         & "WIPTAX REAL NULL DEFAULT(0)," & vbCrLf _
         & "WIPUNCOSTED TINYINT DEFAULT(0))"
      Execute True, sSql

      sSql = "CREATE UNIQUE CLUSTERED INDEX WipReport ON EsReportWIP " & vbCrLf _
         & "(WIPRUNREF,WIPRUNNO) WITH  FILLFACTOR = 80"
      Execute True, sSql

      'set version
      Execute False, "update Version set Version = " & newver
   End If

   newver = 47
   If ver < newver Then
      ver = newver
       Execute False, "DROP TABLE CCLog"
       sSql = "CREATE TABLE CCLog" & vbCrLf _
         & "(" & vbCrLf _
         & "   CCREF varchar(20) NOT NULL default ''," & vbCrLf _
         & "   PARTREF varchar(30) NOT NULL default ''," & vbCrLf _
         & "   LOTNUMBER varchar(15) NULL," & vbCrLf _
         & "   ERRORTYPE varchar(10) NOT NULL default 'FATAL'," & vbCrLf _
         & "   LOGTEXT varchar(80) NOT NULL default ''" & vbCrLf _
         & ")"
      Execute True, sSql

      Execute False, "DROP PROCEDURE AllocateCycleCountLots"
      sSql = "create procedure AllocateCycleCountLots" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   @CCID varchar(20)    -- cycle count ID" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "/* test" & vbCrLf
      sSql = sSql & "   AllocateCycleCountLots '20080820-S'" & vbCrLf
      sSql = sSql & "   Select * from CCLog where CCREF = '20080820-S'" & vbCrLf
      sSql = sSql & "*/" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "delete from CCLog where CCREF = @CCID" & vbCrLf
      sSql = sSql & "delete from CCLotAlloc where CCREF = @CCID" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- make sure all counts have been entered" & vbCrLf
      sSql = sSql & "declare @CountsRequired int, @TotalItems int, @CountsEntered int, @NoLots int" & vbCrLf
      sSql = sSql & "select @TotalItems = count(*) from CcltTable where CLREF = @CCID" & vbCrLf
      sSql = sSql & "select @CountsEntered = count(*) from CcltTable where CLREF = @CCID" & vbCrLf
      sSql = sSql & "and CLENTERED = 1" & vbCrLf
      sSql = sSql & "select @NoLots = count(*) from CcltTable" & vbCrLf
      sSql = sSql & "join CcitTable on CIREF = CLREF and CIPARTREF = CLPARTREF" & vbCrLf
      sSql = sSql & "and CILOTTRACK = 1 and rtrim(CLLOTNUMBER) = ''" & vbCrLf
      sSql = sSql & "where CLREF = @CCID" & vbCrLf
      sSql = sSql & "set @CountsRequired = @TotalItems - @CountsEntered - @NoLots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "if @CountsRequired <> 0" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   insert into CCLog ( CCREF, PARTREF, LOTNUMBER, ERRORTYPE, LOGTEXT )" & vbCrLf
      sSql = sSql & "   values( @CCID, '', '', 'FATAL', cast(@CountsRequired as varchar(10)) + ' counts need to be entered.')" & vbCrLf
      sSql = sSql & "   return" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- insert info for lot-tracked parts" & vbCrLf
      sSql = sSql & "insert into CCLOTALLOC" & vbCrLf
      sSql = sSql & "select @CCID, rtrim(CLPARTREF), rtrim(CLLOTNUMBER), CLLOTREMAININGQTY, CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER <> '' and CLLOTREMAININGQTY <> CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "order by CLPARTREF, CLLOTNUMBER" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- determine whether fifo or lifo" & vbCrLf
      sSql = sSql & "declare @fifo bit" & vbCrLf
      sSql = sSql & "select @fifo = COLOTSFIFO from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "---------------------------------------" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- find all non-lot-tracked parts with increasing quantity" & vbCrLf
      sSql = sSql & "-- apply increases to newest lot if fifo = 0 and oldest lot if fifo = 1" & vbCrLf
      sSql = sSql & "-- select a lot with LOTREMAININGQTY > 0 if one is available" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare @partRef varchar(30)" & vbCrLf
      sSql = sSql & "declare @initialQty decimal(15,4)" & vbCrLf
      sSql = sSql & "declare @countQty decimal(15,4)" & vbCrLf
      sSql = sSql & "declare @adjustQty decimal(15,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare more cursor for" & vbCrLf
      sSql = sSql & "select rtrim(CLPARTREF), CLLOTREMAININGQTY, CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER = '' and CLLOTREMAININGQTY < CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "order by CLPARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare @lotNumber varchar(15)" & vbCrLf
      sSql = sSql & "declare @lotRemaining decimal(15,4)" & vbCrLf
      sSql = sSql & "declare @availQty decimal(15,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "open more" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "while (1 = 1)" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   fetch next from more into @partRef, @initialQty, @countQty" & vbCrLf
      sSql = sSql & "   if (@@FETCH_STATUS <> 0)break" & vbCrLf
      sSql = sSql & "   set @adjustQty = @countQty - @initialQty" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- if fifo find oldest lot with nonzero quantity remaining" & vbCrLf
      sSql = sSql & "   -- if none, find oldeset lot  with zero quantity remaining" & vbCrLf
      sSql = sSql & "   if @fifo = 1" & vbCrLf
      sSql = sSql & "   begin" & vbCrLf
      sSql = sSql & "      select top 1 @lotNumber = LOTNUMBER," & vbCrLf
      sSql = sSql & "         @lotRemaining = LOTREMAININGQTY," & vbCrLf
      sSql = sSql & "         @availQty = LOTORIGINALQTY - LOTREMAININGQTY" & vbCrLf
      sSql = sSql & "      from LohdTable" & vbCrLf
      sSql = sSql & "      where LOTPARTREF = @partRef and LOTREMAININGQTY > 0" & vbCrLf
      sSql = sSql & "      order by LOTADATE desc" & vbCrLf
      sSql = sSql & "      if @lotNumber is null" & vbCrLf
      sSql = sSql & "      begin" & vbCrLf
      sSql = sSql & "         select top 1 @lotNumber = LOTNUMBER," & vbCrLf
      sSql = sSql & "            @lotRemaining = LOTREMAININGQTY," & vbCrLf
      sSql = sSql & "            @availQty = LOTORIGINALQTY - LOTREMAININGQTY" & vbCrLf
      sSql = sSql & "         from LohdTable" & vbCrLf
      sSql = sSql & "         where LOTPARTREF = @partRef" & vbCrLf
      sSql = sSql & "         order by LOTADATE desc" & vbCrLf
      sSql = sSql & "      end" & vbCrLf
      sSql = sSql & "   else" & vbCrLf
      sSql = sSql & "      select top 1 @lotNumber = LOTNUMBER," & vbCrLf
      sSql = sSql & "         @lotRemaining = LOTREMAININGQTY," & vbCrLf
      sSql = sSql & "         @availQty = LOTORIGINALQTY - LOTREMAININGQTY" & vbCrLf
      sSql = sSql & "      from LohdTable" & vbCrLf
      sSql = sSql & "      where LOTPARTREF = @partRef and LOTREMAININGQTY > 0" & vbCrLf
      sSql = sSql & "      order by LOTADATE asc" & vbCrLf
      sSql = sSql & "      if @lotNumber is null" & vbCrLf
      sSql = sSql & "      begin" & vbCrLf
      sSql = sSql & "         select top 1 @lotNumber = LOTNUMBER," & vbCrLf
      sSql = sSql & "            @lotRemaining = LOTREMAININGQTY," & vbCrLf
      sSql = sSql & "            @availQty = LOTORIGINALQTY - LOTREMAININGQTY" & vbCrLf
      sSql = sSql & "         from LohdTable" & vbCrLf
      sSql = sSql & "         where LOTPARTREF = @partRef and LOTREMAININGQTY > 0" & vbCrLf
      sSql = sSql & "         order by LOTADATE asc" & vbCrLf
      sSql = sSql & "      end" & vbCrLf
      sSql = sSql & "   end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   if @lotNumber is null" & vbCrLf
      sSql = sSql & "   begin" & vbCrLf
      sSql = sSql & "      insert into CCLog ( CCREF, PARTREF, LOTNUMBER, ERRORTYPE, LOGTEXT )" & vbCrLf
      sSql = sSql & "         values ( @CCID, @partRef, '', 'FATAL', 'No lots for this non-lot-tracked part' )" & vbCrLf
      sSql = sSql & "   end" & vbCrLf
      sSql = sSql & "   else" & vbCrLf
      sSql = sSql & "   begin" & vbCrLf
      sSql = sSql & "      insert into CCLotAlloc ( CCREF, PARTREF, LOTNUMBER, OLDQTY, NEWQTY )" & vbCrLf
      sSql = sSql & "         values( @CCID, @partRef, @lotNumber, @initialQty, @countQty)" & vbCrLf
      sSql = sSql & "   end" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "close more" & vbCrLf
      sSql = sSql & "deallocate more" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "------------------------------------" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- find all non-lot-tracked parts with decreasing quantity" & vbCrLf
      sSql = sSql & "-- apply decreases to oldest lots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare less cursor for" & vbCrLf
      sSql = sSql & "select rtrim(CLPARTREF), CLLOTREMAININGQTY, CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER = '' and CLLOTREMAININGQTY > 0" & vbCrLf
      sSql = sSql & "order by CLPARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "open less" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "while (1 = 1)" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   fetch next from less into @partRef, @initialQty, @countQty" & vbCrLf
      sSql = sSql & "   if (@@FETCH_STATUS <> 0)break" & vbCrLf
      sSql = sSql & "   set @adjustQty = @initialQty - @countQty" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- find available lots to add to" & vbCrLf
      sSql = sSql & "   declare fifoLots cursor for" & vbCrLf
      sSql = sSql & "   select LOTNUMBER, LOTREMAININGQTY, LOTREMAININGQTY" & vbCrLf
      sSql = sSql & "   from LohdTable" & vbCrLf
      sSql = sSql & "   where LOTPARTREF = @partRef" & vbCrLf
      sSql = sSql & "   and LOTREMAININGQTY > 0" & vbCrLf
      sSql = sSql & "   order by LOTADATE ASC" & vbCrLf
      sSql = sSql & "   open fifoLots" & vbCrLf
      sSql = sSql & "   while(@adjustQty > 0)" & vbCrLf
      sSql = sSql & "   begin" & vbCrLf
      sSql = sSql & "      fetch next from fifoLots into @lotNumber, @lotRemaining, @availQty" & vbCrLf
      sSql = sSql & "      if (@@FETCH_STATUS <> 0)break" & vbCrLf
      sSql = sSql & "      declare @sub decimal(15,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      if @availQty <= @adjustQty" & vbCrLf
      sSql = sSql & "         set @sub = @availQty" & vbCrLf
      sSql = sSql & "      else" & vbCrLf
      sSql = sSql & "         set @sub = @adjustQty" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      set @adjustQty = @adjustQty - @sub" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "      insert into CCLotAlloc ( CCREF, PARTREF, LOTNUMBER, OLDQTY, NEWQTY )" & vbCrLf
      sSql = sSql & "         values( @CCID, @partRef, @lotNumber, @lotRemaining, @lotRemaining - @sub)" & vbCrLf
      sSql = sSql & "   end" & vbCrLf
      sSql = sSql & "   close fifoLots" & vbCrLf
      sSql = sSql & "   deallocate fifoLots" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- if insufficient lots, say so in log" & vbCrLf
      sSql = sSql & "   if @adjustQty > 0" & vbCrLf
      sSql = sSql & "      insert into CCLog ( CCREF, PARTREF, LOTNUMBER, ERRORTYPE, LOGTEXT )" & vbCrLf
      sSql = sSql & "         values ( @CCID, @partRef, '', 'FATAL', 'no room in lots for adjustment of -' " & vbCrLf
      sSql = sSql & "         + cast(@adjustQty as varchar(15)) + ' of ' + cast(@countQty - @initialQty as varchar(15)) )" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "close less" & vbCrLf
      sSql = sSql & "deallocate less" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- now generate warnings for positive adjustments to lot-tracked parts" & vbCrLf
      sSql = sSql & "-- increasing quantity above LOTORIGINALQTY" & vbCrLf
      sSql = sSql & "insert into CCLog ( CCREF, PARTREF, LOTNUMBER, ERRORTYPE, LOGTEXT )" & vbCrLf
      sSql = sSql & "select @CCID, rtrim(CLPARTREF), rtrim(LOTNUMBER), 'WARNING', 'lot adjustment of +' " & vbCrLf
      sSql = sSql & " + cast(CLLOTADJUSTQTY - CLLOTREMAININGQTY as varchar(15)) + ' makes lot qty > inital qty'" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "join LohdTable on CLLOTNUMBER = LOTNUMBER" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER <> '' and CLLOTREMAININGQTY < CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "and LOTORIGINALQTY - LOTREMAININGQTY < CLLOTADJUSTQTY - CLLOTREMAININGQTY" & vbCrLf
      sSql = sSql & "order by CLPARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- now find problems with negative adjustments to lot-tracked parts" & vbCrLf
      sSql = sSql & "insert into CCLog ( CCREF, PARTREF, LOTNUMBER, ERRORTYPE, LOGTEXT )" & vbCrLf
      sSql = sSql & "select @CCID, rtrim(CLPARTREF), rtrim(LOTNUMBER), 'FATAL', 'lot adjustment of ' " & vbCrLf
      sSql = sSql & " + cast(CLLOTADJUSTQTY - CLLOTREMAININGQTY as varchar(15)) + 'makes lot qty < 0'" & vbCrLf
      sSql = sSql & "from CcltTable" & vbCrLf
      sSql = sSql & "join LohdTable on CLLOTNUMBER = LOTNUMBER" & vbCrLf
      sSql = sSql & "where CLREF = @CCID and CLENTERED = 1" & vbCrLf
      sSql = sSql & "and CLLOTNUMBER <> ''" & vbCrLf
      sSql = sSql & "and LOTREMAININGQTY < CLLOTREMAININGQTY - CLLOTADJUSTQTY" & vbCrLf
      sSql = sSql & "order by CLPARTREF" & vbCrLf
      Execute True, sSql
      
      Execute False, "drop procedure UpdateCycleCount" & vbCrLf
      sSql = "create procedure UpdateCycleCount" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   @CCID varchar(20)," & vbCrLf
      sSql = sSql & "   @user varchar(10)" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- make sure all adjustments can be accomodated by lots" & vbCrLf
      sSql = sSql & "exec AllocateCycleCountLots @CCID" & vbCrLf
      sSql = sSql & "declare @problemCount int" & vbCrLf
      sSql = sSql & "select @problemCount = count(*) from CCLog where CCREF = @CCID and ERRORTYPE = 'FATAL'" & vbCrLf
      sSql = sSql & "if @problemCount <> 0" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   print 'UpdateCycleCount ' + @CCID + ' cannot proceed.  See CCLog table.'" & vbCrLf
      sSql = sSql & "   return" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "begin transaction" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare @InvAdjAcct varchar(12)" & vbCrLf
      sSql = sSql & "declare @CountDate datetime" & vbCrLf
      sSql = sSql & "declare @PlanDate datetime" & vbCrLf
      sSql = sSql & "declare @NextCountDate datetime" & vbCrLf
      sSql = sSql & "declare @ABCCode varchar(2)" & vbCrLf
      sSql = sSql & "declare @ABCFrequency int" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "SET NOCOUNT ON" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "select @InvAdjAcct = isnull(COADJACCT,'?') from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "select @CountDate = CCCOUNTLOCKEDDATE," & vbCrLf
      sSql = sSql & "@PlanDate = CCPLANDATE," & vbCrLf
      sSql = sSql & "@ABCCode = CCABCCODE " & vbCrLf
      sSql = sSql & "from CchdTable where CCREF = @CCID" & vbCrLf
      sSql = sSql & "select @ABCFrequency = isnull(COABCFREQUENCY,90) from CabcTable where COABCCODE = @ABCCode" & vbCrLf
      sSql = sSql & "select @NextCountDate = dateadd(d, @ABCFrequency, @PlanDate)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- create inventory activities" & vbCrLf
      sSql = sSql & "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INLOTNUMBER,INPDATE,INADATE," & vbCrLf
      sSql = sSql & "INPQTY,INAQTY,INAMT,INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS," & vbCrLf
      sSql = sSql & "INDEBITACCT,INCREDITACCT,INUSER)" & vbCrLf
      sSql = sSql & "select 30, cl.PARTREF, 'ABC Cycle Count', @CCID, cl.LOTNUMBER, @CountDate, @CountDate," & vbCrLf
      sSql = sSql & "cl.NEWQTY - cl.OLDQTY, cl.NEWQTY - cl.OLDQTY, lh.LOTUNITCOST," & vbCrLf
      sSql = sSql & "abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTMATL / lh.LOTORIGINALQTY," & vbCrLf
      sSql = sSql & "abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTLABOR / lh.LOTORIGINALQTY," & vbCrLf
      sSql = sSql & "abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTEXP / lh.LOTORIGINALQTY," & vbCrLf
      sSql = sSql & "abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTOH / lh.LOTORIGINALQTY," & vbCrLf
      sSql = sSql & "abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTHRS / lh.LOTORIGINALQTY," & vbCrLf
      sSql = sSql & "case when cl.NEWQTY > cl.OLDQTY then dbo.fnGetPartInvAccount(cl.PARTREF) else @InvAdjAcct end," & vbCrLf
      sSql = sSql & "case when cl.NEWQTY > cl.OLDQTY then @InvAdjAcct else dbo.fnGetPartInvAccount(cl.PARTREF) end," & vbCrLf
      sSql = sSql & "@user" & vbCrLf
      sSql = sSql & "from CCLotAlloc cl" & vbCrLf
      sSql = sSql & "join LohdTable lh on cl.LOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "where cl.CCREF = @CCID" & vbCrLf
      sSql = sSql & "order by cl.PARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- create corresponding lot item records" & vbCrLf
      sSql = sSql & "insert into LoitTable" & vbCrLf
      sSql = sSql & "(LOINUMBER,LOIRECORD,LOITYPE,LOIPARTREF," & vbCrLf
      sSql = sSql & "LOIPDATE,LOIADATE,LOIQUANTITY,LOIACTIVITY,LOICOMMENT)" & vbCrLf
      sSql = sSql & "select lh.LOTNUMBER,(select isnull(max(LOIRECORD),0) + 1 from LoitTable" & vbCrLf
      sSql = sSql & "where LOINUMBER = lh.LOTNUMBER),30,cl.PARTREF," & vbCrLf
      sSql = sSql & "@CountDate, @CountDate, cl.NEWQTY - cl.OLDQTY, ia.INNUMBER, @CCID" & vbCrLf
      sSql = sSql & "from CCLotAlloc cl" & vbCrLf
      sSql = sSql & "join LohdTable lh on cl.LOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "join InvaTable ia on cl.PARTREF = ia.INPART and ia.INLOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "and ia.INTYPE = 30 and ia.INREF2 = @CCID" & vbCrLf
      sSql = sSql & "where cl.CCREF = @CCID" & vbCrLf
      sSql = sSql & "order by cl.PARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- update lot header remaining quantity" & vbCrLf
      sSql = sSql & "update LohdTable" & vbCrLf
      sSql = sSql & "set LOTAVAILABLE = case when (lh.LOTREMAININGQTY + cl.NEWQTY - cl.OLDQTY) = 0" & vbCrLf
      sSql = sSql & "then 0 else 1 end," & vbCrLf
      sSql = sSql & "LOTREMAININGQTY = lh.LOTREMAININGQTY +  cl.NEWQTY - cl.OLDQTY" & vbCrLf
      sSql = sSql & "from LohdTable lh" & vbCrLf
      sSql = sSql & "join CCLotAlloc cl on cl.LOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "where cl.CCREF = @CCID" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- update part table" & vbCrLf
      sSql = sSql & "update PartTable" & vbCrLf
      sSql = sSql & "set PAQOH = x.LotSum, PALOTQTYREMAINING = x.LotSum," & vbCrLf
      sSql = sSql & "PANEXTCYCLEDATE = @NextCountDate" & vbCrLf
      sSql = sSql & "from PartTable pt" & vbCrLf
      sSql = sSql & "join (select LOTPARTREF, sum(LOTREMAININGQTY) as LotSum from LohdTable lh" & vbCrLf
      sSql = sSql & "where LOTPARTREF in (select PARTREF from CCLotAlloc where CCREF = @CCID)" & vbCrLf
      sSql = sSql & "group by LOTPARTREF) x on pt.PARTREF = x.LOTPARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- set status of cycle count = Updated" & vbCrLf
      sSql = sSql & "update CchdTable" & vbCrLf
      sSql = sSql & "set CCUPDATED = 1," & vbCrLf
      sSql = sSql & "CCUPDATEDDATE = getdate()" & vbCrLf
      sSql = sSql & "where CCREF = @CCID" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "commit transaction" & vbCrLf
      Execute True, sSql
      
      'set version
      Execute False, "update Version set Version = " & newver
   End If

   newver = 48
   If ver < newver Then
      ver = newver
   
      'make discount amount same precision so ITDOLLARS = ITORIGINAL - ITDISCAMOUNT
      DropColumnDefault "SoitTable", "ITDISCAMOUNT"
      Execute True, "alter table SoitTable alter column ITDISCAMOUNT decimal(12,4) null"
      Execute True, "alter table SoitTable add constraint DF_SoitTable_ITDISCAMOUNT default 0 for ITDISCAMOUNT"

      Execute False, "drop procedure MaintSoDiscountInfo"
      sSql = "create procedure MaintSoDiscountInfo" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "/*" & vbCrLf
      sSql = sSql & "Script to fix discount info for R60.  Logic is" & vbCrLf
      sSql = sSql & "1. Always assume ITDOLLARS is correct." & vbCrLf
      sSql = sSql & "2. Always assume ITDISCRATE is correct except for conflicts with ITDOLLARS." & vbCrLf
      sSql = sSql & "3. If ITDISCRATE = 100 % and ITDOLLARS <> 0, set ITDISCRATE = 0 AND ITDOLLORIG = ITDOLLARS" & vbCrLf
      sSql = sSql & "4. If ITDISCRATE = 0 % then ITDOLLORIG = ITDOLLARS and ITDISCAMOUNT = 0" & vbCrLf
      sSql = sSql & "5. For any other discount rate ITDOLLORIG = 100 * ITDOLLARS / ( 100 – ITDISCRATE )" & vbCrLf
      sSql = sSql & "6. ITDISCAMOUNT = ITDOLLORIG –ITDOLLARS" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "test:" & vbCrLf
      sSql = sSql & "MaintSoDiscountInfo" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- diagnose problems" & vbCrLf
      sSql = sSql & "-- 100% discount with nonzero price" & vbCrLf
      sSql = sSql & "select ITDOLLORIG, ITDISCRATE, ITDISCAMOUNT, ITDOLLARS, * from SoitTable join CihdTable on ITINVOICE = INVNO" & vbCrLf
      sSql = sSql & "where ITDISCRATE = 100 and ITDOLLARS <> 0" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- if no discount, set ITDOLLORIG = ITDOLLARS" & vbCrLf
      sSql = sSql & "select ITDOLLORIG, ITDISCRATE, ITDISCAMOUNT, ITDOLLARS, * from SoitTable" & vbCrLf
      sSql = sSql & "where ITDISCRATE = 0 and ITDOLLORIG <> ITDOLLARS" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- for all other discount rates" & vbCrLf
      sSql = sSql & "select ITDOLLORIG, ITDISCRATE, ITDOLLARS, * from SoitTable " & vbCrLf
      sSql = sSql & "where ITDISCRATE <> 0 and ITDISCRATE <> 100 " & vbCrLf
      sSql = sSql & "and (cast(ITDOLLORIG * ITDISCRATE / 100 as decimal(12,4)) <> ITDISCAMOUNT or ITDOLLORIG - ITDISCAMOUNT <> ITDOLLARS)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "*/" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- fix problems" & vbCrLf
      sSql = sSql & "begin transaction" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- discount = 100% & nonzero price should be impossible (8)" & vbCrLf
      sSql = sSql & "update SoitTable set ITDISCRATE = 0, ITDOLLORIG = ITDOLLARS, ITDISCAMOUNT = 0 " & vbCrLf
      sSql = sSql & "where ITDISCRATE = 100 and ITDOLLARS <> 0" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- discount = 0% & ITDOLLORIG <> ITDOLLARS (1260)" & vbCrLf
      sSql = sSql & "update SoitTable set ITDISCRATE = 0, ITDOLLORIG = ITDOLLARS, ITDISCAMOUNT = 0 " & vbCrLf
      sSql = sSql & "where ITDISCRATE = 0 and ITDOLLORIG <> ITDOLLARS" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- for other discount rates, set calculate correct ITDOLLORIG" & vbCrLf
      sSql = sSql & "update SoitTable set ITDOLLORIG = 100 * ITDOLLARS / ( 100 - ITDISCRATE )" & vbCrLf
      sSql = sSql & "where ITDISCRATE <> 0 AND ITDISCRATE <> 100" & vbCrLf
      sSql = sSql & "and ITDOLLORIG <> cast(100 * ITDOLLARS / ( 100 - ITDISCRATE ) as decimal(12,4))" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- for other discount rates, calculate discount amount" & vbCrLf
      sSql = sSql & "update SoitTable set ITDISCAMOUNT = ITDOLLORIG - ITDOLLARS" & vbCrLf
      sSql = sSql & "where ITDISCRATE <> 0 AND ITDISCRATE <> 100" & vbCrLf
      sSql = sSql & "and ITDISCAMOUNT <> ITDOLLORIG - ITDOLLARS" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "commit transaction" & vbCrLf
      Execute True, sSql
      
      Execute True, "exec MaintSoDiscountInfo"

      'set version
      Execute False, "update Version set Version = " & newver
   End If
   
   newver = 49
   If ver < newver Then
      ver = newver
      ' can't drop this table.  must retain current POM logins
      sSql = "CREATE TABLE IstcTable" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & " ISEMPLOYEE int NOT NULL DEFAULT 0," & vbCrLf
      sSql = sSql & " ISMO char(30) NOT NULL DEFAULT ''," & vbCrLf
      sSql = sSql & " ISRUN int NOT NULL DEFAULT 0," & vbCrLf
      sSql = sSql & " ISOP int NOT NULL DEFAULT 0," & vbCrLf
      sSql = sSql & " ISMOSTART smalldatetime NULL," & vbCrLf
      sSql = sSql & " ISMOEND smalldatetime NULL," & vbCrLf
      sSql = sSql & " ISMOTYPE char(2) NULL," & vbCrLf
      sSql = sSql & " ISOPYIELD real NULL," & vbCrLf
      sSql = sSql & " ISOPCOMMENT char(20) NULL," & vbCrLf
      sSql = sSql & " ISOPCOMPLETE tinyint NULL DEFAULT 0," & vbCrLf
      sSql = sSql & " ISOPLOGOFF tinyint NULL DEFAULT 0," & vbCrLf
      sSql = sSql & " ISBREAK tinyint NULL DEFAULT 0," & vbCrLf
      sSql = sSql & " ISOPACCEPT real NULL," & vbCrLf
      sSql = sSql & " ISOPREJECT real NULL," & vbCrLf
      sSql = sSql & " ISOPSCRAP real NULL," & vbCrLf
      sSql = sSql & " ISOPREJTAG char(12) NULL," & vbCrLf
      sSql = sSql & " ISOPSUCOMPLETE tinyint NULL DEFAULT 0," & vbCrLf
      sSql = sSql & " ISSHOP char(12) NULL DEFAULT ''," & vbCrLf
      sSql = sSql & " ISWCNT char(12) NULL DEFAULT ''," & vbCrLf
      sSql = sSql & " ReadyToDelete tinyint NULL DEFAULT 0" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      Execute False, sSql

      Execute False, "drop function dbo.fnGetOpenJournalID"
      sSql = "create function fnGetOpenJournalID" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & " @Type char(2),    -- 'TJ', 'SJ', etc." & vbCrLf
      sSql = sSql & " @Date datetime    -- get a journal including this date (time is truncated)" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "returns varchar(12)     -- null if not found" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "/* test" & vbCrLf
      sSql = sSql & " select dbo.fnGetOpenJournalID( 'TJ', getdate())" & vbCrLf
      sSql = sSql & "*/" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & " declare @id varchar(12)" & vbCrLf
      sSql = sSql & " declare @truncatedDate datetime" & vbCrLf
      sSql = sSql & " set @truncatedDate = cast(convert(varchar(10),@Date,101) as datetime)" & vbCrLf
      sSql = sSql & " select top 1@id = MJGLJRNL from JrhdTable" & vbCrLf
      sSql = sSql & "    where MJTYPE = @Type and @truncatedDate between MJSTART and MJEND" & vbCrLf
      sSql = sSql & "    and MJCLOSED is null" & vbCrLf
      sSql = sSql & " return @id" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      Execute True, sSql

      Execute False, "drop function dbo.fnGetJournalID"
      sSql = "create function fnGetJournalID" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & " @Type char(2),    -- 'TJ', 'SJ', etc." & vbCrLf
      sSql = sSql & " @Date datetime    -- get a journal including this date (time is truncated)" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "returns varchar(12)     -- null if not found" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "/* test" & vbCrLf
      sSql = sSql & " select dbo.fnGetJournalID( 'TJ', getdate())" & vbCrLf
      sSql = sSql & "*/" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & " declare @id varchar(12)" & vbCrLf
      sSql = sSql & " declare @truncatedDate datetime" & vbCrLf
      sSql = sSql & " set @truncatedDate = cast(convert(varchar(10),@Date,101) as datetime)" & vbCrLf
      sSql = sSql & " select top 1@id = MJGLJRNL from JrhdTable" & vbCrLf
      sSql = sSql & "    where MJTYPE = @Type and @truncatedDate between MJSTART and MJEND" & vbCrLf
      sSql = sSql & " return @id" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      Execute True, sSql
      
      ' add missing column to old POM time entries
      sSql = "begin transaction" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- get regular time coe" & vbCrLf
      sSql = sSql & "declare @RegularTimeCode varchar(4)" & vbCrLf
      sSql = sSql & "select top 1 @RegularTimeCode = TYPECODE from TmcdTable where TYPETYPE = 'R'" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- set indirect indicator" & vbCrLf
      sSql = sSql & "update TcitTable set TCSURUN = 'I' where isnull(TCPARTREF,'') = '' and TCSURUN <> 'I'" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- set parameters where no TCCODE" & vbCrLf
      sSql = sSql & "update TcitTable set TCCODE = @RegularTimeCode," & vbCrLf
      sSql = sSql & " TCRATE = emp.PREMPAYRATE," & vbCrLf
      sSql = sSql & " TCOHRATE = 0,           -- indirect" & vbCrLf
      sSql = sSql & " TCRATENO = 1," & vbCrLf
      sSql = sSql & " TCACCT = emp.PREMACCTS," & vbCrLf
      sSql = sSql & " TCACCOUNT = emp.PREMACCTS," & vbCrLf
      sSql = sSql & " TCGLJOURNAL = dbo.fnGetJournalID( 'TJ', tm.TMDAY )" & vbCrLf
      sSql = sSql & "from TcitTable tc" & vbCrLf
      sSql = sSql & "join TchdTable tm on tc.TCCARD = tm.TMCARD" & vbCrLf
      sSql = sSql & "join EmplTable emp on emp.PREMNUMBER = tc.TCEMP" & vbCrLf
      sSql = sSql & "where TCSURUN = 'I' and TCCODE = ''" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- add journals where there was already a TCCODE but no TCGLJOURNAL" & vbCrLf
      sSql = sSql & "update TcitTable set TCGLJOURNAL = isnull(dbo.fnGetJournalID( 'TJ', tm.TMDAY ), '')" & vbCrLf
      sSql = sSql & "from TcitTable tc" & vbCrLf
      sSql = sSql & "join TchdTable tm on tc.TCCARD = tm.TMCARD" & vbCrLf
      sSql = sSql & "where isnull(TCGLJOURNAL, '') = ''" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "commit transaction" & vbCrLf
      Execute True, sSql
      
      'set version
      Execute False, "update Version set Version = " & newver
   End If

   newver = 50
   If ver < newver Then
      ver = newver

      Execute False, "drop table EsReportBomTable"
      sSql = "CREATE TABLE EsReportBomTable" & vbCrLf _
         & "(" & vbCrLf _
         & "  BomUser VARCHAR(4) NOT NULL DEFAULT('')," & vbCrLf _
         & "  BomRow INT NULL DEFAULT(0)," & vbCrLf _
         & "  BomLevel TINYINT NULL DEFAULT(0)," & vbCrLf _
         & "  BomAssembly CHAR(30) NULL DEFAULT('')," & vbCrLf _
         & "  BomPartRef CHAR(30) NULL DEFAULT('')," & vbCrLf _
         & "  BomRevision CHAR(4) NULL DEFAULT('')," & vbCrLf _
         & "  BomQuantity DECIMAL(15,4) NULL DEFAULT(0)," & vbCrLf _
         & "  BomUnits CHAR(2) NULL DEFAULT('')," & vbCrLf _
         & "  BomConversion SMALLINT NULL DEFAULT(0)," & vbCrLf _
         & "  BomSequence SMALLINT NULL DEFAULT(0)," & vbCrLf _
         & "  BomSortKey VARCHAR(60) NULL DEFAULT ('')," & vbCrLf _
         & "  ExplodedQty DECIMAL(15,4) NULL DEFAULT (0)," & vbCrLf _
         & "  MostRecentCost DECIMAL(15,4) NULL DEFAULT(0)" & vbCrLf _
         & ")"
      Execute True, sSql
      
      sSql = "CREATE INDEX IX_EsReportBomTable_BomRow ON " _
           & "EsReportBomTable(BomUser,BomRow) WITH FILLFACTOR = 80 "
      Execute True, sSql
      
      sSql = "CREATE INDEX IX_EsReportBomTable_BomPartRef ON " _
           & "EsReportBomTable(BomUser,BomPartRef) WITH FILLFACTOR = 80 "
      Execute True, sSql

      Execute False, "drop procedure UpdateCycleCount"
      sSql = "create procedure UpdateCycleCount" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   @CCID varchar(20)," & vbCrLf
      sSql = sSql & "   @user varchar(10)" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "/* test" & vbCrLf
      sSql = sSql & "exec UpdateCycleCount '20080802-A', 'TEL'" & vbCrLf
      sSql = sSql & "*/" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- terminate query on error" & vbCrLf
      sSql = sSql & "set ARITHABORT ON" & vbCrLf
      sSql = sSql & "set ANSI_WARNINGS ON" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- make sure all adjustments can be accomodated by lots" & vbCrLf
      sSql = sSql & "exec AllocateCycleCountLots @CCID" & vbCrLf
      sSql = sSql & "declare @problemCount int" & vbCrLf
      sSql = sSql & "select @problemCount = count(*) from CCLog where CCREF = @CCID and ERRORTYPE = 'FATAL'" & vbCrLf
      sSql = sSql & "if @problemCount <> 0" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   print 'UpdateCycleCount ' + @CCID + ' cannot proceed.  See CCLog table.'" & vbCrLf
      sSql = sSql & "   return" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "begin transaction" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare @InvAdjAcct varchar(12)" & vbCrLf
      sSql = sSql & "declare @CountDate datetime" & vbCrLf
      sSql = sSql & "declare @PlanDate datetime" & vbCrLf
      sSql = sSql & "declare @NextCountDate datetime" & vbCrLf
      sSql = sSql & "declare @ABCCode varchar(2)" & vbCrLf
      sSql = sSql & "declare @ABCFrequency int" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "SET NOCOUNT ON" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "select @InvAdjAcct = isnull(COADJACCT,'?') from ComnTable where COREF = 1" & vbCrLf
      sSql = sSql & "select @CountDate = CCCOUNTLOCKEDDATE," & vbCrLf
      sSql = sSql & "@PlanDate = CCPLANDATE," & vbCrLf
      sSql = sSql & "@ABCCode = CCABCCODE " & vbCrLf
      sSql = sSql & "from CchdTable where CCREF = @CCID" & vbCrLf
      sSql = sSql & "select @ABCFrequency = isnull(COABCFREQUENCY,90) from CabcTable where COABCCODE = @ABCCode" & vbCrLf
      sSql = sSql & "select @NextCountDate = dateadd(d, @ABCFrequency, @PlanDate)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- create inventory activities" & vbCrLf
      sSql = sSql & "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INLOTNUMBER,INPDATE,INADATE," & vbCrLf
      sSql = sSql & "INPQTY,INAQTY,INAMT,INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS," & vbCrLf
      sSql = sSql & "INDEBITACCT,INCREDITACCT,INUSER)" & vbCrLf
      sSql = sSql & "select 30, cl.PARTREF, 'ABC Cycle Count', @CCID, cl.LOTNUMBER, @CountDate, @CountDate," & vbCrLf
      sSql = sSql & "cl.NEWQTY - cl.OLDQTY, cl.NEWQTY - cl.OLDQTY, lh.LOTUNITCOST," & vbCrLf
      sSql = sSql & "cast(abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTMATL / case when lh.LOTORIGINALQTY = 0 then 1 else lh.LOTORIGINALQTY end as decimal(12,4))," & vbCrLf
      sSql = sSql & "cast(abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTLABOR / case when lh.LOTORIGINALQTY = 0 then 1 else lh.LOTORIGINALQTY end as decimal(12,4))," & vbCrLf
      sSql = sSql & "cast(abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTEXP / case when lh.LOTORIGINALQTY = 0 then 1 else lh.LOTORIGINALQTY end as decimal(12,4))," & vbCrLf
      sSql = sSql & "cast(abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTOH / case when lh.LOTORIGINALQTY = 0 then 1 else lh.LOTORIGINALQTY end as decimal(12,4))," & vbCrLf
      sSql = sSql & "cast(abs(cl.NEWQTY - cl.OLDQTY) * lh.LOTTOTHRS / case when lh.LOTORIGINALQTY = 0 then 1 else lh.LOTORIGINALQTY end as decimal(12,4))," & vbCrLf
      sSql = sSql & "case when cl.NEWQTY > cl.OLDQTY then dbo.fnGetPartInvAccount(cl.PARTREF) else @InvAdjAcct end," & vbCrLf
      sSql = sSql & "case when cl.NEWQTY > cl.OLDQTY then @InvAdjAcct else dbo.fnGetPartInvAccount(cl.PARTREF) end," & vbCrLf
      sSql = sSql & "@user" & vbCrLf
      sSql = sSql & "from CCLotAlloc cl" & vbCrLf
      sSql = sSql & "join LohdTable lh on cl.LOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "where cl.CCREF = @CCID" & vbCrLf
      sSql = sSql & "order by cl.PARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- create corresponding lot item records" & vbCrLf
      sSql = sSql & "insert into LoitTable" & vbCrLf
      sSql = sSql & "(LOINUMBER,LOIRECORD,LOITYPE,LOIPARTREF," & vbCrLf
      sSql = sSql & "LOIPDATE,LOIADATE,LOIQUANTITY,LOIACTIVITY,LOICOMMENT)" & vbCrLf
      sSql = sSql & "select lh.LOTNUMBER,(select isnull(max(LOIRECORD),0) + 1 from LoitTable" & vbCrLf
      sSql = sSql & "where LOINUMBER = lh.LOTNUMBER),30,cl.PARTREF," & vbCrLf
      sSql = sSql & "@CountDate, @CountDate, cl.NEWQTY - cl.OLDQTY, ia.INNUMBER, @CCID" & vbCrLf
      sSql = sSql & "from CCLotAlloc cl" & vbCrLf
      sSql = sSql & "join LohdTable lh on cl.LOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "join InvaTable ia on cl.PARTREF = ia.INPART and ia.INLOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "and ia.INTYPE = 30 and ia.INREF2 = @CCID" & vbCrLf
      sSql = sSql & "where cl.CCREF = @CCID" & vbCrLf
      sSql = sSql & "order by cl.PARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- update lot header remaining quantity" & vbCrLf
      sSql = sSql & "update LohdTable" & vbCrLf
      sSql = sSql & "set LOTAVAILABLE = case when (lh.LOTREMAININGQTY + cl.NEWQTY - cl.OLDQTY) = 0" & vbCrLf
      sSql = sSql & "then 0 else 1 end," & vbCrLf
      sSql = sSql & "LOTREMAININGQTY = lh.LOTREMAININGQTY +  cl.NEWQTY - cl.OLDQTY" & vbCrLf
      sSql = sSql & "from LohdTable lh" & vbCrLf
      sSql = sSql & "join CCLotAlloc cl on cl.LOTNUMBER = lh.LOTNUMBER" & vbCrLf
      sSql = sSql & "where cl.CCREF = @CCID" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- update part table" & vbCrLf
      sSql = sSql & "update PartTable" & vbCrLf
      sSql = sSql & "set PAQOH = x.LotSum, PALOTQTYREMAINING = x.LotSum," & vbCrLf
      sSql = sSql & "PANEXTCYCLEDATE = @NextCountDate" & vbCrLf
      sSql = sSql & "from PartTable pt" & vbCrLf
      sSql = sSql & "join (select LOTPARTREF, sum(LOTREMAININGQTY) as LotSum from LohdTable lh" & vbCrLf
      sSql = sSql & "where LOTPARTREF in (select PARTREF from CCLotAlloc where CCREF = @CCID)" & vbCrLf
      sSql = sSql & "group by LOTPARTREF) x on pt.PARTREF = x.LOTPARTREF" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- set status of cycle count = Updated" & vbCrLf
      sSql = sSql & "update CchdTable" & vbCrLf
      sSql = sSql & "set CCUPDATED = 1," & vbCrLf
      sSql = sSql & "CCUPDATEDDATE = getdate()" & vbCrLf
      sSql = sSql & "where CCREF = @CCID" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "commit transaction" & vbCrLf
      Execute True, sSql
      
      'set version
      Execute False, "update Version set Version = " & newver
   End If

   newver = 51
   If ver < newver Then
      ver = newver
      DropColumnDefault "Alerts", "ALERTMSG"
      Execute True, "alter table Alerts alter column ALERTMSG varchar(255) null"
      Execute True, "alter table Alerts add constraint DF_Alerts_ALERTMSG default '' for ALERTMSG"

      'set version
      Execute False, "update Version set Version = " & newver
   End If

   newver = 52
   If ver < newver Then
      ver = newver
      
      Execute False, "drop procedure UpdateTimeCardTotals"
      sSql = "create procedure UpdateTimeCardTotals" & vbCrLf
      sSql = sSql & " @EmpNo int," & vbCrLf
      sSql = sSql & " @Date datetime" & vbCrLf
      sSql = sSql & "as " & vbCrLf
      sSql = sSql & "/* test" & vbCrLf
      sSql = sSql & "    UpdateTimeCardTotals 52, '9/8/2008'" & vbCrLf
      sSql = sSql & "*/" & vbCrLf
''      sSql = sSql & "update TchdTable " & vbCrLf
''      sSql = sSql & " set TMSTART = (select top 1 TCSTART From TcitTable " & vbCrLf
''      sSql = sSql & " join TchdTable on TCCARD = TMCARD" & vbCrLf
''      sSql = sSql & " where TCEMP = @EmpNo and TMDAY = @Date and rtrim(TCSTOP) <> ''" & vbCrLf
''      sSql = sSql & " order by cast(rtrim(TCSTART) + 'm' as datetime))," & vbCrLf
''      sSql = sSql & " TMSTOP = (select top 1 TCSTOP From TcitTable" & vbCrLf
''      sSql = sSql & " join TchdTable on TCCARD = TMCARD" & vbCrLf
''      sSql = sSql & " where TCEMP = @EmpNo and TMDAY = @Date and rtrim(TCSTOP) <> ''" & vbCrLf
''      sSql = sSql & " order by case when datediff( n, cast(rtrim(TCSTART) + 'm' as datetime)," & vbCrLf
''      sSql = sSql & " cast(rtrim(TCSTOP) + 'm' as datetime) ) >= 0  " & vbCrLf
''      sSql = sSql & " then cast(rtrim(TCSTOP) + 'm' as datetime) " & vbCrLf
''      sSql = sSql & " else dateadd(day, 1, cast(rtrim(TCSTOP) + 'm' as datetime)) end desc)" & vbCrLf
''      sSql = sSql & "where TMEMP = @EmpNo and TMDAY = @Date" & vbCrLf
''      sSql = sSql & "and ISDATE(TMSTART + 'm') = 1 and ISDATE(TMSTOP + 'm') = 1" & vbCrLf
      
      sSql = sSql & "update TchdTable " & vbCrLf
      sSql = sSql & "set TMSTART = (select top 1 TCSTART From TcitTable " & vbCrLf
      sSql = sSql & " join TchdTable on TCCARD = TMCARD" & vbCrLf
      sSql = sSql & " where TCEMP = @EmpNo and TMDAY = @Date and TCSTOP <> ''" & vbCrLf
      sSql = sSql & " and ISDATE(TCSTART + 'm') = 1 and ISDATE(TCSTOP + 'm') = 1" & vbCrLf
      sSql = sSql & " order by cast(rtrim(TCSTART) + 'm' as datetime))," & vbCrLf
      sSql = sSql & "TMSTOP = (select top 1 TCSTOP From TcitTable" & vbCrLf
      sSql = sSql & " join TchdTable on TCCARD = TMCARD" & vbCrLf
      sSql = sSql & " where TCEMP = @EmpNo and TMDAY = @Date and TCSTOP <> ''" & vbCrLf
      sSql = sSql & " and ISDATE(TCSTART + 'm') = 1 and ISDATE(TCSTOP + 'm') = 1" & vbCrLf
      sSql = sSql & " order by case when datediff( n, cast(rtrim(TCSTART) + 'm' as datetime)," & vbCrLf
      sSql = sSql & " cast(rtrim(TCSTOP) + 'm' as datetime) ) >= 0  " & vbCrLf
      sSql = sSql & " then cast(rtrim(TCSTOP) + 'm' as datetime) " & vbCrLf
      sSql = sSql & " else dateadd(day, 1, cast(rtrim(TCSTOP) + 'm' as datetime)) end desc)" & vbCrLf
      sSql = sSql & "where TMEMP = @EmpNo and TMDAY = @Date" & vbCrLf
      sSql = sSql & vbCrLf
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
      sSql = sSql & "where TMEMP = @EmpNo and TMDAY = @Date" & vbCrLf
      Execute True, sSql
      
      'delete bad conversion time charges with no start time and negative hours
      Execute True, "delete from tcittable where TCSTART = ''"
      
      'recalculate timecard start and stop times
      Execute True, "UpdateAllTimeCardTotals"

      'delete orphan vendor invoice items with no corresponding vendor invoice header
      '(these were mostly, if not all, brought forward in a conversion).  LUMICOR had 48K
      sSql = "delete from viittable where vitno not in (select vino from vihdtable" & vbCrLf _
         & "where vivendor = vitvendor and vitno = vino)"
      Execute True, sSql
      
      'now create a PK/FK relationship so this can't happen again
      Execute False, "ALTER TABLE ViitTable DROP CONSTRAINT FK_ViitTable_VihdTable"
      sSql = "ALTER TABLE ViitTable WITH CHECK ADD CONSTRAINT FK_ViitTable_VihdTable" & vbCrLf _
         & "FOREIGN KEY(VITNO, VITVENDOR)" & vbCrLf _
         & "REFERENCES VihdTable(VINO, VIVENDOR)"
      Execute True, sSql
      
''      Execute False, "drop view viewWipMaterialDetail"
''      sSql = "create view viewWipMaterialCosts" & vbCrLf
''      sSql = sSql & "as" & vbCrLf
''      sSql = sSql & "select LOIMOPARTREF, LOIMORUNNO, LOITYPE, LOIADATE, LOIQUANTITY, LOTUNITCOST," & vbCrLf
''      sSql = sSql & "cast(-LOIQUANTITY * LOTUNITCOST as decimal(12,2)) as ExtendedCost" & vbCrLf
''      sSql = sSql & "from LoitTable" & vbCrLf
''      sSql = sSql & "join LohdTable on LOINUMBER = LOTNUMBER" & vbCrLf
''      sSql = sSql & "where LOITYPE in (10, 12) and LOTUNITCOST > 0" & vbCrLf
''      Execute True, sSql
''
      'set version
      Execute False, "update Version set Version = " & newver
   End If

   newver = 53
   If ver < newver Then
      ver = newver
      'Execute True, sSql

      Execute True, "ALTER TABLE TcitTable ADD TCSOURCE varchar(6) NOT NULL default ''"
      Execute True, "ALTER TABLE TcitTable ADD TCSTARTTIME smalldatetime NULL"
      Execute True, "ALTER TABLE TcitTable ADD TCSTOPTIME smalldatetime NULL"
      Execute True, "ALTER TABLE TcitTable ADD TCENTERED smalldatetime NOT NULL default getdate()"
      
      sSql = "update TcitTable set TCSOURCE = 'POM'" & vbCrLf _
         & "where TCCARD in (SELECT TMCARD from TchdTable where TMWEEK is null)"
      Execute True, sSql
      
      sSql = "update TcitTable set TCSOURCE = 'TS'" & vbCrLf _
         & "where TCSOURCE = ''"
      Execute True, sSql
      
      Execute True, "ALTER TABLE IstcTable ADD ISINDIRECT tinyint NOT NULL default 0"
      sSql = "insert into IstcTable (ISEMPLOYEE, ISMO, ISRUN,ISOP, ISMOSTART, ISINDIRECT) " & vbCrLf _
         & "SELECT TCEMP, '', 0, 0, TCTIME, 1 FROM TcitTable tc1" & vbCrLf _
         & "where TCSTOP = ''" & vbCrLf _
         & "and TCTIME = (select max(TCTIME) FROM TcitTable tc2" & vbCrLf _
         & "where tc2.TCEMP = tc1.TCEMP and TCSTOP = '')" & vbCrLf _
         & "and TCTIME >= DATEADD(day,-7,getdate())"
      Execute True, sSql
      
      'delete old format open indirect time charges
      Execute True, "delete from TcitTable where TCSTOP = ''"
      
      'delete old overlapping temp time charges
      Execute True, "delete from IstcTable where ReadyToDelete = 1"
      
      'populate newly created TCSTARTTIME and TCENDTIME columns with datetimes
      RdoCon.BeginTrans
      Execute True, "delete from TcitTable where TCHOURS = 0"
      
      sSql = "Update TcitTable" & vbCrLf _
         & "set TCSTARTTIME = cast( convert(varchar(10),TMDAY,101) + ' ' + TCSTART + 'm' as datetime)" & vbCrLf _
         & "from TcitTable tc join TchdTable on TMCARD = TCCARD" & vbCrLf _
         & "where ISDATE(tc.TCSTART + 'm') = 1" & vbCrLf _
         & "and tc.TCSTARTTIME is null"
      Execute True, sSql

      sSql = "Update TcitTable" & vbCrLf _
         & "set TCSTOPTIME = cast( convert(varchar(10),TMDAY,101) + ' ' + TCSTOP + 'm' as datetime)" & vbCrLf _
         & "from TcitTable tc join TchdTable on TMCARD = TCCARD" & vbCrLf _
         & "where ISDATE(tc.TCSTOP + 'm') = 1" & vbCrLf _
         & "and tc.TCSTOPTIME is null"
      Execute True, sSql

      sSql = "Update TcitTable" & vbCrLf _
         & "Set TCSTOPTIME = DateAdd(d, 1, TCSTOPTIME)" & vbCrLf _
         & "Where TCSTOPTIME <= TCSTARTTIME"
      Execute True, sSql
      
      RdoCon.CommitTrans
      
      'for finance > closing > inventory adjustments report (fincl14.rpt)
      Execute False, "drop view viewRptInventoryAdjustments"
      sSql = "CREATE VIEW viewRptInventoryAdjustments" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "SELECT     dbo.LoitTable.LOINUMBER, dbo.LoitTable.LOITYPE, dbo.PartTable.PADESC, dbo.PartTable.PACLASS, dbo.LoitTable.LOIPARTREF, " & vbCrLf
      sSql = sSql & " dbo.LoitTable.LOIADATE, dbo.LoitTable.LOIQUANTITY, dbo.LohdTable.LOTORIGINALQTY, dbo.LohdTable.LOTUNITCOST, dbo.LohdTable.LOTTOTMATL, " & vbCrLf
      sSql = sSql & "    dbo.LohdTable.LOTTOTLABOR, dbo.LohdTable.LOTTOTEXP" & vbCrLf
      sSql = sSql & "FROM         dbo.LoitTable LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "    dbo.LohdTable ON dbo.LoitTable.LOINUMBER = dbo.LohdTable.LOTNUMBER LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "    dbo.PartTable ON dbo.LoitTable.LOIPARTREF = dbo.PartTable.PARTREF AND dbo.LohdTable.LOTPARTREF = dbo.PartTable.PARTREF" & vbCrLf
      sSql = sSql & "WHERE     (dbo.LoitTable.LOITYPE = 19) OR" & vbCrLf
      sSql = sSql & "    (dbo.LoitTable.LOITYPE = 1)" & vbCrLf
      Execute True, sSql
      
      Execute False, "drop procedure RptInventoryAdjustments"
      sSql = "create PROCEDURE RptInventoryAdjustments" & vbCrLf
      sSql = sSql & " @StartDate as varchar(16), @EndDate as Varchar(16), @PartClass as Varchar(16)" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "BEGIN" & vbCrLf
      sSql = sSql & " SET NOCOUNT ON;" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & " IF (@PartClass = 'ALL')" & vbCrLf
      sSql = sSql & " BEGIN" & vbCrLf
      sSql = sSql & "    SET @PartClass = ''" & vbCrLf
      sSql = sSql & " END" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & " SELECT  PADESC as PartDesc, LOIPARTREF as PartRef, LOIQUANTITY as Quantity," & vbCrLf
      sSql = sSql & "    (LOIQUANTITY * LOTUNITCOST) as Material, PACLASS as PartClass" & vbCrLf
      sSql = sSql & " FROM viewRptInventoryAdjustments " & vbCrLf
      sSql = sSql & " WHERE (LOIADATE BETWEEN @StartDate AND @EndDate)" & vbCrLf
      sSql = sSql & " AND PACLASS LIKE '%' + @PartClass + '%'" & vbCrLf
      sSql = sSql & "END" & vbCrLf
      Execute True, sSql
      
      'set version
      Execute False, "update Version set Version = " & newver
   End If

   newver = 54
   If ver < newver Then
      ver = newver
      
      'this probably already exists
      Execute False, "alter table Version add TestDatabase tinyint not null default 0"

      'set version
      Execute False, "update Version set Version = " & newver
   End If

   'R67
   newver = 55
   If ver < newver Then
      ver = newver
      
      sSql = "ALTER TABLE RtopTable DROP CONSTRAINT FK_RtopTable_RthdTable"
      Execute False, sSql
      sSql = "ALTER TABLE RtopTable WITH CHECK ADD  CONSTRAINT FK_RtopTable_RthdTable FOREIGN KEY(OPREF)" & vbCrLf _
         & "References RthdTable(RTREF)" & vbCrLf _
         & "ON UPDATE CASCADE"
      Execute True, sSql

      
      Execute True, "alter table Alerts alter column ALERTMSG varchar(1024) null"
      
      Execute False, "drop table ChHdrTable"
      sSql = "create table ChHdrTable" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   ChkNum varchar(12) null," & vbCrLf
      sSql = sSql & "   ChkVnd varchar(10) null," & vbCrLf
      sSql = sSql & "   ChkNme varchar(40) null," & vbCrLf
      sSql = sSql & "   ChkAdd varchar(160) null," & vbCrLf
      sSql = sSql & "   ChkCty varchar(18) null," & vbCrLf
      sSql = sSql & "   ChkSte varchar(4) null," & vbCrLf
      sSql = sSql & "   ChkZip varchar(10) null," & vbCrLf
      sSql = sSql & "   ChkAmt decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   ChkPAmt decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   ChkDte smalldatetime null," & vbCrLf
      sSql = sSql & "   ChkTxt varchar(80) null," & vbCrLf
      sSql = sSql & "   ChkMem varchar(40) null," & vbCrLf
      sSql = sSql & "   ChkAcct varchar(12) null" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      Execute True, sSql
      
      Execute False, "drop table ChDetTable"
      sSql = "create table ChDetTable" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   DetNum01 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv01 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte01 datetime null," & vbCrLf
      sSql = sSql & "   DetDis01 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt01 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt01 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum02 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv02 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte02 datetime null," & vbCrLf
      sSql = sSql & "   DetDis02 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt02 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt02 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum03 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv03 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte03 datetime null," & vbCrLf
      sSql = sSql & "   DetDis03 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt03 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt03 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum04 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv04 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte04 datetime null," & vbCrLf
      sSql = sSql & "   DetDis04 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt04 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt04 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum05 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv05 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte05 datetime null," & vbCrLf
      sSql = sSql & "   DetDis05 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt05 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt05 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum06 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv06 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte06 datetime null," & vbCrLf
      sSql = sSql & "   DetDis06 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt06 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt06 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum07 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv07 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte07 datetime null," & vbCrLf
      sSql = sSql & "   DetDis07 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt07 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt07 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum08 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv08 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte08 datetime null," & vbCrLf
      sSql = sSql & "   DetDis08 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt08 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt08 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum09 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv09 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte09 datetime null," & vbCrLf
      sSql = sSql & "   DetDis09 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt09 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt09 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum10 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv10 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte10 datetime null," & vbCrLf
      sSql = sSql & "   DetDis10 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt10 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt10 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum11 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv11 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte11 datetime null," & vbCrLf
      sSql = sSql & "   DetDis11 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt11 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt11 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum12 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv12 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte12 datetime null," & vbCrLf
      sSql = sSql & "   DetDis12 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt12 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt12 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum13 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv13 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte13 datetime null," & vbCrLf
      sSql = sSql & "   DetDis13 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt13 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt13 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum14 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv14 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte14 datetime null," & vbCrLf
      sSql = sSql & "   DetDis14 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt14 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt14 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum15 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv15 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte15 datetime null," & vbCrLf
      sSql = sSql & "   DetDis15 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt15 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt15 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum16 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv16 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte16 datetime null," & vbCrLf
      sSql = sSql & "   DetDis16 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt16 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt16 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum17 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv17 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte17 datetime null," & vbCrLf
      sSql = sSql & "   DetDis17 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt17 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt17 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum18 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv18 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte18 datetime null," & vbCrLf
      sSql = sSql & "   DetDis18 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt18 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt18 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum19 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv19 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte19 datetime null," & vbCrLf
      sSql = sSql & "   DetDis19 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt19 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt19 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum20 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv20 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte20 datetime null," & vbCrLf
      sSql = sSql & "   DetDis20 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt20 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt20 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetNum21 varchar(12) null," & vbCrLf
      sSql = sSql & "   DetInv21 varchar(20) null," & vbCrLf
      sSql = sSql & "   DetDte21 datetime null," & vbCrLf
      sSql = sSql & "   DetDis21 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetPAmt21 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & "   DetAmt21 decimal(12,2) not null default 0," & vbCrLf
      sSql = sSql & ")"
      Execute True, sSql
      
      'ALLOW REPORT NAMES UP TO 30 CHARACTERS
      'Execute False, "DROP INDEX CustomReports.ReportRef"   'SQL2000
      'Execute False, "DROP INDEX ReportRef ON CustomReports" 'SQL2005 & SQL2008
      DropIndex "CustomReports", "ReportRef"
      AlterStringColumn "CustomReports", "REPORT_REF", "varchar(30)"
      Execute True, "CREATE INDEX ReportRef ON CustomReports (REPORT_REF)"

      AlterStringColumn "CustomReports", "REPORT_NAME", "varchar(30)"
      AlterStringColumn "CustomReports", "REPORT_CUSTOMREPORT", "varchar(30)"
      
      'set version
      Execute False, "update Version set Version = " & newver
   End If

   newver = 56
   If ver < newver Then
      ver = newver
      
      Execute False, "DROP TABLE Updates"
      sSql = "CREATE TABLE Updates" & vbCrLf _
         & "(" & vbCrLf _
         & "   UpdateID int IDENTITY(1,1) NOT NULL," & vbCrLf _
         & "   UpdateDate smalldatetime NOT NULL," & vbCrLf _
         & "   AppDirectory varchar(100) NOT NULL," & vbCrLf _
         & "   OldRelease int NOT NULL," & vbCrLf _
         & "   NewRelease int NOT NULL," & vbCrLf _
         & "   UserInitials varchar(4) NOT NULL," & vbCrLf _
         & "   OldDbVersion int NOT NULL," & vbCrLf _
         & "   NewDbVersion int NOT NULL," & vbCrLf _
         & "   OldDbType varchar(10) NOT NULL," & vbCrLf _
         & "   NewDbType varchar(10) NOT NULL" & vbCrLf _
         & ")"
      Execute True, sSql

      sSql = "ALTER TABLE Updates ADD CONSTRAINT DF_Updates_UpdateDate  DEFAULT (getdate()) FOR UpdateDate"
      Execute True, sSql
      
      sSql = "ALTER TABLE ComnTable ADD COLinesPerCheckStub tinyint not null default 15"
      Execute False, sSql
      
      Execute False, "drop view viewRptInventoryAdjustments"
      sSql = "CREATE VIEW viewRptInventoryAdjustments" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "SELECT     InvaTable.INTYPE, InvaTable.INPART, InvaTable.INADATE, InvaTable.INREF1, InvaTable.INAQTY, InvaTable.INAMT, " & vbCrLf
      sSql = sSql & "                      InvaTable.INTOTMATL, InvaTable.INTOTLABOR, InvaTable.INTOTEXP, InvaTable.INTOTOH, InvaTable.INCREDITACCT, " & vbCrLf
      sSql = sSql & "                      InvaTable.INDEBITACCT, LohdTable.LOTUNITCOST, LohdTable.LOTTOTMATL, LohdTable.LOTTOTLABOR, LohdTable.LOTTOTEXP, " & vbCrLf
      sSql = sSql & "                      LohdTable.LOTTOTOH" & vbCrLf
      sSql = sSql & "FROM         InvaTable LEFT OUTER JOIN" & vbCrLf
      sSql = sSql & "                      LohdTable ON InvaTable.INLOTNUMBER = LohdTable.LOTNUMBER AND InvaTable.INPART = LohdTable.LOTPARTREF" & vbCrLf
      sSql = sSql & "WHERE     (InvaTable.INTYPE = 19) OR" & vbCrLf
      sSql = sSql & "                      (InvaTable.INTYPE = 30)"
      Execute True, sSql
      
      Execute False, "drop procedure RptInventoryAdjustments"
      sSql = "CREATE PROCEDURE RptInventoryAdjustments" & vbCrLf
      sSql = sSql & "@StartDate as varchar(16), @EndDate as Varchar(16), @PartClass as Varchar(16),@PartCode as varchar(8)" & vbCrLf
      sSql = sSql & "AS" & vbCrLf
      sSql = sSql & "BEGIN" & vbCrLf
      sSql = sSql & "    --SET NOCOUNT ON;" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   declare @partRef as varchar(30)" & vbCrLf
      sSql = sSql & "   declare @partNum as varchar(30)" & vbCrLf
      sSql = sSql & "   declare @partDesc as varchar(30)" & vbCrLf
      sSql = sSql & "   declare @partExDesc as varchar(3072)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   declare @actualDt as varchar(30)" & vbCrLf
      sSql = sSql & "   declare @qty as int" & vbCrLf
      sSql = sSql & "   declare @invAmt as decimal(12,4)" & vbCrLf
      sSql = sSql & "   declare @invTotMatl decimal(12,4)" & vbCrLf
      sSql = sSql & "   declare @invTotLabor decimal(12,4)" & vbCrLf
      sSql = sSql & "   declare @invTotExp decimal(12,4)" & vbCrLf
      sSql = sSql & "   declare @invTotOH decimal(12,4)" & vbCrLf
      sSql = sSql & "   declare @creditAcc int" & vbCrLf
      sSql = sSql & "   declare @debitAcc int" & vbCrLf
      sSql = sSql & "   declare @lotTotMatl decimal(12,4)" & vbCrLf
      sSql = sSql & "   declare @lotTotLabor decimal(12,4)" & vbCrLf
      sSql = sSql & "   declare @lotTotExp decimal(12,4)" & vbCrLf
      sSql = sSql & "   declare @lotTotOH decimal(12,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   declare @totMatlCost decimal(12,4)" & vbCrLf
      sSql = sSql & "   declare @totLaborCost decimal(12,4)" & vbCrLf
      sSql = sSql & "   declare @totExpCost decimal(12,4)" & vbCrLf
      sSql = sSql & "   declare @totOHCost decimal(12,4)" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   declare @fInvMatl int" & vbCrLf
      sSql = sSql & "   declare @fInvLabor int" & vbCrLf
      sSql = sSql & "   declare @fInvExp int" & vbCrLf
      sSql = sSql & "   declare @fInvOH int" & vbCrLf
      sSql = sSql & "                  " & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "    IF (@PartClass = 'ALL')" & vbCrLf
      sSql = sSql & "    BEGIN" & vbCrLf
      sSql = sSql & "      SET @PartClass = ''" & vbCrLf
      sSql = sSql & "    END" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "    IF (@PartCode = 'ALL')" & vbCrLf
      sSql = sSql & "    BEGIN" & vbCrLf
      sSql = sSql & "      SET @PartCode = ''" & vbCrLf
      sSql = sSql & "    END" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   CREATE TABLE #tempINVReport" & vbCrLf
      sSql = sSql & "   (" & vbCrLf
      sSql = sSql & "      PARTNUM Varchar(30) NULL," & vbCrLf
      sSql = sSql & "      PADESC varchar(30) NULL ," & vbCrLf
      sSql = sSql & "      PAEXTDESC varchar(3072) NULL ," & vbCrLf
      sSql = sSql & "      INADATE varchar(30) NULL," & vbCrLf
      sSql = sSql & "      INAQTY decimal(12,4) NULL," & vbCrLf
      sSql = sSql & "      INAMT decimal (12,4) NULL," & vbCrLf
      sSql = sSql & "       TOTMATL decimal(12,4) NULL," & vbCrLf
      sSql = sSql & "       TOTLABOR decimal(12,4) NULL," & vbCrLf
      sSql = sSql & "       TOTEXP decimal(12,4) NULL," & vbCrLf
      sSql = sSql & "       TOTOH decimal(12,4) NULL," & vbCrLf
      sSql = sSql & "      CREDITACCT int NULL," & vbCrLf
      sSql = sSql & "      DEBITACCT int NULL," & vbCrLf
      sSql = sSql & "      flgMatl int NULL, " & vbCrLf
      sSql = sSql & "      flgLabor int NULL, " & vbCrLf
      sSql = sSql & "      flgExp int NULL, " & vbCrLf
      sSql = sSql & "      flgOH int NULL" & vbCrLf
      sSql = sSql & "   )" & vbCrLf
      sSql = sSql & "   " & vbCrLf
      sSql = sSql & "   DECLARE curInvRpt CURSOR   FOR" & vbCrLf
      sSql = sSql & "      SELECT INPART, INADATE, " & vbCrLf
      sSql = sSql & "         INAQTY, INAMT, INTOTMATL, " & vbCrLf
      sSql = sSql & "         INTOTLABOR, INTOTEXP, INTOTOH, " & vbCrLf
      sSql = sSql & "         INCREDITACCT, INDEBITACCT, LOTTOTMATL, " & vbCrLf
      sSql = sSql & "         LOTTOTLABOR, LOTTOTEXP, LOTTOTOH, " & vbCrLf
      sSql = sSql & "         PARTNUM, PADESC, PAEXTDESC" & vbCrLf
      sSql = sSql & "      FROM  viewRptInventoryAdjustments, PartTable " & vbCrLf
      sSql = sSql & "      WHERE viewRptInventoryAdjustments.INPART = PartTable.PARTREF" & vbCrLf
      sSql = sSql & "         AND viewRptInventoryAdjustments.INADATE BETWEEN @StartDate AND @EndDate" & vbCrLf
      sSql = sSql & "         AND PartTable.PACLASS LIKE '%' + @PartClass + '%'" & vbCrLf
      sSql = sSql & "         AND PartTable.PAPRODCODE LIKE '%' + @PartCode + '%'" & vbCrLf
      sSql = sSql & "         --AND INPART = '9148177X48X96'" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   OPEN curInvRpt" & vbCrLf
      sSql = sSql & "   FETCH NEXT FROM curInvRpt INTO @partRef, @actualDt, @qty, " & vbCrLf
      sSql = sSql & "                  @invAmt, @invTotMatl, @invTotLabor, " & vbCrLf
      sSql = sSql & "                  @invTotExp,@invTotOH, @creditAcc, @debitAcc, " & vbCrLf
      sSql = sSql & "                  @lotTotMatl, @lotTotLabor, @lotTotExp, @lotTotOH, " & vbCrLf
      sSql = sSql & "                  @partNum, @partDesc, @partExDesc" & vbCrLf
      sSql = sSql & "   WHILE (@@FETCH_STATUS <> -1)" & vbCrLf
      sSql = sSql & "   BEGIN" & vbCrLf
      sSql = sSql & "      IF (@@FETCH_STATUS <> -2)" & vbCrLf
      sSql = sSql & "      BEGIN" & vbCrLf
      sSql = sSql & "         " & vbCrLf
      sSql = sSql & "         -- Get the costed values from Lothd table" & vbCrLf
      sSql = sSql & "         -- if the Inv table does not have the cost for" & vbCrLf
      sSql = sSql & "         -- material, expenses, OH and Labour." & vbCrLf
      sSql = sSql & "         SET @totMatlCost = @invTotMatl" & vbCrLf
      sSql = sSql & "         SET @fInvMatl = 1" & vbCrLf
      sSql = sSql & "         IF (@invTotMatl = 0.0000)" & vbCrLf
      sSql = sSql & "         BEGIN" & vbCrLf
      sSql = sSql & "            IF (@lotTotMatl IS NOT NULL)" & vbCrLf
      sSql = sSql & "            BEGIN" & vbCrLf
      sSql = sSql & "               SET @totMatlCost = @lotTotMatl" & vbCrLf
      sSql = sSql & "               SET @fInvMatl = 0" & vbCrLf
      sSql = sSql & "            END" & vbCrLf
      sSql = sSql & "         END" & vbCrLf
      sSql = sSql & "         -- Labour" & vbCrLf
      sSql = sSql & "         SET @totLaborCost = @invTotLabor" & vbCrLf
      sSql = sSql & "         SET @fInvLabor = 1" & vbCrLf
      sSql = sSql & "         IF (@invTotLabor = 0.0000)" & vbCrLf
      sSql = sSql & "         BEGIN" & vbCrLf
      sSql = sSql & "            IF (@lotTotLabor IS NOT NULL)" & vbCrLf
      sSql = sSql & "            BEGIN" & vbCrLf
      sSql = sSql & "               SET @totLaborCost = @lotTotLabor" & vbCrLf
      sSql = sSql & "               SET @fInvLabor = 0" & vbCrLf
      sSql = sSql & "            END" & vbCrLf
      sSql = sSql & "         END" & vbCrLf
      sSql = sSql & "         -- Exp" & vbCrLf
      sSql = sSql & "         SET @totExpCost = @invTotExp" & vbCrLf
      sSql = sSql & "         SET @fInvExp = 1" & vbCrLf
      sSql = sSql & "         IF (@invTotExp = 0.0000)" & vbCrLf
      sSql = sSql & "         BEGIN" & vbCrLf
      sSql = sSql & "            IF (@lotTotExp IS NOT NULL)" & vbCrLf
      sSql = sSql & "            BEGIN" & vbCrLf
      sSql = sSql & "               SET @totExpCost = @lotTotExp" & vbCrLf
      sSql = sSql & "               SET @fInvExp = 0" & vbCrLf
      sSql = sSql & "            END" & vbCrLf
      sSql = sSql & "         END" & vbCrLf
      sSql = sSql & "         -- OH" & vbCrLf
      sSql = sSql & "         SET @totOHCost = @invTotOH" & vbCrLf
      sSql = sSql & "         SET @fInvOH = 1" & vbCrLf
      sSql = sSql & "         IF (@invTotOH = 0.0000)" & vbCrLf
      sSql = sSql & "         BEGIN" & vbCrLf
      sSql = sSql & "            IF (@lotTotOH IS NOT NULL)" & vbCrLf
      sSql = sSql & "            BEGIN" & vbCrLf
      sSql = sSql & "               SET @totOHCost = @lotTotOH" & vbCrLf
      sSql = sSql & "               SET @fInvOH = 0" & vbCrLf
      sSql = sSql & "            END" & vbCrLf
      sSql = sSql & "         END" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "         -- Insert to the temp table" & vbCrLf
      sSql = sSql & "         INSERT INTO #tempINVReport (PARTNUM, PADESC, PAEXTDESC, INADATE," & vbCrLf
      sSql = sSql & "               INAQTY,INAMT, TOTMATL,TOTLABOR,TOTEXP,TOTOH, CREDITACCT, DEBITACCT," & vbCrLf
      sSql = sSql & "               flgMatl, flgLabor, flgExp, flgOH)" & vbCrLf
      sSql = sSql & "         VALUES (@partNum, @partDesc, @partExDesc, @actualDt,@qty," & vbCrLf
      sSql = sSql & "               @invAmt,@totMatlCost,@totLaborCost,@totExpCost,@totOHCost," & vbCrLf
      sSql = sSql & "               @creditAcc,@debitAcc,@fInvMatl,@fInvLabor,@fInvExp,@fInvOH)" & vbCrLf
      sSql = sSql & "      END" & vbCrLf
      sSql = sSql & "      " & vbCrLf
      sSql = sSql & "      FETCH NEXT FROM curInvRpt INTO @partRef, @actualDt, @qty, " & vbCrLf
      sSql = sSql & "               @invAmt, @invTotMatl, @invTotLabor, " & vbCrLf
      sSql = sSql & "               @invTotExp,@invTotOH, @creditAcc, @debitAcc, " & vbCrLf
      sSql = sSql & "               @lotTotMatl, @lotTotLabor, @lotTotExp, @lotTotOH, " & vbCrLf
      sSql = sSql & "               @partNum, @partDesc, @partExDesc" & vbCrLf
      sSql = sSql & "   END" & vbCrLf
      sSql = sSql & "         " & vbCrLf
      sSql = sSql & "   CLOSE curInvRpt   --// close the cursor" & vbCrLf
      sSql = sSql & "   DEALLOCATE curInvRpt    " & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   -- select data for the report" & vbCrLf
      sSql = sSql & "   SELECT PARTNUM, PADESC, PAEXTDESC, INADATE," & vbCrLf
      sSql = sSql & "         INAQTY,INAMT, TOTMATL,TOTLABOR,TOTEXP,TOTOH, " & vbCrLf
      sSql = sSql & "         CREDITACCT, DEBITACCT,flgMatl, flgLabor, " & vbCrLf
      sSql = sSql & "         flgExp, flgOH" & vbCrLf
      sSql = sSql & "      FROM #tempINVReport" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- drop the temp table" & vbCrLf
      sSql = sSql & "DROP table #tempINVReport" & vbCrLf
      sSql = sSql & "END"
      Execute True, sSql
      
      'set version
      Execute False, "update Version set Version = " & newver
   End If

   newver = 57
   If ver < newver Then
      ver = newver

      If Not ColumnExists("ComnTable", "COLUSELOGO") Then
         Execute True, "ALTER TABLE ComnTable ADD COLUSELOGO Integer NOT NULL DEFAULT 0"
         'Execute True, sSql
      End If

      Execute False, "DROP PROCEDURE Qry_ImageData"
      sSql = "CREATE PROCEDURE Qry_ImageData" & vbCrLf _
         & "(@RowID  int=1)" & vbCrLf _
         & "AS" & vbCrLf _
         & "SELECT IMageRecord, ImageStored FROM BitImage" & vbCrLf _
         & "   WHERE IMageRecord = @RowID" & vbCrLf
      Execute True, sSql

      'set version
      Execute False, "update Version set Version = " & newver
   End If

'GET "PROCEDURE TOO LARGE ERROR BEYOND THIS POINT, SO CALL A SEPARATE SUB FOR ADDITIONAL DB UPDATES
   UpdateDatabase2

' 2/20/2009 Added scripts and procedures for next release
   UpdateDatabase3
   
' MM 7/24/2009 database datatyp and SP fixes.
   UpdateDatabase4

' MM 9/2/2009 database Hide module buttons.
   UpdateDatabase5

' MM 11/15/2009 database Hide module buttons.
   UpdateDatabase6

' MM 1/24/2010 Vendor evaluation logic
   UpdateDatabase7

' MM 3/28/2010
   UpdateDatabase8

' MM 5/6/2010
   UpdateDatabase9

' MM 6/27/2010
   UpdateDatabase10

UpdateDatabase11

   'restore normal timeout
   RdoCon.QueryTimeout = timeout
   
   'record update
   'don't do this if running in VB IDE and no update was required
   If UpdateReqd Then
      sSql = "insert Updates" & vbCrLf _
         & "(AppDirectory,OldRelease,NewRelease,UserInitials," & vbCrLf _
         & "OldDbVersion, NewDbVersion,OldDbType,NewDbType)" & vbCrLf _
         & "values('" & App.Path & "'," & oldRelease & "," & newRelease & ",'" & sInitials & "'," & vbCrLf _
         & OldDbVersion & "," & DB_VERSION & ",'" & oldType & "','" & NewType & "')"
      Execute True, sSql
      
      'update test/live indicator if necessary
      If oldType <> NewType Then
         sSql = "update Version set TestDatabase = " & IIf(NewType = "Live", 0, 1)
         Execute True, sSql
      End If
   End If
   
   Unload SysMsgBox
   MouseCursor ccDefault

End Sub

Private Sub UpdateDatabase2()
   'continuation of UpdateDatabase
   'required to avoid "procedure too large" error
   
   newver = 58
   If ver < newver Then
      ver = newver

      'Execute False, "DROP INDEX IX_InvaTable_INGLDATE"
      DropIndex "InvaTable", "IX_InvaTable_INGLDATE"
      sSql = "CREATE NONCLUSTERED INDEX IX_InvaTable_INGLDATE" & vbCrLf _
         & "ON InvaTable (INGLDATE,INADATE,INNO)"
      Execute True, sSql

      'Execute False, "DROP INDEX IX_InvaTable_INGLJOURNAL"
      DropIndex "InvaTable", "IX_InvaTable_INGLJOURNAL"
      sSql = "CREATE NONCLUSTERED INDEX IX_InvaTable_INGLJOURNAL" & vbCrLf _
         & "ON InvaTable (INGLJOURNAL,INGLDATE,INNO)"
      Execute True, sSql
 
      'Execute False, "DROP INDEX IX_InvaTable_INTYPE_INPONUMBER"
      DropIndex "InvaTable", "INDEX IX_InvaTable_INTYPE_INPONUMBER"
      sSql = "CREATE NONCLUSTERED INDEX IX_InvaTable_INTYPE_INPONUMBER" & vbCrLf _
         & "ON InvaTable (INTYPE,INPONUMBER,INPORELEASE,INPOITEM,INPOREV)"
      Execute True, sSql

      ' close all rows where there is a closed IJ" & vbCrLf
      sSql = "update InvaTable" & vbCrLf
      sSql = sSql & "set INGLDATE = MJCLOSED," & vbCrLf
      sSql = sSql & "INGLJOURNAL = MJGLJRNL," & vbCrLf
      sSql = sSql & "INGLPOSTED = 1" & vbCrLf
      sSql = sSql & "from InvaTable ia" & vbCrLf
      sSql = sSql & "join JrhdTable on INADATE between MJSTART and MJEND" & vbCrLf
      sSql = sSql & "and MJCLOSED is not null" & vbCrLf
      sSql = sSql & "and INGLDATE is null" & vbCrLf
      sSql = sSql & "and MJTYPE = 'IJ'"
      Execute True, sSql

      ' re-open rows where there is an open IJ" & vbCrLf
      sSql = "update InvaTable" & vbCrLf
      sSql = sSql & "set INGLDATE = null," & vbCrLf
      sSql = sSql & "INGLJOURNAL = ''," & vbCrLf
      sSql = sSql & "INGLPOSTED = 0" & vbCrLf
      sSql = sSql & "from InvaTable ia" & vbCrLf
      sSql = sSql & "join JrhdTable on INADATE between MJSTART and MJEND" & vbCrLf
      sSql = sSql & "and MJCLOSED is null" & vbCrLf
      sSql = sSql & "and INGLDATE is not null" & vbCrLf
      sSql = sSql & "and MJTYPE = 'IJ'"
      Execute True, sSql
      
      Execute False, "drop procedure UpdatePackingSlipCosts"
      sSql = "create procedure UpdatePackingSlipCosts" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & " @PackingSlip varchar(8)," & vbCrLf
      sSql = sSql & " @UpdateIaEvenIfJournalClosed bit    -- = 1 to update items for closed journals" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "update InvaTable" & vbCrLf
      sSql = sSql & "set INAMT = LOTUNITCOST," & vbCrLf
      sSql = sSql & "INTOTMATL = cast ( abs( INAQTY ) * LOTTOTMATL / LOTORIGINALQTY as decimal(12,4))," & vbCrLf
      sSql = sSql & "INTOTLABOR = cast ( abs( INAQTY ) * LOTTOTLABOR / LOTORIGINALQTY as decimal(12,4))," & vbCrLf
      sSql = sSql & "INTOTEXP = cast ( abs( INAQTY ) * LOTTOTEXP / LOTORIGINALQTY as decimal(12,4)), " & vbCrLf
      sSql = sSql & "INTOTOH = cast ( abs( INAQTY ) * LOTTOTOH / LOTORIGINALQTY as decimal(12,4))," & vbCrLf
      sSql = sSql & "INTOTHRS = cast ( abs( INAQTY ) * LOTTOTHRS / LOTORIGINALQTY as decimal(12,4))" & vbCrLf
      sSql = sSql & "from LoitTable " & vbCrLf
      sSql = sSql & "join LohdTable ON LOINUMBER = LOTNUMBER" & vbCrLf
      sSql = sSql & "join InvaTable ia2 ON INNUMBER = LOIACTIVITY" & vbCrLf
      sSql = sSql & "where ia2.INPSNUMBER = @PackingSlip" & vbCrLf
      sSql = sSql & "and LOTORIGINALQTY <> 0" & vbCrLf
      sSql = sSql & "and (@UpdateIaEvenIfJournalClosed = 1 or ia2.INGLPOSTED = 0)"
      Execute True, sSql
      
      'set version
      Execute False, "update Version set Version = " & newver
   End If


   newver = 59
   If ver < newver Then
      ver = newver
      
      'indices to speed up LoitInvaMatchup.sql maintenance script
      DropIndex "InvaTable", "IX_InvaTable_INNUMBER_INTYPE"
      Execute True, "create index IX_InvaTable_INNUMBER_INTYPE on InvaTable(INNUMBER,INTYPE)"
      
      DropIndex "LoitTable", "IX_LoitTable_LOIADATE"
      Execute True, "create index IX_LoitTable_LOIADATE ON LoitTable(LOIADATE)"
      
      DropIndex "InvaTable", "IX_InvaTable_INADATE"
      Execute True, "create index IX_InvaTable_INADATE ON InvaTable(INADATE)"

      DropIndex "InvaTable", "IX_InvaTable_INPART_INAQTY_INAMT"
      Execute True, "CREATE INDEX IX_InvaTable_INPART_INAQTY_INAMT ON InvaTable (INPART,INAQTY,INAMT)"
      
      'set version
      Execute False, "update Version set Version = " & newver
   End If

   newver = 60
   If ver < newver Then
      ver = newver
      
      'set DCTYPE for AP check debits and credits
      sSql = "Update JritTable" & vbCrLf _
         & "Set DCTYPE = DCREF + 10" & vbCrLf _
         & "where DCTYPE = 0 and DCREF IN (1,2,3)" & vbCrLf _
         & "and DCVENDOR <> '' AND DCCHECKNO <> ''"
     Execute True, sSql
     
      'sp to support Exploded Used On Report
      Execute False, "drop procedure ExplodedUsedOn"
      sSql = "create procedure ExplodedUsedOn" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   @PartRef varchar(30),      -- explode places this part is used" & vbCrLf
      sSql = sSql & "   @BomRev varchar(4)   " & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "as" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "/* test" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "To find a good test case:" & vbCrLf
      sSql = sSql & "select bm3.BMPARTREF as TestPart, bm3.BMREV as TestRev from BmplTable bm1" & vbCrLf
      sSql = sSql & "join BmplTable bm2 on bm1.BMPARTREF = bm2.BMASSYPART and bm1.BMREV = bm2.BMREV" & vbCrLf
      sSql = sSql & "join BmplTable bm3 on bm2.BMPARTREF = bm3.BMASSYPART and bm2.BMREV = bm3.BMREV" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "Sample calls:" & vbCrLf
      sSql = sSql & "ExplodedUsedOn '104256', 'A'                   " & vbCrLf
      sSql = sSql & "ExplodedUsedOn '04A08204D014BM', ''                " & vbCrLf
      sSql = sSql & "*/" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "create table #temp" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   UsedOnID int identity," & vbCrLf
      sSql = sSql & "   UsedOnLevel int," & vbCrLf
      sSql = sSql & "   IndentedLevel varchar(10)," & vbCrLf
      sSql = sSql & "   UsedOnPartRef varchar(30)," & vbCrLf
      sSql = sSql & "   UsedOnPartNum varchar(30)," & vbCrLf
      sSql = sSql & "   UsedOnQty decimal(12,3)," & vbCrLf
      sSql = sSql & "   UsedOnUnits varchar(2)," & vbCrLf
      sSql = sSql & "   ExplodedQty decimal(12,3)," & vbCrLf
      sSql = sSql & "   ExplodedUnits varchar(2)," & vbCrLf
      sSql = sSql & "   UsedOnConversion decimal(12,3)," & vbCrLf
      sSql = sSql & "   UsedOnPartType int," & vbCrLf
      sSql = sSql & "   ChildPartRef varchar(30)," & vbCrLf
      sSql = sSql & "   SortKey varchar(80)" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "declare @level int" & vbCrLf
      sSql = sSql & "set @level = 0" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "insert #temp (UsedOnLevel, IndentedLevel, UsedOnPartRef, UsedOnPartNum, UsedOnQty, UsedOnUnits," & vbCrLf
      sSql = sSql & "   ExplodedQty, ExplodedUnits, UsedOnConversion, UsedonPartType, SortKey)" & vbCrLf
      sSql = sSql & "   select @level, '0         '," & vbCrLf
      sSql = sSql & "   PARTREF, PARTNUM, 1, PAUNITS, 1, PAUNITS, 0, PALEVEL, '00000'" & vbCrLf
      sSql = sSql & "   from PartTable where PARTREF = @PartRef" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "-- keep exploding until no more" & vbCrLf
      sSql = sSql & "declare @LevelPrefix varchar(10)" & vbCrLf
      sSql = sSql & "set @LevelPrefix = '>'" & vbCrLf
      sSql = sSql & "declare @ct int" & vbCrLf
      sSql = sSql & "set @ct = 1" & vbCrLf
      sSql = sSql & "while @ct > 0 and @level < 10" & vbCrLf
      sSql = sSql & "begin" & vbCrLf
      sSql = sSql & "   print 'level ' + cast(@level as varchar(10)) + ' rows: ' + cast(@ct as varchar(10))" & vbCrLf
      sSql = sSql & "   insert #temp(UsedOnLevel, IndentedLevel, UsedOnPartRef, UsedOnPartNum, UsedOnQty, UsedOnUnits, " & vbCrLf
      sSql = sSql & "   ExplodedQty, ExplodedUnits," & vbCrLf
      sSql = sSql & "   UsedOnConversion, UsedOnPartType, ChildPartRef, SortKey)" & vbCrLf
      sSql = sSql & "   select @level + 1,  @LevelPrefix + CAST(@level + 1 as CHAR(1)) + REPLICATE(' ', 9 - len(@LevelPrefix))," & vbCrLf
      sSql = sSql & "   BMASSYPART, BMASSYPART, BMQTYREQD, BMUNITS," & vbCrLf
      sSql = sSql & "   ExplodedQty * BMQTYREQD / case when BMCONVERSION = 0 then 1 else BMCONVERSION end, ExplodedUnits, " & vbCrLf
      sSql = sSql & "   BMCONVERSION, PALEVEL, BMPARTREF, SortKey" & vbCrLf
      sSql = sSql & "   from BmplTable" & vbCrLf
      sSql = sSql & "   join #temp on BMPARTREF = UsedOnPartRef and BMREV = @BomRev" & vbCrLf
      sSql = sSql & "   and UsedOnLevel = @level" & vbCrLf
      sSql = sSql & "   join PartTable on BMPARTREF = PARTREF" & vbCrLf
      sSql = sSql & "   order by BMPARTREF" & vbCrLf
      sSql = sSql & "   " & vbCrLf
      sSql = sSql & "   set @ct = @@ROWCOUNT" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   update #temp set SortKey = SortKey + + replicate('0',5-len(cast(UsedOnID as varchar(5))))" & vbCrLf
      sSql = sSql & "   + cast(UsedOnID as varchar(5))" & vbCrLf
      sSql = sSql & "   where UsedOnLevel = @level + 1" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "   set @level = @level + 1" & vbCrLf
      sSql = sSql & "   set @LevelPrefix = @LevelPrefix + '>'" & vbCrLf
      sSql = sSql & "end" & vbCrLf
      sSql = sSql & "" & vbCrLf
      sSql = sSql & "select * from #temp order by SortKey" & vbCrLf
      sSql = sSql & "drop table #temp" & vbCrLf
      Execute True, sSql

      'set version
      Execute False, "update Version set Version = " & newver
   End If

   newver = 61
   If ver < newver Then
      ver = newver
      Execute True, "DROP TRIGGER InsertPart"

'alter table RunsTable alter column RUNBUDHRS decimal(12,3)
'alter table RunsTable alter column RUNCHRS decimal(12,3)
'alter table RunsTable alter column RUNSCRAP decimal(12,4)
'alter table RunsTable alter column RUNREWORK decimal(12,4)

      AlterNumericColumn "RunsTable", "RUNBUDHRS", "decimal(12,3)"
      AlterNumericColumn "RunsTable", "RUNCHRS", "decimal(12,3)"
      AlterNumericColumn "RunsTable", "RUNSCRAP", "decimal(12,4)"
      AlterNumericColumn "RunsTable", "RUNREWORK", "decimal(12,4)"

      'set version
      Execute False, "update Version set Version = " & newver
   End If

'   newver = 62
'   If ver < newver Then
'      ver = newver
'      Execute true, sSql
'
'      'set version
'      Execute false, "update Version set Version = " & newver
'   End If

'   newver = 63
'   If ver < newver Then
'      ver = newver
'      Execute true, sSql
'
'      'set version
'      Execute false, "update Version set Version = " & newver
'   End If

'   newver = 64
'   If ver < newver Then
'      ver = newver
'      Execute true, sSql
'
'      'set version
'      Execute false, "update Version set Version = " & newver
'   End If

'   newver = 65
'   If ver < newver Then
'      ver = newver
'      Execute true, sSql
'
'      'set version
'      Execute false, "update Version set Version = " & newver
'   End If

'   newver = 66
'   If ver < newver Then
'      ver = newver
'      Execute true, sSql
'
'      'set version
'      Execute false, "update Version set Version = " & newver
'   End If

'   newver = 67
'   If ver < newver Then
'      ver = newver
'      Execute true, sSql
'
'      'set version
'      Execute false, "update Version set Version = " & newver
'   End If

'   newver = 68
'   If ver < newver Then
'      ver = newver
'      Execute true, sSql
'
'      'set version
'      Execute false, "update Version set Version = " & newver
'   End If

'   newver = 69
'   If ver < newver Then
'      ver = newver
'      Execute true, sSql
'
'      'set version
'      Execute false, "update Version set Version = " & newver
'   End If

'   newver = 70
'   If ver < newver Then
'      ver = newver
'      Execute true, sSql
'
'      'set version
'      Execute false, "update Version set Version = " & newver
'   End If

End Sub

Private Function UpdateDatabase3()
    newver = 62
    If ver < newver Then
        ver = newver
       
       ' Drop the temp table and create the table
        Execute False, "DROP TABLE tempRawMatFinishGoods"
        
        sSql = "CREATE TABLE [tempRawMatFinishGoods]" & vbCrLf _
            & "([LOTNUMBER] [varchar](15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & vbCrLf _
            & "[PARTNUM] [varchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & vbCrLf _
            & "[PADESC] [varchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & vbCrLf _
            & "[PAEXTDESC] [varchar](3072) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & vbCrLf _
            & "[LOTUSERLOTID] [char](40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & vbCrLf _
            & "[ORGINNUMBER] [int] NULL,[RPTINNUMBER] [int] NULL," & vbCrLf _
            & "[CURINNUMBER] [int] NULL,[ACTUALDATE] [smalldatetime] NULL," & vbCrLf _
            & "[RPTDATEQTY] [decimal](12, 4) NULL,[UNITCOST] [decimal](12, 4) NULL," & vbCrLf _
            & "[PASTDCOST] [decimal](12, 4) NULL,[LOTUNITCOST] [decimal](12, 4) NULL," & vbCrLf _
            & "[INAMT] [decimal](12, 4) NULL,[ORGCOST] [decimal](12, 4) NULL," & vbCrLf _
            & "[STDCOST] [decimal](12, 4) NULL,[LSTACOST] [decimal](12, 4) NULL," & vbCrLf _
            & "[RPTCOST] [decimal](12, 4) NULL,[CURCOST] [decimal](12, 4) NULL," & vbCrLf _
            & "[COSTEDDATE] [smalldatetime] NULL,[RPTACCOUNT] [int] NULL," & vbCrLf _
            & "[ORIGINALACC] [int] NULL,[LASTACTVITYACC] [int] NULL," & vbCrLf _
            & "[CURRENTACC] [int] NULL,[PACLASS] [char](4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & vbCrLf _
            & "[PAPRODCODE] [char](6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & vbCrLf _
            & "[flgStdCost] [int] NULL,[flgLdCost] [int] NULL," & vbCrLf _
            & "[flgInvCost] [int] NULL,[flgLdRQErr] [int] NULL," & vbCrLf _
            & "[flgRptAcc] [int] NULL,[flgOrgAcc] [int] NULL," & vbCrLf _
            & "[flgLastAcc] [int] NULL) ON [PRIMARY]"
      
        Execute True, sSql
      
        ' Drop the view and create the view
        Execute False, "DROP VIEW ViewLohdPartTable"
        
        sSql = "CREATE VIEW ViewLohdPartTable" & vbCrLf _
                        & "AS " & vbCrLf _
                    & "SELECT LohdTable.LOTNUMBER, PartTable.PARTREF, PartTable.PADESC, PartTable.PAEXTDESC, LohdTable.LOTUSERLOTID," & vbCrLf _
                              & "LohdTable.LOTADATE, LohdTable.LOTORIGINALQTY, LohdTable.LOTREMAININGQTY, LohdTable.LOTUNITCOST," & vbCrLf _
                              & "LohdTable.LOTDATECOSTED, PartTable.PACLASS, PartTable.PAPRODCODE, PartTable.PASTDCOST, LohdTable.LOTTOTMATL," & vbCrLf _
                              & "LohdTable.LOTTOTLABOR, LohdTable.LOTTOTEXP, LohdTable.LOTTOTOH, PartTable.PAUSEACTUALCOST," & vbCrLf _
                              & "PartTable.PALOTTRACK, PartTable.PATOTOH, PartTable.PALABOR, PartTable.PATOTEXP, PartTable.PATOTMATL," & vbCrLf _
                              & "PartTable.PALEVEL" & vbCrLf _
                        & "FROM LohdTable LEFT OUTER JOIN" & vbCrLf _
                              & "PartTable ON LohdTable.LOTPARTREF = PartTable.PARTREF" & vbCrLf _
                    & "WHERE (PartTable.PALEVEL <= 4)"
        
        Execute True, sSql

        Execute False, "DROP PROCEDURE RawMaterialFinishGoods"
        
        sSql = "CREATE PROCEDURE [dbo].[RawMaterialFinishGoods]" & vbCrLf _
                    & "@ReportDate as varchar(16), @PartClass as Varchar(16)," & vbCrLf _
                    & "@PartCode as varchar(8), @lotHDOnly as int " & vbCrLf _
                & "AS " & vbCrLf _
        & "BEGIN " & vbCrLf _
            & "declare @partRef as varchar(30)" & vbCrLf _
            & "declare @partDesc as varchar(30)" & vbCrLf _
            & "declare @partExDesc as varchar(3072)" & vbCrLf _
            & "declare @rptRemQty as decimal(12,4)" & vbCrLf _
            & "declare @deltaQty as decimal(12,4)" & vbCrLf _
            & "declare @lotRemQty as decimal(12,4)" & vbCrLf _
            & "declare @lotOrgQty as decimal(12,4)" & vbCrLf _
            & "declare @tmpInvQty as decimal(12,4)" & vbCrLf _
            & "declare @tmpLastInvQty as decimal(12,4)" & vbCrLf _
            & "declare @rptInvcost decimal(12,4)" & vbCrLf _
            & "declare @lastInvcost decimal(12,4)" & vbCrLf _
            & "declare @orgInvcost decimal(12,4)" & vbCrLf _
            & "declare @tmpInvCost decimal(12,4)" & vbCrLf _
            & "declare @tmpLastInvCost decimal(12,4)" & vbCrLf _
            & "declare @lastQty as decimal(12,4)" & vbCrLf _
            & "declare @orgQty as decimal(12,4)" & vbCrLf _
            & "declare @rptQty as decimal(12,4)" & vbCrLf _
            
        sSql = sSql & "declare @rptCreditACC varchar(12)" & vbCrLf _
            & "declare @rptDebitACC varchar(12)" & vbCrLf _
            & "declare @rptACC varchar(12)" & vbCrLf _
            & "declare @CurACC varchar(12)" & vbCrLf _
            & "declare @tmpCreditAcc varchar(12)" & vbCrLf _
            & "declare @tmpDebitAcc varchar(12)" & vbCrLf _
            & "declare @tmplastCreditAcc varchar(12)" & vbCrLf _
            & "declare @tmpLastDebitAcc varchar(12)" & vbCrLf _
            & "declare @lastCreditACC varchar(12)" & vbCrLf _
            & "declare @lastDebitACC varchar(12)" & vbCrLf _
            & "declare @lastACC varchar(12)" & vbCrLf _
            & "declare @orgCreditACC varchar(12)" & vbCrLf _
            & "declare @orgDebitACC varchar(12)" & vbCrLf _
            & "declare @orgACC varchar(12)" & vbCrLf _
            & "declare @OrgInvNum as int " & vbCrLf _
            & "declare @rptInvNum as int " & vbCrLf _
            & "declare @LastInvNum as int" & vbCrLf _
            & "declare @tmpInvNum as int" & vbCrLf _
            & "declare @tmpLastInvNum as int" & vbCrLf _
            & "declare @LotNumber varchar(51)" & vbCrLf _
            & "declare @LotUserID varchar(51)" & vbCrLf _
            & "declare @lotAcualDate as smalldatetime" & vbCrLf _
            & "declare @lotCostedDate as smalldatetime" & vbCrLf _
            & "declare @curDate as smalldatetime" & vbCrLf _

            sSql = sSql & "declare @AcualDate as smalldatetime" & vbCrLf _
            & "declare @CostedDate as smalldatetime" & vbCrLf _
            & "declare @tmpINVAdate  as smalldatetime" & vbCrLf _
            & "declare @tmpLastINVAdate  as smalldatetime" & vbCrLf _
            & "declare @unitcost as decimal(12,4)" & vbCrLf _
            & "declare @partStdCost as decimal(12,4)" & vbCrLf _
            & "declare @LotUnitCost as decimal(12,4)" & vbCrLf _
            & "declare @lotTotMatl as decimal(12,4)" & vbCrLf _
            & "declare @lotTotLabor as decimal(12,4)" & vbCrLf _
            & "declare @lotTotExp as decimal(12,4)" & vbCrLf _
            & "declare @lotTotOH as decimal(12,4)" & vbCrLf _
            & "declare @partActCost as int" & vbCrLf _
            & "declare @partLotTrack as int" & vbCrLf _
            & "declare @flgStdCost as int" & vbCrLf _
            & "declare @flgLdCost as int" & vbCrLf _
            & "declare @flgInvCost as int" & vbCrLf _
            & "declare @flgLdRQErr as int" & vbCrLf _
            & "declare @flgOrgAcc as int" & vbCrLf _
            & "declare @flgRptAcc as int" & vbCrLf _
            & "declare @flgLastAcc as int" & vbCrLf _
    & "DELETE FROM tempRawMatFinishGoods" & vbCrLf _

    sSql = sSql & "IF (@PartClass = 'ALL')" & vbCrLf _
                & "BEGIN " & vbCrLf _
                  & "SET @PartClass = ''" & vbCrLf _
                & "End" & vbCrLf _
                & "IF (@PartCode = 'ALL')" & vbCrLf _
                & "BEGIN " & vbCrLf _
                  & "SET @PartCode = ''" & vbCrLf _
                & "End" & vbCrLf _
                    & "DECLARE curLotHd CURSOR LOCAL" & vbCrLf _
                     & "FOR" & vbCrLf _
                       & "SELECT LOTNUMBER, LOTUSERLOTID, PARTREF, PADESC, PAEXTDESC," & vbCrLf _
                         & "LOTADATE,LOTORIGINALQTY, LOTREMAININGQTY, LOTUNITCOST, LOTDATECOSTED," & vbCrLf _
                         & "PACLASS , PAPRODCODE, PASTDCOST, PAUSEACTUALCOST, PALOTTRACK" & vbCrLf _
                       & "From ViewLohdPartTable" & vbCrLf _
                       & "WHERE ViewLohdPartTable.LOTADATE  < DATEADD(dd, 1 , @ReportDate)" & vbCrLf _
                          & "AND ViewLohdPartTable.PACLASS LIKE '%' + @PartClass + '%'" & vbCrLf _
                          & "AND ViewLohdPartTable.PAPRODCODE LIKE '%' + @PartCode + '%'" & vbCrLf _
                          & "AND ViewLohdPartTable.PALEVEL <= 4" & vbCrLf _

    sSql = sSql & "OPEN curLotHd " & vbCrLf _
                & "FETCH NEXT FROM curLotHd INTO @LotNumber, @LotUserID, @partRef," & vbCrLf _
                         & "@partDesc, @partExDesc, @lotAcualDate, @lotOrgQty," & vbCrLf _
                               & "@lotRemQty,@LotUnitCost, @lotCostedDate," & vbCrLf _
                               & "@PartClass, @PartCode, @partStdCost, @partActCost, @partLotTrack" & vbCrLf _
                 & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf _
                 & "BEGIN" & vbCrLf _
                     & "IF (@@FETCH_STATUS <> -2)" & vbCrLf _
                     & "BEGIN" & vbCrLf _
                & "SET @flgLdRQErr = 0" & vbCrLf _
            & "IF (@lotRemQty < 0.0000)" & vbCrLf _
            & "BEGIN" & vbCrLf _
                & "SET @lotRemQty = 0.0000" & vbCrLf _
                & "SET @flgLdRQErr  = 1" & vbCrLf _
            & "End" & vbCrLf _
            & "SET @curDate = GETDATE()" & vbCrLf _
            & "SELECT @deltaQty = ISNULL(SUM(LOIQUANTITY), 0.0000)" & vbCrLf _
                    & "From LoitTable" & vbCrLf _
                & "WHERE LOIADATE BETWEEN DATEADD(dd, 1 ,@ReportDate) AND DATEADD(dd, 1 ,@curDate)" & vbCrLf _
                & "AND LOIPARTREF = @partRef" & vbCrLf _
                & "AND LOINUMBER = @LotNumber" & vbCrLf _
            & "SET @rptRemQty = @lotRemQty + (@deltaQty * -1)" & vbCrLf _
            & "IF @rptRemQty < 0.0000" & vbCrLf _
                & "SET @rptRemQty = @rptRemQty * -1" & vbCrLf _

sSql = sSql & "SET @flgStdCost = 0" & vbCrLf _
            & "SET @flgLdCost = 0" & vbCrLf _
            & "SET @flgInvCost = 0" & vbCrLf _
            & "SET @flgOrgAcc = 0" & vbCrLf _
            & "SET @flgRptAcc = 0" & vbCrLf _
            & "SET @flgLastAcc = 0" & vbCrLf _
            & "DECLARE curInv CURSOR" & vbCrLf _
            & "LOCAL" & vbCrLf _
            & "Scroll" & vbCrLf _
            & "FOR" & vbCrLf _
                & "SELECT INNUMBER, INAMT, INAQTY, INCREDITACCT, INDEBITACCT, INADATE" & vbCrLf _
                    & "From InvaTable, LoitTable" & vbCrLf _
                & "WHERE InvaTable.INPART = @partRef" & vbCrLf _
                    & "AND LoitTable.LOINUMBER = @LotNumber" & vbCrLf _
                    & "AND InvaTable.INPART = LoitTable.LOIPARTREF" & vbCrLf _
                    & "AND InvaTable.INNUMBER = LoitTable.LOIACTIVITY" & vbCrLf _
                & "ORDER BY INADATE ASC" & vbCrLf _
            & "OPEN curInv" & vbCrLf _

sSql = sSql & "FETCH FIRST FROM curInv INTO @tmpInvNum, @tmpInvCost, @tmpInvQty," & vbCrLf _
                            & "@tmpCreditAcc, @tmpDebitAcc, @tmpINVAdate" & vbCrLf _
            & "IF (@@FETCH_STATUS <> -1)" & vbCrLf _
            & "BEGIN" & vbCrLf _
                & "IF (@@FETCH_STATUS <> -2)" & vbCrLf _
                & "BEGIN" & vbCrLf _
                    & "SET @orgInvcost = @tmpInvCost" & vbCrLf _
                    & "SET @orgQty = @tmpInvQty" & vbCrLf _
                    & "SET @orgCreditACC = @tmpCreditAcc" & vbCrLf _
                    & "SET @orgDebitACC = @tmpDebitAcc" & vbCrLf _
                    & "SET @OrgInvNum = @tmpInvNum" & vbCrLf _
                & "End" & vbCrLf _
                & "FETCH LAST FROM curInv INTO @tmpLastInvNum, @tmpLastInvCost, @tmpLastInvQty," & vbCrLf _
                                & "@tmplastCreditAcc, @tmpLastDebitAcc, @tmpLastINVAdate" & vbCrLf _
                & "IF (@@FETCH_STATUS <> -2)" & vbCrLf _
                & "BEGIN" & vbCrLf _
                    & "SET @lastInvcost = @tmpLastInvCost" & vbCrLf _
                    & "SET @lastQty = @tmpLastInvQty" & vbCrLf _
                    & "SET @lastCreditACC = @tmplastCreditAcc" & vbCrLf _
                    & "SET @lastDebitACC = @tmpLastDebitAcc" & vbCrLf _
                    & "SET @LastInvNum = @tmpLastInvNum" & vbCrLf _
                    
        sSql = sSql & "IF @tmpLastINVAdate > DATEADD(dd, 1, @ReportDate)" & vbCrLf _
                    & "BEGIN" & vbCrLf _
                        & "SELECT TOP 1 @rptInvNum = INNUMBER, @rptInvcost = INAMT," & vbCrLf _
                                & "@rptQty = INAQTY, @rptCreditACC = INCREDITACCT," & vbCrLf _
                                & "@rptDebitACC = INDEBITACCT" & vbCrLf _
                            & "From InvaTable, LoitTable" & vbCrLf _
                        & "WHERE INADATE < DATEADD(dd, 1, @ReportDate)" & vbCrLf _
                                & "AND InvaTable.INPART = @partRef" & vbCrLf _
                                & "AND LoitTable.LOINUMBER = @LotNumber" & vbCrLf _
                                & "AND InvaTable.INPART = LoitTable.LOIPARTREF" & vbCrLf _
                                & "AND InvaTable.INNUMBER = LoitTable.LOIACTIVITY" & vbCrLf _
                        & "ORDER BY INADATE DESC" & vbCrLf _
                    & "End" & vbCrLf _
                    & "Else" & vbCrLf _
                    & "BEGIN" & vbCrLf _
                        & "SET @rptInvNum = @tmpLastInvNum" & vbCrLf _
                        & "SET @rptInvcost = @tmpLastInvCost" & vbCrLf _
                        & "SET @rptQty = @tmpLastInvQty" & vbCrLf _
                        & "SET @rptCreditACC = @tmplastCreditAcc" & vbCrLf _
                        & "SET @rptDebitACC = @tmpLastDebitAcc" & vbCrLf _
                    & "End" & vbCrLf _
                & "End" & vbCrLf _
            & "End" & vbCrLf _
            & "CLOSE curInv   --// close the cursor" & vbCrLf _

            sSql = sSql & "DEALLOCATE curInv" & vbCrLf _
            & "IF (@partActCost = 0)" & vbCrLf _
            & "BEGIN" & vbCrLf _
                & "SET @unitcost = @partStdCost" & vbCrLf _
                & "SET @flgStdCost = 1" & vbCrLf _
            & "End" & vbCrLf _
            & "Else" & vbCrLf _
            & "BEGIN" & vbCrLf _
                & "IF @lotHDOnly = 1" & vbCrLf _
                & "BEGIN" & vbCrLf _
                    & "SET @unitcost = @LotUnitCost" & vbCrLf _
                    & "SET @flgLdCost = 1" & vbCrLf _
                & "End" & vbCrLf _
                & "Else" & vbCrLf _
                & "BEGIN" & vbCrLf _
                    & "IF @lotCostedDate < DATEADD(dd, 1, @ReportDate)" & vbCrLf _
                    & "BEGIN" & vbCrLf _
                        & "SET @unitcost = @LotUnitCost" & vbCrLf _
                        & "SET @flgLdCost = 1" & vbCrLf _
                    & "End" & vbCrLf _
                    & "Else" & vbCrLf _
                    & "BEGIN" & vbCrLf _
                        & "SET @unitcost = @rptInvcost" & vbCrLf _
                        & "SET @flgInvCost = 1" & vbCrLf _

        sSql = sSql & "End" & vbCrLf _
                & "END --LotHD only" & vbCrLf _
            & "END --Part Cost" & vbCrLf _
            & "SELECT @CurACC = dbo.fnGetPartInvAccount(@partRef)" & vbCrLf _
            & "-- Lastest < report account#" & vbCrLf _
            & "IF @rptQty >= 0.0000" & vbCrLf _
                & "SET @rptACC = @rptDebitACC" & vbCrLf _
            & "Else" & vbCrLf _
                & "SET @rptACC = @rptCreditACC" & vbCrLf _
            & "IF ((@rptACC = '') OR (@rptACC = NULL))" & vbCrLf _
            & "BEGIN" & vbCrLf _
                & "SET @rptACC = @CurACC" & vbCrLf _
                & "SET @flgRptAcc = 1" & vbCrLf _
            & "End" & vbCrLf _
            & "-- last record" & vbCrLf _
            & "IF @lastQty >= 0.0000" & vbCrLf _
                & "SET @lastACC = @lastDebitACC" & vbCrLf _
            & "Else" & vbCrLf _
                & "SET @lastACC = @lastCreditACC" & vbCrLf _
            & "IF ((@lastACC = '') OR (@lastACC = NULL))" & vbCrLf _
            & "BEGIN" & vbCrLf _
                & "SET @lastACC = @CurACC" & vbCrLf _
                & "SET @flgLastAcc = 1" & vbCrLf _
            & "End" & vbCrLf _

            sSql = sSql & "-- Lastest < report account#" & vbCrLf _
            & "IF @orgQty >= 0.0000" & vbCrLf _
                & "SET @orgACC = @orgDebitACC" & vbCrLf _
            & "Else" & vbCrLf _
                & "SET @orgACC = @orgCreditACC" & vbCrLf _
            & "IF ((@orgACC = '') OR (@orgACC = NULL))" & vbCrLf _
            & "BEGIN" & vbCrLf _
                & "SET @orgACC = @CurACC" & vbCrLf _
                & "SET @flgOrgAcc = 1" & vbCrLf _
            & "End" & vbCrLf _
         & "-- Insert to the temp table" & vbCrLf _

         sSql = sSql & "INSERT INTO tempRawMatFinishGoods" & vbCrLf _
               & "(PARTNUM, PADESC, PAEXTDESC, LOTNUMBER, LOTUSERLOTID," & vbCrLf _
               & "ORGINNUMBER, RPTINNUMBER, CURINNUMBER, ACTUALDATE,RPTDATEQTY, COSTEDDATE," & vbCrLf _
               & "UNITCOST,PASTDCOST, LOTUNITCOST, ORGCOST, STDCOST," & vbCrLf _
               & "LSTACOST, RPTCOST, CURCOST, RPTACCOUNT,ORIGINALACC," & vbCrLf _
               & "LASTACTVITYACC, CURRENTACC, PACLASS,PAPRODCODE," & vbCrLf _
               & "flgStdCost, flgLdCost, flgInvCost, flgLdRQErr," & vbCrLf _
               & "flgRptAcc, flgOrgAcc, flgLastAcc)" & vbCrLf _
         & "VALUES (@partRef, @partDesc, @partExDesc, @LotNumber,@LotUserID, @OrgInvNum," & vbCrLf _
               & "@rptInvNum, @LastInvNum, @lotAcualDate,@rptRemQty,@lotCostedDate," & vbCrLf _
               & "@unitcost,@partStdCost, @LotUnitCost, @orgInvcost, @partStdCost," & vbCrLf _
                & "@lastInvcost,@rptInvcost, @LotUnitCost, @rptACC, @orgACC, @lastACC," & vbCrLf _
               & "@CurACC, @PartClass,@PartCode,@flgStdCost, @flgLdCost, @flgInvCost," & vbCrLf _
               & "@flgLdRQErr, @flgRptAcc, @flgOrgAcc, @flgLastAcc)" & vbCrLf _


            sSql = sSql & "SET @rptRemQty = NULL" & vbCrLf _
            & "SET @deltaQty = NULL" & vbCrLf _
            & "SET @lotRemQty = NULL" & vbCrLf _
            & "SET @lotOrgQty = NULL" & vbCrLf _
            & "SET @tmpInvQty = NULL" & vbCrLf _
            & "SET @tmpLastInvQty  = NULL" & vbCrLf _
            & "SET @rptInvcost = NULL" & vbCrLf _
            & "SET @lastInvcost = NULL" & vbCrLf _
            & "SET @orgInvcost = NULL" & vbCrLf _
            & "SET @tmpInvCost = NULL" & vbCrLf _
            & "SET @unitcost = NULL" & vbCrLf _
            & "SET @tmpLastInvCost = NULL" & vbCrLf _
            & "SET @lastQty = NULL" & vbCrLf _
            & "SET @orgQty = NULL" & vbCrLf _
            & "SET @rptQty  = NULL" & vbCrLf _
            & "SET @rptCreditACC = NULL" & vbCrLf _
            & "SET @rptDebitACC  = NULL" & vbCrLf _
            & "SET @rptACC  = NULL" & vbCrLf _
            & "SET @tmpCreditAcc  = NULL" & vbCrLf _
            & "SET @tmpDebitAcc = NULL" & vbCrLf _
            & "SET @tmplastCreditAcc = NULL" & vbCrLf _
            & "SET @tmpLastDebitAcc = NULL" & vbCrLf _

            sSql = sSql & "SET @lastCreditACC = NULL" & vbCrLf _
            & "SET @lastDebitACC = NULL" & vbCrLf _
            & "SET @lastACC  = NULL" & vbCrLf _
            & "SET @orgCreditACC  = NULL" & vbCrLf _
            & "SET @orgDebitACC  = NULL" & vbCrLf _
            & "SET @orgACC  = NULL" & vbCrLf _
            & "SET @OrgInvNum = NULL" & vbCrLf _
            & "SET @rptInvNum = NULL" & vbCrLf _
            & "SET @LastInvNum = NULL" & vbCrLf _
                & "End" & vbCrLf _
               & "FETCH NEXT FROM curLotHd INTO @LotNumber, @LotUserID, @partRef," & vbCrLf _
                              & "@partDesc, @partExDesc, @lotAcualDate, @lotOrgQty," & vbCrLf _
                              & "@lotRemQty,@LotUnitCost, @lotCostedDate," & vbCrLf _
                              & "@PartClass, @PartCode, @partStdCost, @partActCost, @partLotTrack" & vbCrLf _
           & "End" & vbCrLf _
           & "CLOSE curLotHd   --// close the cursor" & vbCrLf _
           & "DEALLOCATE curLotHd" & vbCrLf _
        & "End"
        
        ' Execute the sql to create the store procedure
        Execute True, sSql
      
        ' Drop the view and create the view
        Execute False, "DROP VIEW viewPohdPoit"

        sSql = "CREATE VIEW viewPohdPoit" & vbCrLf _
                          & " AS" & vbCrLf _
                    & "SELECT PoitTable.PIITEM, PoitTable.PIREV, PohdTable.*, PoitTable.PILOT" & vbCrLf _
                            & "FROM PohdTable INNER JOIN" & vbCrLf _
                     & "PoitTable ON PohdTable.PONUMBER = PoitTable.PINUMBER AND PohdTable.PORELEASE = PoitTable.PIRELEASE"
        ' Execute the sql to create the store procedure
        Execute True, sSql

        Execute False, "DROP VIEW viewOpenAPTerms"
        
        sSql = "CREATE VIEW [dbo].[viewOpenAPTerms]" & vbCrLf _
                                & " AS" & vbCrLf _
                    & "SELECT dbo.VihdTable.VIVENDOR, dbo.VndrTable.VENICKNAME, dbo.VndrTable.VEBNAME, dbo.VihdTable.VINO, dbo.VihdTable.VIDATE," & vbCrLf _
                              & "dbo.VihdTable.VIDUE AS InvTotal, dbo.VihdTable.VIDUEDATE, dbo.VihdTable.VIFREIGHT, dbo.VihdTable.VITAX, dbo.VihdTable.VIPAY," & vbCrLf _
                              & "ISNULL(ViitTable_1.VITPO, 0) AS VITPO, ISNULL(ViitTable_1.VITPORELEASE, 0) AS VITPORELEASE," & vbCrLf _
                              & "CASE WHEN PONETDAYS <> 0 THEN PODISCOUNT ELSE VEDISCOUNT END AS DiscRate," & vbCrLf _
                              & "CASE WHEN PONETDAYS <> 0 THEN PODDAYS ELSE VEDDAYS END AS DiscDays," & vbCrLf _
                              & "CASE WHEN PONETDAYS <> 0 THEN PONETDAYS ELSE VENETDAYS END AS NetDays, dbo.VihdTable.VIDUE - dbo.VihdTable.VIPAY AS AmountDue," & vbCrLf _
                              & "dbo.VihdTable.VIFREIGHT + dbo.VihdTable.VITAX +" & vbCrLf _
                                  & "(SELECT     CAST(SUM(ROUND(CAST(dbo.ViitTable.VITQTY AS decimal(12, 3)) / (CASE WHEN ISNULL(dbo.viewPohdPoit.PILOT, 0)" & vbCrLf _
                                                           & "<> 0 THEN CAST(dbo.viewPohdPoit.PILOT AS decimal(12, 3)) ELSE 1 END) * CAST(dbo.ViitTable.VITCOST AS decimal(12, 4))" & vbCrLf _
                                                           & "+ CAST(dbo.ViitTable.VITADDERS AS decimal(12, 2)), 2)) AS decimal(12, 2)) AS Expr1" & vbCrLf _
                                    & "FROM          dbo.ViitTable INNER JOIN" & vbCrLf _
                                                           & "dbo.viewPohdPoit ON dbo.ViitTable.VITPO = dbo.viewPohdPoit.PONUMBER AND" & vbCrLf _
                                                           & "dbo.ViitTable.VITPORELEASE = dbo.viewPohdPoit.PORELEASE AND dbo.ViitTable.VITPOITEM = dbo.viewPohdPoit.PIITEM AND" & vbCrLf _
                                                           & "dbo.ViitTable.VITPOITEMREV = dbo.viewPohdPoit.PIREV" & vbCrLf _
                                    & "WHERE      (dbo.ViitTable.VITVENDOR = dbo.VihdTable.VIVENDOR) AND (dbo.ViitTable.VITNO = dbo.VihdTable.VINO)) AS CalcTotal," & vbCrLf _
                              & "ViitTable_1.VITCOST , viewPohdPoit_1.PILOT" & vbCrLf _
                              
                              
                sSql = sSql & "FROM dbo.VihdTable INNER JOIN" & vbCrLf _
                              & "dbo.VndrTable ON dbo.VihdTable.VIVENDOR = dbo.VndrTable.VEREF LEFT OUTER JOIN" & vbCrLf _
                              & "dbo.viewPohdPoit AS viewPohdPoit_1 RIGHT OUTER JOIN" & vbCrLf _
                              & "dbo.ViitTable AS ViitTable_1 ON viewPohdPoit_1.PIREV = ViitTable_1.VITPOITEMREV AND viewPohdPoit_1.PIITEM = ViitTable_1.VITPOITEM AND" & vbCrLf _
                              & "viewPohdPoit_1.PONUMBER = ViitTable_1.VITPO AND viewPohdPoit_1.PORELEASE = ViitTable_1.VITPORELEASE ON" & vbCrLf _
                              & "ViitTable_1.VITNO = dbo.VihdTable.VINO AND ViitTable_1.VITVENDOR = dbo.VihdTable.VIVENDOR AND ViitTable_1.VITITEM =" & vbCrLf _
                                 & " (SELECT     MIN(VITITEM) AS Expr1" & vbCrLf _
                                    & "From dbo.ViitTable" & vbCrLf _
                                    & "WHERE      (VITVENDOR = dbo.VndrTable.VEREF) AND (VITNO = dbo.VihdTable.VINO) AND (VITPO IS NOT NULL) AND (VITPO <> 0))" & vbCrLf _
        & "Where (dbo.VihdTable.VIPIF <> 1) And (ISNULL(ViitTable_1.VITPO, 0) <> 0)" & vbCrLf _
        & "Union" & vbCrLf _
        & "SELECT     VihdTable_1.VIVENDOR, VndrTable_1.VENICKNAME, VndrTable_1.VEBNAME, VihdTable_1.VINO, VihdTable_1.VIDATE, VihdTable_1.VIDUE AS InvTotal," & vbCrLf _
                              & "VihdTable_1.VIDUEDATE, VihdTable_1.VIFREIGHT, VihdTable_1.VITAX, VihdTable_1.VIPAY, ISNULL(ViitTable_1.VITPO, 0) AS VITPO," & vbCrLf _
                              & "ISNULL(ViitTable_1.VITPORELEASE, 0) AS VITPORELEASE, CASE WHEN PONETDAYS <> 0 THEN PODISCOUNT ELSE VEDISCOUNT END AS DiscRate," & vbCrLf _
                              & "CASE WHEN PONETDAYS <> 0 THEN PODDAYS ELSE VEDDAYS END AS DiscDays," & vbCrLf _
                              & "CASE WHEN PONETDAYS <> 0 THEN PONETDAYS ELSE VENETDAYS END AS NetDays, VihdTable_1.VIDUE - VihdTable_1.VIPAY AS AmountDue," & vbCrLf _
                              & "VihdTable_1.VIFREIGHT + VihdTable_1.VITAX +" & vbCrLf _
                                  & "(SELECT     CAST(SUM(ROUND(CAST(VITQTY AS decimal(12, 3)) * CAST(VITCOST AS decimal(12, 4)) + CAST(VITADDERS AS decimal(12, 2)), 2))" & vbCrLf _
                                                           & "AS decimal(12, 2)) AS Expr1" & vbCrLf _
                                    & "FROM          dbo.ViitTable AS ViitTable_2" & vbCrLf _
                                    & "WHERE      (VITVENDOR = VihdTable_1.VIVENDOR) AND (VITNO = VihdTable_1.VINO)) AS CalcTotal, ViitTable_1.VITCOST," & vbCrLf _
                              & "viewPohdPoit_2.PILOT" & vbCrLf _
        
                sSql = sSql & "FROM  dbo.VihdTable AS VihdTable_1 INNER JOIN" & vbCrLf _
                              & "dbo.VndrTable AS VndrTable_1 ON VihdTable_1.VIVENDOR = VndrTable_1.VEREF LEFT OUTER JOIN" & vbCrLf _
                              & "dbo.viewPohdPoit AS viewPohdPoit_2 RIGHT OUTER JOIN" & vbCrLf _
                              & "dbo.ViitTable AS ViitTable_1 ON viewPohdPoit_2.PIREV = ViitTable_1.VITPOITEMREV AND viewPohdPoit_2.PIITEM = ViitTable_1.VITPOITEM AND" & vbCrLf _
                              & "viewPohdPoit_2.PONUMBER = ViitTable_1.VITPO AND viewPohdPoit_2.PORELEASE = ViitTable_1.VITPORELEASE ON" & vbCrLf _
                              & "ViitTable_1.VITNO = VihdTable_1.VINO AND ViitTable_1.VITVENDOR = VihdTable_1.VIVENDOR AND ViitTable_1.VITITEM =" & vbCrLf _
                                  & "(SELECT     MIN(VITITEM) AS Expr1" & vbCrLf _
                                    & "From dbo.ViitTable" & vbCrLf _
                                    & "WHERE      (VITVENDOR = VndrTable_1.VEREF) AND (VITNO = VihdTable_1.VINO) AND (VITPO IS NOT NULL) AND (VITPO <> 0))" & vbCrLf _
        & "Where (VihdTable_1.VIPIF <> 1) And (ISNULL(ViitTable_1.VITPO, 0) = 0)"
        
        ' Execute the sql to create the store procedure
        Execute True, sSql
        
        Execute False, "DROP PROCEDURE UpdateFADOCTable"

        sSql = "CREATE PROCEDURE UpdateFADOCTable" & vbCrLf _
                            & "AS" & vbCrLf _
                        & "BEGIN" & vbCrLf _
                            & "declare @iRcdCnt as int" & vbCrLf _
                            & "declare @FaDocNum as varchar(30)" & vbCrLf _
                            & "declare @FaDocRev as int" & vbCrLf _
                            & "DECLARE curFADocNum CURSOR   FOR" & vbCrLf _
                                & "SELECT DISTINCT FA_DOCNUMBER, FA_DOCREVISION" & vbCrLf _
                                    & "From FadcTable" & vbCrLf _
                            & "OPEN curFADocNum" & vbCrLf _
                            & "FETCH NEXT FROM curFADocNum INTO @FaDocNum, @FaDocRev" & vbCrLf _
                            & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf _
                            & "BEGIN" & vbCrLf _
                                & "IF (@@FETCH_STATUS <> -2)" & vbCrLf _
                                & "BEGIN" & vbCrLf _
                                     & "SELECT @iRcdCnt = (COUNT(FA_DOCNUMBER) + 1)" & vbCrLf _
                                                & "From FadcTable" & vbCrLf _
                                            & "WHERE FA_DOCNUMBER = @FaDocNum AND" & vbCrLf _
                                                & "FA_DOCREVISION = @FaDocRev" & vbCrLf _
                                    
                    sSql = sSql & "WHILE @iRcdCnt < 11" & vbCrLf _
                                    & "BEGIN" & vbCrLf _
                                    & "INSERT INTO FadcTable(FA_DOCNUMBER, FA_DOCREVISION, FA_DOCITEM)" & vbCrLf _
                                        & "VALUES (@FaDocNum, @FaDocRev, @iRcdCnt)" & vbCrLf _
                                    & "SET @iRcdCnt = @iRcdCnt + 1" & vbCrLf _
                                & "End" & vbCrLf _
                            & "End" & vbCrLf _
                            & "FETCH NEXT FROM curFADocNum INTO @FaDocNum, @FaDocRev" & vbCrLf _
                       & "End" & vbCrLf _
                       & "CLOSE curFADocNum   --// close the cursor" & vbCrLf _
                       & "DEALLOCATE curFADocNum" & vbCrLf _
                    & "End"

        ' Execute the sql to create the store procedure
        Execute True, sSql
        
        Execute False, "DROP PROCEDURE Qry_GetMRPSysBuyer"

        sSql = "CREATE PROCEDURE Qry_GetMRPSysBuyer" & vbCrLf _
                        & "(@partref as varchar(30))" & vbCrLf _
                        & "AS" & vbCrLf _
                        & "BEGIN" & vbCrLf _
                            & "declare @PartCodeBuyer as varchar(20)" & vbCrLf _
                            & "declare @PartBuyer as varchar(20)" & vbCrLf _
                            & "SELECT DISTINCT BuycTable.BYREF PRODCODEBUYER," & vbCrLf _
                                & "BuypTable.BYREF PABUYER" & vbCrLf _
                            & "FROM  PartTable LEFT OUTER JOIN" & vbCrLf _
                               & "BuypTable ON PartTable.PARTREF = BuypTable.BYPARTNUMBER LEFT OUTER JOIN" & vbCrLf _
                                & "BuycTable ON PartTable.PAPRODCODE = BuycTable.BYPRODCODE" & vbCrLf _
                            & "WHERE (PartTable.PARTREF = @partref)" & vbCrLf _
                        & "End"
        ' Execute the sql to create the store procedure
        Execute True, sSql

        sSql = "--// Truncate the BuypTable" & vbCrLf _
                & "DELETE FROM BuypTable" & vbCrLf _
                & "--// Query to polulate the BuyerPartnumber association table" & vbCrLf _
                & "INSERT INTO BuypTable" & vbCrLf _
                & "SELECT DISTINCT PABUYER, PARTREF" & vbCrLf _
                    & "From PartTable" & vbCrLf _
                & "WHERE PABUYER <> '' AND PABUYER IN (SELECT DISTINCT BYREF FROM BuyrTable)" & vbCrLf _
                & "--// Truncte the BuycTable" & vbCrLf _
                & "DELETE FROM BuycTable" & vbCrLf _
                & "--// Query to polulate the BuyerPartCode association" & vbCrLf _
                & "--// table." & vbCrLf _
                & "INSERT INTO BuycTable" & vbCrLf _
                    & "SELECT PCBUYERREF, PCCODE FROM dbo.PcodTable" & vbCrLf _
                  & "WHERE PCBUYERREF <> ''  AND PCBUYERREF IN (SELECT DISTINCT BYREF FROM BuyrTable)"

        ' Execute the sql to create the store procedure
        Execute True, sSql
        
        
        Execute False, "DROP TABLE [dbo].StcodeTable"
        ' Execute the sql to create Internal StatusCode
        
        sSql = "CREATE TABLE [dbo].[StcodeTable]" & vbCrLf _
                & "([STATUS_REF] [char](4)  NOT NULL," & vbCrLf _
            & "[STATUS_CODE] [char](20) NOT NULL) ON [PRIMARY]"
            
        Execute True, sSql
        
        sSql = "ALTER TABLE [dbo].[StcodeTable] WITH NOCHECK ADD" & vbCrLf _
                & "CONSTRAINT [PK_StcodeTable] PRIMARY KEY CLUSTERED" & vbCrLf _
                & "([STATUS_REF] Asc) ON [PRIMARY]"
        
        Execute True, sSql
        
        
        Execute False, "DROP TABLE [dbo].StatCdType"
        ' Execute the sql to create Internal StatusCode Type table
        
        sSql = "CREATE TABLE [dbo].[StatCdType] (" & vbCrLf _
            & "[STATCODE_TYPE_REF] [char] (20) NOT NULL ," & vbCrLf _
            & "[STATCODE_TYPE_NAME] [varchar] (50) NOT NULL ," & vbCrLf _
            & "[STATCODE_TYPE_UNIQUE_KEY] [int] NULL" & vbCrLf _
        & ") ON [PRIMARY]"
        
        Execute True, sSql
        
        sSql = "ALTER TABLE [dbo].[StatCdType] WITH NOCHECK ADD" & vbCrLf _
            & "CONSTRAINT [PK_StatCdType] PRIMARY KEY  CLUSTERED" & vbCrLf _
            & "([STATCODE_TYPE_REF])  ON [PRIMARY]"
        
        Execute True, sSql
        
        ' Insert seed data
        sSql = "INSERT INTO StatCdType(STATCODE_TYPE_REF, STATCODE_TYPE_NAME, STATCODE_TYPE_UNIQUE_KEY) " & _
                    " VALUES('SO', 'Sales Order', '1')"
        Execute True, sSql
        
        sSql = "INSERT INTO StatCdType(STATCODE_TYPE_REF, STATCODE_TYPE_NAME, STATCODE_TYPE_UNIQUE_KEY) " & _
                    " VALUES('SOI', 'Sales Order Item', '3')"
        Execute True, sSql
        
        
        Execute False, "DROP TABLE [dbo].StCmtTable"
        ' Execute the sql to create Internal StatusCode Association table
        
        sSql = "CREATE TABLE [dbo].[StCmtTable]" & vbCrLf _
                & "([STATUS_CMT_REF] [int] NOT NULL," & vbCrLf _
                & "[STATUS_CMT_REF1] [int] NOT NULL," & vbCrLf _
                & "[STATUS_CMT_REF2] [char](20) NOT NULL," & vbCrLf _
                & "[STATUS_REF] [char](4) NOT NULL," & vbCrLf _
                & "[STATCODE_TYPE_REF] [char](20) NOT NULL," & vbCrLf _
                & "[COMMENT] [varchar](1024)  NULL," & vbCrLf _
                & "[STATUS_ORG_USER] [varchar](50)  NOT NULL," & vbCrLf _
                & "[STATUS_ORG_DATE] [datetime] NOT NULL," & vbCrLf _
                & "[STATUS_CUR_USER] [varchar](50)  NOT NULL," & vbCrLf _
                & "[STATUS_CUR_DATE] [datetime] NOT NULL," & vbCrLf _
                & "[STATUS_ACT_STATE] [tinyint] NOT NULL ) ON [PRIMARY]"

        Execute True, sSql


        sSql = "ALTER TABLE [dbo].[StCmtTable] WITH NOCHECK ADD" & vbCrLf _
                    & "CONSTRAINT  [PK_StCmtTable_1] PRIMARY KEY CLUSTERED" & vbCrLf _
                        & "([STATUS_CMT_REF] ASC," & vbCrLf _
                        & "[STATUS_CMT_REF1] ASC," & vbCrLf _
                        & "[STATUS_CMT_REF2] ASC," & vbCrLf _
                        & "[STATUS_REF] ASC," & vbCrLf _
                        & "[STATCODE_TYPE_REF] Asc) ON [PRIMARY]"
        
        Execute True, sSql
      
      
        Execute False, "DROP PROCEDURE BackLogBySchedDate"
        
        sSql = "CREATE PROCEDURE [dbo].[BackLogBySchedDate]" & vbCrLf _
                    & "@CutoffDate as varchar(16), @Customer as varchar(10), " & vbCrLf _
                    & "@PartClass as Varchar(16),@PartCode as varchar(8) " & vbCrLf _
                    & "AS " & vbCrLf _
                    & "BEGIN " & vbCrLf _
                        & "declare @SoType as varchar(1) " & vbCrLf _
                        & "declare @SoText as varchar(6) " & vbCrLf _
                        & "declare @ItSo as int " & vbCrLf _
                        & "declare @ItRev as char(2) " & vbCrLf _
                        & "declare @ItNum as int " & vbCrLf _
                        & "declare @ItQty as decimal(12,4) " & vbCrLf _
                        & "declare @PaLotRemQty as decimal(12,4) " & vbCrLf _
                        & "declare @PartRem as decimal(12,4) " & vbCrLf _
                        & "declare @RunningTot as decimal(12,4) " & vbCrLf _
                        & "declare @ItDollars as decimal(12,4) " & vbCrLf _
                        & "declare @ItSched as smalldatetime " & vbCrLf _
                        & "declare @CusName as varchar(10) " & vbCrLf _
                        & "declare @PartNum as varchar(30) " & vbCrLf _
                        & "declare @CurPartNum as varchar(30) " & vbCrLf _
                        & "declare @PartDesc as varchar(30) " & vbCrLf _
                        & "declare @PartLoc as varchar(4) " & vbCrLf _
                        & "declare @PartExDesc as varchar(3072) " & vbCrLf _
                        & "declare @ItCanceled as tinyint " & vbCrLf _
                        & "declare @ItPSNum as varchar(8) " & vbCrLf _
                        & "declare @ItInvoice as int "
                        
                        
            sSql = sSql & "declare @ItPSShipped as tinyint" & vbCrLf _
                        & "IF (@Customer = 'ALL') " & vbCrLf _
                          & "SET @Customer = '' " & vbCrLf _
                        & "IF (@PartClass = 'ALL') " & vbCrLf _
                          & "SET @PartClass = '' " & vbCrLf _
                        & "IF (@PartCode = 'ALL') " & vbCrLf _
                          & "SET @PartCode = '' " & vbCrLf _
                       & "CREATE TABLE #tempBackLogInfo " & vbCrLf _
                        & "(SOTYPE varchar(1) NULL, " & vbCrLf _
                        & "SOTEXT varchar(6) NULL, " & vbCrLf _
                        & "ITSO Int NULL, " & vbCrLf _
                        & "ITREV char(2) NULL, " & vbCrLf _
                        & "ITNUMBER int NULL, " & vbCrLf _
                        & "ITQTY decimal(12,4) NULL, " & vbCrLf _
                        & "PALOTQTYREMAINING decimal(12,4) NULL, " & vbCrLf _
                        & "RUNQTYTOT decimal(12,4) NULL, " & vbCrLf _
                        & "ITDOLLARS decimal(12,4) NULL, " & vbCrLf _
                        & "ITSCHED smalldatetime NULL, " & vbCrLf _
                        & "CUNICKNAME varchar(10) NULL, " & vbCrLf _
                        & "PARTNUM varchar(30) NULL, " & vbCrLf _
                        & "PADESC varchar(30) NULL, " & vbCrLf _
                        & "PAEXTDESC varchar(3072) NULL, " & vbCrLf _
                        & "PALOCATION varchar(4) NULL, " & vbCrLf _
                        & "ITCANCELED tinyint NULL, " & vbCrLf _
                        & "ITPSNUMBER varchar(8) NULL, "
                        
            sSql = sSql & "ITINVOICE int NULL," & vbCrLf _
                        & "ITPSSHIPPED tinyint NULL)" & vbCrLf _
                       & "DECLARE curbackLog CURSOR   FOR " & vbCrLf _
                        & "SELECT SohdTable.SOTYPE, SohdTable.SOTEXT, " & vbCrLf _
                            & "SoitTable.ITSO, SoitTable.ITREV, SoitTable.ITNUMBER, " & vbCrLf _
                            & "SoitTable.ITQTY, PartTable.PALOTQTYREMAINING, " & vbCrLf _
                            & "SoitTable.ITDOLLARS,SoitTable.ITSCHED, CustTable.CUNICKNAME, " & vbCrLf _
                            & "PartTable.PARTNUM, PartTable.PADESC, PartTable.PAEXTDESC, " & vbCrLf _
                            & "PartTable.PALOCATION, SoitTable.ITCANCELED, " & vbCrLf _
                            & "SoitTable.ITPSNUMBER , SoitTable.ITINVOICE, SoitTable.ITPSSHIPPED " & vbCrLf _
                        & "From SohdTable, SoitTable, CustTable, PartTable " & vbCrLf _
                        & "WHERE SohdTable.SOCUST = CustTable.CUREF AND " & vbCrLf _
                            & "SohdTable.SONUMBER =SoitTable.ITSO AND " & vbCrLf _
                            & "SoitTable.ITPART=PartTable.PARTREF AND " & vbCrLf _
                            & "SoitTable.ITCANCELED=0 AND SoitTable.ITPSNUMBER='' " & vbCrLf _
                            & "AND SoitTable.ITINVOICE=0 AND SoitTable.ITPSSHIPPED=0 " & vbCrLf _
                            & "AND CUREF LIKE '%' + @Customer + '%' " & vbCrLf _
                            & "AND SoitTable.ITSCHED <= @CutoffDate " & vbCrLf _
                            & "AND PartTable.PACLASS LIKE '%' + @PartClass + '%' " & vbCrLf _
                            & "AND PartTable.PAPRODCODE LIKE '%' + @PartCode + '%' " & vbCrLf _
                        & "ORDER BY partnum, ITSCHED " & vbCrLf _
                       & "OPEN curbackLog " & vbCrLf _
                       & "FETCH NEXT FROM curbackLog INTO @SoType, @SoText, @ItSo, "
                       
            sSql = sSql & " @ItRev, @ItNum, @ItQty, @PaLotRemQty, " & vbCrLf _
                                      & "@ItDollars,@ItSched, @CusName, @PartNum," & vbCrLf _
                                      & "@PartDesc, @PartExDesc, @PartLoc, @ItCanceled," & vbCrLf _
                                      & "@ItPSNum, @ItInvoice, @ItPSShipped" & vbCrLf _
                        & "SET @CurPartNum = @PartNum" & vbCrLf _
                        & "SET @RunningTot = 0" & vbCrLf _
                        & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf _
                        & "BEGIN" & vbCrLf _
                          & "IF (@@FETCH_STATUS <> -2)" & vbCrLf _
                          & "BEGIN" & vbCrLf _
                            & "IF  @CurPartNum <> @PartNum" & vbCrLf _
                            & "BEGIN" & vbCrLf _
                                & "SET @RunningTot = @ItQty" & vbCrLf _
                                & "set @CurPartNum = @PartNum" & vbCrLf _
                            & "End" & vbCrLf _
                            & "Else" & vbCrLf _
                            & "BEGIN" & vbCrLf _
                                & "SET @RunningTot = @RunningTot + @ItQty" & vbCrLf _
                            & "End" & vbCrLf _
                            & "SET @PartRem = @PaLotRemQty - @RunningTot " & vbCrLf _
                            & "INSERT INTO #tempBackLogInfo (SOTYPE, SOTEXT, ITSO, ITREV, " & vbCrLf _
                               & "ITNUMBER,ITQTY, PALOTQTYREMAINING,RUNQTYTOT, ITDOLLARS, " & vbCrLf _
                                & "ITSCHED,CUNICKNAME, PARTNUM, PADESC,PAEXTDESC,PALOCATION, " & vbCrLf _
                               & "ITCANCELED, ITPSNUMBER, ITINVOICE, ITPSSHIPPED) " & vbCrLf _
                            & "VALUES (@SoType, @SoText, @ItSo, @ItRev,@ItNum, "
                            
            sSql = sSql & "@ItQty,@PaLotRemQty,@PartRem, @ItDollars,@ItSched,@CusName, " & vbCrLf _
                               & "@PartNum,@PartDesc,@PartExDesc,@PartLoc, @ItCanceled,@ItPSNum,@ItInvoice,@ItPSShipped) " & vbCrLf _
                            & "End " & vbCrLf _
                            & "FETCH NEXT FROM curbackLog INTO @SoType, @SoText, @ItSo, " & vbCrLf _
                                          & "@ItRev, @ItNum, @ItQty, @PaLotRemQty, " & vbCrLf _
                                          & "@ItDollars,@ItSched, @CusName, @PartNum, " & vbCrLf _
                                          & "@PartDesc, @PartExDesc, @PartLoc, @ItCanceled, " & vbCrLf _
                                          & "@ItPSNum, @ItInvoice, @ItPSShipped " & vbCrLf _
                        & "End " & vbCrLf _
                        & "CLOSE curbackLog   --// close the cursor " & vbCrLf _
                        & "DEALLOCATE curbackLog " & vbCrLf _
                       & "-- select data for the report " & vbCrLf _
                        & "SELECT SOTYPE, SOTEXT, ITSO, ITREV, " & vbCrLf _
                           & "ITNUMBER,ITQTY, PALOTQTYREMAINING,RUNQTYTOT, ITDOLLARS, " & vbCrLf _
                            & "ITSCHED,CUNICKNAME, PARTNUM, PADESC,PAEXTDESC, PALOCATION, " & vbCrLf _
                           & "ITCANCELED , ITPSNUMBER, ITINVOICE, ITPSSHIPPED " & vbCrLf _
                        & "FROM #tempBackLogInfo " & vbCrLf _
                        & "ORDER BY ITSCHED " & vbCrLf _
                    & "-- drop the temp table " & vbCrLf _
                    & "DROP table #tempBackLogInfo " & vbCrLf _
                    & "End "
        ' Create the store procedure to create run time part quantity.
        Execute True, sSql
      
        Execute False, "DROP VIEW viewStatusCodeComments"
        ' Create the view to status code comments.
        sSql = "CREATE VIEW [dbo].[viewStatusCodeComments] AS " & vbCrLf _
                    & "SELECT  StCmtTable.STATUS_CMT_REF, StCmtTable.STATUS_CMT_REF1, " & vbCrLf _
                            & "StCmtTable.STATUS_CMT_REF2, StCmtTable.STATUS_REF, " & vbCrLf _
                            & "StCmtTable.STATCODE_TYPE_REF , StcodeTable.STATUS_CODE, " & vbCrLf _
                            & "StCmtTable.comment , StCmtTable.STATUS_ACT_STATE " & vbCrLf _
                    & "FROM StcodeTable INNER JOIN " & vbCrLf _
                        & "StCmtTable ON StcodeTable.STATUS_REF = StCmtTable.STATUS_REF "

        Execute True, sSql
      
        Execute False, "DROP PROCEDURE Qry_FillStatCode"
        sSql = "CREATE PROCEDURE [dbo].[Qry_FillStatCode] AS " & vbCrLf _
                & "SELECT STATUS_REF,STATUS_CODE " & vbCrLf _
                & "From STCODETABLE " & vbCrLf _
                & "ORDER BY STATUS_REF "
        Execute True, sSql
        
        Execute False, "DROP PROCEDURE Qry_UpdateStatusCode"
        ' Create procedure to add Status Code.
        sSql = "CREATE PROCEDURE [dbo].[Qry_UpdateStatusCode] " & vbCrLf _
                    & "(@StatRef varchar(4), @StatCode varchar(20)) " & vbCrLf _
                    & "AS " & vbCrLf _
                    & "BEGIN " & vbCrLf _
                    & "SELECT * FROM StcodeTable " & vbCrLf _
                        & "WHERE STATUS_REF = @StatRef " & vbCrLf _
                    & "IF @@ROWCOUNT = 0 " & vbCrLf _
                    & "BEGIN " & vbCrLf _
                        & "INSERT INTO StcodeTable (STATUS_REF, STATUS_CODE) " & vbCrLf _
                                & "VALUES (@StatRef, @StatCode) " & vbCrLf _
                    & "End " & vbCrLf _
                    & "Else " & vbCrLf _
                    & "BEGIN " & vbCrLf _
                        & "UPDATE StcodeTable SET STATUS_CODE = @StatCode " & vbCrLf _
                            & "WHERE STATUS_REF = @StatRef " & vbCrLf _
                    & "End " & vbCrLf _
                    & "End "
        
        Execute True, sSql

        
        Execute False, "DROP PROCEDURE Qry_AddInternStatCode"
        ' Create the store procedure to add Internal Status Code.
        sSql = "CREATE PROCEDURE [dbo].[Qry_AddInternStatCode] " & vbCrLf _
                 & "(@StatCmtRef int,@StatCmtRef1 int, @StatCmtRef2 varchar(20), " & vbCrLf _
                 & "@StatRef varchar(4), @StatCodeTypeRef varchar(3), @user varchar(50), " & vbCrLf _
                & "@comments as varchar(1024), @ActStat as int) " & vbCrLf _
                & "AS " & vbCrLf _
                & "BEGIN " & vbCrLf _
                & "declare @StatCdTypeKey as int " & vbCrLf _
                & "declare @curDate as datetime " & vbCrLf _
                & "SET @StatCdTypeKey = 0 " & vbCrLf _
                & "SELECT @StatCdTypeKey = ISNULL(STATCODE_TYPE_UNIQUE_KEY, 0) FROM StatCdType " & vbCrLf _
                & "SET @curDate = GetDATE() " & vbCrLf _
                & "IF @StatCdTypeKey = 1 " & vbCrLf _
                & "BEGIN " & vbCrLf _
                    & "SELECT * FROM StCmtTable WHERE " & vbCrLf _
                        & "STATUS_REF = @StatRef AND STATCODE_TYPE_REF = @StatCodeTypeRef " & vbCrLf _
                        & "AND STATUS_CMT_REF = @StatCmtRef " & vbCrLf _
                    & "SET @StatCmtRef1 = NULL " & vbCrLf _
                    & "SET @StatCmtRef2 = NULL " & vbCrLf _
                & "End " & vbCrLf _
                & "IF @StatCdTypeKey = 3 " & vbCrLf _
                & "BEGIN " & vbCrLf _
                    & "SELECT * FROM StCmtTable WHERE " & vbCrLf _
                        & "STATUS_REF = @StatRef AND STATCODE_TYPE_REF = @StatCodeTypeRef " & vbCrLf _
                        & "AND STATUS_CMT_REF = @StatCmtRef " & vbCrLf _
                        & "AND STATUS_CMT_REF1 = @StatCmtRef1 "
                        
        sSql = sSql & " AND STATUS_CMT_REF2 = @StatCmtRef2 " & vbCrLf _
                        & "End " & vbCrLf _
                        & "IF @@ROWCOUNT = 0 " & vbCrLf _
                            & "INSERT INTO StCmtTable (STATUS_REF, STATCODE_TYPE_REF, STATUS_CMT_REF, " & vbCrLf _
                                & "STATUS_CMT_REF1, STATUS_CMT_REF2, STATUS_ORG_USER, STATUS_ORG_DATE, " & vbCrLf _
                                & "COMMENT, STATUS_CUR_USER, STATUS_CUR_DATE, STATUS_ACT_STATE) " & vbCrLf _
                            & "VALUES (@StatRef, @StatCodeTypeRef, @StatCmtRef, @StatCmtRef1, @StatCmtRef2, " & vbCrLf _
                                & "@user, @curDate,@comments, @user, @curDate, @ActStat) " & vbCrLf _
                        & "Else " & vbCrLf _
                            & "UPDATE StCmtTable SET STATUS_CUR_USER = @user, STATUS_CUR_DATE = @curDate, " & vbCrLf _
                                & "COMMENT = @comments, STATUS_ACT_STATE = @ActStat " & vbCrLf _
                            & "WHERE STATUS_REF = @StatRef AND STATCODE_TYPE_REF = @StatCodeTypeRef " & vbCrLf _
                                & "AND STATUS_CMT_REF = @StatCmtRef " & vbCrLf _
                                & "AND STATUS_CMT_REF1 = @StatCmtRef1 " & vbCrLf _
                                & "AND STATUS_CMT_REF2 = @StatCmtRef2 " & vbCrLf _
                        & "End "
      
        ' Exec the store procedure to add Internal Status Code.
        Execute True, sSql
      
      'set version
      Execute False, "update Version set Version = " & newver
        
    End If
End Function

Private Function UpdateDatabase4()
    newver = 63
    If ver < newver Then
        ver = newver
        
       ' Drop the temp table and create the table
        Execute False, "DROP TABLE EsReportCapa15a"
        
        sSql = "CREATE TABLE [EsReportCapa15a] " & vbCrLf _
            & "([RPTSHOPREF] [char](12) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTSHOPNUM] [char](12) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTSHOPDESC] [char](30) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTWCNREF] [char](12) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTWCNNUM] [char](12) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTWCNDESC] [char](30) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTBEGDATE1] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTENDDATE1] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS1] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTUSEDHOURS1] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTBEGDATE2] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTENDDATE2] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS2] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTUSEDHOURS2] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTBEGDATE3] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTENDDATE3] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS3] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTUSEDHOURS3] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTBEGDATE4] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTENDDATE4] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS4] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTUSEDHOURS4] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTBEGDATE5] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTENDDATE5] [char](8) NULL DEFAULT (''), " & vbCrLf
            
    sSql = sSql & "[RPTHOURS5] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTUSEDHOURS5] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTBEGDATE6] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTENDDATE6] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS6] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTUSEDHOURS6] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTBEGDATE7] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTENDDATE7] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS7] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTUSEDHOURS7] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTBEGDATE8] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTENDDATE8] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS8] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTUSEDHOURS8] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTBEGDATE9] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTENDDATE9] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS9] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTUSEDHOURS9] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTBEGDATE10] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTENDDATE10] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS10] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTUSEDHOURS10] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTBEGDATE11] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTENDDATE11] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS11] [real] NULL DEFAULT ((0)), " & vbCrLf
            
        sSql = sSql & "[RPTUSEDHOURS11] [real] NULL DEFAULT ((0)) " & vbCrLf _
            & ") ON [PRIMARY] " & vbCrLf
            
        Execute True, sSql
        
        Execute False, "CREATE CLUSTERED INDEX ReportRef ON dbo.EsReportCapa15a(RPTSHOPREF,RPTWCNREF) WITH  FILLFACTOR = 80"
        Execute False, "CREATE INDEX ShopRef ON dbo.EsReportCapa15a(RPTSHOPREF) WITH  FILLFACTOR = 80"
        Execute False, "CREATE INDEX WcnRef ON dbo.EsReportCapa15a(RPTWCNREF) WITH  FILLFACTOR = 80"
      


       ' Drop the temp table and create the table
        Execute False, "DROP TABLE EsReportCapa15b"
        
        sSql = "CREATE TABLE [EsReportCapa15b] " & vbCrLf _
            & "([RPTRECORD] [int] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTPARTREF] [char](30) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTPARTNUM] [char](30) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTPARTDESC] [char](30) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTSHOPREF] [char](12) DEFAULT (''), " & vbCrLf _
            & "[RPTWCNREF] [char](12) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTRUNNO] [int] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTOPNO] [char](5) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTRUNSTATUS] [char](2) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTSCHEDCOMPL] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTREMAININGQTY] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTENDDATE1] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS1] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTENDDATE2] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS2] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTENDDATE3] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS3] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTENDDATE4] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS4] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTENDDATE5] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS5] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTENDDATE6] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS6] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTENDDATE7] [char](8) NULL DEFAULT (''), " & vbCrLf
            
        sSql = sSql & "[RPTHOURS7] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTENDDATE8] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS8] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTENDDATE9] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS9] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTENDDATE10] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS10] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTENDDATE11] [char](8) NULL DEFAULT (''), " & vbCrLf _
            & "[RPTHOURS11] [real] NULL DEFAULT ((0)), " & vbCrLf _
            & "[RPTPASTDUE] [char](1) NULL DEFAULT ('') " & vbCrLf _
            & ") ON [PRIMARY] "
            
        Execute True, sSql
        
        Execute False, "CREATE CLUSTERED INDEX ReportRef ON dbo.EsReportCapa15b(RPTRECORD) WITH  FILLFACTOR = 80"
        Execute False, "CREATE INDEX ShopRef ON dbo.EsReportCapa15b(RPTSHOPREF) WITH  FILLFACTOR = 80"
        Execute False, "CREATE INDEX WcnRef ON dbo.EsReportCapa15b(RPTWCNREF) WITH  FILLFACTOR = 80"


        ' Fixe Account field to Varchar
        Execute False, "DROP TABLE [dbo].[tempRawMatFinishGoods]"
        
        sSql = "CREATE TABLE [dbo].[tempRawMatFinishGoods]( " & vbCrLf _
            & "[LOTNUMBER] [varchar](15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
            & "[PARTNUM] [varchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
            & "[PADESC] [varchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
            & "[PAEXTDESC] [varchar](3072) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
            & "[LOTUSERLOTID] [char](40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
            & "[ORGINNUMBER] [int] NULL, " & vbCrLf _
            & "[RPTINNUMBER] [int] NULL, " & vbCrLf _
            & "[CURINNUMBER] [int] NULL, " & vbCrLf _
            & "[ACTUALDATE] [smalldatetime] NULL, " & vbCrLf _
            & "[RPTDATEQTY] [decimal](12, 4) NULL, " & vbCrLf _
            & "[UNITCOST] [decimal](12, 4) NULL, " & vbCrLf _
            & "[PASTDCOST] [decimal](12, 4) NULL, " & vbCrLf _
            & "[LOTUNITCOST] [decimal](12, 4) NULL, " & vbCrLf _
            & "[INAMT] [decimal](12, 4) NULL, " & vbCrLf _
            & "[ORGCOST] [decimal](12, 4) NULL, " & vbCrLf _
            & "[STDCOST] [decimal](12, 4) NULL, " & vbCrLf _
            & "[LSTACOST] [decimal](12, 4) NULL, " & vbCrLf _
            & "[RPTCOST] [decimal](12, 4) NULL, " & vbCrLf _
            & "[CURCOST] [decimal](12, 4) NULL, " & vbCrLf _
            & "[COSTEDDATE] [smalldatetime] NULL, " & vbCrLf _
            & "[RPTACCOUNT] [char](12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
            & "[ORIGINALACC] [char](12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
            & "[LASTACTVITYACC] [char](12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
            & "[CURRENTACC] [char](12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf
        
        sSql = sSql & "[PACLASS] [char](4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
            & "[PAPRODCODE] [char](6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
            & "[flgStdCost] [int] NULL, " & vbCrLf _
            & "[flgLdCost] [int] NULL, " & vbCrLf _
            & "[flgInvCost] [int] NULL, " & vbCrLf _
            & "[flgLdRQErr] [int] NULL, " & vbCrLf _
            & "[flgRptAcc] [int] NULL, " & vbCrLf _
            & "[flgOrgAcc] [int] NULL, " & vbCrLf _
            & "[flgLastAcc] [int] NULL " & vbCrLf _
                & ") ON [PRIMARY] "
                
        Execute True, sSql
                
                
    sSql = "ALTER PROCEDURE [dbo].[RawMaterialFinishGoods]" & vbCrLf _
        & "@ReportDate as varchar(16), @PartClass as Varchar(16)," & vbCrLf _
        & "@PartCode as varchar(8), @lotHDOnly as int" & vbCrLf _
        & "AS" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "declare @partRef as varchar(30)" & vbCrLf _
        & "declare @partDesc as varchar(30)" & vbCrLf _
        & "declare @partExDesc as varchar(3072)" & vbCrLf _
        & "declare @rptRemQty as decimal(12,4)" & vbCrLf _
        & "declare @deltaQty as decimal(12,4)" & vbCrLf _
        & "declare @lotRemQty as decimal(12,4)" & vbCrLf _
        & "declare @lotOrgQty as decimal(12,4)" & vbCrLf _
        & "declare @tmpInvQty as decimal(12,4)" & vbCrLf _
        & "declare @tmpLastInvQty as decimal(12,4)" & vbCrLf _
        & "declare @rptInvcost decimal(12,4)" & vbCrLf _
        & "declare @lastInvcost decimal(12,4)" & vbCrLf _
        & "declare @orgInvcost decimal(12,4)" & vbCrLf _
        & "declare @tmpInvCost decimal(12,4)" & vbCrLf _
        & "declare @tmpLastInvCost decimal(12,4)" & vbCrLf _
        & "declare @lastQty as decimal(12,4)" & vbCrLf _
        & "declare @orgQty as decimal(12,4)" & vbCrLf _
        & "declare @rptQty as decimal(12,4)" & vbCrLf _
        & "declare @rptCreditACC varchar(12)" & vbCrLf _
        & "declare @rptDebitACC varchar(12)" & vbCrLf _
        & "declare @rptACC varchar(12)" & vbCrLf
    sSql = sSql & "declare @CurACC varchar(12)" & vbCrLf _
        & "declare @tmpCreditAcc varchar(12)" & vbCrLf _
        & "declare @tmpDebitAcc varchar(12)" & vbCrLf _
        & "declare @tmplastCreditAcc varchar(12)" & vbCrLf _
        & "declare @tmpLastDebitAcc varchar(12)" & vbCrLf _
        & "declare @lastCreditACC varchar(12)" & vbCrLf _
        & "declare @lastDebitACC varchar(12)" & vbCrLf _
        & "declare @lastACC varchar(12)" & vbCrLf _
        & "declare @orgCreditACC varchar(12)" & vbCrLf _
        & "declare @orgDebitACC varchar(12)" & vbCrLf _
        & "declare @orgACC varchar(12)" & vbCrLf _
        & "declare @OrgInvNum as int" & vbCrLf _
        & "declare @rptInvNum as int" & vbCrLf _
        & "declare @LastInvNum as int" & vbCrLf _
        & "declare @tmpInvNum as int" & vbCrLf _
        & "declare @tmpLastInvNum as int" & vbCrLf _
        & "declare @LotNumber varchar(51)" & vbCrLf _
        & "declare @LotUserID varchar(51)" & vbCrLf _
        & "declare @lotAcualDate as smalldatetime" & vbCrLf _
        & "declare @lotCostedDate as smalldatetime" & vbCrLf _
        & "declare @curDate as smalldatetime" & vbCrLf _
        & "declare @AcualDate as smalldatetime" & vbCrLf _
        & "declare @CostedDate as smalldatetime" & vbCrLf _
        & "declare @tmpINVAdate  as smalldatetime" & vbCrLf _
        & "declare @tmpLastINVAdate  as smalldatetime" & vbCrLf
    sSql = sSql & "declare @unitcost as decimal(12,4)" & vbCrLf _
        & "declare @partStdCost as decimal(12,4)" & vbCrLf _
        & "declare @LotUnitCost as decimal(12,4)" & vbCrLf _
        & "declare @lotTotMatl as decimal(12,4)" & vbCrLf _
        & "declare @lotTotLabor as decimal(12,4)" & vbCrLf _
        & "declare @lotTotExp as decimal(12,4)" & vbCrLf _
        & "declare @lotTotOH as decimal(12,4)" & vbCrLf _
        & "declare @partActCost as int" & vbCrLf _
        & "declare @partLotTrack as int" & vbCrLf _
        & "declare @flgStdCost as int" & vbCrLf _
        & "declare @flgLdCost as int" & vbCrLf _
        & "declare @flgInvCost as int" & vbCrLf _
        & "declare @flgLdRQErr as int" & vbCrLf _
        & "declare @flgOrgAcc as int" & vbCrLf _
        & "declare @flgRptAcc as int" & vbCrLf _
        & "declare @flgLastAcc as int" & vbCrLf _
        & "DELETE FROM tempRawMatFinishGoods" & vbCrLf _
        & "IF (@PartClass = 'ALL')" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @PartClass = ''" & vbCrLf _
        & "End" & vbCrLf _
        & "IF (@PartCode = 'ALL')" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @PartCode = ''" & vbCrLf _
        & "End" & vbCrLf
    sSql = sSql & "DECLARE curLotHd CURSOR LOCAL" & vbCrLf _
        & "FOR" & vbCrLf _
        & "SELECT LOTNUMBER, LOTUSERLOTID, PARTREF, PADESC, PAEXTDESC," & vbCrLf _
        & "LOTADATE,LOTORIGINALQTY, LOTREMAININGQTY, LOTUNITCOST, LOTDATECOSTED," & vbCrLf _
        & "PACLASS , PAPRODCODE, PASTDCOST, PAUSEACTUALCOST, PALOTTRACK" & vbCrLf _
        & "From ViewLohdPartTable" & vbCrLf _
        & "WHERE ViewLohdPartTable.LOTADATE  < DATEADD(dd, 1 , @ReportDate)" & vbCrLf _
        & "AND ViewLohdPartTable.PACLASS LIKE '%' + @PartClass + '%'" & vbCrLf _
        & "AND ViewLohdPartTable.PAPRODCODE LIKE '%' + @PartCode + '%'" & vbCrLf _
        & "AND ViewLohdPartTable.PALEVEL <= 4" & vbCrLf _
        & "OPEN curLotHd" & vbCrLf _
        & "FETCH NEXT FROM curLotHd INTO @LotNumber, @LotUserID, @partRef," & vbCrLf _
        & "@partDesc, @partExDesc, @lotAcualDate, @lotOrgQty," & vbCrLf _
        & "@lotRemQty,@LotUnitCost, @lotCostedDate," & vbCrLf _
        & "@PartClass, @PartCode, @partStdCost, @partActCost, @partLotTrack" & vbCrLf _
        & "WHILE (@@FETCH_STATUS <> -1)" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "IF (@@FETCH_STATUS <> -2)" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @flgLdRQErr = 0" & vbCrLf _
        & "IF (@lotRemQty < 0.0000)" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @lotRemQty = 0.0000" & vbCrLf _
        & "SET @flgLdRQErr  = 1" & vbCrLf _
        & "End" & vbCrLf
    sSql = sSql & "SET @curDate = GETDATE()" & vbCrLf _
        & "SELECT @deltaQty = ISNULL(SUM(LOIQUANTITY), 0.0000)" & vbCrLf _
        & "From LoitTable" & vbCrLf _
        & "WHERE LOIADATE BETWEEN DATEADD(dd, 1 ,@ReportDate) AND DATEADD(dd, 1 ,@curDate)" & vbCrLf _
        & "AND LOIPARTREF = @partRef" & vbCrLf _
        & "AND LOINUMBER = @LotNumber" & vbCrLf _
        & "SET @rptRemQty = @lotRemQty + (@deltaQty * -1)" & vbCrLf _
        & "IF @rptRemQty < 0.0000" & vbCrLf _
        & "SET @rptRemQty = @rptRemQty * -1" & vbCrLf _
        & "SET @flgStdCost = 0" & vbCrLf _
        & "SET @flgLdCost = 0" & vbCrLf _
        & "SET @flgInvCost = 0" & vbCrLf _
        & "SET @flgOrgAcc = 0" & vbCrLf _
        & "SET @flgRptAcc = 0" & vbCrLf _
        & "SET @flgLastAcc = 0" & vbCrLf _
        & "DECLARE curInv CURSOR" & vbCrLf _
        & "LOCAL" & vbCrLf _
        & "Scroll" & vbCrLf _
        & "FOR" & vbCrLf _
        & "SELECT INNUMBER, INAMT, INAQTY, ISNULL(INCREDITACCT,0), ISNULL(INDEBITACCT, 0), INADATE" & vbCrLf _
        & "From InvaTable, LoitTable" & vbCrLf _
        & "WHERE InvaTable.INPART = @partRef" & vbCrLf _
        & "AND LoitTable.LOINUMBER = @LotNumber" & vbCrLf _
        & "AND InvaTable.INPART = LoitTable.LOIPARTREF" & vbCrLf _
        & "AND InvaTable.INNUMBER = LoitTable.LOIACTIVITY" & vbCrLf
    sSql = sSql & "ORDER BY INADATE ASC" & vbCrLf _
        & "OPEN curInv" & vbCrLf _
        & "FETCH FIRST FROM curInv INTO @tmpInvNum, @tmpInvCost, @tmpInvQty," & vbCrLf _
        & "@tmpCreditAcc, @tmpDebitAcc, @tmpINVAdate" & vbCrLf _
        & "IF (@@FETCH_STATUS <> -1)" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "IF (@@FETCH_STATUS <> -2)" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @orgInvcost = @tmpInvCost" & vbCrLf _
        & "SET @orgQty = @tmpInvQty" & vbCrLf _
        & "SET @orgCreditACC = @tmpCreditAcc" & vbCrLf _
        & "SET @orgDebitACC = @tmpDebitAcc" & vbCrLf _
        & "SET @OrgInvNum = @tmpInvNum" & vbCrLf _
        & "End" & vbCrLf _
        & "FETCH LAST FROM curInv INTO @tmpLastInvNum, @tmpLastInvCost, @tmpLastInvQty," & vbCrLf _
        & "@tmplastCreditAcc, @tmpLastDebitAcc, @tmpLastINVAdate" & vbCrLf _
        & "IF (@@FETCH_STATUS <> -2)" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @lastInvcost = @tmpLastInvCost" & vbCrLf _
        & "SET @lastQty = @tmpLastInvQty" & vbCrLf _
        & "SET @lastCreditACC = @tmplastCreditAcc" & vbCrLf _
        & "SET @lastDebitACC = @tmpLastDebitAcc" & vbCrLf _
        & "SET @LastInvNum = @tmpLastInvNum" & vbCrLf _
        & "IF @tmpLastINVAdate > DATEADD(dd, 1, @ReportDate)" & vbCrLf _
        & "BEGIN" & vbCrLf
    sSql = sSql & "SELECT TOP 1 @rptInvNum = INNUMBER, @rptInvcost = INAMT," & vbCrLf _
        & "@rptQty = INAQTY, @rptCreditACC = ISNULL(INCREDITACCT, 0)," & vbCrLf _
        & "@rptDebitACC = ISNULL(INDEBITACCT, 0)" & vbCrLf _
        & "From InvaTable, LoitTable" & vbCrLf _
        & "WHERE INADATE < DATEADD(dd, 1, @ReportDate)" & vbCrLf _
        & "AND InvaTable.INPART = @partRef" & vbCrLf _
        & "AND LoitTable.LOINUMBER = @LotNumber" & vbCrLf _
        & "AND InvaTable.INPART = LoitTable.LOIPARTREF" & vbCrLf _
        & "AND InvaTable.INNUMBER = LoitTable.LOIACTIVITY" & vbCrLf _
        & "ORDER BY INADATE DESC" & vbCrLf _
        & "End" & vbCrLf _
        & "Else" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @rptInvNum = @tmpLastInvNum" & vbCrLf _
        & "SET @rptInvcost = @tmpLastInvCost" & vbCrLf _
        & "SET @rptQty = @tmpLastInvQty" & vbCrLf _
        & "SET @rptCreditACC = @tmplastCreditAcc" & vbCrLf _
        & "SET @rptDebitACC = @tmpLastDebitAcc" & vbCrLf _
        & "End" & vbCrLf _
        & "End" & vbCrLf _
        & "End" & vbCrLf _
        & "CLOSE curInv   --// close the cursor" & vbCrLf _
        & "DEALLOCATE curInv" & vbCrLf _
        & "IF (@partActCost = 0)" & vbCrLf _
        & "BEGIN" & vbCrLf
    sSql = sSql & "SET @unitcost = @partStdCost" & vbCrLf _
        & "SET @flgStdCost = 1" & vbCrLf _
        & "End" & vbCrLf _
        & "Else" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "IF @lotHDOnly = 1" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @unitcost = @LotUnitCost" & vbCrLf _
        & "SET @flgLdCost = 1" & vbCrLf _
        & "End" & vbCrLf _
        & "Else" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "IF @lotCostedDate < DATEADD(dd, 1, @ReportDate)" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @unitcost = @LotUnitCost" & vbCrLf
    sSql = sSql & "SET @flgLdCost = 1" & vbCrLf _
        & "End" & vbCrLf _
        & "Else" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @unitcost = @rptInvcost" & vbCrLf _
        & "SET @flgInvCost = 1" & vbCrLf _
        & "End" & vbCrLf _
        & "END --LotHD only" & vbCrLf _
        & "END --Part Cost" & vbCrLf _
        & "SELECT @CurACC = dbo.fnGetPartInvAccount(@partRef)" & vbCrLf _
        & "-- Lastest < report account#" & vbCrLf _
        & "IF @rptQty >= 0.0000" & vbCrLf _
        & "SET @rptACC = @rptDebitACC" & vbCrLf _
        & "Else" & vbCrLf _
        & "SET @rptACC = @rptCreditACC" & vbCrLf _
        & "IF ((@rptACC = '') OR (@rptACC = NULL))" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @rptACC = @CurACC" & vbCrLf _
        & "SET @flgRptAcc = 1" & vbCrLf _
        & "End" & vbCrLf _
        & "-- last record" & vbCrLf _
        & "IF @lastQty >= 0.0000" & vbCrLf _
        & "SET @lastACC = @lastDebitACC" & vbCrLf _
        & "Else" & vbCrLf _
        & "SET @lastACC = @lastCreditACC" & vbCrLf
    sSql = sSql & "IF ((@lastACC = '') OR (@lastACC = NULL))" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @lastACC = @CurACC" & vbCrLf _
        & "SET @flgLastAcc = 1" & vbCrLf _
        & "End" & vbCrLf _
        & "-- Lastest < report account#" & vbCrLf _
        & "IF @orgQty >= 0.0000" & vbCrLf _
        & "SET @orgACC = @orgDebitACC" & vbCrLf _
        & "Else" & vbCrLf _
        & "SET @orgACC = @orgCreditACC" & vbCrLf _
        & "IF ((@orgACC = '') OR (@orgACC = NULL))" & vbCrLf _
        & "BEGIN" & vbCrLf _
        & "SET @orgACC = @CurACC" & vbCrLf _
        & "SET @flgOrgAcc = 1" & vbCrLf _
        & "End" & vbCrLf _
        & "-- Insert to the temp table" & vbCrLf _
        & "INSERT INTO tempRawMatFinishGoods" & vbCrLf _
        & "(PARTNUM, PADESC, PAEXTDESC, LOTNUMBER, LOTUSERLOTID," & vbCrLf _
        & "ORGINNUMBER, RPTINNUMBER, CURINNUMBER, ACTUALDATE,RPTDATEQTY, COSTEDDATE," & vbCrLf _
        & "UNITCOST,PASTDCOST, LOTUNITCOST, ORGCOST, STDCOST," & vbCrLf _
        & "LSTACOST, RPTCOST, CURCOST, RPTACCOUNT,ORIGINALACC," & vbCrLf _
        & "LASTACTVITYACC, CURRENTACC, PACLASS,PAPRODCODE," & vbCrLf _
        & "flgStdCost, flgLdCost, flgInvCost, flgLdRQErr," & vbCrLf _
        & "flgRptAcc, flgOrgAcc, flgLastAcc)" & vbCrLf _
        & "VALUES (@partRef, @partDesc, @partExDesc, @LotNumber,@LotUserID, @OrgInvNum," & vbCrLf
    sSql = sSql & "@rptInvNum, @LastInvNum, @lotAcualDate,@rptRemQty,@lotCostedDate," & vbCrLf _
        & "@unitcost,@partStdCost, @LotUnitCost, @orgInvcost, @partStdCost," & vbCrLf _
        & "@lastInvcost,@rptInvcost, @LotUnitCost, @rptACC, @orgACC, @lastACC," & vbCrLf _
        & "@CurACC, @PartClass,@PartCode,@flgStdCost, @flgLdCost, @flgInvCost," & vbCrLf _
        & "@flgLdRQErr, @flgRptAcc, @flgOrgAcc, @flgLastAcc)" & vbCrLf _
        & "SET @rptRemQty = NULL" & vbCrLf _
        & "SET @deltaQty = NULL" & vbCrLf _
        & "SET @lotRemQty = NULL" & vbCrLf _
        & "SET @lotOrgQty = NULL" & vbCrLf _
        & "SET @tmpInvQty = NULL" & vbCrLf _
        & "SET @tmpLastInvQty  = NULL" & vbCrLf _
        & "SET @rptInvcost = NULL" & vbCrLf _
        & "SET @lastInvcost = NULL" & vbCrLf _
        & "SET @orgInvcost = NULL" & vbCrLf _
        & "SET @tmpInvCost = NULL" & vbCrLf _
        & "SET @unitcost = NULL" & vbCrLf _
        & "SET @tmpLastInvCost = NULL" & vbCrLf _
        & "SET @lastQty = NULL" & vbCrLf _
        & "SET @orgQty = NULL" & vbCrLf _
        & "SET @rptQty  = NULL" & vbCrLf _
        & "SET @rptCreditACC = NULL" & vbCrLf _
        & "SET @rptDebitACC  = NULL" & vbCrLf _
        & "SET @rptACC  = NULL" & vbCrLf _
        & "SET @tmpCreditAcc  = NULL" & vbCrLf _
        & "SET @tmpDebitAcc = NULL" & vbCrLf
    sSql = sSql & "SET @tmplastCreditAcc = NULL" & vbCrLf _
        & "SET @tmpLastDebitAcc = NULL" & vbCrLf _
        & "SET @lastCreditACC = NULL" & vbCrLf _
        & "SET @lastDebitACC = NULL" & vbCrLf _
        & "SET @lastACC  = NULL" & vbCrLf _
        & "SET @orgCreditACC  = NULL" & vbCrLf _
        & "SET @orgDebitACC  = NULL" & vbCrLf _
        & "SET @orgACC  = NULL" & vbCrLf _
        & "SET @OrgInvNum = NULL" & vbCrLf _
        & "SET @rptInvNum = NULL" & vbCrLf _
        & "SET @LastInvNum = NULL" & vbCrLf _
        & "End" & vbCrLf _
        & "FETCH NEXT FROM curLotHd INTO @LotNumber, @LotUserID, @partRef," & vbCrLf _
        & "@partDesc, @partExDesc, @lotAcualDate, @lotOrgQty," & vbCrLf _
        & "@lotRemQty,@LotUnitCost, @lotCostedDate," & vbCrLf _
        & "@PartClass, @PartCode, @partStdCost, @partActCost, @partLotTrack" & vbCrLf _
        & "End" & vbCrLf _
        & "CLOSE curLotHd   --// close the cursor" & vbCrLf _
        & "DEALLOCATE curLotHd" & vbCrLf _
        & "SELECT * FROM tempRawMatFinishGoods" & vbCrLf _
        & "End"
        
        ' Alter the RMTG's sp
        Execute True, sSql
        
        
        sSql = "ALTER VIEW [dbo].[viewOpenAP]" & vbCrLf _
                & "AS" & vbCrLf _
                & "SELECT     VIVENDOR, VENICKNAME, VEBNAME, VINO, VIDATE, InvTotal, VIDUEDATE, VIFREIGHT, VITAX, VIPAY, VITPO, VITPORELEASE, DiscRate, DiscDays," & vbCrLf _
                & "       NetDays, AmountDue, CalcTotal, VITCOST, PILOT, CAST(AmountDue * DiscRate / 100 AS decimal(12, 2)) AS DiscAmount," & vbCrLf _
                & "       CASE WHEN DiscRate > 0 THEN DATEADD(day, DiscDays, VIDATE) ELSE NULL END AS DiscCutoff" & vbCrLf _
                & " From dbo.viewOpenAPTerms"
                
        ' Execute the sql to alter the view viewOpenAP
        Execute True, sSql
            
        
        ' Drop the view and create the view
        Execute False, "DROP VIEW viewOpenAPWrapper"

        sSql = "CREATE VIEW [dbo].[viewOpenAPWrapper]" & vbCrLf _
                & "AS " & vbCrLf _
                & "SELECT VIVENDOR, VENICKNAME, VEBNAME, VINO, VIDATE, InvTotal, " & vbCrLf
                
        sSql = sSql & " VIDUEDATE, VIFREIGHT, VITAX, VIPAY, VITPO, VITPORELEASE, DiscRate, DiscDays, " & vbCrLf _
                & " NetDays , AmountDue, CalcTotal, VITCOST, PILOT, DiscAmount, DiscCutoff " & vbCrLf _
                & "     From dbo.viewOpenAP "

        ' Execute the sql to create the ViewOpenAP
        Execute True, sSql



        sSql = "ALTER VIEW [dbo].[viewOpenAPTerms]" & vbCrLf _
                & "AS" & vbCrLf _
                & "SELECT dbo.VihdTable.VIVENDOR, dbo.VndrTable.VENICKNAME, dbo.VndrTable.VEBNAME, dbo.VihdTable.VINO, dbo.VihdTable.VIDATE," & vbCrLf _
                & "    dbo.VihdTable.VIDUE AS InvTotal, dbo.VihdTable.VIDUEDATE, dbo.VihdTable.VIFREIGHT, dbo.VihdTable.VITAX, dbo.VihdTable.VIPAY," & vbCrLf _
                & "    ISNULL(ViitTable_1.VITPO, 0) AS VITPO, ISNULL(ViitTable_1.VITPORELEASE, 0) AS VITPORELEASE," & vbCrLf _
                & "    CASE WHEN PONETDAYS <> 0 THEN PODISCOUNT ELSE VEDISCOUNT END AS DiscRate," & vbCrLf _
                & "    CASE WHEN PONETDAYS <> 0 THEN PODDAYS ELSE VEDDAYS END AS DiscDays," & vbCrLf _
                & "    CASE WHEN PONETDAYS <> 0 THEN PONETDAYS ELSE VENETDAYS END AS NetDays, dbo.VihdTable.VIDUE - dbo.VihdTable.VIPAY AS AmountDue," & vbCrLf _
                & "    dbo.VihdTable.VIFREIGHT + dbo.VihdTable.VITAX +" & vbCrLf _
                & "          (SELECT     CAST(SUM(ROUND(CAST(dbo.ViitTable.VITQTY AS decimal(12, 3)) / (CASE WHEN ISNULL(dbo.viewPohdPoit.PILOT, 0)" & vbCrLf _
                & "                 <> 0 THEN CAST(dbo.viewPohdPoit.PILOT AS decimal(12, 3)) ELSE 1 END) * CAST(dbo.ViitTable.VITCOST AS decimal(12, 4))" & vbCrLf _
                & "                     + CAST(dbo.ViitTable.VITADDERS AS decimal(12, 2)), 2)) AS decimal(12, 2)) AS Expr1" & vbCrLf _
                & "                         FROM   dbo.ViitTable LEFT OUTER JOIN" & vbCrLf _
                & "                                                   dbo.viewPohdPoit ON dbo.ViitTable.VITPO = dbo.viewPohdPoit.PONUMBER AND" & vbCrLf _
                & "                                                   dbo.ViitTable.VITPORELEASE = dbo.viewPohdPoit.PORELEASE AND dbo.ViitTable.VITPOITEM = dbo.viewPohdPoit.PIITEM AND" & vbCrLf _
                & "                                                   dbo.ViitTable.VITPOITEMREV = dbo.viewPohdPoit.PIREV" & vbCrLf _
                & "                            WHERE      (dbo.ViitTable.VITVENDOR = dbo.VihdTable.VIVENDOR) AND (dbo.ViitTable.VITNO = dbo.VihdTable.VINO)) AS CalcTotal," & vbCrLf _
                & "                      ViitTable_1.VITCOST , viewPohdPoit_1.PILOT" & vbCrLf _
                & "             FROM  dbo.VihdTable INNER JOIN" & vbCrLf _
                & "                      dbo.VndrTable ON dbo.VihdTable.VIVENDOR = dbo.VndrTable.VEREF LEFT OUTER JOIN" & vbCrLf _
                & "                      dbo.viewPohdPoit AS viewPohdPoit_1 RIGHT OUTER JOIN" & vbCrLf _
                & "                      dbo.ViitTable AS ViitTable_1 ON viewPohdPoit_1.PIREV = ViitTable_1.VITPOITEMREV AND viewPohdPoit_1.PIITEM = ViitTable_1.VITPOITEM AND" & vbCrLf _
                & "                      viewPohdPoit_1.PONUMBER = ViitTable_1.VITPO AND viewPohdPoit_1.PORELEASE = ViitTable_1.VITPORELEASE ON" & vbCrLf _
                & "                      ViitTable_1.VITNO = dbo.VihdTable.VINO AND ViitTable_1.VITVENDOR = dbo.VihdTable.VIVENDOR AND ViitTable_1.VITITEM =" & vbCrLf _
                & "                          (SELECT     MIN(VITITEM) AS Expr1" & vbCrLf
                
                
            sSql = sSql & "                            From dbo.ViitTable" & vbCrLf _
                   & "                         WHERE      (VITVENDOR = dbo.VndrTable.VEREF) AND (VITNO = dbo.VihdTable.VINO) AND (VITPO IS NOT NULL) AND (VITPO <> 0))" & vbCrLf _
                   & "        Where (dbo.VihdTable.VIPIF <> 1) And (ISNULL(ViitTable_1.VITPO, 0) <> 0)" & vbCrLf _
                   & "         Union" & vbCrLf _
                   & "        SELECT     VihdTable_1.VIVENDOR, VndrTable_1.VENICKNAME, VndrTable_1.VEBNAME, VihdTable_1.VINO, VihdTable_1.VIDATE, VihdTable_1.VIDUE AS InvTotal," & vbCrLf _
                   & "                               VihdTable_1.VIDUEDATE, VihdTable_1.VIFREIGHT, VihdTable_1.VITAX, VihdTable_1.VIPAY, ISNULL(ViitTable_1.VITPO, 0) AS VITPO," & vbCrLf _
                   & "                               ISNULL(ViitTable_1.VITPORELEASE, 0) AS VITPORELEASE, CASE WHEN PONETDAYS <> 0 THEN PODISCOUNT ELSE VEDISCOUNT END AS DiscRate," & vbCrLf _
                   & "                               CASE WHEN PONETDAYS <> 0 THEN PODDAYS ELSE VEDDAYS END AS DiscDays," & vbCrLf _
                   & "                               CASE WHEN PONETDAYS <> 0 THEN PONETDAYS ELSE VENETDAYS END AS NetDays, VihdTable_1.VIDUE - VihdTable_1.VIPAY AS AmountDue," & vbCrLf _
                   & "                               VihdTable_1.VIFREIGHT + VihdTable_1.VITAX +" & vbCrLf _
                   & "                                   (SELECT     CAST(SUM(ROUND(CAST(VITQTY AS decimal(12, 3)) * CAST(VITCOST AS decimal(12, 4)) + CAST(VITADDERS AS decimal(12, 2)), 2))" & vbCrLf _
                   & "                                                            AS decimal(12, 2)) AS Expr1" & vbCrLf _
                   & "                                    FROM          dbo.ViitTable AS ViitTable_2" & vbCrLf _
                   & "                                     WHERE      (VITVENDOR = VihdTable_1.VIVENDOR) AND (VITNO = VihdTable_1.VINO)) AS CalcTotal, ViitTable_1.VITCOST," & vbCrLf _
                   & "                               viewPohdPoit_2.PILOT" & vbCrLf _
                   & "         FROM         dbo.VihdTable AS VihdTable_1 INNER JOIN" & vbCrLf _
                   & "                               dbo.VndrTable AS VndrTable_1 ON VihdTable_1.VIVENDOR = VndrTable_1.VEREF LEFT OUTER JOIN" & vbCrLf _
                   & "                               dbo.viewPohdPoit AS viewPohdPoit_2 RIGHT OUTER JOIN" & vbCrLf _
                   & "                               dbo.ViitTable AS ViitTable_1 ON viewPohdPoit_2.PIREV = ViitTable_1.VITPOITEMREV AND viewPohdPoit_2.PIITEM = ViitTable_1.VITPOITEM AND" & vbCrLf _
                   & "                               viewPohdPoit_2.PONUMBER = ViitTable_1.VITPO AND viewPohdPoit_2.PORELEASE = ViitTable_1.VITPORELEASE ON" & vbCrLf _
                   & "                               ViitTable_1.VITNO = VihdTable_1.VINO AND ViitTable_1.VITVENDOR = VihdTable_1.VIVENDOR AND ViitTable_1.VITITEM =" & vbCrLf _
                   & "                                   (SELECT     MIN(VITITEM) AS Expr1" & vbCrLf _
                   & "                                     From dbo.ViitTable" & vbCrLf _
                   & "                                     WHERE      (VITVENDOR = VndrTable_1.VEREF) AND (VITNO = VihdTable_1.VINO) AND (VITPO IS NOT NULL) AND (VITPO <> 0))" & vbCrLf _
                   & "         Where (VihdTable_1.VIPIF <> 1) And (ISNULL(ViitTable_1.VITPO, 0) = 0)"


            ' Execute the sql to alter the view viewOpenAPTerms
            Execute True, sSql
            
            
            
            sSql = "ALTER function [dbo].[fnGetOpenAP]" & vbCrLf _
                    & "(" & vbCrLf _
                    & "   @MinDiscountDate datetime, -- show discounts available on or after this date" & vbCrLf _
                    & "   @MaxDueDate datetime    -- show invoices due on or before this date" & vbCrLf _
                    & ")" & vbCrLf _
                    & "returns Table" & vbCrLf _
                    & "as" & vbCrLf _
                    & "Return" & vbCrLf _
                    & "(" & vbCrLf _
                    & "select *," & vbCrLf _
                    & "cast(case when DiscCutoff is null then null" & vbCrLf _
                    & "when datediff(d, DiscCutoff, @MinDiscountDate) > 0 then null" & vbCrLf _
                    & "else DiscCutoff end as datetime) as TakeDiscByDate," & vbCrLf _
                    & "case when DiscCutoff is null then 0.00" & vbCrLf _
                    & "when datediff(d, DiscCutoff, @MinDiscountDate) > 0 then 0.00" & vbCrLf _
                    & "else DiscAmount end as DiscAvail" & vbCrLf _
                    & "From viewOpenAPWrapper" & vbCrLf _
                    & "where VIDUEDATE <= @MaxDueDate" & vbCrLf _
                    & "or DiscCutoff >= @MinDiscountDate" & vbCrLf _
                    & ")"

            ' Execute the sql to alter the view fnGetOpenAP
            Execute True, sSql
            
            'set version
            Execute False, "update Version set Version = " & newver



    End If
End Function


Private Function UpdateDatabase5()
    newver = 64
    If ver < newver Then
        ver = newver

            sSql = "ALTER TABLE [dbo].[Preferences] ADD [HideModuleButton] [tinyint] NULL"
            ' Execute the sql to alter the view fnGetOpenAP
            Execute True, sSql
        
            ' Drop the temp packslip Lot table
            Execute False, "DROP TABLE [dbo].[TempPsLots]"
            
            sSql = "CREATE TABLE [dbo].[TempPsLots](" & vbCrLf _
                        & "[PsNumber] [char](8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf _
                        & "[PsItem] [smallint] NOT NULL," & vbCrLf _
                        & "[LotID] [char](15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf _
                        & "[LotQty] [decimal](12, 4) NOT NULL," & vbCrLf _
                        & "[WhenCreated] [datetime] NOT NULL CONSTRAINT [DF_TempPsLots_WhenCreated]  DEFAULT (getdate())," & vbCrLf _
                        & "[PartRef] [varchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL CONSTRAINT [DF_TempPsLots_PartRef]  DEFAULT ('')," & vbCrLf _
                        & "[LotItemID] [smallint] NULL" & vbCrLf _
                    & ") ON [PRIMARY]"

            ' Execute the sql to alter the view fnGetOpenAP
            Execute True, sSql

            'set version
            Execute False, "UPDATE Version set Version = " & newver
    End If
End Function

Private Function UpdateDatabase6()
    newver = 65
    If ver < newver Then
        ver = newver

            ' Alter the Sohdtable to increate the filed size
            sSql = "ALTER TABLE [dbo].[SohdTable] ALTER COLUMN [SOREMARKS] varchar(6000)"
            ' Execute the sql to alter the view fnGetOpenAP
            Execute True, sSql
        
            ' Alter the Pohdtable to increate the filed size
            sSql = "ALTER TABLE [dbo].[PoHdTable] ALTER COLUMN [POREMARKS] varchar(6000)"
            ' Execute the sql to alter the view fnGetOpenAP
            Execute True, sSql
        
            ' Alter the report lot header to increate the filed size
            sSql = "ALTER TABLE [dbo].[EsReportLots01h] ALTER COLUMN [LotComment] varchar(1020)"
            ' Execute the sql to alter the view fnGetOpenAP
            Execute True, sSql
        
            ' Alter the lot header to increate the filed size
            sSql = "ALTER TABLE [dbo].[EsReportLots01d] ALTER COLUMN [LoiComments] varchar(1020)"
            ' Execute the sql to alter the view fnGetOpenAP
            Execute True, sSql
            
            
            ' Alter the sales view
            sSql = "ALTER VIEW [dbo].[Vw_Sales] " & vbCrLf _
                    & " AS " & vbCrLf _
                    & " SELECT DISTINCT " & vbCrLf _
                    & "            TOP 100 PERCENT dbo.CihdTable.INVCANCELED, dbo.CihdTable.INVTYPE, dbo.CihdTable.INVPRE, dbo.CihdTable.INVNO, dbo.LoitTable.LOICUSTINVNO, " & vbCrLf _
                    & "            dbo.CihdTable.INVCUST, dbo.CustTable.CUNUMBER, dbo.CustTable.CUNAME, dbo.CihdTable.INVDATE, dbo.CihdTable.INVPIF, dbo.PshdTable.PSCANCELED, " & vbCrLf _
                    & "            dbo.PshdTable.PSTYPE, dbo.PsitTable.PIPACKSLIP, dbo.PsitTable.PIITNO, dbo.InvaTable.INPSNUMBER, dbo.InvaTable.INPSITEM, dbo.LoitTable.LOIPSNUMBER, " & vbCrLf _
                    & "            dbo.LoitTable.LOIPSITEM, dbo.SoitTable.ITPSSHIPPED, dbo.SohdTable.SOSALESMAN AS PSSoSalesman, SohdTable_1.SOSALESMAN AS SOSoSlsmn, " & vbCrLf _
                    & "            dbo.SohdTable.SOPO AS PSSoPo, SohdTable_1.SOPO AS SOSoPo, dbo.SohdTable.SODIVISION AS PSSoDiv, SohdTable_1.SODIVISION AS SOSoDiv, " & vbCrLf _
                    & "            dbo.SohdTable.SOREGION AS PSSoReg, SohdTable_1.SOREGION AS SOSoReg, dbo.SohdTable.SOBUSUNIT AS PSSoBu, SohdTable_1.SOBUSUNIT AS SOSoBu, " & vbCrLf _
                    & "            dbo.SprsTable.SPLAST AS PSSlsmnLast, SprsTable_1.SPLAST AS SOSlsmnLast, dbo.SprsTable.SPFIRST AS PSSlsmnFirst, " & vbCrLf _
                    & "            SprsTable_1.SPFIRST AS SOSlsmnFirst, dbo.SprsTable.SPMIDD AS PSSlsmnInit, SprsTable_1.SPMIDD AS SOSlsmnInit, dbo.SohdTable.SOTYPE AS SOSoType, " & vbCrLf _
                    & "            SohdTable_1.SOTYPE AS PSSoType, dbo.SoitTable.ITSO, dbo.SoitTable.ITNUMBER, dbo.SoitTable.ITREV, dbo.SoitTable.ITPART, dbo.PartTable.PARTNUM, " & vbCrLf _
                    & "            dbo.PartTable.PADESC, dbo.PartTable.PALEVEL, dbo.PartTable.PALOTTRACK, dbo.PartTable.PAUSEACTUALCOST, dbo.PartTable.PAUNITS, " & vbCrLf _
                    & "            dbo.PartTable.PAMAKEBUY, dbo.PartTable.PAFAMILY, dbo.PartTable.PAPRODCODE, dbo.PcodTable.PCDESC, dbo.PartTable.PACLASS, dbo.PclsTable.CCDESC, " & vbCrLf _
                    & "            dbo.SoitTable.ITQTY, dbo.SoitTable.ITDOLLARS, dbo.SoitTable.ITADJUST, dbo.SoitTable.ITDISCAMOUNT, dbo.SoitTable.ITCOMMISSION, " & vbCrLf _
                    & "            dbo.SoitTable.ITBOOKDATE, dbo.SoitTable.ITSCHED, dbo.SoitTable.ITACTUAL, dbo.SoitTable.ITCANCELDATE, dbo.SoitTable.ITCANCELED, " & vbCrLf _
                    & "            dbo.SoitTable.ITREVACCT, dbo.GlacTable.GLACCTNO AS RevAcct, dbo.GlacTable.GLDESCR AS RevAcctDesc, dbo.SoitTable.ITCGSACCT, " & vbCrLf _
                    & "            dbo.SoitTable.ITDISACCT, dbo.SoitTable.ITSTATE, dbo.SoitTable.ITTAXCODE, dbo.InvaTable.INTYPE, dbo.InvaTable.INAQTY, dbo.InvaTable.INAMT, " & vbCrLf _
                    & "            dbo.InvaTable.INTOTLABOR, dbo.InvaTable.INTOTMATL, dbo.InvaTable.INTOTEXP, dbo.InvaTable.INTOTOH, dbo.PartTable.PASTDCOST, " & vbCrLf _
                    & "            dbo.PartTable.PATOTCOST, dbo.PartTable.PATOTLABOR, dbo.PartTable.PATOTMATL, dbo.PartTable.PATOTEXP, dbo.PartTable.PATOTOH, " & vbCrLf _
                    & "            dbo.InvaTable.INDRLABACCT, dbo.InvaTable.INDRMATACCT, dbo.InvaTable.INDREXPACCT, dbo.InvaTable.INDROHDACCT, dbo.InvaTable.INCRLABACCT, " & vbCrLf _
                    & "            dbo.InvaTable.INCRMATACCT, dbo.InvaTable.INCREXPACCT, dbo.InvaTable.INCROHDACCT, dbo.LoitTable.LOITYPE, dbo.LoitTable.LOINUMBER, " & vbCrLf _
                    & "            dbo.LoitTable.LOIRECORD, dbo.LoitTable.LOIQUANTITY, dbo.LoitTable.LOIADATE, dbo.LohdTable.LOTNUMBER, dbo.LohdTable.LOTUSERLOTID, " & vbCrLf _
                    & "            dbo.LohdTable.LOTPARTREF, dbo.LohdTable.LOTDATECOSTED, dbo.LohdTable.LOTTOTLABOR, dbo.LohdTable.LOTTOTMATL, dbo.LohdTable.LOTTOTEXP, " & vbCrLf _
                    & "            dbo.LohdTable.LOTTOTOH , dbo.LohdTable.LOTUNITCOST "

        sSql = sSql & "  FROM  dbo.GlacTable RIGHT OUTER JOIN " & vbCrLf _
                    & "            dbo.SprsTable RIGHT OUTER JOIN " & vbCrLf _
                    & "            dbo.SoitTable LEFT OUTER JOIN " & vbCrLf _
                    & "            dbo.SohdTable ON dbo.SoitTable.ITSO = dbo.SohdTable.SONUMBER ON dbo.SprsTable.SPNUMBER = dbo.SohdTable.SOSALESMAN RIGHT OUTER JOIN " & vbCrLf _
                    & "            dbo.CustTable RIGHT OUTER JOIN " & vbCrLf _
                    & "            dbo.CihdTable ON dbo.CustTable.CUREF = dbo.CihdTable.INVCUST ON dbo.SoitTable.ITINVOICE = dbo.CihdTable.INVNO LEFT OUTER JOIN " & vbCrLf _
                    & "            dbo.SprsTable AS SprsTable_1 RIGHT OUTER JOIN " & vbCrLf _
                    & "            dbo.SohdTable AS SohdTable_1 ON SprsTable_1.SPNUMBER = SohdTable_1.SOSALESMAN ON " & vbCrLf _
                    & "            dbo.CihdTable.INVSO = SohdTable_1.SONUMBER LEFT OUTER JOIN " & vbCrLf _
                    & "            dbo.PcodTable RIGHT OUTER JOIN " & vbCrLf _
                    & "            dbo.PclsTable RIGHT OUTER JOIN " & vbCrLf _
                    & "            dbo.PartTable ON dbo.PclsTable.CCREF = dbo.PartTable.PACLASS ON dbo.PcodTable.PCREF = dbo.PartTable.PAPRODCODE ON " & vbCrLf _
                    & "            dbo.SoitTable.ITPART = dbo.PartTable.PARTREF ON dbo.GlacTable.GLACCTREF = dbo.SoitTable.ITREVACCT LEFT OUTER JOIN " & vbCrLf _
                    & "            dbo.PshdTable RIGHT OUTER JOIN " & vbCrLf _
                    & "            dbo.PsitTable ON dbo.PshdTable.PSNUMBER = dbo.PsitTable.PIPACKSLIP LEFT OUTER JOIN " & vbCrLf _
                    & "            dbo.InvaTable LEFT OUTER JOIN " & vbCrLf _
                    & "            dbo.LoitTable INNER JOIN " & vbCrLf _
                    & "            dbo.LohdTable ON dbo.LoitTable.LOINUMBER = dbo.LohdTable.LOTNUMBER ON dbo.InvaTable.INAQTY = dbo.LoitTable.LOIQUANTITY AND " & vbCrLf _
                    & "            dbo.InvaTable.INPSNUMBER = dbo.LoitTable.LOIPSNUMBER AND dbo.InvaTable.INPSITEM = dbo.LoitTable.LOIPSITEM ON " & vbCrLf _
                    & "            dbo.PsitTable.PIPACKSLIP = dbo.InvaTable.INPSNUMBER AND dbo.PsitTable.PIITNO = dbo.InvaTable.INPSITEM ON " & vbCrLf _
                    & "            dbo.SoitTable.ITSO = dbo.PsitTable.PISONUMBER AND dbo.SoitTable.ITNUMBER = dbo.PsitTable.PISOITEM AND " & vbCrLf _
                    & "            dbo.SoitTable.ITREV = dbo.PsitTable.PISOREV " & vbCrLf _
                    & " ORDER BY dbo.CihdTable.INVDATE, dbo.CihdTable.INVNO, dbo.InvaTable.INPSNUMBER, dbo.InvaTable.INPSITEM, dbo.SoitTable.ITSO, dbo.SoitTable.ITNUMBER, " & vbCrLf _
                    & "            dbo.SoitTable.ITREV "

            ' Execute the sql to alter the view Vw_Sales
            Execute True, sSql
            
            ' Drop the maintain part QOH
            Execute False, "DROP TABLE [dbo].[MaintPAQOH]"
            
            sSql = "CREATE TABLE [dbo].[MaintPAQOH]( " & vbCrLf _
                   & "  [PARTREF] [char](30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & vbCrLf _
                   & "  [PARTNUM] [char](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
                   & "  [CURPAQOH] [decimal](12, 4) NULL, " & vbCrLf _
                   & "  [PREPAQOH] [decimal](12, 4) NULL, " & vbCrLf _
                   & "  [PALOTREMAININGQTY] [decimal](12, 4) NULL " & vbCrLf _
                & "  ) ON [PRIMARY]"

            ' Execute to create the table
            Execute True, sSql
            
            
            If (Not TableExists("sfcdTable")) Then
                sSql = "CREATE TABLE [dbo].[sfcdTable](" & vbCrLf _
                            & "  [SFREF] [varchar](6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & vbCrLf _
                            & "  [SFCODE] [varchar](6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & vbCrLf _
                            & "  [SFDESC] [varchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
                            & "  [SFSTHR] [varchar](6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
                            & "  [SFENHR] [varchar](6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
                            & "  [SFLUNSTHR] [varchar](6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
                            & "  [SFLUNENHR] [varchar](6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & vbCrLf _
                            & "  [SFADJHR] [decimal](7, 2) NULL, " & vbCrLf _
                            & "  [SFRNDHR] [decimal](7, 2) NULL, " & vbCrLf _
                            & "   CONSTRAINT [PK_sfcdTable] PRIMARY KEY CLUSTERED " & vbCrLf _
                            & "  ( " & vbCrLf _
                            & "      [SFREF] Asc " & vbCrLf _
                            & "  ) ON [PRIMARY] " & vbCrLf _
                            & "  ) ON [PRIMARY] "
                        
                ' Execute to create the table
                Execute True, sSql
            End If

            If (Not TableExists("sfempTable")) Then
                sSql = "CREATE TABLE [dbo].[sfempTable]( " & vbCrLf _
                    & "  [SFREF] [varchar](6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & vbCrLf _
                    & "  [PREMNUMBER] [int] NOT NULL, " & vbCrLf _
                    & "  [STARTDATE] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & vbCrLf _
                    & "  [ENDDATE] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL " & vbCrLf _
                    & "     ) ON [PRIMARY]"
                
                ' Execute to create the table
                Execute True, sSql
            End If
            
            
            ' Drop the shift code view
            Execute False, "DROP VIEW [dbo].[viewShiftCdEmployeeDetail]"
            
            sSql = "CREATE VIEW [dbo].[viewShiftCdEmployeeDetail]" & vbCrLf _
            & "  AS " & vbCrLf _
            & "  SELECT dbo.sfcdTable.SFREF, dbo.sfcdTable.SFCODE, dbo.sfempTable.PREMNUMBER, dbo.EmplTable.PREMLSTNAME, dbo.EmplTable.PREMFSTNAME, " & vbCrLf _
                 & "    dbo.sfcdTable.SFDESC, dbo.sfcdTable.SFSTHR, dbo.sfcdTable.SFENHR, dbo.sfcdTable.SFLUNSTHR, dbo.sfcdTable.SFLUNENHR, " & vbCrLf _
                 & "    dbo.sfcdTable.SFADJHR , dbo.sfcdTable.SFRNDHR " & vbCrLf _
            & "  FROM dbo.EmplTable INNER JOIN " & vbCrLf _
                 & "       dbo.sfempTable ON dbo.EmplTable.PREMNUMBER = dbo.sfempTable.PREMNUMBER LEFT OUTER JOIN " & vbCrLf _
                 & "       dbo.sfcdTable ON dbo.sfempTable.SFREF = dbo.sfcdTable.SFREF "

            ' Execute to create the table
            Execute True, sSql
            
            
            ' Drop the shift code view
            Execute False, "DROP TABLE [dbo].[EsReportClosedRunsLog]"
            
            sSql = "CREATE TABLE [dbo].[EsReportClosedRunsLog]( " & vbCrLf _
                        & "   [LOG_NUMBER] [smallint] NULL CONSTRAINT [DF__EsReportC__LOG_N__46829927]  DEFAULT ((0)), " & vbCrLf _
                        & "   [LOG_TEXT] [varchar](80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL CONSTRAINT [DF__EsReportC__LOG_T__4776BD60]  DEFAULT ('') " & vbCrLf _
                        & ") ON [PRIMARY] "
            ' Execute to create the table
            Execute True, sSql
            
            If Not ColumnExists("ComnTable", "COPOTIMESERVOP") Then
               Execute True, "ALTER TABLE [dbo].[ComnTable] ADD [COPOTIMESERVOP] [int] NULL"
               'Execute True, sSql
            End If

            'set version
            Execute False, "UPDATE Version set Version = " & newver
    End If
End Function


Private Function UpdateDatabase7()
    newver = 66
    If ver < newver Then
        ver = newver
            
            ' Alter Lohdtable to include new date
            sSql = "ALTER VIEW [dbo].[ViewLohdPartTable] " & vbCrLf _
                    & " AS " & vbCrLf _
                    & " SELECT dbo.LohdTable.LOTNUMBER, dbo.PartTable.PARTREF, dbo.PartTable.PADESC, dbo.PartTable.PAEXTDESC, dbo.LohdTable.LOTUSERLOTID, " & vbCrLf _
                    & "        dbo.LohdTable.LOTADATE, dbo.LohdTable.LOTORIGINALQTY, dbo.LohdTable.LOTREMAININGQTY, dbo.LohdTable.LOTUNITCOST, " & vbCrLf _
                    & "        dbo.LohdTable.LOTDATECOSTED, dbo.PartTable.PACLASS, dbo.PartTable.PAPRODCODE, dbo.PartTable.PASTDCOST, dbo.LohdTable.LOTTOTMATL, " & vbCrLf _
                    & "        dbo.LohdTable.LOTTOTLABOR, dbo.LohdTable.LOTTOTEXP, dbo.LohdTable.LOTTOTOH, dbo.PartTable.PAUSEACTUALCOST, " & vbCrLf _
                    & "        dbo.PartTable.PALOTTRACK, dbo.PartTable.PATOTOH, dbo.PartTable.PALABOR, dbo.PartTable.PATOTEXP, dbo.PartTable.PATOTMATL, " & vbCrLf _
                    & "        dbo.PartTable.PALEVEL , dbo.PartTable.PARTNUM " & vbCrLf _
                    & " FROM  dbo.LohdTable LEFT OUTER JOIN " & vbCrLf _
                    & "        dbo.PartTable ON dbo.LohdTable.LOTPARTREF = dbo.PartTable.PARTREF " & vbCrLf _
                    & " Where (dbo.PartTable.PALEVEL <= 4) "

            ' Execute the sql to alter the view LohdPartTable
            Execute True, sSql
            
            
        Execute False, "DROP PROCEDURE InvMRPExcessReport"
            
        sSql = "CREATE PROCEDURE [dbo].[InvMRPExcessReport] " & vbCrLf _
                & " @PartClass as Varchar(16), @PartCode as varchar(8), @PartType1 as Integer, " & vbCrLf _
                & " @PartType2 as Integer, @PartType3 as Integer, @PartType4 as Integer " & vbCrLf _
                & " AS  " & vbCrLf _
                & " BEGIN " & vbCrLf _
                & "     declare @mrpPartRef as varchar(30) " & vbCrLf _
                & "     declare @QtyRem as integer " & vbCrLf & vbCrLf _
                & "    IF (@PartClass = 'ALL') " & vbCrLf _
                & "     BEGIN                   " & vbCrLf _
                & "         SET @PartClass = '' " & vbCrLf _
                & "     End                     " & vbCrLf & vbCrLf _
                & "     IF (@PartCode = 'ALL')  " & vbCrLf _
                & "     BEGIN                   " & vbCrLf _
                & "         SET @PartCode = ''  " & vbCrLf _
                & "     End                     " & vbCrLf _
                & "     IF (@PartType1 = 1)     " & vbCrLf _
                & "         SET @PartType1 = 1  " & vbCrLf _
                & "     Else                    " & vbCrLf _
                & "         SET @PartType1 = 0  " & vbCrLf _
                & "     IF (@PartType2 = 1)     " & vbCrLf _
                & "         SET @PartType2 = 2  " & vbCrLf _
                & "     Else                    " & vbCrLf _
                & "         SET @PartType2 = 0  " & vbCrLf _
                & "     IF (@PartType3 = 1)     " & vbCrLf _
                & "         SET @PartType3 = 3  " & vbCrLf
    sSql = sSql & "    Else                    " & vbCrLf _
                & "         SET @PartType3 = 0  " & vbCrLf _
                & "     IF (@PartType4 = 1)     " & vbCrLf _
                & "         SET @PartType4 = 4  " & vbCrLf _
                & "     Else                    " & vbCrLf _
                & "         SET @PartType4 = 0  " & vbCrLf _
                & "     CREATE TABLE #tempMrpExRpt          " & vbCrLf _
                & "     (                                   " & vbCrLf _
                & "         PACLASS varchar(4) NULL ,       " & vbCrLf _
                & "         PAPRODCODE varchar(6) NULL ,    " & vbCrLf _
                & "         PALEVEL tinyint NULL ,          " & vbCrLf _
                & "         PARTREF varchar(30) NULL ,      " & vbCrLf _
                & "         PARTNUM varchar(30) NULL ,      " & vbCrLf _
                & "         PADESC varchar(30) NULL ,       " & vbCrLf _
                & "         PAEXTDESC varchar(3072) NULL ,  " & vbCrLf _
                & "         LOTNUMBER varchar(15) NULL,     " & vbCrLf _
                & "         LOTUSERLOTID varchar(40) NULL,  " & vbCrLf _
                & "         MRP_QTYREM int NULL,            " & vbCrLf _
                & "         LOTUNITCOST decimal(12,4) NULL , " & vbCrLf _
                & "         PASTDCOST decimal(12,4) NULL ,  " & vbCrLf _
                & "         PAUSEACTUALCOST tinyint NULL ,  " & vbCrLf _
                & "         PALOTTRACK tinyint NULL         " & vbCrLf _
                & "     )                                   " & vbCrLf _
                & "     DECLARE curInv CURSOR               " & vbCrLf _
                & "     LOCAL                               " & vbCrLf
    sSql = sSql & "     Scroll                              " & vbCrLf _
                & "    FOR                                     " & vbCrLf _
                & "        SELECT mrp_partref, SUM(mrp_partqtyrqd) as rem from " & vbCrLf _
                & "            MrplTable                                       " & vbCrLf _
                & "        --WHERE mrp_partref LIKE 'BNPL%'                    " & vbCrLf _
                & "        GROUP BY mrp_partref                                " & vbCrLf _
                & "            Having Sum(mrp_partqtyrqd) >= 1                 " & vbCrLf _
                & "    OPEN curInv                                             " & vbCrLf _
                & "    FETCH NEXT FROM curInv INTO @mrpPartRef, @QtyRem        " & vbCrLf _
                & "    WHILE (@@FETCH_STATUS <> -1)                            " & vbCrLf _
                & "    BEGIN                                                   " & vbCrLf _
                & "        IF (@@FETCH_STATUS <> -2)                           " & vbCrLf _
                & "        BEGIN                                               " & vbCrLf _
                & "        INSERT INTO #tempMrpExRpt (PACLASS, PAPRODCODE, PALEVEL,                " & vbCrLf _
                & "            PARTREF, PARTNUM, PADESC, PAEXTDESC, LOTNUMBER, LOTUSERLOTID,       " & vbCrLf _
                & "            MRP_QTYREM, LOTUNITCOST, PASTDCOST, PAUSEACTUALCOST, PALOTTRACK)    " & vbCrLf _
                & "            SELECT top 1 PACLASS, PAPRODCODE, PALEVEL, PARTREF, PARTNUM, PADESC, " & vbCrLf _
                & "                PAEXTDESC, LOTNUMBER, LOTUSERLOTID, @QtyRem,LOTUNITCOST,        " & vbCrLf _
                & "                PASTDCOST , PAUSEACTUALCOST, PALOTTRACK                         " & vbCrLf _
                & "            From ViewLohdPartTable                                              " & vbCrLf _
                & "                WHERE partRef = @mrpPartRef                                     " & vbCrLf _
                & "                AND PACLASS LIKE '%' + @PartClass + '%'                         " & vbCrLf _
                & "                AND PAPRODCODE LIKE '%' + @PartCode + '%'                       " & vbCrLf _
                & "                AND PALEVEL IN (@PartType1, @PartType2, @PartType3, @PartType4) " & vbCrLf _
                & "        End                                                                     " & vbCrLf
    sSql = sSql & "        FETCH NEXT FROM curInv INTO @mrpPartRef, @QtyRem                     " & vbCrLf _
                & "    End                                                                         " & vbCrLf _
                & "    CLOSE curInv   --// close the cursor                                        " & vbCrLf _
                & "    DEALLOCATE curInv                                                           " & vbCrLf _
                & "    SELECT PACLASS, PAPRODCODE, PALEVEL,                                        " & vbCrLf _
                & "            PARTREF, PARTNUM, PADESC, PAEXTDESC, LOTNUMBER, LOTUSERLOTID,       " & vbCrLf _
                & "            MRP_QTYREM , LOTUNITCOST, PASTDCOST, PAUSEACTUALCOST, PALOTTRACK    " & vbCrLf _
                & "    FROM #tempMrpExRpt                                                          " & vbCrLf _
                & "    DROP table #tempMrpExRpt                                                    " & vbCrLf _
                & " End                                                                             " & vbCrLf
            
            ' Execute to create the Inventory Excess stpre procedure
            Execute True, sSql
            

            ' Execute to create the Inventory Excess stpre procedure
            Execute False, "DROP PROCEDURE InventoryExcessReport"
    sSql = " CREATE PROCEDURE [dbo].[InventoryExcessReport] " & vbCrLf _
           & "          @BeginDate as varchar(16), @EndDate as varchar(16), @PartClass as Varchar(16)," & vbCrLf _
           & "          @PartCode as varchar(8), @InclZQty as Integer, @PartType1 as Integer," & vbCrLf _
           & "          @PartType2 as Integer, @PartType3 as Integer, @PartType4 as Integer" & vbCrLf _
           & "      AS                                   " & vbCrLf _
           & "      BEGIN                                " & vbCrLf _
           & "                                           " & vbCrLf _
           & "          declare @sqlZQty as varchar(12)  " & vbCrLf _
           & "                                           " & vbCrLf _
           & "          IF (@PartClass = 'ALL')          " & vbCrLf _
           & "          BEGIN                            " & vbCrLf _
           & "              SET @PartClass = ''          " & vbCrLf _
           & "          End                              " & vbCrLf _
           & "          IF (@PartCode = 'ALL')           " & vbCrLf _
           & "          BEGIN                            " & vbCrLf _
           & "              SET @PartCode = ''           " & vbCrLf _
           & "          End                              " & vbCrLf _
           & "                                           " & vbCrLf _
           & "          IF (@PartType1 = 1)              " & vbCrLf _
           & "              SET @PartType1 = 1           " & vbCrLf _
           & "          Else                             " & vbCrLf _
           & "              SET @PartType1 = 0           " & vbCrLf _
           & "                                           " & vbCrLf _
           & "          IF (@PartType2 = 1)              " & vbCrLf _
           & "              SET @PartType2 = 2           " & vbCrLf
           
    sSql = sSql & "     Else                             " & vbCrLf _
                & "        SET @PartType2 = 0               " & vbCrLf _
                & "                                         " & vbCrLf _
                & "    IF (@PartType3 = 1)                  " & vbCrLf _
                & "        SET @PartType3 = 3               " & vbCrLf _
                & "    Else                                 " & vbCrLf _
                & "        SET @PartType3 = 0               " & vbCrLf _
                & "                                         " & vbCrLf _
                & "    IF (@PartType4 = 1)                  " & vbCrLf _
                & "        SET @PartType4 = 4               " & vbCrLf _
                & "    Else                                 " & vbCrLf _
                & "        SET @PartType4 = 0               " & vbCrLf _
                & "                                         " & vbCrLf _
                & "    CREATE TABLE #tempExRpt              " & vbCrLf _
                & "    (                                    " & vbCrLf _
                & "        PACLASS varchar(4) NULL ,        " & vbCrLf _
                & "        PAPRODCODE varchar(6) NULL ,     " & vbCrLf _
                & "        PALEVEL tinyint NULL ,           " & vbCrLf _
                & "        PADESC varchar(30) NULL ,        " & vbCrLf _
                & "        PAEXTDESC varchar(3072) NULL ,   " & vbCrLf _
                & "        INPART varchar(30) NULL ,        " & vbCrLf _
                & "        INNUMBER int NULL ,              " & vbCrLf _
                & "        INTYPE int NULL ,                " & vbCrLf _
                & "        INAMT decimal(12,4) NULL ,       " & vbCrLf _
                & "        LOTUNITCOST decimal(12,4) NULL , " & vbCrLf
                        
    sSql = sSql & "        LOTNUMBER varchar(15) NULL,      " & vbCrLf _
                & "        LOTUSERLOTID varchar(40) NULL,      " & vbCrLf _
                & "        LOIQUANTITY decimal(12,4) NULL ,     " & vbCrLf _
                & "        LOTREMAININGQTY decimal(12,4) NULL , " & vbCrLf _
                & "        INADATE smalldatetime NULL ,         " & vbCrLf _
                & "        LOTADATE smalldatetime NULL ,        " & vbCrLf _
                & "        LOIMOPARTREF varchar(30) NULL        " & vbCrLf _
                & "    )                                         " & vbCrLf _
                & "                                             " & vbCrLf _
                & "    IF (@InclZQty = 1)                       " & vbCrLf _
                & "                                             " & vbCrLf _
                & "        INSERT INTO #tempExRpt (PACLASS, PAPRODCODE, PALEVEL,                            " & vbCrLf _
                & "            PADESC, PAEXTDESC, INPART , INNUMBER, INTYPE, INAMT,                         " & vbCrLf _
                & "            LOTUNITCOST, LOTNUMBER, LOTUSERLOTID, LOIQUANTITY,                           " & vbCrLf _
                & "            LOTREMAININGQTY, INADATE, LOTADATE, LOIMOPARTREF)                            " & vbCrLf _
                & "        SELECT PACLASS, PAPRODCODE, PALEVEL, PADESC, PAEXTDESC, a.INPART , a.INNUMBER,   " & vbCrLf _
                & "            a.INTYPE, a.INAMT, LOTUNITCOST, LOTNUMBER, LOTUSERLOTID,                     " & vbCrLf _
                & "            LOIQUANTITY , LOTREMAININGQTY, a.INADATE, LOTADATE, LOIMOPARTREF             " & vbCrLf _
                & "        FROM invaTable a, LoitTable, ViewLohdPartTable                                   " & vbCrLf _
                & "        Where a.INPART = LoitTable.LOIPARTREF                                            " & vbCrLf _
                & "            AND ViewLohdPartTable.partref = a.INPART                                     " & vbCrLf _
                & "            AND LoitTable.LOIPARTREF = ViewLohdPartTable.partref                         " & vbCrLf _
                & "            AND ViewLohdPartTable.LOTNUMBER = LoitTable.LOINUMBER                        " & vbCrLf _
                & "            AND a.INNUMBER = LoitTable.LOIACTIVITY                                       " & vbCrLf _
                & "            AND ViewLohdPartTable.PACLASS LIKE '%' + @PartClass + '%'                    " & vbCrLf
                
    sSql = sSql & "            AND ViewLohdPartTable.PAPRODCODE LIKE '%' + @PartCode + '%'                      " & vbCrLf _
                & "            AND ViewLohdPartTable.PALEVEL IN (@PartType1, @PartType2, @PartType3, @PartType4) " & vbCrLf _
                & "            AND a.INPART NOT IN                         " & vbCrLf _
                & "                (SELECT INPART FROM invaTable b         " & vbCrLf _
                & "                Where a.INPART = b.INPART               " & vbCrLf _
                & "                    AND INADATE BETWEEN @BeginDate and @EndDate " & vbCrLf _
                & "                    AND INTYPE IN (1, 3, 4, 6, 7, 9, 10, 11,15,17,19,23,25,26,32))  " & vbCrLf _
                & "        AND a.INADATE =                                 " & vbCrLf _
                & "                (SELECT MAX(INADATE) FROM invaTable c   " & vbCrLf _
                & "                Where C.INPART = a.INPART               " & vbCrLf _
                & "                    AND c.INADATE < DATEADD(dd, -1 , @BeginDate)    " & vbCrLf _
                & "                Group by c.INPART)                      " & vbCrLf _
                & "        order by a.INPART                               " & vbCrLf _
                & "    Else                                                " & vbCrLf _
                & "        INSERT INTO #tempExRpt (PACLASS, PAPRODCODE, PALEVEL,PADESC, PAEXTDESC, " & vbCrLf _
                & "            INPART, INNUMBER, INTYPE, INAMT, LOTUNITCOST, LOTNUMBER,            " & vbCrLf _
                & "            LOTUSERLOTID, LOIQUANTITY, LOTREMAININGQTY, INADATE,                " & vbCrLf _
                & "            LOTADATE, LOIMOPARTREF)                                             " & vbCrLf _
                & "        SELECT PACLASS, PAPRODCODE, PALEVEL, PADESC, PAEXTDESC, a.INPART , a.INNUMBER,  " & vbCrLf _
                & "            a.INTYPE, a.INAMT, LOTUNITCOST, LOTNUMBER, LOTUSERLOTID,           " & vbCrLf _
                & "            LOIQUANTITY , LOTREMAININGQTY, a.INADATE, LOTADATE, LOIMOPARTREF    " & vbCrLf _
                & "        FROM invaTable a, LoitTable, ViewLohdPartTable                          " & vbCrLf _
                & "        Where a.INPART = LoitTable.LOIPARTREF                                   " & vbCrLf _
                & "            AND ViewLohdPartTable.partref = a.INPART                            " & vbCrLf _
                & "            AND LoitTable.LOIPARTREF = ViewLohdPartTable.partref                " & vbCrLf
                            
   sSql = sSql & "            AND ViewLohdPartTable.LOTNUMBER = LoitTable.LOINUMBER               " & vbCrLf _
                & "            AND a.INNUMBER = LoitTable.LOIACTIVITY                              " & vbCrLf _
                & "            AND ViewLohdPartTable.PACLASS LIKE '%' + @PartClass + '%'           " & vbCrLf _
                & "            AND ViewLohdPartTable.PAPRODCODE LIKE '%' + @PartCode + '%'         " & vbCrLf _
                & "            AND ViewLohdPartTable.PALEVEL IN (@PartType1, @PartType2, @PartType3, @PartType4)   " & vbCrLf _
                & "            AND a.INPART NOT IN                                     " & vbCrLf _
                & "                (SELECT INPART FROM invaTable b                     " & vbCrLf _
                & "                Where a.INPART = b.INPART                           " & vbCrLf _
                & "                    AND INADATE BETWEEN @BeginDate and @EndDate     " & vbCrLf _
                & "                    AND INTYPE IN (1, 3, 4, 6, 7, 9, 10, 11,15,17,19,23,25,26,32))  " & vbCrLf _
                & "        AND a.INADATE =                                             " & vbCrLf _
                & "                (SELECT MAX(INADATE) FROM invaTable c               " & vbCrLf _
                & "                Where C.INPART = a.INPART                           " & vbCrLf _
                & "                    AND c.INADATE < DATEADD(dd, -1 , @BeginDate)    " & vbCrLf _
                & "                Group by c.INPART)                                  " & vbCrLf _
                & "        AND a.INAQTY > 0                                            " & vbCrLf _
                & "            order by a.INPART                                       " & vbCrLf _
                & "                                                                    " & vbCrLf _
                & "    SELECT PACLASS, PAPRODCODE, PALEVEL, PADESC, PAEXTDESC, INPART, " & vbCrLf _
                & "        INPART, INNUMBER, INTYPE, INAMT, LOTUNITCOST,               " & vbCrLf _
                & "        LOTNUMBER, LOTUSERLOTID, LOIQUANTITY,                       " & vbCrLf _
                & "        LOTREMAININGQTY , INADATE, LOTADATE, LOIMOPARTREF           " & vbCrLf _
                & "    FROM #tempExRpt                                                 " & vbCrLf _
                & "        WHERE INPART NOT IN                                         " & vbCrLf _
                & "                (SELECT DISTINCT mrp_Partref FROM dbo.MrplTable     " & vbCrLf
                                    
    sSql = sSql & "         WHERE mrp_type IN (2, 3, 4, 11, 12, 17)             " & vbCrLf _
                & "                        AND mrp_partDateRQD < DATEADD(dd, +1 , @EndDate))   " & vbCrLf _
                & "    DROP table #tempExRpt                                                   " & vbCrLf _
                & " End                                                                         " & vbCrLf

            ' Execute to create the Inventory Excess stpre procedure
            Execute True, sSql
            
            
            ' Modified the Vendor table
            If Not ColumnExists("VndrTable", "VEAPPROVREQ") Then
               Execute True, "ALTER TABLE [dbo].[VndrTable] ADD VEAPPROVREQ tinyint NULL"
            End If
            If Not ColumnExists("VndrTable", "VESURVEY") Then
               Execute True, "ALTER TABLE [dbo].[VndrTable] ADD VESURVEY tinyint NULL"
            End If
            
            If Not ColumnExists("VndrTable", "VESURVSENT") Then
               Execute True, "ALTER TABLE [dbo].[VndrTable] ADD VESURVSENT smalldatetime NULL"
            End If
            
            If Not ColumnExists("VndrTable", "VESURVREC") Then
               Execute True, "ALTER TABLE [dbo].[VndrTable] ADD VESURVREC smalldatetime NULL"
            End If
            
            If Not ColumnExists("VndrTable", "VEAPPDATE") Then
               Execute True, "ALTER TABLE [dbo].[VndrTable] ADD VEAPPDATE smalldatetime NULL"
            End If
            
            If Not ColumnExists("VndrTable", "VEREVIEWDT") Then
               Execute True, "ALTER TABLE [dbo].[VndrTable] ADD VEREVIEWDT smalldatetime NULL"
            End If
            
            
            sSql = " INSERT INTO StatCdType " & vbCrLf _
                & "    (STATCODE_TYPE_REF, STATCODE_TYPE_NAME, STATCODE_TYPE_UNIQUE_KEY) " & vbCrLf _
                & "  VALUES('VE', 'Vendor Record', 1)"
            ' Insert Vendor ID
            Execute True, sSql
            
            
            sSql = " ALTER PROCEDURE [dbo].[Qry_GetVendorBasics] " & vbCrLf _
                   & "     (@vendornick char(10))                " & vbCrLf _
                   & " AS                                            " & vbCrLf _
                   & " SELECT VEREF,VENICKNAME,VEBNAME,VEAPPROVREQ,VESURVEY, VESURVSENT,VESURVREC,VEAPPDATE,VEREVIEWDT FROM VndrTable  " & vbCrLf _
                   & " WHERE VEREF=@vendornick " & vbCrLf

            ' Alter Vendor detail SP
            Execute True, sSql
            
            
            sSql = " ALTER PROCEDURE [dbo].[Qry_AddInternStatCode]                      " & vbCrLf _
                   & "  (@StatCmtRef int,@StatCmtRef1 int, @StatCmtRef2 varchar(20),        " & vbCrLf _
                   & "  @StatRef varchar(4), @StatCodeTypeRef varchar(3), @user varchar(50), " & vbCrLf _
                   & "  @comments as varchar(1024), @ActStat as int)    " & vbCrLf _
                   & "  AS                                              " & vbCrLf _
                   & "  BEGIN                                           " & vbCrLf _
                   & "                                                  " & vbCrLf _
                   & "      declare @StatCdTypeKey as int               " & vbCrLf _
                   & "      declare @curDate as datetime                " & vbCrLf _
                   & "      declare @count as integer                   " & vbCrLf _
                   & "                                                  " & vbCrLf _
                   & "      SET @StatCdTypeKey = 0                      " & vbCrLf _
                   & "      SELECT @StatCdTypeKey = ISNULL(STATCODE_TYPE_UNIQUE_KEY, 0) FROM StatCdType " & vbCrLf _
                   & "          WHERE STATCODE_TYPE_REF = @StatCodeTypeRef                              " & vbCrLf _
                   & "                                  " & vbCrLf _
                   & "      SET @curDate = GetDATE()    " & vbCrLf _
                   & "                                  " & vbCrLf _
                   & "      IF @StatCdTypeKey = 1       " & vbCrLf _
                   & "      BEGIN                       " & vbCrLf _
                   & "          SET @StatCmtRef1 = 0    " & vbCrLf _
                   & "          SET @StatCmtRef2 = ''   " & vbCrLf _
                   & "                                  " & vbCrLf _
                   & "          SELECT @count = COunt(*) FROM StCmtTable WHERE  " & vbCrLf _
                   & "          STATUS_REF = @StatRef AND STATCODE_TYPE_REF = @StatCodeTypeRef  " & vbCrLf _
                   & "          AND STATUS_CMT_REF = @StatCmtRef    " & vbCrLf
            sSql = sSql & " End                                    " & vbCrLf _
                        & " IF @StatCdTypeKey = 3                       " & vbCrLf _
                        & " BEGIN                                       " & vbCrLf _
                        & "     SELECT @count = COunt(*) FROM StCmtTable WHERE  " & vbCrLf _
                        & "     STATUS_REF = @StatRef AND STATCODE_TYPE_REF = @StatCodeTypeRef          " & vbCrLf _
                        & "     AND STATUS_CMT_REF = @StatCmtRef                                        " & vbCrLf _
                        & "     AND STATUS_CMT_REF1 = @StatCmtRef1  AND STATUS_CMT_REF2 = @StatCmtRef2  " & vbCrLf _
                        & " End                                                                         " & vbCrLf _
                        & "                                                                             " & vbCrLf _
                        & " IF  @count = 0                                                              " & vbCrLf _
                        & "     INSERT INTO StCmtTable (STATUS_REF, STATCODE_TYPE_REF, STATUS_CMT_REF,  " & vbCrLf _
                        & "     STATUS_CMT_REF1, STATUS_CMT_REF2, STATUS_ORG_USER, STATUS_ORG_DATE,     " & vbCrLf _
                        & "     COMMENT, STATUS_CUR_USER, STATUS_CUR_DATE, STATUS_ACT_STATE)            " & vbCrLf _
                        & "     VALUES (@StatRef, @StatCodeTypeRef, @StatCmtRef, @StatCmtRef1, @StatCmtRef2,    " & vbCrLf _
                        & "     @user, @curDate,@comments, @user, @curDate, @ActStat)                   " & vbCrLf _
                        & " Else                                                                        " & vbCrLf _
                        & "     UPDATE StCmtTable SET STATUS_CUR_USER = @user, STATUS_CUR_DATE = @curDate,  " & vbCrLf _
                        & "         COMMENT = @comments, STATUS_ACT_STATE = @ActStat                        " & vbCrLf _
                        & "         WHERE STATUS_REF = @StatRef AND STATCODE_TYPE_REF = @StatCodeTypeRef    " & vbCrLf _
                        & "         AND STATUS_CMT_REF = @StatCmtRef    " & vbCrLf _
                        & "         AND STATUS_CMT_REF1 = @StatCmtRef1  " & vbCrLf _
                        & "         AND STATUS_CMT_REF2 = @StatCmtRef2  " & vbCrLf _
                        & "                                             " & vbCrLf _
                        & " SELECT @@ROWCOUNT                           " & vbCrLf _
                    & " End                                             " & vbCrLf
                
            ' Alter Vendor detail SP
            Execute True, sSql
            
            ' Add new Column to the Preference table. Global flag to enable vendor approval
            If Not ColumnExists("Preferences", "ReqVendorApproval") Then
               Execute True, "ALTER TABLE [dbo].[Preferences] ADD ReqVendorApproval tinyint NULL"
            End If
        

            'set version
            Execute False, "UPDATE Version set Version = " & newver
    End If
End Function

Private Function UpdateDatabase8()
    newver = 67
    If ver < newver Then
        ver = newver

        ' Add new Column to the PIPORIGDATE to PoitTable table
        If Not ColumnExists("PoitTable", "PIPORIGDATE") Then
           Execute True, "ALTER TABLE dbo.PoitTable ADD PIPORIGDATE smalldatetime NULL"
        End If
        
        ' Update PoitTable
        sSql = "UPDATE PoitTable SET PIPORIGDATE=PIPDATE WHERE PIPORIGDATE IS NULL"
        Execute True, sSql
        
        ' Update Vendor table
        sSql = "Update VndrTable " & vbCrLf _
                & "  SET VEAPPROVREQ = Case When (Len(VEAPPDATE) > 0) OR (Len(VEREVIEWDT)>0) Then 1 Else VEAPPROVREQ END, " & vbCrLf _
                & "  VESURVEY = Case When (Len(VESURVSENT) > 0) Or (Len(VESURVREC) > 0) Then 1 Else VESURVEY END " & vbCrLf _
                & "  Where " & vbCrLf _
                & "  ((Len(VEAPPDATE)>0 Or Len(VEREVIEWDT)>0) And VEAPPROVREQ <> 1) Or  " & vbCrLf _
                & "  ((Len(VESURVSENT)>0 Or Len(VESURVREC)>0) And VESURVEY <> 1) "

        Execute True, sSql

        ' Add new Column to the PIPORIGDATE to PoitTable table
        If Not ColumnExists("EsReportBook18", "Col8") Then
           Execute True, "ALTER TABLE dbo.EsReportBook18 ADD Col8 char(20) NULL"
        End If
        

        ' Alter the BackLogBy Sched Date proc to include the Begin and End dates.
        sSql = "ALTER PROCEDURE [dbo].[BackLogBySchedDate]" & vbCrLf _
                & " @BegDate as varchar(16), @EndDate as varchar(16), @Customer as varchar(10)," & vbCrLf _
                & " @PartClass as Varchar(16),@PartCode as varchar(8)" & vbCrLf _
                & " AS                                               " & vbCrLf _
                & " BEGIN                           " & vbCrLf _
                & " declare @SoType as varchar(1)   " & vbCrLf _
                & " declare @SoText as varchar(6)   " & vbCrLf _
                & " declare @ItSo as int            " & vbCrLf _
                & " declare @ItRev as char(2)       " & vbCrLf _
                & " declare @ItNum as int           " & vbCrLf _
                & " declare @ItQty as decimal(12,4) " & vbCrLf _
                & " declare @PaLotRemQty as decimal(12,4)   " & vbCrLf _
                & " declare @PartRem as decimal(12,4)       " & vbCrLf _
                & " declare @RunningTot as decimal(12,4)    " & vbCrLf _
                & " declare @ItDollars as decimal(12,4)     " & vbCrLf _
                & " declare @ItSched as smalldatetime       " & vbCrLf _
                & " declare @CusName as varchar(10)         " & vbCrLf _
                & " declare @PartNum as varchar(30)         " & vbCrLf _
                & " declare @CurPartNum as varchar(30)      " & vbCrLf _
                & " declare @PartDesc as varchar(30)        " & vbCrLf _
                & " declare @PartLoc as varchar(4)          " & vbCrLf _
                & " declare @PartExDesc as varchar(3072)    " & vbCrLf _
                & " declare @ItCanceled as tinyint          " & vbCrLf _
                & " declare @ItPSNum as varchar(8)          " & vbCrLf _
                & " declare @ItInvoice as int declare @ItPSShipped as tinyint   " & vbCrLf
    
    sSql = sSql & " IF (@Customer = 'ALL')                      " & vbCrLf _
            & "         SET @Customer = ''                      " & vbCrLf _
            & "     IF (@PartClass = 'ALL')                     " & vbCrLf _
            & "         SET @PartClass = ''                     " & vbCrLf _
            & "     IF (@PartCode = 'ALL')                      " & vbCrLf _
            & "        SET @PartCode = ''                      " & vbCrLf _
            & "     CREATE TABLE #tempBackLogInfo               " & vbCrLf _
            & "                 (SOTYPE varchar(1) NULL,        " & vbCrLf _
            & "                 SOTEXT varchar(6) NULL,         " & vbCrLf _
            & "                 ITSO Int NULL,                  " & vbCrLf _
            & "                ITREV char(2) NULL,              " & vbCrLf _
            & "                 ITNUMBER int NULL,              " & vbCrLf _
            & "                 ITQTY decimal(12,4) NULL,       " & vbCrLf _
            & "                 PALOTQTYREMAINING decimal(12,4) NULL,   " & vbCrLf _
            & "                 RUNQTYTOT decimal(12,4) NULL,   " & vbCrLf _
            & "                 ITDOLLARS decimal(12,4) NULL,   " & vbCrLf _
            & "                 ITSCHED smalldatetime NULL,     " & vbCrLf _
            & "                 CUNICKNAME varchar(10) NULL,    " & vbCrLf _
            & "                 PARTNUM varchar(30) NULL,       " & vbCrLf _
            & "                PADESC varchar(30) NULL,     " & vbCrLf _
            & "                 PAEXTDESC varchar(3072) NULL,   " & vbCrLf _
            & "                 PALOCATION varchar(4) NULL,     " & vbCrLf _
            & "                 ITCANCELED tinyint NULL,        " & vbCrLf _
            & "                 ITPSNUMBER varchar(8) NULL, ITINVOICE int NULL, " & vbCrLf
                
    sSql = sSql & "                 ITPSSHIPPED tinyint NULL)       " & vbCrLf & vbCrLf _
            & "     DECLARE curbackLog CURSOR   FOR                             " & vbCrLf _
            & "     SELECT SohdTable.SOTYPE, SohdTable.SOTEXT,                  " & vbCrLf _
            & "         SoitTable.ITSO, SoitTable.ITREV, SoitTable.ITNUMBER,    " & vbCrLf _
            & "         SoitTable.ITQTY, PartTable.PALOTQTYREMAINING,           " & vbCrLf _
            & "         SoitTable.ITDOLLARS,SoitTable.ITSCHED, CustTable.CUNICKNAME,    " & vbCrLf _
            & "         PartTable.PARTNUM, PartTable.PADESC, PartTable.PAEXTDESC,       " & vbCrLf _
            & "         PartTable.PALOCATION, SoitTable.ITCANCELED,                     " & vbCrLf _
            & "         SoitTable.ITPSNUMBER , SoitTable.ITINVOICE, SoitTable.ITPSSHIPPED   " & vbCrLf _
            & "     From SohdTable, SoitTable, CustTable, PartTable             " & vbCrLf _
            & "     WHERE SohdTable.SOCUST = CustTable.CUREF AND                " & vbCrLf _
            & "         SohdTable.SONUMBER =SoitTable.ITSO AND                  " & vbCrLf _
            & "         SoitTable.ITPART=PartTable.PARTREF AND                  " & vbCrLf _
            & "         SoitTable.ITCANCELED=0 AND SoitTable.ITPSNUMBER=''      " & vbCrLf _
            & "         AND SoitTable.ITINVOICE=0 AND SoitTable.ITPSSHIPPED=0   " & vbCrLf _
            & "         AND CUREF LIKE '%' + @Customer + '%'                    " & vbCrLf _
            & "         AND SoitTable.ITSCHED BETWEEN @BegDate AND @EndDate     " & vbCrLf _
            & "         AND PartTable.PACLASS LIKE '%' + @PartClass + '%'       " & vbCrLf _
            & "         AND PartTable.PAPRODCODE LIKE '%' + @PartCode + '%'     " & vbCrLf _
            & "     ORDER BY partnum, ITSCHED                                   " & vbCrLf & vbCrLf _

     sSql = sSql & "     OPEN curbackLog                                                " & vbCrLf _
            & "     FETCH NEXT FROM curbackLog INTO @SoType, @SoText, @ItSo,  @ItRev, @ItNum, @ItQty, @PaLotRemQty, " & vbCrLf _
            & "                     @ItDollars,@ItSched, @CusName, @PartNum,        " & vbCrLf _
            & "                     @PartDesc, @PartExDesc, @PartLoc, @ItCanceled,  " & vbCrLf _
            & "                     @ItPSNum, @ItInvoice, @ItPSShipped              " & vbCrLf _
            & "     SET @CurPartNum = @PartNum                                      " & vbCrLf _
            & "     SET @RunningTot = 0                                             " & vbCrLf _
            & "     WHILE (@@FETCH_STATUS <> -1)                                    " & vbCrLf _
            & "     BEGIN                                                           " & vbCrLf _
            & "         IF (@@FETCH_STATUS <> -2)                                   " & vbCrLf _
            & "         BEGIN                                                       " & vbCrLf _
            & "             IF  @CurPartNum <> @PartNum                             " & vbCrLf _
            & "            BEGIN                                                    " & vbCrLf _
            & "                 SET @RunningTot = @ItQty                            " & vbCrLf _
            & "                 set @CurPartNum = @PartNum                          " & vbCrLf _
            & "             End                                                     " & vbCrLf _
            & "             Else                                                    " & vbCrLf _
            & "             BEGIN                                                   " & vbCrLf _
            & "                 SET @RunningTot = @RunningTot + @ItQty              " & vbCrLf _
            & "             End                                                     " & vbCrLf & vbCrLf _
            & "             SET @PartRem = @PaLotRemQty - @RunningTot                       " & vbCrLf _
            & "             INSERT INTO #tempBackLogInfo (SOTYPE, SOTEXT, ITSO, ITREV,      " & vbCrLf _
            & "                 ITNUMBER,ITQTY, PALOTQTYREMAINING,RUNQTYTOT, ITDOLLARS,     " & vbCrLf _
            & "                 ITSCHED,CUNICKNAME, PARTNUM, PADESC,PAEXTDESC,PALOCATION,   " & vbCrLf

    sSql = sSql & "                      ITCANCELED, ITPSNUMBER, ITINVOICE, ITPSSHIPPED)             " & vbCrLf _
            & "             VALUES (@SoType, @SoText, @ItSo, @ItRev,@ItNum, @ItQty,@PaLotRemQty,@PartRem, @ItDollars,@ItSched,@CusName, " & vbCrLf _
            & "                 @PartNum,@PartDesc,@PartExDesc,@PartLoc, @ItCanceled,@ItPSNum,@ItInvoice,@ItPSShipped)  " & vbCrLf _
            & "         End                                                                 " & vbCrLf _
            & "         FETCH NEXT FROM curbackLog INTO @SoType, @SoText, @ItSo,            " & vbCrLf _
            & "             @ItRev, @ItNum, @ItQty, @PaLotRemQty,                           " & vbCrLf _
            & "             @ItDollars,@ItSched, @CusName, @PartNum,                        " & vbCrLf _
            & "             @PartDesc, @PartExDesc, @PartLoc, @ItCanceled,                  " & vbCrLf _
            & "             @ItPSNum, @ItInvoice, @ItPSShipped                              " & vbCrLf _
            & "     End                                                                     " & vbCrLf _
            & "     CLOSE curbackLog   --// close the cursor                                " & vbCrLf _
            & "     DEALLOCATE curbackLog                                                   " & vbCrLf _
            & "     -- select data for the report                                           " & vbCrLf _
            & "     SELECT SOTYPE, SOTEXT, ITSO, ITREV,                                     " & vbCrLf _
            & "         ITNUMBER,ITQTY, PALOTQTYREMAINING,RUNQTYTOT, ITDOLLARS,             " & vbCrLf _
            & "         ITSCHED,CUNICKNAME, PARTNUM, PADESC,PAEXTDESC, PALOCATION,          " & vbCrLf _
            & "         ITCANCELED , ITPSNUMBER, ITINVOICE, ITPSSHIPPED                     " & vbCrLf _
            & "     FROM #tempBackLogInfo                                                   " & vbCrLf _
            & "     ORDER BY ITSCHED                                                        " & vbCrLf _
            & "     -- drop the temp table                                                  " & vbCrLf _
            & "     DROP table #tempBackLogInfo                                             " & vbCrLf _
            & "     End                                                                     "
        
        ' Excute the query
        Execute True, sSql

 
        sSql = "ALTER PROC [dbo].[Qry_FillSalesOrders] (@DateRange as smalldatetime) AS " & vbCrLf _
                & " SELECT SONUMBER,SOTYPE,SOPO FROM SohdTable WHERE (SOCANCELED=0 AND SOLOCKED=0 AND SODATE>@DateRange)ORDER BY SONUMBER DESC "
        ' Excute the query
        Execute True, sSql


    
        'set version
        Execute False, "UPDATE Version set Version = " & newver
    End If
End Function


Private Function UpdateDatabase9()
   newver = 68
   If ver < newver Then
       ver = newver
      
      ' Alter the TciTable comment field
      If Not ColumnExists("TcitTable", "TCCOMMENTS") Then
         Execute True, "ALTER TABLE dbo.TcitTable ADD TCCOMMENTS varchar(1024) NULL"
         'Execute True, sSql
      End If
         
      sSql = "ALTER VIEW [dbo].[viewOpenAPTerms] " & vbCrLf _
                     & " AS  " & vbCrLf _
               & " SELECT dbo.VihdTable.VIVENDOR, dbo.VndrTable.VENICKNAME, dbo.VndrTable.VEBNAME, dbo.VihdTable.VINO, dbo.VihdTable.VIDATE," & vbCrLf _
                   & " dbo.VihdTable.VIDUE AS InvTotal, dbo.VihdTable.VIDUEDATE, dbo.VihdTable.VIFREIGHT, dbo.VihdTable.VITAX, dbo.VihdTable.VIPAY," & vbCrLf _
                   & " ISNULL(ViitTable_1.VITPO, 0) AS VITPO, ISNULL(ViitTable_1.VITPORELEASE, 0) AS VITPORELEASE, " & vbCrLf _
                   & " CASE WHEN PONETDAYS <> 0 THEN PODISCOUNT ELSE VEDISCOUNT END AS DiscRate, " & vbCrLf _
                   & " CASE WHEN PONETDAYS <> 0 THEN PODDAYS ELSE VEDDAYS END AS DiscDays, " & vbCrLf _
                   & " CASE WHEN PONETDAYS <> 0 THEN PONETDAYS ELSE VENETDAYS END AS NetDays, dbo.VihdTable.VIDUE - dbo.VihdTable.VIPAY AS AmountDue, " & vbCrLf _
                   & " dbo.VihdTable.VIFREIGHT + dbo.VihdTable.VITAX + " & vbCrLf _
                   & "      (SELECT     CAST(SUM(ROUND(CAST(dbo.ViitTable.VITQTY AS decimal(12, 3)) * CAST(dbo.ViitTable.VITCOST AS decimal(12, 4)) " & vbCrLf _
                   & "                 + CAST(dbo.ViitTable.VITADDERS AS decimal(12, 2)), 2)) AS decimal(12, 2)) AS Expr1 " & vbCrLf _
                   & "                     FROM   dbo.ViitTable LEFT OUTER JOIN " & vbCrLf _
                   & "                                               dbo.viewPohdPoit ON dbo.ViitTable.VITPO = dbo.viewPohdPoit.PONUMBER AND " & vbCrLf _
                   & "                                               dbo.ViitTable.VITPORELEASE = dbo.viewPohdPoit.PORELEASE AND dbo.ViitTable.VITPOITEM = dbo.viewPohdPoit.PIITEM AND " & vbCrLf _
                   & "                                               dbo.ViitTable.VITPOITEMREV = dbo.viewPohdPoit.PIREV " & vbCrLf _
                   & "                        WHERE      (dbo.ViitTable.VITVENDOR = dbo.VihdTable.VIVENDOR) AND (dbo.ViitTable.VITNO = dbo.VihdTable.VINO)) AS CalcTotal," & vbCrLf _
                   & "                  ViitTable_1.VITCOST , viewPohdPoit_1.PILOT " & vbCrLf _
                   & "         FROM  dbo.VihdTable INNER JOIN " & vbCrLf _
                   & "                  dbo.VndrTable ON dbo.VihdTable.VIVENDOR = dbo.VndrTable.VEREF LEFT OUTER JOIN " & vbCrLf _
                   & "                  dbo.viewPohdPoit AS viewPohdPoit_1 RIGHT OUTER JOIN " & vbCrLf _
                   & "                  dbo.ViitTable AS ViitTable_1 ON viewPohdPoit_1.PIREV = ViitTable_1.VITPOITEMREV AND viewPohdPoit_1.PIITEM = ViitTable_1.VITPOITEM AND " & vbCrLf _
                   & "                  viewPohdPoit_1.PONUMBER = ViitTable_1.VITPO AND viewPohdPoit_1.PORELEASE = ViitTable_1.VITPORELEASE ON " & vbCrLf _
                   & "                  ViitTable_1.VITNO = dbo.VihdTable.VINO AND ViitTable_1.VITVENDOR = dbo.VihdTable.VIVENDOR AND ViitTable_1.VITITEM = " & vbCrLf _
                   & "                      (SELECT     MIN(VITITEM) AS Expr1 " & vbCrLf _
                   & "                        From dbo.ViitTable" & vbCrLf
                   
            sSql = sSql & "             WHERE (VITVENDOR = dbo.VndrTable.VEREF) AND (VITNO = dbo.VihdTable.VINO) AND (VITPO IS NOT NULL) AND (VITPO <> 0))" & vbCrLf _
                       & " WHERE (dbo.VihdTable.VIPIF <> 1) And (ISNULL(ViitTable_1.VITPO, 0) <> 0)" & vbCrLf _
                       & "    UNION " & vbCrLf _
                       & " SELECT  VihdTable_1.VIVENDOR, VndrTable_1.VENICKNAME, VndrTable_1.VEBNAME, VihdTable_1.VINO, VihdTable_1.VIDATE, VihdTable_1.VIDUE AS InvTotal, " & vbCrLf _
                       & "                        VihdTable_1.VIDUEDATE, VihdTable_1.VIFREIGHT, VihdTable_1.VITAX, VihdTable_1.VIPAY, ISNULL(ViitTable_1.VITPO, 0) AS VITPO, " & vbCrLf _
                       & "                       ISNULL(ViitTable_1.VITPORELEASE, 0) AS VITPORELEASE, CASE WHEN PONETDAYS <> 0 THEN PODISCOUNT ELSE VEDISCOUNT END AS DiscRate, " & vbCrLf _
                       & "                        CASE WHEN PONETDAYS <> 0 THEN PODDAYS ELSE VEDDAYS END AS DiscDays, " & vbCrLf _
                       & "                        CASE WHEN PONETDAYS <> 0 THEN PONETDAYS ELSE VENETDAYS END AS NetDays, VihdTable_1.VIDUE - VihdTable_1.VIPAY AS AmountDue, " & vbCrLf _
                       & "                        VihdTable_1.VIFREIGHT + VihdTable_1.VITAX + " & vbCrLf _
                       & "                           (SELECT     CAST(SUM(ROUND(CAST(VITQTY AS decimal(12, 3)) * CAST(VITCOST AS decimal(12, 4)) + CAST(VITADDERS AS decimal(12, 2)), 2)) " & vbCrLf _
                       & "                                                    AS decimal(12, 2)) AS Expr1 " & vbCrLf _
                       & "                            FROM          dbo.ViitTable AS ViitTable_2 " & vbCrLf _
                       & "                             WHERE      (VITVENDOR = VihdTable_1.VIVENDOR) AND (VITNO = VihdTable_1.VINO)) AS CalcTotal, ViitTable_1.VITCOST, " & vbCrLf _
                       & "                       viewPohdPoit_2.PILOT " & vbCrLf _
                       & "    FROM  dbo.VihdTable AS VihdTable_1 INNER JOIN " & vbCrLf _
                       & "                       dbo.VndrTable AS VndrTable_1 ON VihdTable_1.VIVENDOR = VndrTable_1.VEREF LEFT OUTER JOIN " & vbCrLf _
                       & "                       dbo.viewPohdPoit AS viewPohdPoit_2 RIGHT OUTER JOIN " & vbCrLf _
                       & "                       dbo.ViitTable AS ViitTable_1 ON viewPohdPoit_2.PIREV = ViitTable_1.VITPOITEMREV AND viewPohdPoit_2.PIITEM = ViitTable_1.VITPOITEM AND " & vbCrLf _
                       & "                       viewPohdPoit_2.PONUMBER = ViitTable_1.VITPO AND viewPohdPoit_2.PORELEASE = ViitTable_1.VITPORELEASE ON " & vbCrLf _
                       & "                       ViitTable_1.VITNO = VihdTable_1.VINO AND ViitTable_1.VITVENDOR = VihdTable_1.VIVENDOR AND ViitTable_1.VITITEM = " & vbCrLf _
                       & "                           (SELECT     MIN(VITITEM) AS Expr1 " & vbCrLf _
                       & "                             From dbo.ViitTable " & vbCrLf _
                       & "                             WHERE      (VITVENDOR = VndrTable_1.VEREF) AND (VITNO = VihdTable_1.VINO) AND (VITPO IS NOT NULL) AND (VITPO <> 0)) " & vbCrLf _
                       & " Where (VihdTable_1.VIPIF <> 1) And (ISNULL(ViitTable_1.VITPO, 0) = 0) " & vbCrLf

         ' Execute the sql
         Execute True, sSql
    
    
         sSql = "CREATE PROCEDURE [dbo].[RptRMFGoods] " & vbCrLf _
                  & " @ReportDate as varchar(16), @PartClass as Varchar(16), " & vbCrLf _
                  & " @PartCode as varchar(8), @lotHDOnly as int, @PartType1 as Integer, " & vbCrLf _
                  & " @PartType2 as Integer, @PartType3 as Integer, @PartType4 as Integer " & vbCrLf _
                  & " AS        " & vbCrLf _
                  & " BEGIN      " & vbCrLf _
                  & "   declare @partRef as varchar(30)   " & vbCrLf _
                  & "   declare @partType as int           " & vbCrLf _
                  & "   declare @partDesc as varchar(30)     " & vbCrLf _
                  & "   declare @partExDesc as varchar(3072)    " & vbCrLf _
                  & "   declare @rptRemQty as decimal(12,4)  " & vbCrLf _
                  & "   declare @deltaQty as decimal(12,4)   " & vbCrLf _
                  & "   declare @lotRemQty as decimal(12,4)  " & vbCrLf _
                  & "   declare @lotOrgQty as decimal(12,4)  " & vbCrLf _
                  & "   declare @tmpInvQty as decimal(12,4)  " & vbCrLf _
                  & "   declare @tmpLastInvQty as decimal(12,4)    " & vbCrLf _
                  & "   declare @rptInvcost decimal(12,4)    " & vbCrLf _
                  & "   declare @lastInvcost decimal(12,4)   " & vbCrLf _
                  & "   declare @orgInvcost decimal(12,4)    " & vbCrLf _
                  & "   declare @tmpInvCost decimal(12,4)    " & vbCrLf _
                  & "   declare @tmpLastInvCost decimal(12,4)   " & vbCrLf _
                  & "   declare @lastQty as decimal(12,4)    " & vbCrLf _
                  & "   declare @orgQty as decimal(12,4)  " & vbCrLf _
                  & "   declare @rptQty as decimal(12,4)  " & vbCrLf _
                  & "   declare @rptCreditACC varchar(12)    " & vbCrLf
         
         sSql = sSql & " declare @rptDebitACC varchar(12)" & vbCrLf _
                     & " declare @rptACC varchar(12)   " & vbCrLf _
                     & " declare @CurACC varchar(12)   " & vbCrLf _
                     & " declare @tmpCreditAcc varchar(12)   " & vbCrLf _
                     & " declare @tmpDebitAcc varchar(12) " & vbCrLf _
                     & " declare @tmplastCreditAcc varchar(12)  " & vbCrLf _
                     & " declare @tmpLastDebitAcc varchar(12)   " & vbCrLf _
                     & " declare @lastCreditACC varchar(12)  " & vbCrLf _
                     & " declare @lastDebitACC varchar(12)   " & vbCrLf _
                     & " declare @lastACC varchar(12)  " & vbCrLf _
                     & " declare @orgCreditACC varchar(12)" & vbCrLf _
                     & " declare @orgDebitACC varchar(12) " & vbCrLf _
                     & " declare @orgACC varchar(12)   " & vbCrLf _
                     & " declare @OrgInvNum as int  " & vbCrLf _
                     & " declare @rptInvNum as int  " & vbCrLf _
                     & " declare @LastInvNum as int " & vbCrLf _
                     & " declare @tmpInvNum as int  " & vbCrLf _
                     & " declare @tmpLastInvNum as int " & vbCrLf _
                     & " declare @LotNumber varchar(51)   " & vbCrLf _
                     & " declare @LotUserID varchar(51)   " & vbCrLf _
                     & " declare @lotAcualDate as smalldatetime " & vbCrLf _
                     & " declare @lotCostedDate as smalldatetime   " & vbCrLf _
                     & " declare @curDate as smalldatetime   " & vbCrLf _
                     & " declare @AcualDate as smalldatetime " & vbCrLf _
                     & " declare @CostedDate as smalldatetime   " & vbCrLf
                     
         sSql = sSql & " declare @tmpINVAdate  as smalldatetime   " & vbCrLf _
                     & " declare @tmpLastINVAdate  as smalldatetime   " & vbCrLf _
                     & " declare @unitcost as decimal(12,4)  " & vbCrLf _
                     & " declare @partStdCost as decimal(12,4)  " & vbCrLf _
                     & " declare @LotUnitCost as decimal(12,4)  " & vbCrLf _
                     & " declare @lotTotMatl as decimal(12,4)   " & vbCrLf _
                     & " declare @lotTotLabor as decimal(12,4)  " & vbCrLf _
                     & " declare @lotTotExp as decimal(12,4) " & vbCrLf _
                     & " declare @lotTotOH as decimal(12,4)  " & vbCrLf _
                     & " declare @partActCost as int   " & vbCrLf _
                     & " declare @partLotTrack as int  " & vbCrLf _
                     & " declare @flgStdCost as int " & vbCrLf _
                     & " declare @flgLdCost as int  " & vbCrLf _
                     & " declare @flgInvCost as int " & vbCrLf _
                     & " declare @flgLdRQErr as int " & vbCrLf _
                     & " declare @flgOrgAcc as int  " & vbCrLf _
                     & " declare @flgRptAcc as int  " & vbCrLf _
                     & " declare @flgLastAcc as int " & vbCrLf _
                     & "   IF (@PartClass = 'ALL')" & vbCrLf _
                     & "   BEGIN " & vbCrLf _
                     & "      SET @PartClass = ''  " & vbCrLf _
                     & "   End            " & vbCrLf _
                     & "   IF (@PartCode = 'ALL')  " & vbCrLf _
                     & "   BEGIN " & vbCrLf _
                     & "      SET @PartCode = ''      " & vbCrLf
                           
         sSql = sSql & " End                     " & vbCrLf _
                     & "    IF (@PartType1 = 1)     " & vbCrLf _
                     & "      SET @PartType1 = 1    " & vbCrLf _
                     & "   Else                    " & vbCrLf _
                     & "     SET @PartType1 = 0    " & vbCrLf _
                     & "   IF (@PartType2 = 1)     " & vbCrLf _
                     & "     SET @PartType2 = 2    " & vbCrLf _
                     & "   Else                    " & vbCrLf _
                     & "     SET @PartType2 = 0    " & vbCrLf _
                     & "    IF (@PartType3 = 1)    " & vbCrLf _
                     & "        SET @PartType3 = 3 " & vbCrLf _
                     & "    Else                   " & vbCrLf _
                     & "        SET @PartType3 = 0 " & vbCrLf _
                     & "                           " & vbCrLf _
                     & "    IF (@PartType4 = 1)    " & vbCrLf _
                     & "        SET @PartType4 = 4 " & vbCrLf _
                     & "    Else                   " & vbCrLf _
                     & "        SET @PartType4 = 0 " & vbCrLf _
                     & "      " & vbCrLf _
                     & "   CREATE TABLE #tempRMFGoods(   " & vbCrLf _
                     & "   [LOTNUMBER] [varchar](15) NULL,  " & vbCrLf _
                     & "   [PARTNUM] [varchar](30) NULL, " & vbCrLf _
                     & "   [PALEVEL] [int] NULL,         " & vbCrLf _
                     & "   [PADESC] [varchar](30) NULL,  " & vbCrLf
                        
         sSql = sSql & "   [PAEXTDESC] [varchar](3072) NULL,   " & vbCrLf _
                     & "   [LOTUSERLOTID] [char](40) NULL,     " & vbCrLf _
                     & "   [ORGINNUMBER] [int] NULL,           " & vbCrLf _
                     & "   [RPTINNUMBER] [int] NULL,           " & vbCrLf _
                     & "   [CURINNUMBER] [int] NULL,           " & vbCrLf _
                     & "   [ACTUALDATE] [smalldatetime] NULL,  " & vbCrLf _
                     & "   [RPTDATEQTY] [decimal](12, 4) NULL, " & vbCrLf _
                     & "   [UNITCOST] [decimal](12, 4) NULL,   " & vbCrLf _
                     & "   [PASTDCOST] [decimal](12, 4) NULL,  " & vbCrLf _
                     & "   [LOTUNITCOST] [decimal](12, 4) NULL,   " & vbCrLf _
                     & "   [INAMT] [decimal](12, 4) NULL,         " & vbCrLf _
                     & "   [ORGCOST] [decimal](12, 4) NULL,       " & vbCrLf _
                     & "   [STDCOST] [decimal](12, 4) NULL,       " & vbCrLf _
                     & "   [LSTACOST] [decimal](12, 4) NULL,      " & vbCrLf _
                     & "   [RPTCOST] [decimal](12, 4) NULL,       " & vbCrLf _
                     & "   [CURCOST] [decimal](12, 4) NULL,       " & vbCrLf _
                     & "   [COSTEDDATE] [smalldatetime] NULL,     " & vbCrLf _
                     & "   [RPTACCOUNT] [char](12) NULL,          " & vbCrLf _
                     & "   [ORIGINALACC] [char](12) NULL,         " & vbCrLf _
                     & "   [LASTACTVITYACC] [char](12) NULL,      " & vbCrLf _
                     & "   [CURRENTACC] [char](12) NULL, " & vbCrLf _
                     & "   [PACLASS] [char](4) NULL,  " & vbCrLf _
                     & "   [PAPRODCODE] [char](6) NULL,  " & vbCrLf _
                     & "   [flgStdCost] [int] NULL,   " & vbCrLf _
                     & "   [flgLdCost] [int] NULL,    " & vbCrLf
                        
         sSql = sSql & "   [flgInvCost] [int] NULL,   " & vbCrLf _
                     & "   [flgLdRQErr] [int] NULL,   " & vbCrLf _
                     & "   [flgRptAcc] [int] NULL,    " & vbCrLf _
                     & "   [flgOrgAcc] [int] NULL,    " & vbCrLf _
                     & "   [flgLastAcc] [int] NULL    " & vbCrLf _
                     & ")                             " & vbCrLf _
                     & "                              " & vbCrLf _
                     & " DECLARE curLotHd CURSOR LOCAL " & vbCrLf _
                     & " FOR                           " & vbCrLf _
                     & " SELECT LOTNUMBER, LOTUSERLOTID, PARTREF, PADESC, PAEXTDESC, " & vbCrLf _
                     & "    LOTADATE,LOTORIGINALQTY, LOTREMAININGQTY, LOTUNITCOST, LOTDATECOSTED,   " & vbCrLf _
                     & "    PACLASS , PAPRODCODE, PASTDCOST, PAUSEACTUALCOST, PALOTTRACK, PALEVEL   " & vbCrLf _
                     & " From ViewLohdPartTable  " & vbCrLf _
                     & " WHERE ViewLohdPartTable.LOTADATE  < DATEADD(dd, 1 , @ReportDate)  " & vbCrLf _
                     & "    AND ViewLohdPartTable.PACLASS LIKE '%' + @PartClass + '%'      " & vbCrLf _
                     & "    AND ViewLohdPartTable.PAPRODCODE LIKE '%' + @PartCode + '%'    " & vbCrLf _
                     & "    AND ViewLohdPartTable.PALEVEL IN (@PartType1, @PartType2, @PartType3, @PartType4)   " & vbCrLf _
                     & " OPEN curLotHd                                                     " & vbCrLf _
                     & " FETCH NEXT FROM curLotHd INTO @LotNumber, @LotUserID, @partRef,   " & vbCrLf _
                     & "    @partDesc, @partExDesc, @lotAcualDate, @lotOrgQty,             " & vbCrLf _
                     & "    @lotRemQty,@LotUnitCost, @lotCostedDate,                       " & vbCrLf _
                     & "    @PartClass, @PartCode, @partStdCost, @partActCost, @partLotTrack, @partType   " & vbCrLf _
                     & "                                                                " & vbCrLf _
                     & " WHILE (@@FETCH_STATUS <> -1)                                   " & vbCrLf _
                     & " BEGIN                                                          " & vbCrLf
                     
         sSql = sSql & " IF (@@FETCH_STATUS <> -2)     " & vbCrLf _
                     & " BEGIN                         " & vbCrLf _
                     & "    SET @flgLdRQErr = 0        " & vbCrLf _
                     & " IF (@lotRemQty < 0.0000)      " & vbCrLf _
                     & " BEGIN                         " & vbCrLf _
                     & "    SET @lotRemQty = 0.0000    " & vbCrLf _
                     & "    SET @flgLdRQErr  = 1       " & vbCrLf _
                     & " End                           " & vbCrLf _
                     & " SET @curDate = GETDATE()      " & vbCrLf _
                     & "                               " & vbCrLf _
                     & " SELECT @deltaQty = ISNULL(SUM(LOIQUANTITY), 0.0000)   " & vbCrLf _
                     & " From LoitTable                                        " & vbCrLf _
                     & " WHERE LOIADATE BETWEEN DATEADD(dd, 1 ,@ReportDate) AND DATEADD(dd, 1 ,@curDate)  " & vbCrLf _
                     & "    AND LOIPARTREF = @partRef                    " & vbCrLf _
                     & "    AND LOINUMBER = @LotNumber                   " & vbCrLf _
                     & " SET @rptRemQty = @lotRemQty + (@deltaQty * -1)  " & vbCrLf _
                     & " IF @rptRemQty < 0.0000           " & vbCrLf _
                     & " SET @rptRemQty = @rptRemQty * -1 " & vbCrLf _
                     & " SET @flgStdCost = 0     " & vbCrLf _
                     & " SET @flgLdCost = 0      " & vbCrLf _
                     & " SET @flgInvCost = 0     " & vbCrLf _
                     & " SET @flgOrgAcc = 0      " & vbCrLf _
                     & " SET @flgRptAcc = 0      " & vbCrLf _
                     & " SET @flgLastAcc = 0     " & vbCrLf _
                     & "                         " & vbCrLf
                                             
         sSql = sSql & " DECLARE curInv CURSOR   " & vbCrLf _
                     & " LOCAL                   " & vbCrLf _
                     & " Scroll                  " & vbCrLf _
                     & " FOR                     " & vbCrLf _
                     & " SELECT INNUMBER, INAMT, INAQTY, ISNULL(INCREDITACCT,0), ISNULL(INDEBITACCT, 0), INADATE   " & vbCrLf _
                     & "     From InvaTable, LoitTable                          " & vbCrLf _
                     & " WHERE InvaTable.INPART = @partRef                  " & vbCrLf _
                     & "    AND LoitTable.LOINUMBER = @LotNumber            " & vbCrLf _
                     & "    AND InvaTable.INPART = LoitTable.LOIPARTREF     " & vbCrLf _
                     & "    AND InvaTable.INNUMBER = LoitTable.LOIACTIVITY  " & vbCrLf _
                     & " ORDER BY INADATE ASC          " & vbCrLf _
                     & "                               " & vbCrLf _
                     & " OPEN curInv                   " & vbCrLf _
                     & " FETCH FIRST FROM curInv INTO @tmpInvNum, @tmpInvCost, @tmpInvQty, " & vbCrLf _
                     & " @tmpCreditAcc, @tmpDebitAcc, @tmpINVAdate " & vbCrLf _
                     & " IF (@@FETCH_STATUS <> -1)  " & vbCrLf _
                     & " BEGIN                      " & vbCrLf _
                     & " IF (@@FETCH_STATUS <> -2)  " & vbCrLf _
                     & " BEGIN                      " & vbCrLf _
                     & "    SET @orgInvcost = @tmpInvCost " & vbCrLf _
                     & "    SET @orgQty = @tmpInvQty      " & vbCrLf _
                     & "    SET @orgCreditACC = @tmpCreditAcc   " & vbCrLf _
                     & "    SET @orgDebitACC = @tmpDebitAcc  " & vbCrLf _
                     & "    SET @OrgInvNum = @tmpInvNum      " & vbCrLf _
                     & " End                                 " & vbCrLf
                     
                     
         sSql = sSql & " FETCH LAST FROM curInv INTO @tmpLastInvNum, @tmpLastInvCost, @tmpLastInvQty,  " & vbCrLf _
                     & " @tmplastCreditAcc, @tmpLastDebitAcc, @tmpLastINVAdate " & vbCrLf _
                     & " IF (@@FETCH_STATUS <> -2)  " & vbCrLf _
                     & " BEGIN                      " & vbCrLf _
                     & " SET @lastInvcost = @tmpLastInvCost  " & vbCrLf _
                     & " SET @lastQty = @tmpLastInvQty    " & vbCrLf _
                     & " SET @lastCreditACC = @tmplastCreditAcc " & vbCrLf _
                     & " SET @lastDebitACC = @tmpLastDebitAcc   " & vbCrLf _
                     & " SET @LastInvNum = @tmpLastInvNum       " & vbCrLf _
                     & " IF @tmpLastINVAdate > DATEADD(dd, 1, @ReportDate)  " & vbCrLf _
                     & " BEGIN                                              " & vbCrLf _
                     & "   SELECT TOP 1 @rptInvNum = INNUMBER, @rptInvcost = INAMT, " & vbCrLf _
                     & "      @rptQty = INAQTY, @rptCreditACC = ISNULL(INCREDITACCT, 0),  " & vbCrLf _
                     & "      @rptDebitACC = ISNULL(INDEBITACCT, 0)     " & vbCrLf _
                     & "   From InvaTable, LoitTable                 " & vbCrLf _
                     & "      WHERE INADATE < DATEADD(dd, 1, @ReportDate)  " & vbCrLf _
                     & "         AND InvaTable.INPART = @partRef              " & vbCrLf _
                     & "         AND LoitTable.LOINUMBER = @LotNumber         " & vbCrLf _
                     & "         AND InvaTable.INPART = LoitTable.LOIPARTREF  " & vbCrLf _
                     & "         AND InvaTable.INNUMBER = LoitTable.LOIACTIVITY  " & vbCrLf _
                     & "         ORDER BY INADATE DESC                           " & vbCrLf _
                     & "  End                                          " & vbCrLf _
                     & "  Else                                         " & vbCrLf _
                     & "  BEGIN                                        " & vbCrLf _
                     & "  SET @rptInvNum = @tmpLastInvNum              " & vbCrLf
        
        sSql = sSql & "                                  " & vbCrLf _
                    & "                                     " & vbCrLf _
                    & " SET @rptInvcost = @tmpLastInvCost   " & vbCrLf _
                    & " SET @rptQty = @tmpLastInvQty        " & vbCrLf _
                    & " SET @rptCreditACC = @tmplastCreditAcc  " & vbCrLf _
                    & " SET @rptDebitACC = @tmpLastDebitAcc    " & vbCrLf _
                    & " End                                 " & vbCrLf _
                    & " End                                 " & vbCrLf _
                    & " End                                 " & vbCrLf _
                    & " CLOSE curInv   --// close the cursor   " & vbCrLf _
                    & " DEALLOCATE curInv          " & vbCrLf _
                    & " IF (@partActCost = 0)      " & vbCrLf _
                    & " BEGIN                      " & vbCrLf _
                    & " SET @unitcost = @partStdCost  " & vbCrLf _
                    & " SET @flgStdCost = 1        " & vbCrLf _
                    & " End                        " & vbCrLf _
                    & " Else                       " & vbCrLf _
                    & " BEGIN                      " & vbCrLf _
                    & " IF @lotHDOnly = 1          " & vbCrLf _
                    & " BEGIN                      " & vbCrLf _
                    & " SET @unitcost = @LotUnitCost  " & vbCrLf _
                    & " SET @flgLdCost = 1            " & vbCrLf _
                    & " End                           " & vbCrLf _
                    & " Else                          " & vbCrLf _
                    & " BEGIN                         " & vbCrLf
                     
      sSql = sSql & " IF @lotCostedDate < DATEADD(dd, 1, @ReportDate) " & vbCrLf _
                  & "    BEGIN                            " & vbCrLf _
                  & "    SET @unitcost = @LotUnitCost     " & vbCrLf _
                  & "    SET @flgLdCost = 1   " & vbCrLf _
                  & "    End                  " & vbCrLf _
                  & "    Else                 " & vbCrLf _
                  & "    BEGIN                " & vbCrLf _
                  & "    SET @unitcost = @rptInvcost   " & vbCrLf _
                  & "    SET @flgInvCost = 1           " & vbCrLf _
                  & "    End                        " & vbCrLf _
                  & "    END --LotHD only           " & vbCrLf _
                  & "    END --Part Cost            " & vbCrLf _
                  & "    SELECT @CurACC = dbo.fnGetPartInvAccount(@partRef) " & vbCrLf _
                  & "    -- Lastest < report account#        " & vbCrLf _
                  & "    IF @rptQty >= 0.0000                " & vbCrLf _
                  & "    SET @rptACC = @rptDebitACC          " & vbCrLf _
                  & "    Else                                " & vbCrLf _
                  & "    SET @rptACC = @rptCreditACC         " & vbCrLf _
                  & "    IF ((@rptACC = '') OR (@rptACC = NULL))      " & vbCrLf _
                  & "    BEGIN                   " & vbCrLf _
                  & "    SET @rptACC = @CurACC   " & vbCrLf _
                  & "    SET @flgRptAcc = 1      " & vbCrLf _
                  & "    End                     " & vbCrLf _
                  & "    -- last record          " & vbCrLf _
                  & "    IF @lastQty >= 0.0000      " & vbCrLf
                     
                     
      sSql = sSql & "    SET @lastACC = @lastDebitACC  " & vbCrLf _
                  & "    Else                          " & vbCrLf _
                  & "    SET @lastACC = @lastCreditACC " & vbCrLf _
                  & "    IF ((@lastACC = '') OR (@lastACC = NULL)) " & vbCrLf _
                  & "    BEGIN                   " & vbCrLf _
                  & "    SET @lastACC = @CurACC  " & vbCrLf _
                  & "    SET @flgLastAcc = 1     " & vbCrLf _
                  & "    End                     " & vbCrLf _
                  & "    -- Lastest < report account#  " & vbCrLf _
                  & "    IF @orgQty >= 0.0000          " & vbCrLf _
                  & "    SET @orgACC = @orgDebitACC    " & vbCrLf _
                  & "    Else                          " & vbCrLf _
                  & "    SET @orgACC = @orgCreditACC      " & vbCrLf _
                  & "    IF ((@orgACC = '') OR (@orgACC = NULL))   " & vbCrLf _
                  & "    BEGIN                    " & vbCrLf _
                  & "    SET @orgACC = @CurACC      " & vbCrLf _
                  & "    SET @flgOrgAcc = 1      " & vbCrLf _
                  & "    End                     " & vbCrLf _
                  & "    -- Insert to the temp table   " & vbCrLf _
                  & "    INSERT INTO #tempRMFGoods     " & vbCrLf _
                  & "    (PARTNUM, PALEVEL, PADESC, PAEXTDESC, LOTNUMBER, LOTUSERLOTID, " & vbCrLf _
                  & "    ORGINNUMBER, RPTINNUMBER, CURINNUMBER, ACTUALDATE,RPTDATEQTY, COSTEDDATE,  " & vbCrLf _
                  & "    UNITCOST,PASTDCOST, LOTUNITCOST, ORGCOST, STDCOST, " & vbCrLf _
                  & "    LSTACOST, RPTCOST, CURCOST, RPTACCOUNT,ORIGINALACC,   " & vbCrLf _
                  & "    LASTACTVITYACC, CURRENTACC, PACLASS,PAPRODCODE,    " & vbCrLf
                     
      sSql = sSql & "    flgStdCost, flgLdCost, flgInvCost, flgLdRQErr,  " & vbCrLf _
                  & "    flgRptAcc, flgOrgAcc, flgLastAcc)      " & vbCrLf _
                  & "    VALUES (@partRef, @partType, @partDesc, @partExDesc, @LotNumber,@LotUserID, @OrgInvNum,   " & vbCrLf _
                  & "    @rptInvNum, @LastInvNum, @lotAcualDate,@rptRemQty,@lotCostedDate, " & vbCrLf _
                  & "    @unitcost,@partStdCost, @LotUnitCost, @orgInvcost, @partStdCost,  " & vbCrLf _
                  & "    @lastInvcost,@rptInvcost, @LotUnitCost, @rptACC, @orgACC, @lastACC,  " & vbCrLf _
                  & "    @CurACC, @PartClass,@PartCode,@flgStdCost, @flgLdCost, @flgInvCost,  " & vbCrLf _
                  & "    @flgLdRQErr, @flgRptAcc, @flgOrgAcc, @flgLastAcc)                    " & vbCrLf _
                  & "    SET @rptRemQty = NULL   " & vbCrLf _
                  & "    SET @deltaQty = NULL    " & vbCrLf _
                  & "    SET @lotRemQty = NULL   " & vbCrLf _
                  & "    SET @lotOrgQty = NULL   " & vbCrLf _
                  & "    SET @tmpInvQty = NULL   " & vbCrLf _
                  & "    SET @tmpLastInvQty  = NULL " & vbCrLf _
                  & "    SET @rptInvcost = NULL  " & vbCrLf _
                  & "    SET @lastInvcost = NULL " & vbCrLf _
                  & "    SET @orgInvcost = NULL  " & vbCrLf _
                  & "    SET @tmpInvCost = NULL  " & vbCrLf _
                  & "    SET @unitcost = NULL    " & vbCrLf _
                  & "    SET @tmpLastInvCost = NULL " & vbCrLf _
                  & "    SET @lastQty = NULL  " & vbCrLf _
                  & "    SET @orgQty = NULL   " & vbCrLf _
                  & "    SET @rptQty  = NULL  " & vbCrLf _
                  & "    SET @rptCreditACC = NULL   " & vbCrLf _
                  & "    SET @rptDebitACC  = NULL   " & vbCrLf
                     
      sSql = sSql & "   SET @rptACC  = NULL     " & vbCrLf _
                  & "   SET @tmpCreditAcc  = NULL  " & vbCrLf _
                  & "   SET @tmpDebitAcc = NULL    " & vbCrLf _
                  & "   SET @tmplastCreditAcc = NULL  " & vbCrLf _
                  & "   SET @tmpLastDebitAcc = NULL   " & vbCrLf _
                  & "   SET @lastCreditACC = NULL     " & vbCrLf _
                  & "   SET @lastDebitACC = NULL      " & vbCrLf _
                  & "   SET @lastACC  = NULL          " & vbCrLf _
                  & "   SET @orgCreditACC  = NULL     " & vbCrLf _
                  & "   SET @orgDebitACC  = NULL      " & vbCrLf _
                  & "   SET @orgACC  = NULL           " & vbCrLf _
                  & "   SET @OrgInvNum = NULL         " & vbCrLf _
                  & "   SET @rptInvNum = NULL         " & vbCrLf _
                  & "   SET @LastInvNum = NULL        " & vbCrLf _
                  & "   End                     " & vbCrLf _
                  & "   FETCH NEXT FROM curLotHd INTO @LotNumber, @LotUserID, @partRef,   " & vbCrLf _
                  & "   @partDesc, @partExDesc, @lotAcualDate, @lotOrgQty,                " & vbCrLf _
                  & "   @lotRemQty,@LotUnitCost, @lotCostedDate,                          " & vbCrLf _
                  & "   @PartClass, @PartCode, @partStdCost, @partActCost, @partLotTrack, @partType   " & vbCrLf _
                  & "   End                                    " & vbCrLf _
                  & "   CLOSE curLotHd   --// close the cursor " & vbCrLf _
                  & "   DEALLOCATE curLotHd                    " & vbCrLf _
                  & "   SELECT * FROM #tempRMFGoods            " & vbCrLf _
                  & "                                           " & vbCrLf _
                  & "   DROP table #tempRMFGoods            " & vbCrLf
                     
      sSql = sSql & "   End " & vbCrLf

            Execute True, sSql
    
        'set version
        Execute False, "UPDATE Version set Version = " & newver
    End If
End Function

Private Function UpdateDatabase10()
   newver = 69
   If ver < newver Then
       ver = newver
         
         ' Alter the LoitTable to add MO Pick canceled status
         If Not ColumnExists("LoitTable", "LOIMOPKCANCEL") Then
            Execute True, "ALTER TABLE dbo.LoitTable ADD LOIMOPKCANCEL smallint NULL"
         End If

         Dim RdoMrpSeed As rdoResultset

         sSql = "SELECT * FROM MrpdTable"
         bSqlRows = GetDataSet(RdoMrpSeed, ES_FORWARD)
         If Not bSqlRows Then
            Dim strDate As String
            strDate = Format(GetServerDateTime(), "mm/dd/yy")
            sSql = "INSERT INTO MrpdTable (MRP_ROW,MRP_CREATEDATE,MRP_THROUGHDATE) " & vbCrLf _
                     & " VALUES(1,'" & strDate & "','" & strDate & "')"
            Execute False, sSql
         End If
         
         ' Clear the database
         ClearResultSet RdoMrpSeed

        'set version
        Execute False, "UPDATE Version set Version = " & newver
    End If
End Function


Private Function UpdateDatabase11()
   newver = 70
   If ver < newver Then
       ver = newver
       
      Execute False, "drop table EsReportUsers"
      Execute False, "drop table EsReportUserPermissions"

      sSql = "CREATE TABLE [dbo].[EsReportUsers]" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   [UserID] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf
      sSql = sSql & "   [UserName] [varchar](40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf
      sSql = sSql & "   [Initials] [varchar](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & vbCrLf
      sSql = sSql & "   [Nickname] [varchar](20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & vbCrLf
      sSql = sSql & "   [Active] [bit] NULL," & vbCrLf
      sSql = sSql & "   [Created] [datetime] NULL," & vbCrLf
      sSql = sSql & "   [Level] [int] NULL," & vbCrLf
      sSql = sSql & "   [Admin] bit NULL," & vbCrLf
      sSql = sSql & "   CONSTRAINT [PK_EsReportUsers] PRIMARY KEY CLUSTERED " & vbCrLf
      sSql = sSql & "   (" & vbCrLf
      sSql = sSql & "      [UserID] ASC" & vbCrLf
      sSql = sSql & "   )" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      Execute False, sSql

      sSql = "CREATE TABLE [dbo].[EsReportUserPermissions]" & vbCrLf
      sSql = sSql & "(" & vbCrLf
      sSql = sSql & "   [UserID] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf
      sSql = sSql & "   [UserName] [varchar](40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf
      sSql = sSql & "   [ModuleName] [varchar](40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf
      sSql = sSql & "   [SectionName] [varchar](40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & vbCrLf
      sSql = sSql & "   [GroupPermission] [bit] NOT NULL," & vbCrLf
      sSql = sSql & "   [EditPermission] [bit] NOT NULL," & vbCrLf
      sSql = sSql & "   [ViewPermission] [bit] NOT NULL," & vbCrLf
      sSql = sSql & "   [FunctionPermission] [bit] NOT NULL" & vbCrLf
      sSql = sSql & ")" & vbCrLf
      Execute False, sSql
      
      sSql = "ALTER TABLE [dbo].[EsReportUserPermissions]  WITH CHECK ADD  CONSTRAINT [FK_EsReportUserPermissions_EsReportUsers] FOREIGN KEY([UserID])" & vbCrLf
      sSql = sSql & "REFERENCES [dbo].[EsReportUsers] ([UserID])" & vbCrLf
      Execute False, sSql
      
      sSql = "ALTER TABLE [dbo].[EsReportUserPermissions] CHECK CONSTRAINT [FK_EsReportUserPermissions_EsReportUsers]" & vbCrLf
      Execute False, sSql
           
      ' Modify the view to get only distinct records.
      sSql = "ALTER VIEW [dbo].[viewLotCostsByMoDetails]" & vbCrLf
      sSql = sSql & "   AS " & vbCrLf
      sSql = sSql & "   SELECT DISTINCT dbo.LoitTable.LOITYPE, RTRIM(dbo.LoitTable.LOIMOPARTREF) AS MoPart, dbo.LoitTable.LOIMORUNNO AS MoRun," & vbCrLf
      sSql = sSql & "      RTRIM(dbo.LohdTable.LOTPARTREF) AS Part, - dbo.LoitTable.LOIQUANTITY AS Quantity, dbo.LohdTable.LOTUNITCOST AS UnitCost," & vbCrLf
      sSql = sSql & "      CAST(- (dbo.LohdTable.LOTUNITCOST * dbo.LoitTable.LOIQUANTITY) AS decimal(15, 4)) AS TotalCost, dbo.LohdTable.LOTNUMBER," & vbCrLf
      sSql = sSql & "      RTRIM(dbo.LohdTable.LOTUSERLOTID) AS LotUserID, dbo.LohdTable.LOTMAINTCOSTED, dbo.LohdTable.LOTDATECOSTED" & vbCrLf
      sSql = sSql & "   FROM  dbo.LohdTable INNER JOIN " & vbCrLf
      sSql = sSql & "      dbo.LoitTable ON dbo.LohdTable.LOTNUMBER = dbo.LoitTable.LOINUMBER AND dbo.LoitTable.LOITYPE = 10 INNER JOIN " & vbCrLf
      sSql = sSql & "      dbo.PartTable ON dbo.LohdTable.LOTPARTREF = dbo.PartTable.PARTREF AND dbo.PartTable.PALEVEL < 5 AND dbo.PartTable.PALOTTRACK = 1 "
      Execute False, sSql
      
      
      'set version
      Execute False, "UPDATE Version set Version = " & newver
    
       
   End If
End Function


Private Function IsUpdateRequired(OldDbVersion As Integer, NewDbVersion As Integer) As Boolean
   'terminates if cannot proceed.
   'returns false if no update required
   'returns true if update required and authorized by and admin

   Dim strFulVer As String
   
   sSql = "select * from Updates" & vbCrLf _
      & "where UpdateID = (select max(UpdateID) from Updates)"
   On Error Resume Next
   Dim rdo As rdoResultset
   If GetDataSet(rdo) Then
      If Err = 0 Then
         oldRelease = rdo!newRelease
      End If
   Else
      oldRelease = 0
   End If
   
   If IsTestDatabase Then
      oldType = "Test"
   Else
      oldType = "Live"
   End If
   
   If App.Minor < 10 Then
    strFulVer = CStr(App.Major) & "0" & CStr(App.Minor)
   Else
    strFulVer = CStr(App.Major) & CStr(App.Minor)
   End If
   newRelease = CInt(strFulVer)
   
   If InTestMode() Then
      NewType = "Test"
   Else
      NewType = "Live"
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

Private Sub Execute(DisplayError As Boolean, sql As String)
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
   
   RdoCon.Execute saveSql
   
   DoEvents

   'display error if required
   'always display a timeout error
   If Err Then
      Debug.Print CStr(Err.Number) & ": " & "  " & Err.Description & vbCrLf & " SQL: " & sql
      Dim msg As String
      If DisplayError Or InStr(1, Err.Description, "timeout expired", vbTextCompare) > 0 Then
         msg = "Database version " & ver & " update failed with error " & CStr(Err.Number) & vbCrLf _
            & Err.Description & vbCrLf
         If InStr(1, Err.Description, "timeout expired", vbTextCompare) > 0 Then
            msg = msg & RdoCon.QueryTimeout & " second timeout occurred performing database update." & vbCrLf
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
   Execute False, "exec DropColumnDefault '" & TableName & "', '" & ColumnName & "'"
   'drop defaults created by sp_binddefault (DEFZERO and DEFBLANK)
   Execute False, "exec sp_unbindefault '" & TableName & "." & ColumnName & "'"
   'alter the column
   Execute False, "alter table " & TableName & " alter column " & ColumnName _
      & " " & NewType
   'add a zero constraint back
   Execute False, "alter table " & TableName & " add constraint DF_" & TableName & "_" _
      & ColumnName & " default 0 for " & ColumnName
   
End Sub

Private Sub AlterStringColumn(TableName As String, ColumnName As String, NewType As String)
   'remove default constraint if any, alter column, and add a default of blank
   'example:
   'AlterStringColumn "RndlTable", "RUNDLSDOCREF", "varchar(30)"
   
   'drop defaults created by alter table, if any
   Execute False, "exec DropColumnDefault '" & TableName & "', '" & ColumnName & "'"
   'drop defaults created by sp_binddefault (DEFZERO and DEFBLANK)
   Execute False, "exec sp_unbindefault '" & TableName & "." & ColumnName & "'"
   'alter the column
   Execute False, "alter table " & TableName & " alter column " & ColumnName _
      & " " & NewType & " NULL"
   'add a blank constraint back
   Execute False, "alter table " & TableName & " add constraint DF_" & TableName & "_" _
      & ColumnName & " default '' for " & ColumnName
   
End Sub

Private Sub DropColumnDefault(TableName As String, ColumnName As String)
   'remove default constraint, if any, from a column
   
   'drop defaults created by alter table, if any
   Execute False, "exec DropColumnDefault '" & TableName & "', '" & ColumnName & "'"
   'drop defaults created by sp_binddefault (DEFZERO and DEFBLANK)
   Execute False, "exec sp_unbindefault '" & TableName & "." & ColumnName & "'"
   
End Sub

Private Function ColumnExists(TableName As String, ColumnName As String) As Boolean
   'returns True if column exists
   
   sSql = "SELECT 1 FROM information_schema.Columns" & vbCrLf _
      & "WHERE COLUMN_NAME = '" & ColumnName & "'" & "AND TABLE_NAME = '" & TableName & "'"
   Dim rdo As rdoResultset
   If GetDataSet(rdo) Then
      ColumnExists = True
   End If
End Function

Private Function TableExists(strTableName As String) As Boolean
   'returns True if column exists
   On Error Resume Next
   Err.Clear
   sSql = "SELECT * FROM " & strTableName
   Dim rdo As rdoResultset
   bSqlRows = GetDataSet(rdo)
   If Err.Number = 0 Then
      TableExists = True
   Else
      TableExists = False
   End If
End Function

Public Function GetSysLogon(bGetLogin As Byte) As String
   'get database login info
   'if bGetLogin = 0 then get the password
   'if bGetLogin = 1 then get the login
   
   Dim a As Integer
   Dim iLen As Integer
   Dim sTest As String
   Dim sNewString As String
   Dim sPassword As String
   
   If bGetLogin <> 0 Then
      'GetSysLogon = GetSetting("UserObjects", "System", "NoReg", GetSysLogon)
      GetSysLogon = GetUserSetting(USERSETTING_SqlLogin)
      If Trim(GetSysLogon) = "" Then GetSysLogon = "sa"
   Else
      '        sPassword = GetSetting("SysCan", "System", "RegOne", sPassword)
      '        sPassword = Trim(sPassword)
      '        If sPassword = "H01E2" Then sPassword = ""
      '        If sPassword <> "" Then
      '            iLen = Len(sPassword)
      '            If iLen > 5 Then
      '                sPassword = Mid(sPassword, 4, iLen - 5)
      '            End If
      '        End If
      sPassword = GetDatabasePassword
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

Public Function GetSystemCaption1() As String
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
End Function

Public Sub GetCurrentDatabase()
   'Database
   'sDataBase = GetSetting("Esi2000", "System", "CurDatabase", sDataBase)
   sDataBase = GetUserSetting(USERSETTING_DatabaseName)
   If Trim(sDataBase = "") Then sDataBase = "Esi2000Db"
   
End Sub


'Note: Skips over KeySets and Dynamic Cursors

Public Sub ClearResultSet1(RdoDataSet As rdoResultset)
   If Not RdoDataSet.Updatable Then
      Do While RdoDataSet.MoreResults
      Loop
      RdoDataSet.Cancel
   End If
   
End Sub


Sub GetCompany(Optional bWantAddress As Byte)
   Dim ActRs As rdoResultset
   Dim bByte As Byte
   Dim a As Integer
   Dim b As Integer
   Dim C As Integer
   Dim d As Integer
   Dim sAddress As String
   
   On Error GoTo ModErr1
   If bWantAddress Then
      sSql = "SELECT COREF,CONAME,COPHONE,COFAX,COGLVERIFY,COADR FROM ComnTable " _
             & "WHERE COREF=1"
   Else
      sSql = "SELECT COREF,CONAME,COPHONE,COFAX,COGLVERIFY FROM ComnTable " _
             & "WHERE COREF=1"
   End If
   bSqlRows = GetDataSet(ActRs, ES_STATIC)
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
      a = InStr(1, sAddress, Chr(13) & Chr(10))
      Co.Addr(1) = Left(sAddress, a - 1)
      
      sAddress = Right(sAddress, Len(sAddress) - (a + 1))
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
   
ModErr1:
   Resume modErr2
modErr2:
   On Error GoTo 0
   
End Sub




'Make sure that the user's DSN is pointed to the
'correct server. If none is registered, then build it

Public Function RegisterSqlDsn(sDataSource As String) As String
   Dim sAttribs As String
   If sDataSource = "" Then sDataSource = "ESI2000"
   sAttribs = "Description=" _
              & "ES/2000ERP SQL Server Data " _
              & vbCrLf & "OemToAnsi=No" _
              & vbCrLf & "SERVER=" & sserver _
              & vbCrLf & "Database=" & sDataBase
   'Create new DSN or revise registered DSN.
   rdoEngine.rdoRegisterDataSource sDataSource, _
      "SQL Server", True, sAttribs
   RegisterSqlDsn = sDataSource
   Exit Function
   
ModErr1:
   On Error GoTo 0
   RegisterSqlDsn = sDataSource
   
End Function



'Code  GETSERVERDATETIME() = Format(GetServerDateTime,"mm/dd/yy") etal
'11/21/06 Revised for clarity

Public Function GetServerDateTime() As Variant
   Dim RdoTme As rdoResultset
   On Error GoTo ModErr1
   sSql = "SELECT GETDATE() AS ServerTime"
   bSqlRows = GetDataSet(RdoTme, ES_FORWARD)
   If bSqlRows Then GetServerDateTime = RdoTme!ServerTime
   Set RdoTme = Nothing
   Exit Function
   
ModErr1:
   GetServerDateTime = Format(Now, "mm/dd/yy")
   
End Function

Public Function GetServerDate() As Variant
   Dim RdoTme As rdoResultset
   On Error Resume Next
   GetServerDate = Format(Now, "mm/dd/yy")
   sSql = "select getdate() AS ServerTime"
   If GetDataSet(RdoTme, ES_FORWARD) Then
      GetServerDate = Format(RdoTme!ServerTime, "mm/dd/yy")
   End If
End Function

Private Sub AddNonNullColumnWithDefault(TableName As String, ColumnName As String, _
   TypeName As String, DefaultValue As String)

   'if column already exists, just return
   If ColumnExists(TableName, ColumnName) Then
      Exit Sub
   End If
   
   'add a non-null column and do the gyrations to give it a default
   Execute True, "ALTER TABLE " & TableName & " ADD " & ColumnName & " " & TypeName & " NULL"
   Execute True, "UPDATE " & TableName & " SET " & ColumnName & " = " & DefaultValue
   Execute True, "ALTER TABLE " & TableName & " ALTER COLUMN " & ColumnName & " " & TypeName & " NOT NULL"
   Execute True, "ALTER TABLE " & TableName & " ADD CONSTRAINT DF_" & TableName & "_" & ColumnName & " DEFAULT " & DefaultValue & " FOR " & ColumnName


End Sub

Public Function IsTestDatabase() As Boolean
   'returns true if using a test database
   
   On Error Resume Next
   sSql = "select TestDatabase from Version"
   Dim rdo As rdoResultset
   Dim TestDatabase As Boolean
   If GetDataSet(rdo) Then
      If Err = 0 Then
         IsTestDatabase = IIf(rdo.rdoColumns(0) = 1, True, False)
      Else
         RdoCon.Execute "alter table Version add TestDatabase tinyint not null default 0"
      End If
   End If
   
End Function

Private Sub DropIndex(TableName As String, IndexName As String)
   'drop index that works with SQL2000, SQL2005, and SQL2008
   
      Execute False, "DROP INDEX " & TableName & "." & IndexName & " -- SQL2000"
      Execute False, "DROP INDEX " & IndexName & " ON " & TableName & " -- SQL2005 & SQL2008"
End Sub
