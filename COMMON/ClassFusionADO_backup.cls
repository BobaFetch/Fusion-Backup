VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFusionADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Enumeration - Connection Drivers ("DRIVER=XXXX;")
Public Enum ConnDrivers
  cd_NotSpecified = 0
  cd_Access = 1
  cd_Excel = 2
  cd_Excel_97 = 3
  cd_Oracle = 4
  cd_Paradox = 5
  cd_SQL = 6
  cd_Text = 7
  cd_VFoxPro = 8
  cd_VFoxPro_Cont = 9
End Enum

' Enumeration - Connection Providers ("PROVIDER=XXXX;")
Public Enum ConnProviders
  cp_NotSpecified = 0
  cp_Access = 1
  cp_dBaseIII = 2
  cp_dBaseIV = 3
  cp_dBaseV = 4
  cp_Excel = 5
  cp_HTML = 6
  cp_MsIndexingSrv = 7
  cp_Oracle = 8
  cp_SQL = 9
  cp_Text = 10
End Enum

' Enumeration - Connection Network Protocols ("NETWORK LIBRARY=XXXX:")
Public Enum ConnNetProtocols
  cn_NotSpecified = 0
  cn_AppleTalk = 1
  cn_BanyanVINES = 2
  cn_IPX_SPX = 3
  cn_Multiprotocol = 4
  cn_NamedPipes = 5
  cn_TCP_IP = 6
  cn_VIA_Giganet = 7
End Enum

' Constants - Max Numbers
Private Const MAX_DRIVERS = 9
Private Const MAX_PROVIDERS = 10
Private Const MAX_NET_PROTOCOLS = 7

' Private variable declarations
Private ConnIsolationLvl As IsolationLevelEnum
Private ConnMode         As ConnectModeEnum
Private lngRecAffected   As Long
Private lngConnTimeout   As Long
Private lngCmdTimeout    As Long
Private strConnString    As String
Private blnUseTrans      As Boolean
Private strConnDriv()    As String
Private strConnProv()    As String
Private strConnNetProt() As String
Private blnConnInfoBuilt As Boolean

' Private varaible declarations (with EVENTS)
Private WithEvents conConnection  As ADODB.Connection
Attribute conConnection.VB_VarHelpID = -1

' Event Declarations
Public Event ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)


Private Sub Class_Initialize()
On Error Resume Next
  
  ' Set the defaults (these are defined as the defaults in the MSDN for ADO)
  lngConnTimeout = 60
  lngCmdTimeout = 120
  ConnIsolationLvl = adXactCursorStability
  ConnMode = adModeUnknown
  
End Sub

Private Sub Class_Terminate()
On Error Resume Next
  
  CleanupConnection conConnection
  
End Sub


' This event fires whenever a SQL Statement or Stored Procedure in this class module finishes executing
Private Sub conConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error Resume Next
  
  lngRecAffected = RecordsAffected
  RaiseEvent ExecuteComplete(RecordsAffected, pError, adStatus, pCommand, pRecordset, pConnection)
  
End Sub


' Sets the timeout (in seconds) for all queries run through a command object
Public Property Get CommandTimeout() As Long
  CommandTimeout = lngCmdTimeout
End Property

Public Property Let CommandTimeout(ByVal NewValue As Long)
  If NewValue <> lngCmdTimeout Then
    lngCmdTimeout = NewValue
    CleanupConnection conConnection
    CheckConnection
  End If
End Property

Public Property Get ConnectionIsolationLevel() As IsolationLevelEnum
  ConnectionIsolationLevel = ConnIsolationLvl
End Property


Public Property Let ConnectionIsolationLevel(ByVal NewValue As IsolationLevelEnum)
  If NewValue <> ConnIsolationLvl Then
    ConnIsolationLvl = NewValue
    CleanupConnection conConnection
    CheckConnection
  End If
End Property


Public Property Get ConnectionMode() As ConnectModeEnum
  ConnectionMode = ConnMode
End Property
Public Property Let ConnectionMode(ByVal NewValue As ConnectModeEnum)
  If NewValue <> ConnMode Then
    ConnMode = NewValue
    CleanupConnection conConnection
    CheckConnection
  End If
End Property

Public Property Get ConnectionObject() As ADODB.Connection
On Error Resume Next
  Set ConnectionObject = conConnection
End Property


Public Function BeginTrans()

   If (Not conConnection Is Nothing) Then
      If (conConnection.State = adStateOpen) Then
         conConnection.BeginTrans
      End If
   End If
   conConnection
End Function

Public Function CommitTrans()


   If (Not conConnection Is Nothing) Then
      If (conConnection.State = adStateOpen) Then
         conConnection.CommitTrans
      End If
   End If
End Function

Public Function RollbackTrans()

   If (Not conConnection Is Nothing) Then
      If (conConnection.State = adStateOpen) Then
         conConnection.RollbackTrans
      End If
   End If
End Function


' Returns the current connection string.  To set this value, use the "OpenConnection" method
Public Property Get ConnectionString(Optional blnRealConnString As Boolean = False) As String
  
  If blnRealConnString = True Then
    If CheckConnection = False Then
      ConnectionString = ""
    Else
      ConnectionString = conConnection.ConnectionString
    End If
  Else
    ConnectionString = strConnString
  End If
  
End Property

Public Property Get ConnectionTimeout() As Long
  ConnectionTimeout = lngConnTimeout
End Property


Public Property Let ConnectionTimeout(ByVal NewValue As Long)
  If NewValue <> lngConnTimeout Then
    lngConnTimeout = NewValue
    CleanupConnection conConnection
    CheckConnection
  End If
End Property

Public Property Get UseTransactions() As Boolean
  UseTransactions = blnUseTrans
End Property

Public Property Let UseTransactions(ByVal NewValue As Boolean)
  blnUseTrans = NewValue
End Property

Public Function CleanupCommand(ByRef objCommand As ADODB.Command) As Boolean
On Error Resume Next
  
  If Not objCommand Is Nothing Then
    If Not objCommand.ActiveConnection Is Nothing Then Set objCommand.ActiveConnection = Nothing
    Set objCommand = Nothing
    Err.Clear
  End If
  CleanupCommand = True
  
End Function


Public Function CleanupConnection(ByRef objConnection As ADODB.Connection) As Boolean
On Error Resume Next
  
  If Not objConnection Is Nothing Then
    If objConnection.State <> adStateClosed Then objConnection.Close
    Set objConnection = Nothing
    Err.Clear
  End If
  CleanupConnection = True
  
End Function


Public Function CleanupRecordset(ByRef objRecordset As ADODB.Recordset, _
                                 Optional ByVal blnCloseRS As Boolean = True) As Boolean
On Error Resume Next
  
  If Not objRecordset Is Nothing Then
    Set objRecordset.ActiveConnection = Nothing
    If blnCloseRS = True Then If objRecordset.State <> adStateClosed Then objRecordset.Close
    Set objRecordset = Nothing
    Err.Clear
  End If
  CleanupRecordset = True
  
End Function


Public Function GetDataSet(ByVal strSQL As String, _
               ByRef rsDataSet As ADODB.Recordset, _
               Optional iCursorType As Integer = ES_FORWARD) As Integer
   
   On Error GoTo ErrorTrap
  
   Err.Clear
   CleanupRecordset rsDataSet
   
   Dim Return_ErrNum As Long
   Dim Return_ErrDesc As String
   Return_ErrNum = -1
   Return_ErrDesc = ""
   ' Clear variables
   Err.Clear
   CleanupRecordset rsDataSet
   
   strSQL = Trim(strSQL)
   If strSQL = "" Then
      Return_ErrNum = -1
      Return_ErrDesc = "No SQL statement specified to execute"
      Exit Function
   ' Check the connection to the database
   ElseIf CheckConnection(Return_ErrNum, Return_ErrDesc) = False Then
      Exit Function
   End If
   
   Set rsDataSet = New ADODB.Recordset
   rsDataSet.CursorLocation = adUseServer
   
   If iCursorType = ES_FORWARD Then
      rsDataSet.CursorType = adOpenForwardOnly
      'Forward only "cursor" (not a cursor)
      rsDataSet.Open strSQL, conConnection, , , adCmdText
      If Not rsDataSet.BOF And Not rsDataSet.EOF Then
         GetRecordSet = 1
      Else
         GetRecordSet = 0
      End If
   Else
         If iCursorType = ES_KEYSET Then
         rsDataSet.CursorType = adOpenKeyset
         'Keyset cursor for Editing
         rsDataSet.Open strSQL, conConnection, , , adCmdText
         If Not rsDataSet.BOF And Not rsDataSet.EOF Then
            GetRecordSet = 1
         Else
            GetRecordSet = 0
         End If
      ElseIf iCursorType = ES_DYNAMIC Then
         'Dynamic
         rsDataSet.CursorType = adOpenDynamic
         rsDataSet.Open strSQL, conConnection, , , adCmdText
         If Not rsDataSet.BOF And Not rsDataSet.EOF Then
            GetRecordSet = 1
         Else
            GetRecordSet = 0
         End If
      ElseIf iCursorType = ES_STATIC Then
         'Static Cursor - Note: Needed for BLOBS
         rsDataSet.CursorType = adOpenStatic
         rsDataSet.Open strSQL, conConnection, , , adCmdText
         
         If Not rsDataSet.BOF And Not rsDataSet.EOF Then
            GetRecordSet = 1
         Else
            GetRecordSet = 0
         End If
      End If
   End If
   If Err > 0 Then GetRecordSet = 0

   Exit Function
ErrorTrap:
  
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  
  CleanupRecordset rsDataSet
End Function


Public Function GetQuerySet(ByRef rsDataSet As ADODB.Recordset, _
         ByVal cmdObj As ADODB.Command, Optional iCursorType As Integer = ES_FORWARD) As Boolean
   
   On Error GoTo ErrorTrap
   Err.Clear
   CleanupRecordset rsDataSet
   
   If (cmdObj.CommandText <> "") Then
      
      cmdObj.ActiveConnection = conConnection
      cmdObj.CommandType = adCmdText
      cmdObj.CommandTimeout = 15
      Set rsDataSet = cmdObj.Execute
      rsDataSet.CursorLocation = adUseServer
      
      If Not rsDataSet.BOF And Not rsDataSet.EOF Then
         GetQuerySet = 1
      Else
         GetQuerySet = 0
      End If
      
      If iCursorType = ES_FORWARD Then
         rsDataSet.CursorType = adOpenForwardOnly
      ElseIf iCursorType = ES_KEYSET Then
         rsDataSet.CursorType = adOpenKeyset
      Else
         rsDataSet.CursorType = adOpenStatic
            'Static Cursor - Note: Needed for BLOBS
      End If
      
      If Err > 0 Then GetQuerySet = 0
   End If
   
   Exit Function
ErrorTrap:
  
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  
  CleanupRecordset rsDataSet
End Function

Public Function GetRecordSet(ByVal strSQL As String, Optional iCursorType As Integer = rdOpenForwardOnly) As ADODB.Recordset
   
   On Error GoTo ErrorTrap
   Err.Clear
   
   Dim strErrNum As Long
   Dim strErrDesc As String
   
   ' Check the SQL statement
   strSQL = Trim(strSQL)
   If strSQL = "" Then
      strErrNum = -1
      strErrDesc = "No SQL statement specified to execute"
      Exit Function
   ElseIf CheckConnection(strErrNum, strErrDesc) = False Then
     Exit Function
   End If
  
   ' Setup the recordset to use
   Set rsDataSet = New ADODB.Recordset
   rsDataSet.CursorLocation = adUseServer
   rsDataSet.CursorType = iCursorType
   
   ' Execute
   rsDataSet.Open strSQL, conConnection, , , adCmdText
   
   GetRecordSet = rsDataSet
   
   Exit Function
ErrorTrap:
  
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  
  CleanupRecordset rsDataSet
End Function



Public Function ExecuteSQL(ByVal strSQL As String, _
                           Optional ByVal blnStoredProc As Boolean = False, _
                           Optional ByVal RsCursorLocation As ADODB.CursorLocationEnum = adUseServer, _
                           Optional ByVal RsCursorType As ADODB.CursorTypeEnum = adOpenKeyset, _
                           Optional ByVal RsLockType As ADODB.LockTypeEnum = adLockPessimistic, _
                           Optional ByRef Return_Affected As Long, _
                           Optional ByRef Return_RecordCount As Long = -1, _
                           Optional ByRef Return_ErrNum As Long, _
                           Optional ByRef Return_ErrDesc As String, _
                           Optional ByRef Return_Recordset As ADODB.Recordset, _
                           Optional ByVal blnDisconnectTheRS As Boolean = True, _
                           Optional ByVal lngRecordsetCacheSize As Long = 1, _
                           Optional ByVal blnDeleteRsIfBlank As Boolean = True) As Boolean
On Error GoTo ErrorTrap
  
  ' Clear variables
  Err.Clear
  CleanupRecordset Return_Recordset
  lngRecAffected = -1
  Return_Affected = -1
  Return_ErrNum = -1
  Return_ErrDesc = ""
  
  ' Check the SQL statement
  strSQL = Trim(strSQL)
  If strSQL = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "No SQL statement specified to execute"
    Exit Function
    
  ' Check the connection to the database
  ElseIf CheckConnection(Return_ErrNum, Return_ErrDesc) = False Then
    Exit Function
  End If
  
  ' Execute the stored procedure
  If blnStoredProc = True Then
    
    ' If you're executing a stored procedure with NO PARAMETERS, and you include the open/close
    ' parentheses "()", it will give you an error... so take them off here
    If Right(strSQL, 2) = "()" Then strSQL = Left(strSQL, Len(strSQL) - 2)
    
    ' Setup the recordset to use
    Set Return_Recordset = New ADODB.Recordset
    Return_Recordset.CursorLocation = RsCursorLocation
    Return_Recordset.CursorType = RsCursorType
    Return_Recordset.LockType = RsLockType
    Return_Recordset.CacheSize = lngRecordsetCacheSize
    
    ' If user specified to use transactions, start the transaction
    If blnUseTrans = True Then conConnection.BeginTrans
    
    ' Execute
    Return_Recordset.Open strSQL, conConnection, , , adCmdStoredProc
    
    ' If user specified to use transactions, commit the transaction
    If blnUseTrans = True Then conConnection.CommitTrans
    
    ' If the user specified to do so, disconnect the RS from the database
    If blnDisconnectTheRS = True Then Set Return_Recordset.ActiveConnection = Nothing
    
    ' Check if the recordset returned is valid.
    If IsRecordsetValid(Return_Recordset, True, blnDeleteRsIfBlank) = True Then
      
      ' If the user wants the record count back, loop through it and count the records (this works on all cursor types and locations)
      If Return_RecordCount <> -1 Then
        Return_RecordCount = 0
        Return_Recordset.MoveFirst
        Do While Return_Recordset.EOF = False
          Return_RecordCount = Return_RecordCount + 1
          Return_Recordset.MoveNext
        Loop
        Return_Recordset.MoveFirst
      End If
      
    Else
      Return_RecordCount = 0
    End If
    
    ' Return the number of records affected
    Return_Affected = lngRecAffected
    
  ' Execute the SQL statement
  Else
    
    ' If the SQL statement is a SELECT statement, use the ADODB.Recordset object
    If UCase(Left(strSQL, 7)) = "SELECT " Then
      
      ' Setup the recordset to use
      Set Return_Recordset = New ADODB.Recordset
      Return_Recordset.CursorLocation = RsCursorLocation
      Return_Recordset.CursorType = RsCursorType
      Return_Recordset.LockType = RsLockType
      Return_Recordset.CacheSize = lngRecordsetCacheSize
      
      ' If user specified to use transactions, start the transaction
      If blnUseTrans = True Then conConnection.BeginTrans
      
      ' Execute
      Return_Recordset.Open strSQL, conConnection, , , adCmdText
      
      ' If user specified to use transactions, commit the transaction
      If blnUseTrans = True Then conConnection.CommitTrans
      
      ' If the user specified to do so, disconnect the RS from the database
      If blnDisconnectTheRS = True Then Set Return_Recordset.ActiveConnection = Nothing
      
      ' Check if the recordset returned is valid.
      If IsRecordsetValid(Return_Recordset, True, blnDeleteRsIfBlank) = True Then
        
        ' If the user wants the record count back, loop through it and count the records (this works on all cursor types and locations)
        If Return_RecordCount <> -1 Then
          Return_RecordCount = 0
          Return_Recordset.MoveFirst
          Do While Return_Recordset.EOF = False
            Return_RecordCount = Return_RecordCount + 1
            Return_Recordset.MoveNext
          Loop
          Return_Recordset.MoveFirst
        End If
        
      Else
        Return_RecordCount = 0
      End If
      
      ' Nothing affected when executing a SELECT statement
      Return_Affected = 0
      
    ' The SQL statement is not a SELECT statement, so use the Connection object in order to get back how many records were affected
    Else
      
      ' If user specified to use transactions, start the transaction
      If blnUseTrans = True Then conConnection.BeginTrans
      
      ' Execute
      conConnection.Execute strSQL, Return_Affected, adCmdText
      
      ' If user specified to use transactions, commit the transaction
      If blnUseTrans = True Then conConnection.CommitTrans
      
      ' No records are returned on a non-SELECT statement
      Return_RecordCount = 0
      
    End If
  End If
  
  ExecuteSQL = True
  Exit Function
  
ErrorTrap:
  
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  If blnUseTrans = True Then conConnection.RollbackTrans 'Rollback the transaction
  CleanupRecordset Return_Recordset
  
End Function


Public Function ExecuteStoredProc(ByVal strStoredProcName As String, _
                                  ByRef blnParameterIsString() As Boolean, _
                                  ByRef varParameterValues() As Variant, _
                                  ByVal blnParametersPassed As Boolean, _
                                  Optional ByVal RsCursorLocation As ADODB.CursorLocationEnum = adUseClient, _
                                  Optional ByVal RsCursorType As ADODB.CursorTypeEnum = adOpenKeyset, _
                                  Optional ByVal RsLockType As ADODB.LockTypeEnum = adLockPessimistic, _
                                  Optional ByRef Return_Affected As Long, _
                                  Optional ByRef Return_RecordCount As Long = -1, _
                                  Optional ByRef Return_ErrNum As Long, _
                                  Optional ByRef Return_ErrDesc As String, _
                                  Optional ByRef Return_Recordset As ADODB.Recordset, _
                                  Optional ByVal blnDisconnectTheRS As Boolean = True, _
                                  Optional ByVal lngRecordsetCacheSize As Long = 1, _
                                  Optional ByVal blnDeleteRsIfBlank As Boolean = True) As Boolean
On Error GoTo ErrorTrap
  
  Dim BoolUBound   As Long
  Dim BoolLBound   As Long
  Dim ValueUBound  As Long
  Dim ValueLBound  As Long
  Dim lngCounter   As Long
  Dim strSQL       As String
  
  ' Clear variables
  Err.Clear
  CleanupRecordset Return_Recordset
  lngRecAffected = -1
  Return_Affected = -1
  Return_ErrNum = -1
  Return_ErrDesc = ""
  
  ' Check the SQL statement
  strStoredProcName = Trim(strStoredProcName)
  If strStoredProcName = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "No stored procedure name specified to execute"
    Exit Function
    
  ' Check the connection to the database
  ElseIf CheckConnection(Return_ErrNum, Return_ErrDesc) = False Then
    Exit Function
    
  ' Check if there are parameters passed
  ElseIf blnParametersPassed = True Then
    If ArrayIsEmpty_BLN(blnParameterIsString, BoolLBound, BoolUBound) = True Or _
       ArrayIsEmpty_VAR(varParameterValues, ValueLBound, ValueUBound) = True Then
      Return_ErrNum = -1
      Return_ErrDesc = "Parameter count passed to 'ExecuteStoredProc' is greater than zero, but the parameters passed were invalid or missing"
      Exit Function
    ElseIf (BoolLBound <> ValueLBound) Or (BoolUBound <> ValueUBound) Then
      Return_ErrNum = -1
      Return_ErrDesc = "The number of parameter names don't match the number of parameter values passed to the 'ExecuteStoredProc' method"
      Exit Function
    End If
  End If
  
  ' Create the Stored Procedure string
  strSQL = strStoredProcName
  If blnParametersPassed = True Then
    strSQL = strSQL & " ("
    For lngCounter = BoolLBound To BoolUBound
      If blnParameterIsString(lngCounter) = True Then
        If IsEmpty(varParameterValues(lngCounter)) = True Then
          strSQL = strSQL & ", "
        Else
          strSQL = strSQL & "'" & ReplaceChar(CStr(varParameterValues(lngCounter)), "'", "''") & "', "
        End If
      Else
        If IsEmpty(varParameterValues(lngCounter)) = True Then
          strSQL = strSQL & ", "
        Else
          strSQL = strSQL & ReplaceChar(CStr(varParameterValues(lngCounter)), "'", "''") & ", "
        End If
      End If
    Next
    strSQL = Trim(strSQL)
    If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1)
    strSQL = strSQL & ")"
  End If
  
  ' Create the recordset to use
  Set Return_Recordset = New ADODB.Recordset
  Return_Recordset.CursorLocation = RsCursorLocation
  Return_Recordset.CursorType = RsCursorType
  Return_Recordset.LockType = RsLockType
  Return_Recordset.CacheSize = lngRecordsetCacheSize
  
  ' If user specified to use transactions, start the transaction
  If blnUseTrans = True Then conConnection.BeginTrans
  
  ' Execute
  Return_Recordset.Open strSQL, conConnection, , , adCmdStoredProc
  
  ' If user specified to use transactions, commit the transaction
  If blnUseTrans = True Then conConnection.CommitTrans
  
  ' If the user specified to do so, disconnect the RS from the database
  If blnDisconnectTheRS = True Then Set Return_Recordset.ActiveConnection = Nothing
  
  ' Check if the return recordset is valid
  If IsRecordsetValid(Return_Recordset, True, blnDeleteRsIfBlank) = True Then
    
    ' If the user wants the record count back, loop through it and count the records (this works on all cursor types and locations)
    If Return_RecordCount <> -1 Then
      Return_RecordCount = 0
      Return_Recordset.MoveFirst
      Do While Return_Recordset.EOF = False
        Return_RecordCount = Return_RecordCount + 1
        Return_Recordset.MoveNext
      Loop
      Return_Recordset.MoveFirst
    End If
    
  Else
    Return_RecordCount = 0
  End If
  
  Return_Affected = lngRecAffected
  ExecuteStoredProc = True
  
CleanUp:
  
  Exit Function
  
ErrorTrap:
  
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  If blnUseTrans = True Then conConnection.RollbackTrans ' Rollback the transaction
  CleanupRecordset Return_Recordset
  GoTo CleanUp
  
End Function


Public Function ExecuteStoredProcEx(ByVal strStoredProcName As String, _
                                    ByRef strParameterNames() As String, _
                                    ByRef varParameterValues() As Variant, _
                                    ByVal blnParametersPassed As Boolean, _
                                    Optional ByVal RsCursorLocation As ADODB.CursorLocationEnum = adUseClient, _
                                    Optional ByVal RsCursorType As ADODB.CursorTypeEnum = adOpenKeyset, _
                                    Optional ByVal RsLockType As ADODB.LockTypeEnum = adLockPessimistic, _
                                    Optional ByRef Return_Affected As Long, _
                                    Optional ByRef Return_RecordCount As Long = -1, _
                                    Optional ByRef Return_ErrNum As Long, _
                                    Optional ByRef Return_ErrDesc As String, _
                                    Optional ByRef Return_Recordset As ADODB.Recordset, _
                                    Optional ByVal blnDisconnectTheRS As Boolean = True, _
                                    Optional ByVal lngRecordsetCacheSize As Long = 1, _
                                    Optional ByVal blnDeleteRsIfBlank As Boolean = True) As Boolean
On Error GoTo ErrorTrap
  
  Dim objCommand   As ADODB.Command
  Dim NameUBound   As Long
  Dim NameLBound   As Long
  Dim ValueUBound  As Long
  Dim ValueLBound  As Long
  Dim blnErrorParm As Boolean
  Dim strErrorParm As String
  Dim lngCounter   As Long
  
  ' Clear variables
  Err.Clear
  CleanupRecordset Return_Recordset
  lngRecAffected = -1
  Return_Affected = -1
  Return_ErrNum = -1
  Return_ErrDesc = ""
  
  ' Check the SQL statement
  strStoredProcName = Trim(strStoredProcName)
  If strStoredProcName = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "No stored procedure name specified to execute"
    Exit Function
    
  ' Check the connection to the database
  ElseIf CheckConnection(Return_ErrNum, Return_ErrDesc) = False Then
    Exit Function
    
  ' Check if there are parameters passed
  ElseIf blnParametersPassed = True Then
    If ArrayIsEmpty_STR(strParameterNames, NameLBound, NameUBound) = True Or _
       ArrayIsEmpty_VAR(varParameterValues, ValueLBound, ValueUBound) = True Then
      Return_ErrNum = -1
      Return_ErrDesc = "Parameter count passed to 'ExecuteStoredProcEx' is greater than zero, but the parameters passed were invalid or missing"
      Exit Function
    ElseIf (NameLBound <> ValueLBound) Or (NameUBound <> ValueUBound) Then
      Return_ErrNum = -1
      Return_ErrDesc = "The number of parameter names don't match the number of parameter values passed to the 'ExecuteStoredProcEx' method"
      Exit Function
    Else
      For lngCounter = NameLBound To NameUBound
        strParameterNames(lngCounter) = Trim(strParameterNames(lngCounter))
        If strParameterNames(lngCounter) <> "" Then
          If Left(strParameterNames(lngCounter), 1) <> "@" Then strParameterNames(lngCounter) = "@" & strParameterNames(lngCounter)
        End If
      Next
    End If
  End If
  
  ' Create the command object to use
  Set objCommand = New ADODB.Command
  Set objCommand.ActiveConnection = conConnection
  objCommand.CommandTimeout = lngCmdTimeout
  objCommand.CommandType = adCmdStoredProc
  objCommand.CommandText = strStoredProcName
  objCommand.Parameters.Refresh
  
  ' Loop through the parameters passed and
  If blnParametersPassed = True Then
    blnErrorParm = True
    For lngCounter = NameLBound To NameUBound
      strErrorParm = strParameterNames(lngCounter)
      If strParameterNames(lngCounter) <> "" Then
        If UCase(strParameterNames(lngCounter)) <> "@RETURN_VALUE" Then
          objCommand(strParameterNames(lngCounter)) = varParameterValues(lngCounter)
        End If
      End If
    Next
    blnErrorParm = False
  End If
  
  ' Create the recordset to use
  Set Return_Recordset = New ADODB.Recordset
  Return_Recordset.CursorLocation = RsCursorLocation
  Return_Recordset.CursorType = RsCursorType
  Return_Recordset.LockType = RsLockType
  Return_Recordset.CacheSize = lngRecordsetCacheSize
  
  ' If user specified to use transactions, start the transaction
  If blnUseTrans = True Then conConnection.BeginTrans
  
  ' Execute
  Return_Recordset.Open objCommand, , , , adCmdStoredProc
  
  ' If user specified to use transactions, commit the transaction
  If blnUseTrans = True Then conConnection.CommitTrans
  
  ' Function succeeded
  Return_Affected = lngRecAffected
  ExecuteStoredProcEx = True
  
  ' Return the stored parameter OUTPUT parameters
  If blnParametersPassed = True Then AssignReturnValue strParameterNames, varParameterValues, objCommand, NameLBound, NameUBound
  
  ' Check if the return recordset is valid
  If Not Return_Recordset Is Nothing Then
    
    ' If the user specified to do so, disconnect the RS from the database
    If blnDisconnectTheRS = True Then Set Return_Recordset.ActiveConnection = Nothing
    
    If Return_Recordset.State <> adStateClosed Then
      If Return_Recordset.BOF = False Or Return_Recordset.EOF = False Then
        
        ' If the user wants the record count back, loop through it and count the records (this works on all cursor types and locations)
        If Return_RecordCount <> -1 Then
          Return_RecordCount = 0
          Return_Recordset.MoveFirst
          Do While Return_Recordset.EOF = False
            Return_RecordCount = Return_RecordCount + 1
            Return_Recordset.MoveNext
          Loop
          Return_Recordset.MoveFirst
        Else
          Return_RecordCount = 0
          If blnDeleteRsIfBlank = True Then
            Return_Recordset.Close
            Set Return_Recordset = Nothing
          End If
        End If
        
      Else
        Return_RecordCount = 0
        If blnDeleteRsIfBlank = True Then
          Return_Recordset.Close
          Set Return_Recordset = Nothing
        End If
      End If
    Else
      Return_RecordCount = 0
      If blnDeleteRsIfBlank = True Then Set Return_Recordset = Nothing
    End If
  Else
    Return_RecordCount = 0
  End If
  
CleanUp:
  
  CleanupCommand objCommand
  Exit Function
  
ErrorTrap:
  
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  If blnErrorParm = True Then Return_ErrDesc = "Parameter = " & strErrorParm & " : " & Return_ErrDesc
  If blnParametersPassed = True Then AssignReturnValue strParameterNames, varParameterValues, objCommand, NameLBound, NameUBound
  If blnUseTrans = True Then conConnection.RollbackTrans ' Rollback the transaction
  CleanupRecordset Return_Recordset
  GoTo CleanUp
  
End Function


Public Function GetRecordCount(ByRef rsRecordSet As ADODB.Recordset, _
                               ByRef Return_RecordCount As Long, _
                               Optional ByRef Return_ErrNum As Long, _
                               Optional ByRef Return_ErrDesc As String) As Boolean
On Error Resume Next
  
  Dim blnReturn As Boolean
  
  ' Clear variables
  Err.Clear
  Return_RecordCount = -1
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' Validate parameters
  If IsRecordsetValid(rsRecordSet) = False Then
    Return_ErrNum = -1
    Return_ErrDesc = "Invalid ADODB.Recordset object passed to the GetRecordCount function"
    Exit Function
  Else
    Return_RecordCount = 0
  End If
  
  ' Loop through the recordset and get the count
  ' (This works for any type of recordset because they all support MoveNext and MoveFirst)
  rsRecordSet.MoveFirst
  Do While rsRecordSet.EOF = False
    Return_RecordCount = Return_RecordCount + 1
    rsRecordSet.MoveNext
  Loop
  rsRecordSet.MoveFirst
  
  ' Got the count
  GetRecordCount = True
  
End Function


Public Function IsRecordsetValid(ByRef rsRecordSet As ADODB.Recordset, _
                                 Optional ByVal blnCheck_BOF_EOF As Boolean = True, _
                                 Optional ByVal blnDeleteIfInvalid As Boolean = False) As Boolean
On Error Resume Next
  
  ' Recordset is invalid if it's NOTHING
  If rsRecordSet Is Nothing Then
    If blnDeleteIfInvalid = True Then CleanupRecordset rsRecordSet
    
  ' Recordset is invalid if it's state is CLOSED
  ElseIf rsRecordSet.State = adStateClosed Then
    If blnDeleteIfInvalid = True Then CleanupRecordset rsRecordSet
    
  ' If a Recordset has it's "BOF" and "EOF" properties both set to TRUE, there's nothing in the
  ' recordset, but it's still valid... so if the user wants to check for it, do so here.
  ElseIf rsRecordSet.BOF = True And rsRecordSet.EOF = True Then
    If blnCheck_BOF_EOF = True Then
      If blnDeleteIfInvalid = True Then CleanupRecordset rsRecordSet
    Else
      ' Recordset is VALID!
      IsRecordsetValid = True
    End If
    
  Else
    ' Recordset is VALID!
    IsRecordsetValid = True
  End If
  
End Function


Public Function GetConnectionProperty(ByVal strPropertyName As String, _
                                      ByRef Return_PropertyValue As Variant, _
                                      Optional ByRef Return_ErrNum As Long, _
                                      Optional ByRef Return_ErrDesc As String) As Boolean
On Error GoTo ErrorTrap
  
  Dim lngCounter As Long
  Dim blnFoundIt As Boolean
  Dim strExtProp As String
  Dim lngStart   As Long
  Dim lngEnd     As Long
  Dim strName    As String
  Dim vntValue   As Variant
  
  ' Clear variables
  Return_PropertyValue = Empty
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' Make sure there's a valid connection
  If CheckConnection(Return_ErrNum, Return_ErrDesc) = False Then Exit Function
  
  ' Loop through the properties of the current connection and see if the specified property exists
  strPropertyName = UCase(strPropertyName)
  If conConnection.Properties.Count > 0 Then
    For lngCounter = 0 To conConnection.Properties.Count - 1
      
      strName = UCase(conConnection.Properties.Item(lngCounter).Name)
      vntValue = conConnection.Properties.Item(lngCounter).value
      Debug.Print strName & " = " & Trim(vntValue & "")
      
      If strName = strPropertyName Then
        blnFoundIt = True
        Return_PropertyValue = vntValue
        Exit For
      ElseIf strName = "EXTENDED PROPERTIES" Then
        strExtProp = CStr(vntValue)
      End If
    Next
  End If
  
  ' Check the extended properties for the property we're looking for
  If blnFoundIt = False And strExtProp <> "" Then
    lngStart = InStr(1, strExtProp, strPropertyName & "=", vbTextCompare)
    If lngStart > 0 Then
      blnFoundIt = True
      lngStart = lngStart + Len(strPropertyName & "=")
      lngEnd = InStr(lngStart, strExtProp, ";", vbTextCompare)
      If lngEnd < 1 Then
        lngEnd = Len(strExtProp) - lngStart
      Else
        lngEnd = lngEnd - lngStart
      End If
      Return_PropertyValue = Mid(strExtProp, lngStart, lngEnd)
    End If
  End If
  
  ' If we didn't find it, then return that we didn't
  If blnFoundIt = False Then
    Return_ErrNum = -1
    Return_ErrDesc = "The connection property '" & strPropertyName & "' was not found in the current connection."
  Else
    GetConnectionProperty = True
  End If
  Exit Function
  
ErrorTrap:
  
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  
End Function


Public Function GenerateConnectionString(ByRef Return_ConnectionString As String, _
                                         Optional ByVal blnRemoteServer As Boolean = False, _
                                         Optional ByVal strRemoteServerURL As String, _
                                         Optional ByVal lngRemoteServerTimeout As Long = -1, _
                                         Optional ByVal strRemoteServerProvider As String = "", _
                                         Optional ByVal blnDataShape As Boolean = False, _
                                         Optional ByVal strDataShapeProvider As String = "", _
                                         Optional ByVal ConnectionProvider As ConnProviders = cp_NotSpecified, _
                                         Optional ByVal ConnectionDriver As ConnDrivers = cd_NotSpecified, _
                                         Optional ByVal ConnectionNetProtocol As ConnNetProtocols = cn_NotSpecified, _
                                         Optional ByVal strServerName As String, _
                                         Optional ByVal strDatabaseName As String, _
                                         Optional ByVal strDataSource As String, _
                                         Optional ByVal strUserName As String, _
                                         Optional ByVal strPassword As String, _
                                         Optional ByVal strWorkstationID As String = "", _
                                         Optional ByVal lngPacketSize As Long = -1, _
                                         Optional ByVal blnPersistSecurityInfo As String = "", _
                                         Optional ByVal blnUseEncryptionForData As String = "", _
                                         Optional ByVal blnAutoTranslate As String = "", _
                                         Optional ByRef Return_ErrNum As Long, _
                                         Optional ByRef Return_ErrDesc As String) As Boolean
  
  Dim strConn As String
  
  ' Clear variables
  Return_ConnectionString = ""
  Return_ErrNum = 0
  Return_ErrDesc = ""
  strDataShapeProvider = Trim(strDataShapeProvider)
  strRemoteServerURL = Trim(strRemoteServerURL)
  strRemoteServerProvider = Trim(strRemoteServerProvider)
  strServerName = Trim(strServerName)
  strDatabaseName = Trim(strDatabaseName)
  strDataSource = Trim(strDataSource)
  strUserName = Trim(strUserName)
  strPassword = Trim(strPassword)
  strWorkstationID = Trim(strWorkstationID)
  blnPersistSecurityInfo = UCase(Trim(blnPersistSecurityInfo))
  blnUseEncryptionForData = UCase(Trim(blnUseEncryptionForData))
  blnAutoTranslate = UCase(Trim(blnAutoTranslate))
  
  ' Validate parameters
  If blnDataShape = False And blnRemoteServer = False And ConnectionDriver = cd_NotSpecified And ConnectionProvider = cp_NotSpecified Then
    Return_ErrNum = -1
    Return_ErrDesc = "No data provider specified"
    Exit Function
  ElseIf blnDataShape = False And blnRemoteServer = False And strServerName = "" And strDatabaseName = "" And strDataSource = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "No data source specified"
    Exit Function
  ElseIf blnRemoteServer = True And strRemoteServerURL = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "Remove Server specified as the data source, but no URL specified"
    Exit Function
  ElseIf blnDataShape = True And strDataShapeProvider = "" And ConnectionDriver = cd_NotSpecified Then
    Return_ErrNum = -1
    Return_ErrDesc = "Data Shape specified, but no connection driver specified"
    Exit Function
  End If
  If blnRemoteServer = False And blnDataShape = False Then
    If ConnectionProvider <> cp_NotSpecified Then
      If (ConnectionProvider = cp_SQL Or ConnectionProvider = cp_Oracle) And (strServerName = "") Then
        Return_ErrNum = -1
        Return_ErrDesc = "No server name specified"
        Exit Function
      ElseIf (ConnectionProvider = cp_SQL Or ConnectionProvider = cp_Oracle) And (strDatabaseName = "") Then
        Return_ErrNum = -1
        Return_ErrDesc = "No database name specified"
        Exit Function
      ElseIf (ConnectionProvider <> cp_SQL And ConnectionProvider <> cp_Oracle And strDataSource = "") Then
        Return_ErrNum = -1
        Return_ErrDesc = "No data source specified"
        Exit Function
      End If
    ElseIf ConnectionDriver <> cd_NotSpecified Then
      If (ConnectionDriver = cd_SQL Or ConnectionDriver = cd_Oracle) And (strServerName = "") Then
        Return_ErrNum = -1
        Return_ErrDesc = "No server name specified"
        Exit Function
      ElseIf (ConnectionDriver = cd_SQL Or ConnectionDriver = cd_Oracle) And (strDatabaseName = "") Then
        Return_ErrNum = -1
        Return_ErrDesc = "No database name specified"
        Exit Function
      ElseIf (ConnectionDriver <> cd_SQL And ConnectionDriver <> cd_Oracle And strDataSource = "") Then
        Return_ErrNum = -1
        Return_ErrDesc = "No data source specified"
        Exit Function
      End If
    End If
  End If
  
  ' Build the connection string information
  If blnConnInfoBuilt = False Then
    ' Setup the connection driver strings - This information pulled from the MSDN.MICROSOFT.COM
    ReDim strConnDriv(1 To MAX_DRIVERS) As String
    strConnDriv(cd_Access) = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" 'physical path to .MDB file
    strConnDriv(cd_Excel) = "Driver={Microsoft Excel Driver (*.xls)};DriverID=278;DBQ=" 'physical path to .XLS file
    strConnDriv(cd_Excel_97) = "Driver={Microsoft Excel Driver (*.xls)};DriverID=790;DBQ=" 'physical path to .XLS file
    strConnDriv(cd_Paradox) = "Driver={Microsoft Paradox Driver (*.db)};DriverID=26;DBQ=" 'physical path to .DB file
    strConnDriv(cd_Text) = "Driver={Microsoft Text Driver (*.txt;*.csv)};DefaultDir=" 'physical path to .TXT file
    strConnDriv(cd_VFoxPro) = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBC;SourceDb=" 'physical path to .DBF file
    strConnDriv(cd_VFoxPro_Cont) = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDb=" 'physical path to .DBF file
    strConnDriv(cd_SQL) = "Driver={SQL Server};SERVER=" 'path to server
    strConnDriv(cd_Oracle) = "Driver={Microsoft ODBC for Oracle};SERVER=" 'path to server
    
    ' Setup the connection provider strings - This information pulled from the MSDN.MICROSOFT.COM
    ReDim strConnProv(1 To MAX_PROVIDERS) As String
    strConnProv(cp_Access) = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=4;Data Source=" 'physical path to .MDB file
    strConnProv(cp_dBaseIII) = "Provider=MSDASQL.1;Extended Properties=;DRIVER=Microsoft dBase Driver (*.dbf);DriverId=21;FIL=dBase III;MaxBufferSize=2048;PageTimeout=600;DefaultDir=" 'physical path to .DBF file
    strConnProv(cp_dBaseIV) = "Provider=MSDASQL.1;Extended Properties=;DRIVER=Microsoft dBase Driver (*.dbf);DriverId=21;FIL=dBase IV;MaxBufferSize=2048;PageTimeout=600;DefaultDir=" 'physical path to .DBF file
    strConnProv(cp_dBaseV) = "Provider=MSDASQL.1;Extended Properties=;DRIVER=Microsoft dBase Driver (*.dbf);DriverId=21;FIL=dBase V;MaxBufferSize=2048;PageTimeout=600;DefaultDir=" 'physical path to .DBF file
    strConnProv(cp_Excel) = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" 'physical path to .XLS file
    strConnProv(cp_HTML) = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=HTML Import;Data Source=" 'physical path or URL to HTM(L) file
    strConnProv(cp_MsIndexingSrv) = "Provider=MSIDXS.1;Data Source=" 'physical path to file
    strConnProv(cp_Text) = "Driver={Microsoft Text Driver (*.txt;*.csv)};DBQ=" 'physical path to .TXT file
    strConnProv(cp_Oracle) = "Provider=MSDAORA.1;Data Source=" 'path to database on server
    strConnProv(cp_SQL) = "Provider=SQLOLEDB.1;Data Source=" 'path to database on server
    
    ' Setup the connection provider strings - This information pulled from the "SQL Server Client Network Utility" that comes with SQL Server
    ReDim strConnNetProt(1 To MAX_NET_PROTOCOLS) As String
    strConnNetProt(cn_AppleTalk) = "Network Library=DBMSADSN" '.DLL
    strConnNetProt(cn_BanyanVINES) = "Network Library=DBMSVINN" '.DLL
    strConnNetProt(cn_IPX_SPX) = "Network Library=DBNETLIB" '.DLL
    strConnNetProt(cn_Multiprotocol) = "Network Library=DBMSRPCN" '.DLL
    strConnNetProt(cn_NamedPipes) = "Network Library=DBNMPNTW" '.DLL
    strConnNetProt(cn_TCP_IP) = "Network Library=DBNETLIB" '.DLL
    strConnNetProt(cn_VIA_Giganet) = "Network Library=DBMSGNET" '.DLL
    
    blnConnInfoBuilt = True
  End If
  
  ' Setup the data source for REMOTE SERVER
  If blnRemoteServer = True Then
    strConn = "Provider=MS Remote;Remote Server=" & strRemoteServerURL & ";"
    If lngRemoteServerTimeout > 0 Then strConn = strConn & "Internet Timeout=" & CStr(lngRemoteServerTimeout) & ";"
    strConn = strConn & "Data Source=" & strDataSource & ";"
    
  ' Setup the data source for DATA SHAPE
  ElseIf blnDataShape = True Then
    strConn = "Provider=MSDataShape;"
    If strDataShapeProvider <> "" Then
      strConn = strConn & "Data Provider=" & strDataShapeProvider & ";"
    Else
      strConn = strConn & strConnDriv(ConnectionDriver)
      If ConnectionDriver = cd_SQL Or ConnectionDriver = cd_Oracle Then
        strConn = strConn & strServerName & ";Database=" & strDatabaseName
      Else
        strConn = strConn & strDataSource & ";"
      End If
    End If
    
  Else
    ' Setup the data source based on the specified PROVIDER
    If ConnectionProvider <> cp_NotSpecified Then
      Select Case ConnectionProvider
        Case cp_Access, cp_dBaseIII, cp_dBaseIV, cp_dBaseV, cp_Excel, cp_HTML, cp_MsIndexingSrv, cp_Text
          strConn = strConnProv(ConnectionProvider) & strDataSource & ";"
        Case cp_Oracle, cp_SQL
          strConn = strConnProv(ConnectionProvider) & strServerName & ";Database=" & strDatabaseName & ";"
        Case Else
          Return_ErrNum = -1: Return_ErrDesc = "Unknown connection provider specified"
          Exit Function
      End Select
      
    ' Setup the data source based on the specified DRIVER
    ElseIf ConnectionDriver <> cd_NotSpecified Then
      Select Case ConnectionDriver
        Case cd_Access, cd_Excel, cd_Excel_97, cd_Paradox, cd_Text, cd_VFoxPro, cd_VFoxPro_Cont
          strConn = strConnDriv(ConnectionDriver) & strDataSource & ";"
        Case cd_SQL, cd_Oracle
          strConn = strConnDriv(ConnectionDriver) & strServerName & ";Database=" & strDatabaseName & ";"
        Case Else
          Return_ErrNum = -1: Return_ErrDesc = "Unknown connection driver specified"
          Exit Function
      End Select
    End If
  End If
  
  ' Setup the network protocol type if specified
  If ConnectionNetProtocol <> cn_NotSpecified Then
    Select Case ConnectionNetProtocol
      Case cn_AppleTalk, cn_BanyanVINES, cn_IPX_SPX, cn_Multiprotocol, cn_NamedPipes, cn_TCP_IP, cn_VIA_Giganet
        strConn = strConn & strConnNetProt(ConnectionNetProtocol) & ";"
      Case Else
        Return_ErrNum = -1: Return_ErrDesc = "Unknown connection driver specified"
        Exit Function
    End Select
  End If
  
  ' Specify the user name and password
  strConn = strConn & "UID=" & strUserName & ";PWD=" & strPassword & ";"
  
  ' Specify the packet size if the user wants to
  If lngPacketSize > 0 Then
    strConn = strConn & "Packet Size=" & CStr(lngPacketSize) & ";"
  End If
  
  ' Specify whether to persist security information or not
  If blnPersistSecurityInfo = "TRUE" Or blnPersistSecurityInfo = "FALSE" Then
    strConn = strConn & "Persist Security Info=" & blnPersistSecurityInfo & ";"
  End If
  
  ' Specify whether to use encryption for data or not
  If blnUseEncryptionForData = "TRUE" Or blnUseEncryptionForData = "FALSE" Then
    strConn = strConn & "Use Encryption for Data=" & blnUseEncryptionForData & ";"
  End If
  
  ' Specify whether to auto translate or not
  If blnAutoTranslate = "TRUE" Or blnAutoTranslate = "FALSE" Then
    strConn = strConn & "Auto Translate=" & blnAutoTranslate & ";"
  End If
  
  ' Specify the workstation ID (name)
  If strWorkstationID <> "" Then
    strConn = strConn & "Workstation ID=" & strWorkstationID & ";"
  End If
  
  ' Return the results
  Return_ConnectionString = strConn
  GenerateConnectionString = True
  
End Function


Public Function OpenConnection(ByVal strConnectionString As String, _
                                    Optional ByRef Return_ErrNum As Long, _
                                    Optional ByRef Return_ErrDesc As String) As Boolean
On Error GoTo ErrorTrap
  
  ' Clear variables
  Err.Clear
  Return_ErrNum = 0
  Return_ErrDesc = ""
  strConnString = ""
  
  ' Clean up any previous connection before recreating it
  CleanupConnection conConnection
  
  ' Validate parameters
  strConnectionString = Trim(strConnectionString)
  If strConnectionString = "" Then
    OpenConnection = True
    Exit Function
  End If
  
  ' Setup the new connection
  Set conConnection = New ADODB.Connection
  conConnection.CommandTimeout = lngCmdTimeout
  conConnection.ConnectionTimeout = lngConnTimeout
  conConnection.IsolationLevel = ConnIsolationLvl
  conConnection.Mode = ConnMode
  conConnection.Open strConnectionString
  
  ' Exit the function
  strConnString = strConnectionString
  OpenConnection = True
  Exit Function
  
ErrorTrap:
  
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  CleanupConnection conConnection
  
End Function


Public Function ValidateStringForSQL(ByVal strSQL As String) As String
  
  ValidateStringForSQL = ReplaceChar(strSQL, Chr(39), Chr(39) & Chr(39))
  ValidateStringForSQL = ReplaceChar(ValidateStringForSQL, Chr(34), Chr(39) & Chr(39))
  
End Function


Private Function ArrayIsEmpty_STR(ByRef TheArray() As String, _
                                  Optional ByRef Return_LBound As Long, _
                                  Optional ByRef Return_UBound As Long) As Boolean
On Error GoTo ErrorTrap
  
  ' Set the default return values
  Return_LBound = -1
  Return_UBound = -1
  ArrayIsEmpty_STR = True
  
  ' By checking the L and U bound of the array, you'll see if the array is empty
  Return_LBound = LBound(TheArray)
  Return_UBound = UBound(TheArray)
  
  ' If it made it this far, the array isn't empty
  ArrayIsEmpty_STR = False
  Exit Function
  
ErrorTrap:
  Err.Clear
  
End Function

' Takes a BOOLEAN array and checks to see if it has any elements in it by returning the Upper and Lower bounds of the array.
Private Function ArrayIsEmpty_BLN(ByRef TheArray() As Boolean, _
                                  Optional ByRef Return_LBound As Long, _
                                  Optional ByRef Return_UBound As Long) As Boolean
On Error GoTo ErrorTrap
  
  ' Set the default return values
  Return_LBound = -1
  Return_UBound = -1
  ArrayIsEmpty_BLN = True
  
  ' By checking the L and U bound of the array, you'll see if the array is empty
  Return_LBound = LBound(TheArray)
  Return_UBound = UBound(TheArray)
  
  ' If it made it this far, the array isn't empty
  ArrayIsEmpty_BLN = False
  Exit Function
  
ErrorTrap:
  Err.Clear
  
End Function

' Takes a VARIANT array and checks to see if it has any elements in it by returning the Upper and Lower bounds of the array.
Private Function ArrayIsEmpty_VAR(ByRef TheArray() As Variant, _
                                  Optional ByRef Return_LBound As Long, _
                                  Optional ByRef Return_UBound As Long) As Boolean
On Error GoTo ErrorTrap
  
  ' Set the default return values
  Return_LBound = -1
  Return_UBound = -1
  ArrayIsEmpty_VAR = True
  
  ' By checking the L and U bound of the array, you'll see if the array is empty
  Return_LBound = LBound(TheArray)
  Return_UBound = UBound(TheArray)
  
  ' If it made it this far, the array isn't empty
  ArrayIsEmpty_VAR = False
  Exit Function
  
ErrorTrap:
  Err.Clear
  
End Function

' This function takes an ADODB.Command object and returns the values of the parameters within it to parameter's that have the same name.
Private Function AssignReturnValue(ByRef ParameterNames() As String, _
                                   ByRef ParameterValues() As Variant, _
                                   ByRef objCommand As ADODB.Command, _
                                   ByVal TheLBound As Long, _
                                   ByVal TheUBound As Long) As Boolean
On Error Resume Next
  
  Dim objParam   As ADODB.Parameter
  Dim lngCounter As Long
  Dim strParam   As String
  
  If objCommand.Parameters.Count = 0 Then
    AssignReturnValue = True
    Exit Function
  End If
  
  ' Loop through the parameters in the command object's Parameters collection
  For Each objParam In objCommand.Parameters
    
    ' Make sure all the parameters are in the same format
    strParam = UCase(Trim(objParam.Name))
    If strParam <> "" Then
      If Left(strParam, 1) <> "@" Then strParam = "@" & strParam
    End If
    
    ' Look for a parameter that was passed in that matches the name and re-assign the value to it if found
    For lngCounter = TheLBound To TheUBound
      If ParameterNames(lngCounter) <> "" Then
        
        ' Check if the current parameter name is the same... and if it is, get the value back
        If UCase(Trim(ParameterNames(lngCounter))) = strParam Then
          If Trim(objParam.value) & "" = "" Then
            ParameterValues(lngCounter) = Empty
          Else
            ParameterValues(lngCounter) = objParam.value
          End If
          GoTo Continue
        End If
      End If
    Next
Continue:
  Next
  
  AssignReturnValue = True
  
End Function

' Checks to see if a valid connection exists, and if one doesn't, establishes one
Private Function CheckConnection(Optional ByRef Return_ErrNum As Long, _
                                 Optional ByRef Return_ErrDesc As String) As Boolean
On Error GoTo ErrorTrap
  
  ' Clear variables
  Err.Clear
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' Check if the connection string is valid
  If strConnString = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "No connection string specified"
    Exit Function
    
  ' Check if the connection exists, and if it doesn't, create it
  ElseIf conConnection Is Nothing Then
    Set conConnection = New ADODB.Connection
    
  ' Check if the connection is still valid, and re-create it if it's not
  ElseIf conConnection.State <> adStateOpen Then
    CleanupConnection conConnection
    Set conConnection = New ADODB.Connection
    
  ' The connection seems to be fine at this point.  It is declared, initialized, and open.
  Else
    CheckConnection = True
    Exit Function
  End If
  
  ' Setup the connection, then connect to the database
  conConnection.CommandTimeout = lngCmdTimeout
  conConnection.ConnectionTimeout = lngConnTimeout
  conConnection.IsolationLevel = ConnIsolationLvl
  conConnection.Mode = ConnMode
  conConnection.Open strConnString
  
  CheckConnection = True
  
  Exit Function
  
ErrorTrap:
  
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  
End Function

' This function gives VB5 the functionality that VB6 gives via the "Replace" function
Private Function ReplaceChar(ByVal strSearchString As String, _
                             ByVal strSearchChar As String, _
                             ByVal strReplaceChar As String) As String
On Error Resume Next
  
  Dim lngCounter As Long
  Dim CharLeft   As String
  Dim CharRight  As String
  Dim strLeft    As String
  Dim strRight   As String
  
  ReplaceChar = strSearchString
  If strSearchString = "" Then Exit Function
  If strSearchChar = "" Then Exit Function
  If strSearchChar = strReplaceChar Then Exit Function
  If InStr(strSearchString, strSearchChar) <= 0 Then Exit Function
  
  lngCounter = 1
TheBegin:
  For lngCounter = lngCounter To Len(strSearchString)
    CharLeft = Left(strSearchString, lngCounter)
    CharRight = Right(CharLeft, Len(strSearchChar))
    If CharRight = strSearchChar Then
      strLeft = Left(CharLeft, Len(CharLeft) - Len(strSearchChar))
      strRight = Right(strSearchString, (Len(strSearchString) - Len(strLeft)) - Len(strSearchChar))
      strSearchString = strLeft & strReplaceChar & strRight
      lngCounter = Len(strLeft) + Len(strReplaceChar) + 1
      GoTo TheBegin
    End If
  Next
  
  ReplaceChar = strSearchString
  
End Function



