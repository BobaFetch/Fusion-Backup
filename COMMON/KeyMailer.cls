VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyMailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private emailerDatabaseKey As String
Private emailerServerName As String
Private emailerDatabaseName As String
Private emailerLogin As String
Private emailerPassword As String
Private report As String
Private distListKey As String
Private Type Parameter
   Name As String
   Value As String
   ParameterID As Long
   IsDistListKey As Boolean
End Type

Public Enum eVALIDATION
   VAL_NotValid = 0
   VAL_ValidButNoDistList = 1
   VAL_ValidWithDistList
End Enum

'Private adoConn As ClassFusionADO
Private ReportConn As ClassFusionADO
Private KeyMailerDbConn As ClassFusionADO

Private prop_ReportID As Long

Private parameters() As Parameter
Private parameterCount As Integer

Dim ErrNum    As Long
Dim ErrDesc   As String

Dim rs As ADODB.Recordset
Dim query As String


'Private Sub Class_Initialize()
'
'   ' get emailer server and database from registry
'   parameterCount = 0
'
'Top:
'   emailerServerName = GetSetting("Esi2000", "KeyMailer", "ServerName", "")
'   emailerDatabaseName = GetSetting("Esi2000", "KeyMailer", "DatabaseName", "")
'   emailerLogin = GetSetting("Esi2000", "KeyMailer", "Login", "")
'   emailerPassword = GetSetting("Esi2000", "KeyMailer", "Password", "")
'   If Not TestEmailerConnection Then
'      EsiEmailerSettings.Show vbModal
'      Dim newServer As String
'      Dim newDatabase As String
'      Dim newLogin As String
'      Dim NewPassWord As String
'      newServer = GetSetting("Esi2000", "KeyMailer", "ServerName", "")
'      newDatabase = GetSetting("Esi2000", "KeyMailer", "DatabaseName", "")
'      newLogin = GetSetting("Esi2000", "KeyMailer", "Login", "")
'      NewPassWord = GetSetting("Esi2000", "KeyMailer", "Password", "")
'      If newServer <> emailerServerName Or newDatabase <> emailerDatabaseName _
'         Or newLogin <> emailerLogin Or NewPassWord <> emailerPassword Then GoTo Top
'   End If
'
'End Sub

Private Sub Class_Initialize()

   ' get KeyMailer database credentials from The KeyMailer table in the Fusion database
   parameterCount = 0
   
   'Set adoConn = New ClassFusionADO
   Dim rdo As ADODB.Recordset
   Dim query As String
   query = " select DatabaseKey, ServerName, DatabaseName, Login," & vbCrLf _
      & "cast(DECRYPTBYPASSPHRASE('KM Encryption 123!@#', Password) as varchar(50)) as Password" & vbCrLf _
      & "from KeyMailerInfo"
   bSqlRows = clsADOCon.GetDataSet(query, rdo)
   If bSqlRows Then
      With rdo
         emailerDatabaseKey = !DatabaseKey
         emailerServerName = !ServerName
         emailerDatabaseName = !DatabaseName
         emailerLogin = !Login
         emailerPassword = !PassWord
         .Cancel
      End With
      
'   If Not TestEmailerConnection Then
'      MsgBox "KeyMailer not initialized for this database"
'   End If

      Set KeyMailerDbConn = New ClassFusionADO
      If KeyMailerDbConn.OpenConnection(KeyMailerConnectionString, ErrNum, ErrDesc) = False Then
         If emailerServerName = "" Then
            'MsgBox "KeyMailer is either not installed or not initialized for this database."
            KeyMailerDbConn = Null
         Else
            KeyMailerDbConn = Null
            'MsgBox "Unable to connect to KeyMailer server " & emailerServerName & " database " & emailerDatabaseName & vbCrLf _
            '   & "login " & emailerLogin & vbCrLf & ErrDesc & vbCrLf & KeyMailerConnectionStringDisplay()
         End If
         
         Return
      End If
      
   End If


End Sub

Public Property Get KeyMailerConnectionString() As String
   KeyMailerConnectionString = "Driver={SQL Server};Provider='sqloledb';UID=" & emailerLogin & ";PWD=" & _
            emailerPassword & ";SERVER=" & emailerServerName & ";DATABASE=" & emailerDatabaseName & ";"
End Property

Public Property Get KeyMailerConnectionStringDisplay() As String
   KeyMailerConnectionStringDisplay = "Driver={SQL Server};Provider='sqloledb';UID=" & emailerLogin & ";PWD=..." & _
            ";SERVER=" & emailerServerName & ";DATABASE=" & emailerDatabaseName & ";"
End Property

Public Property Get ReportName() As String
   ReportName = report
End Property

Public Property Let ReportName(NewValue As String)
   report = NewValue
End Property

Public Property Get DistributionListKey() As String
   DistributionListKey = distListKey
End Property

Public Property Get ReportID() As Integer
   ReportID = prop_ReportID
End Property

Public Property Let ReportID(NewValue As Integer)
   prop_ReportID = NewValue
End Property

Public Property Let DistributionListKey(NewValue As String)
   distListKey = NewValue
End Property

Private Function TestEmailerConnection(ShowErrorMessage As Boolean) As Boolean
   
   ' ShowErrorMessage = True if error message to be displayed
   TestEmailerConnection = False
   If emailerServerName = "" Or emailerDatabaseName = "" Then
      Exit Function
   End If

   'Set adoConn = New ClassFusionADO
   On Error Resume Next
   Err.Clear
   
   If KeyMailerDbConn.OpenConnection(KeyMailerConnectionString, ErrNum, ErrDesc) = False Then
      If ShowErrorMessage Then
         MsgBox "Unable to connect to KeyMailer database." & vbCrLf _
            & ErrDesc & vbCrLf & KeyMailerConnectionStringDisplay()
      End If
      'Set adoConn = Nothing
      Exit Function
   End If
   
   'Set adoConn = Nothing
   
   TestEmailerConnection = True
End Function

Public Sub AddLongParameter(ParameterName As String, ParameterValue As Long)
   parameterCount = parameterCount + 1
   ReDim Preserve parameters(parameterCount)
   parameters(parameterCount).Name = ParameterName
   parameters(parameterCount).Value = CStr(ParameterValue)
End Sub


Public Sub AddStringParameter(ParameterName As String, ParameterValue As String)
   parameterCount = parameterCount + 1
   ReDim Preserve parameters(parameterCount)
   parameters(parameterCount).Name = ParameterName
   parameters(parameterCount).Value = ParameterValue
End Sub

Public Sub AddBooleanParameter(ParameterName As String, ParameterValue As Boolean)
   parameterCount = parameterCount + 1
   ReDim Preserve parameters(parameterCount)
   parameters(parameterCount).Name = ParameterName
   If ParameterValue Then
      parameters(parameterCount).Value = "True"
   Else
      parameters(parameterCount).Value = "False"
   End If
End Sub

Public Function IsRequestValid() As eVALIDATION
   On Error GoTo whoops
   IsRequestValid = VAL_NotValid
   
   On Error Resume Next
   Err.Clear
   
   If KeyMailerDbConn.OpenConnection(KeyMailerConnectionString, ErrNum, ErrDesc) = False Or Err Then
      MsgBox "KeyMailer unable to connect: " & vbCrLf _
         & ErrDesc & vbCrLf & KeyMailerConnectionStringDisplay()
      Exit Function
   End If

'   If Err Then
'      MsgBox "KeyMailer unable to connect: " & vbCrLf _
'         & Err.Description & vbCrLf & KeyMailerConnectionStringDisplay()
'      Exit Function
'   End If
   
   'look up report id and paramter ids
'   query = "select ReportID from Reports where ReportName = '" & ReportName & "' and DatabaseKey = '" & emailerDatabaseKey & "'"
'   Set rs = adoConn.GetRecordSet(query)
'   If rs.BOF Or rs.EOF Then
'      MsgBox "No such report defined: " & ReportName
'      Exit Function
'   End If
'   reportID = CLng(rs.Fields(0))
'   Set rs = Nothing
'   If Err Then
'      MsgBox "Error from query: " & query
'      Exit Function
'   End If
'
   Dim i As Integer
   Dim distListKey As String
   For i = 1 To parameterCount
      query = "select ReportParameterID, UseToDetermineRecipients from ReportParameters" & vbCrLf _
         & "where ReportID = " & ReportID & vbCrLf _
         & "and ParameterName = '" & parameters(i).Name & "'"
      Set rs = KeyMailerDbConn.GetRecordSet(query)
      If Err Then
         MsgBox "Error from query: " & query
         Exit Function
      End If
      If rs.BOF Or rs.EOF Then
         MsgBox "No such report parameter defined: " & ReportName & "." & parameters(i).Name
         Exit Function
      End If
      parameters(i).ParameterID = CLng(rs.Fields(0))
      parameters(i).IsDistListKey = CBool(rs.Fields(1))
      If parameters(i).IsDistListKey Then
         distListKey = parameters(i).Value
      End If
      Set rs = Nothing
   Next

   'request is valid.  Determine whether there are recipients
   IsRequestValid = VAL_ValidButNoDistList
   If distListKey = "" Then
      Exit Function
   End If
   
   query = "select count(*) from DistributionLists dl" & vbCrLf _
      & "join Recipients rec on rec.DistributionListID = dl.DistributionListID" & vbCrLf _
      & "where ReportID = " & ReportID & vbCrLf _
      & "and DistributionListKey = '" & distListKey & "'"
   Set rs = KeyMailerDbConn.GetRecordSet(query)
   If CLng(rs(0)) = 0 Then
      Set rs = Nothing
      Exit Function
   End If
   Set rs = Nothing
   
   IsRequestValid = VAL_ValidWithDistList
   Exit Function
   
whoops:
   MsgBox "KeyMailer.IsRequestValid error: " & Err.Description & vbCrLf _
      & "Last query: " & query

End Function

Public Function Generate() As Boolean

   Generate = False
   On Error GoTo whoops
   
   'Set adoConn = New rdoConnection
   ''Set adoConn = New ClassFusionADO

   Dim rs As ADODB.Recordset
   Dim query As String
   
   'validate request and initialize
   If IsRequestValid = VAL_NotValid Then
      Exit Function
   End If
   
   If KeyMailerDbConn.OpenConnection(KeyMailerConnectionString, ErrNum, ErrDesc) = False Then
      If emailerServerName = "" Then
         MsgBox "KeyMailer is either not installed or not initialized for this database."
      Else
         MsgBox "Unable to connect to KeyMailer server " & emailerServerName & " database " & emailerDatabaseName & vbCrLf _
            & "login " & emailerLogin & vbCrLf & ErrDesc & vbCrLf & KeyMailerConnectionStringDisplay()
      End If
      
      Return
      
   End If

   KeyMailerDbConn.BeginTrans
   query = "SET NOCOUNT ON" & vbCrLf _
      & "INSERT Requests ( ReportID, RequestedBy, DistributionListKey )" & vbCrLf _
      & "VALUES( " & ReportID & ", '" & sInitials & "', '" & DistributionListKey & "')" & vbCrLf _
      & "SELECT @@IDENTITY" & vbCrLf _
      & "SET NOCOUNT OFF"
   Set rs = KeyMailerDbConn.GetRecordSet(query)
   If rs.BOF Or rs.EOF Then
      MsgBox "Unable to insert report request"
      Set rs = Nothing
      KeyMailerDbConn.RollbackTrans
      Exit Function
   End If
   Dim requestId As Long
   requestId = rs.Fields(0).Value
   Set rs = Nothing
   
   'now insert the request parameters
   Dim i As Integer
   For i = 1 To parameterCount
      query = "INSERT RequestParameters" & vbCrLf _
         & "( ReportParameterID, RequestID, ParameterValue )" & vbCrLf _
         & "VALUES( " & parameters(i).ParameterID & ", " & requestId & ", '" & parameters(i).Value & "' )"
      KeyMailerDbConn.ExecuteSQL query
      If KeyMailerDbConn.RowsAffected = 0 Then
         MsgBox "Failed to insert request parameter: " & parameters(i).Name
         KeyMailerDbConn.RollbackTrans
         Exit Function
      End If
   Next

   KeyMailerDbConn.CommitTrans
   Generate = True
   
   Exit Function
whoops:
   KeyMailerDbConn.RollbackTrans
   MsgBox "KeyMailer.Generate error: " & Err.Description & vbCrLf _
      & "Last query: " & query
End Function

Public Function GetReportInfo(ShowErrorMessage As Boolean) As Boolean
   ' returns false if database or report unavailable
   ' ShowErrorMessage = True if error message to be displayed
   GetReportInfo = False
   If KeyMailerDbConn Is Nothing Then
      If ShowErrorMessage Then MsgBox "KeyMailer not installed"
      Exit Function
   End If
   
   'look up report id and paramter ids
   query = "select ReportID from Reports where ReportName = '" & ReportName & "' and DatabaseKey = '" & emailerDatabaseKey & "'"
   Set rs = KeyMailerDbConn.GetRecordSet(query)
   If rs.BOF Or rs.EOF Or Err Then
      If ShowErrorMessage Then MsgBox "Report not installed: " & ReportName & "." & emailerDatabaseKey
      Exit Function
   End If
   ReportID = CLng(rs.Fields(0))
'   Set rs = Nothing
'   If Err Then
'      If ShowErrorMessage Then MsgBox "Error " & ErrorMessage & vbCrLf & "from query: " & query
'      Exit Function
'   End If
   
   GetReportInfo = True
End Function

'Public Function GetDistListEmails() As String()
'   ' returns a list of email addreses from the KeyMailer Database
'   On Error GoTo whoops
'   Dim nullAdr(0) As String
'   GetDistListEmails = nullAdr
'   If KeyMailerDbConn Is Nothing Then Exit Function
'
'   query = "select rec.EmailAddress, rec.FullName  from Recipients rec" & vbCrLf _
'      & "join DistributionLists dl on dl.DistributionListID = rec.DistributionListID" & vbCrLf _
'      & "join Reports rpt on rpt.ReportID = dl.ReportID" & vbCrLf _
'      & "where dl.DistributionListKey = '" & DistributionListKey & "'" & vbCrLf _
'      & "and rpt.ReportID = " & CStr(ReportID) & vbCrLf
'
'   Set rs = KeyMailerDbConn.GetRecordSet(query)
'   If rs.BOF Or rs.EOF Or Err Then
'      Exit Function
'   End If
'
'   Dim count As Integer
'   count = rs.RecordCount
'   Dim adrs() As String
'   ReDim adrs(count)
'   Do While Not rs.EOF
'      adrs(rs.Index) = rs!EmailAddress
'      .MoveNext
'   Loop
'
'   rs.Close
'
'   GetDistListEmails = adrs
'whoops:
'End Function
'
