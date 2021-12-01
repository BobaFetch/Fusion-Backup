VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EsiEmailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private emailerServerName As String
Private emailerDatabaseName As String
Private emailerLogin As String
Private emailerPassword As String
Private mailingList As String
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
Private adoConn As ClassFusionADO

Private mailingListID As Long

Private parameters() As Parameter
Private parameterCount As Integer

Dim ErrNum    As Long
Dim ErrDesc   As String

Private Sub Class_Initialize()

   ' get emailer server and database from registry
   parameterCount = 0
   
Top:
   emailerServerName = GetSetting("Esi2000", "EsiEmailer", "ServerName", "")
   emailerDatabaseName = GetSetting("Esi2000", "EsiEmailer", "DatabaseName", "")
   emailerLogin = GetSetting("Esi2000", "EsiEmailer", "Login", "")
   emailerPassword = GetSetting("Esi2000", "EsiEmailer", "Password", "")
   If Not TestEmailerConnection Then
      EsiEmailerSettings.Show vbModal
      Dim newServer As String
      Dim newDatabase As String
      Dim newLogin As String
      Dim NewPassWord As String
      newServer = GetSetting("Esi2000", "EsiEmailer", "ServerName", "")
      newDatabase = GetSetting("Esi2000", "EsiEmailer", "DatabaseName", "")
      newLogin = GetSetting("Esi2000", "EsiEmailer", "Login", "")
      NewPassWord = GetSetting("Esi2000", "EsiEmailer", "Password", "")
      If newServer <> emailerServerName Or newDatabase <> emailerDatabaseName _
         Or newLogin <> emailerLogin Or NewPassWord <> emailerPassword Then GoTo Top
   End If
   
End Sub

Public Property Get ConnectionString() As String
   ConnectionString = "Driver={SQL Server};Provider='sqloledb';UID=" & emailerLogin & ";PWD=" & _
            emailerPassword & ";SERVER=" & emailerServerName & ";DATABASE=" & emailerDatabaseName & ";"
End Property

Public Property Get ConnectionStringDisplay() As String
   ConnectionStringDisplay = "Driver={SQL Server};Provider='sqloledb';UID=" & emailerLogin & ";PWD=..." & _
            ";SERVER=" & emailerServerName & ";DATABASE=" & emailerDatabaseName & ";"
End Property

Public Property Get MailingListName() As String
   MailingListName = mailingList
End Property

Public Property Let MailingListName(NewValue As String)
   mailingList = NewValue
End Property


Private Function TestEmailerConnection() As Boolean
   TestEmailerConnection = False
   If emailerServerName = "" Or emailerDatabaseName = "" Then
      Exit Function
   End If

   'Dim adoConn As ClassFusionADO
   Set adoConn = New ClassFusionADO
   On Error Resume Next
   Err.Clear
'   rdo.Connect = "Data Source=" & emailerServerName & ";" _
'      & "Initial Catalog=" + emailerDatabaseName & ";" _
'      & "Integrated Security=SSPI"
'   rdo.Connect = "SERVER=" & emailerServerName & ";" _
'      & "DATABASE=" + emailerDatabaseName & ";" _
'      & "DRIVER={SQL Server};" _
'      & "Integrated Security=SSPI"
'   adoConn.Connect = "SERVER=" & emailerServerName & ";" _
'      & "DATABASE=" & emailerDatabaseName & ";" _
'      & "DRIVER={SQL Server};" _
'      & "Integrated Security=SSPI"
'      & "Trusted_Connection=yes"
   'Dim conString As String
'   conString = "UID=" & emailerLogin & ";PWD=" & emailerPassword & ";DRIVER={SQL Server};" _
'      & "SERVER=" & emailerServerName & ";DATABASE=" & emailerDatabaseName & ";"
'DRIVER=SQL Server;SERVER=localhost\SQL2005;UID=TerryL;PWD=;APP=Visual Basic;WSID=TERRYSHP;DATABASE=Emailer;Trusted_Connection=Yes
'   conString = "DRIVER=SQL Server;SERVER=" & emailerServerName & ";DATABASE=" & emailerDatabaseName & ";Trusted_Connection=Yes"
'   conString = InputBox("Edit connection string:", "Connection String", conString)
   
  ' conString = "DRIVER={SQL Server};SERVER=" & emailerServerName & ";DATABASE=" & emailerDatabaseName & ";Trusted_Connection=Yes"
   
'   conString = "Driver={SQL Server};Provider='sqloledb';UID=" & emailerLogin & ";PWD=" & _
'            emailerPassword & ";SERVER=" & emailerServerName & ";DATABASE=" & emailerDatabaseName & ";"
   
'   adoConn.OpenConnection (conString)
'   'adoConn.Connect = conString
'   'adoConn.EstablishConnection rdDriverNoPrompt
'   If Err Then
'      MsgBox "Unable to connect to emailer server " & emailerServerName & " database " & emailerDatabaseName & vbCrLf _
'         & "login " & emailerLogin & vbCrLf & Err.Description & vbCrLf & conString
'      Exit Function
'   End If
   
   If adoConn.OpenConnection(ConnectionString, ErrNum, ErrDesc) = False Then
      MsgBox "Unable to connect to emailer server " & emailerServerName & " database " & emailerDatabaseName & vbCrLf _
         & "login " & emailerLogin & vbCrLf & ErrDesc & vbCrLf & ConnectionStringDisplay()
      Set adoConn = Nothing
      Exit Function
   End If
   
   Set adoConn = Nothing
   
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
   'Dim adoConn As rdoConnection
   
   'Dim adoConn As ClassFusionADO
   Set adoConn = New ClassFusionADO
   
   'Set adoConn = New ClassFusionADO
   Dim rs As ADODB.Recordset
   Dim query As String
   IsRequestValid = VAL_NotValid
   
'   Dim ErrNum    As Long
'   Dim ErrDesc   As String
   
   On Error Resume Next
   Err.Clear
   'Dim conString As String
'   conString = "DRIVER={SQL Server};SERVER=" & emailerServerName & ";DATABASE=" & emailerDatabaseName & ";Trusted_Connection=Yes"
   'conString = "Driver={SQL Server};Provider='sqloledb';UID=" & emailerLogin & ";PWD=" & _
            emailerPassword & ";SERVER=" & emailerServerName & ";DATABASE=" & emailerDatabaseName & ";"

'   adoConn.OpenConnection (conString)
'   If Err Then
'      MsgBox "Emailer unable to connect: " & vbCrLf _
'         & Err.Description
'      Exit Function
'   End If
   
   If adoConn.OpenConnection(ConnectionString, ErrNum, ErrDesc) = False Then
      MsgBox "Emailer unable to connect: " & vbCrLf _
         & ErrDesc
      Exit Function
   End If

'   adoConn.EstablishConnection rdDriverNoPrompt
   If Err Then
      MsgBox "Generate.EstablishConnection unable to connect: " & vbCrLf _
         & Err.Description
      Exit Function
   End If
   
   'look up mailing list id and paramter ids
   query = "select MailingListID from MailingLists where MailingListName = '" & MailingListName & "'"
   Set rs = adoConn.GetRecordSet(query)
   If rs.BOF Or rs.EOF Then
      MsgBox "No such mailing list defined: " & MailingListName
      Exit Function
   End If
   mailingListID = CLng(rs.Fields(0))
   Set rs = Nothing
   If Err Then
      MsgBox "Error from query: " & query
      Exit Function
   End If
   
   Dim i As Integer
   Dim distListKey As String
   For i = 1 To parameterCount
      query = "select MailingListParameterID, UseToDetermineRecipients from MailingListParameters" & vbCrLf _
         & "where MailingListID = " & mailingListID & vbCrLf _
         & "and ParameterName = '" & parameters(i).Name & "'"
      Set rs = adoConn.GetRecordSet(query)
      If Err Then
         MsgBox "Error from query: " & query
         Exit Function
      End If
      If rs.BOF Or rs.EOF Then
         MsgBox "No such mailing list parameter defined: " & MailingListName & "." & parameters(i).Name
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
      & "where MailingListID = " & mailingListID & vbCrLf _
      & "and DistributionListKey = '" & distListKey & "'"
   Set rs = adoConn.GetRecordSet(query)
   If CLng(rs(0)) = 0 Then
      Set rs = Nothing
      Exit Function
   End If
   Set rs = Nothing
   
   IsRequestValid = VAL_ValidWithDistList
   Exit Function
   
whoops:
   MsgBox "EsiEmailer.IsRequestValid error: " & Err.Description & vbCrLf _
      & "Last query: " & query

End Function

Public Sub Generate()
   On Error GoTo whoops
   'Set adoConn = New rdoConnection
   Set adoConn = New ClassFusionADO

   Dim rs As ADODB.Recordset
   Dim query As String
   
   'validate request and initialize
   If IsRequestValid = VAL_NotValid Then
      Exit Sub
   End If
   
'   On Error Resume Next
'   Err.Clear
'   Dim conString As String
'   conString = "DRIVER={SQL Server};SERVER=" & emailerServerName & ";DATABASE=" & emailerDatabaseName & ";Trusted_Connection=Yes"
'   adoConn.Connect = conString
'   If Err Then
'      MsgBox "Emailer unable to connect: " & adoConn.Connect & vbCrLf _
'         & Err.Description
'      Exit Sub
'   End If
'
'   adoConn.EstablishConnection rdDriverNoPrompt
'   If Err Then
'      MsgBox "Generate.EstablishConnection unable to connect: " & adoConn.Connect & vbCrLf _
'         & Err.Description
'      Exit Sub
'   End If
'
'   'look up mailing list id and paramter ids
'   query = "select MailingListID from MailingLists where MailingListName = '" & MailingListName & "'"
'   Set rs = adoConn.OpenResultset(query)
'   If rs.BOF Or rs.EOF Then
'      MsgBox "No such mailing list defined: " & MailingListName
'      Exit Sub
'   End If
'   Dim mailingListID As Long
'   mailingListID = CLng(rs.Fields(0))
'   rs.Close
'   If Err Then
'      MsgBox "Error from query: " & query
'      Exit Sub
'   End If
'
'   Dim i As Integer
'   For i = 1 To parameterCount
'      query = "select MailingListParameterID from MailingListParameters" & vbCrLf _
'         & "where MailingListID = " & mailingListID & vbCrLf _
'         & "and ParameterName = '" & parameters(i).Name & "'"
'      Set rs = adoConn.OpenResultset(query)
'      If Err Then
'         MsgBox "Error from query: " & query
'         Exit Sub
'      End If
'      If rs.BOF Or rs.EOF Then
'         MsgBox "No such mailing list parameter defined: " & MailingListName & "." & parameters(i).Name
'         Exit Sub
'      End If
'      parameters(i).ParameterID = CLng(rs.Fields(0))
'      rs.Close
'   Next

   If adoConn.OpenConnection(ConnectionString, ErrNum, ErrDesc) = False Then
      MsgBox "Unable to connect to emailer server " & emailerServerName & " database " & emailerDatabaseName & vbCrLf _
         & "login " & emailerLogin & vbCrLf & ErrDesc & vbCrLf & ConnectionStringDisplay()
      Return
   End If

   adoConn.BeginTrans
   query = "SET NOCOUNT ON" & vbCrLf _
      & "INSERT Requests ( MailingListID, RequestedBy )" & vbCrLf _
      & "VALUES( " & mailingListID & ", '" & sInitials & "')" & vbCrLf _
      & "SELECT @@IDENTITY" & vbCrLf _
      & "SET NOCOUNT OFF"
   'adoConn.Execute (query)
   Set rs = adoConn.GetRecordSet(query)
   If rs.BOF Or rs.EOF Then
      MsgBox "Unable to insert mailing list request"
      Set rs = Nothing
      Exit Sub
   End If
   Dim requestId As Long
   requestId = rs.Fields(0).Value
   Set rs = Nothing
   
   'now insert the request parameters
   Dim i As Integer
   For i = 1 To parameterCount
      query = "INSERT RequestParameters" & vbCrLf _
         & "( MailingListParameterID, RequestID, ParameterValue )" & vbCrLf _
         & "VALUES( " & parameters(i).ParameterID & ", " & requestId & ", '" & parameters(i).Value & "' )"
      adoConn.ExecuteSQL query
      If adoConn.RowsAffected = 0 Then
         MsgBox "Failed to insert request parameter: " & parameters(i).Name
         Exit Sub
      End If
   Next

   adoConn.CommitTrans
   Set adoConn = Nothing
   Exit Sub
whoops:
   adoConn.RollbackTrans
   MsgBox "EsiEmailer.Generate error: " & Err.Description & vbCrLf _
      & "Last query: " & query
End Sub
