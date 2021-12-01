Attribute VB_Name = "Customer"
'Customer Permissions 6/27/03
'12/6/05 Redirected ES_CUSTOM
Option Explicit
Public bCustomerGroups(9) As Byte
Public ES_CUSTOM As String 'See Customer.CustomerPermissions
Public sFacility As String

'6/29/03

Public Sub CodingPrincipals()
End Sub

Public Sub SysChannelsDoc()
   
End Sub


'Gets the msdb and only works on SQL 7.0 and later

Public Sub GetCustomerPermissions()
   Dim RdoPrm As ADODB.Recordset
   Dim bFalse As Byte
   Dim bGroup As Byte
   Dim bSect As Byte
   Dim btest As Byte
   Dim iList As Integer
   Dim sCustomerId As String
   
   MouseCursor ccHourglass
   Select Case Left$(sProgName, 4)
      Case "Admi"
         bSect = 1
      Case "Sale"
         bSect = 2
      Case "Engi"
         bSect = 3
      Case "Prod"
         bSect = 4
      Case "Inve"
         bSect = 5
      Case "Qual"
         bSect = 6
      Case Else
         'For finance
         bSect = 7
   End Select
On Error GoTo modErr1:
   sSql = "use msdb"
   'clsAdoCon.ExecuteSql sSql, rdExecDirect
   clsAdoCon.ExecuteSql sSql
   sSql = "SELECT ChannelId FROM syschannels WHERE ChannelRow=1"
   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoPrm, ES_FORWARD)
   If bSqlRows Then
      With RdoPrm
         sCustomerId = "" & Trim(!ChannelId)
         If Trim(sCustomerId) <> "" Then
            ES_CUSTOM = sCustomerId
            sSql = "SELECT "
            For iList = 97 To 103
               bGroup = iList
               sSql = sSql & "Channel" & bSect & Chr$(bGroup) & ","
            Next
            bGroup = iList
            sSql = sSql & "Channel" & bSect & Chr$(bGroup) & " " _
                   & "FROM syschannels WHERE ChannelRow=1"
            bSqlRows = clsAdoCon.GetDataSet(sSql, RdoPrm, ES_FORWARD)
            If bSqlRows Then
               With RdoPrm
                  For iList = 0 To 7
                     btest = Asc(.Fields(iList))
                     If btest > 109 Then btest = 1 Else btest = 0
                     bCustomerGroups(iList + 1) = btest
                     If iList + 1 = 4 And btest = 0 Then
                        Debug.Print "whoops"
                     End If
                  Next
                  ClearResultSet RdoPrm
               End With
            End If
         Else
            bFalse = 1
         End If
         ClearResultSet RdoPrm
      End With
   Else
      bFalse = 1
   End If
   If bFalse = 1 Then
      ES_CUSTOM = sserver
      For iList = 1 To 8
         bCustomerGroups(iList) = 1
      Next
   End If
   sSql = "use " & sDataBase
   clsAdoCon.ExecuteSql sSql 'rdExecDirect
   Set RdoPrm = Nothing
   MouseCursor ccArrow
   Exit Sub
   
modErr1:
   'Either the table doesn't exist yet or there are the
   'user doesn't have permission
   'In either case, we have to let them go
   Resume modErr2
modErr2:
   sSql = "use " & sDataBase
   clsAdoCon.ExecuteSql sSql 'rdExecDirect
   For iList = 1 To 8
      bCustomerGroups(iList) = 1
   Next
   MouseCursor ccArrow
End Sub












'12/7/05 Set ES_CUSTOM to these values (see GetCustomerPermissions)

Private Sub aCustomerIDs()
   'Austin Waterjet        WATERJET
   'Intercoastal           INTCOA
   'JEVCO                  JEVCO
   'Production Plating     PROPLA
   '
   'Development            ESINORTH
End Sub
