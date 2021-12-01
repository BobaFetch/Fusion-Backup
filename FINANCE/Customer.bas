Attribute VB_Name = "Customer"
'Customer Permissions 6/27/03
Option Explicit
Public bCustomerGroups(9) As Byte
'6/29/03

Public Sub CodingPrincipals()
   '
   'syschannels require the following:
   '
   'Each tab follows a sequence from 1 to 8 as viewed
   'across the side panel from the MdiSect. Note the items
   'on the channels form.
   '
   'Add Customer.Bas to your project.
   'GetCustomerPermissions to EsiProj.bas was added at
   'the end after GetCompany
   '
   'The following code will be inserted in Load as noted below:
   '
   '        Else
   '            lstFun.Enabled = False
   '            zm(2).Visible = True
   '        End If
   '    End If
   'New Code here *****:
   '    If bCustomerGroups(1) = 0 Then
   '       'they didn't sign up, but we want them to see but not use
   '        lstEdt.Enabled = False
   '        lstVew.Enabled = False
   '        lstFun.Enabled = False
   '        lblCustomer.ForeColor = ES_BLUE
   '        lblCustomer.Visible = True
   '        User.Group1 = 1
   '    End If
   'End New Code ******
   '    If User.Group1 Then
   '        lstEdt.AddItem "New Manufacturing Order"
   '
   'A Label must be produced called lblCustomer
   'Flat
   'Transparent
   'Fixed Single
   'Width = 2895
   'Visible = False
   'Caption = "This Feature Is Not Available
   'Place it in approx the bottom center of the form
   'Left = 840
   'Top = 4080
   '
   'If the Group hasn't strayed to far, then that should be
   'all that is needed
   '
End Sub

Public Sub SysChannelsDoc()
   'msdb syschannels
   
   'ChannelCreate SMALLDATETIME NULL
   'ChannelRevise SMALLDATETIME NULL
   'ChannelOps CHAR(1) NULL DEFAULT('')
   
   'Admn
   'Channel1a CHAR(1) NULL DEFAULT('o')
   'Channel1b CHAR(1) NULL DEFAULT('r')
   'Channel1c CHAR(1) NULL DEFAULT('v')
   'Channel1d CHAR(1) NULL DEFAULT('a')
   'Channel1e CHAR(1) NULL DEFAULT('u')
   'Channel1f CHAR(1) NULL DEFAULT('s')
   'Channel1g CHAR(1) NULL DEFAULT('c')
   'Channel1h CHAR(1) NULL DEFAULT('d')
   
   'Sale
   'Channel2a CHAR(1) NULL DEFAULT('l')
   'Channel2b CHAR(1) NULL DEFAULT('g')
   'Channel2c CHAR(1) NULL DEFAULT('k')
   'Channel2d CHAR(1) NULL DEFAULT('e')
   'Channel2e CHAR(1) NULL DEFAULT('j')
   'Channel2f CHAR(1) NULL DEFAULT('f')
   'Channel2g CHAR(1) NULL DEFAULT('b')
   'Channel2h CHAR(1) NULL DEFAULT('i')
   
   'Engieering
   'Channel3a CHAR(1) NULL DEFAULT('')
   'Channel3b CHAR(1) NULL DEFAULT('')
   'Channel3c CHAR(1) NULL DEFAULT('')
   'Channel3d CHAR(1) NULL DEFAULT('')
   'Channel3e CHAR(1) NULL DEFAULT('')
   'Channel3f CHAR(1) NULL DEFAULT('')
   'Channel3g CHAR(1) NULL DEFAULT('')
   'Channel3h CHAR(1) NULL DEFAULT('')
   
   'Production
   'Channel4a CHAR(1) NULL DEFAULT('')
   'Channel4b CHAR(1) NULL DEFAULT('')
   'Channel4c CHAR(1) NULL DEFAULT('')
   'Channel4d CHAR(1) NULL DEFAULT('')
   'Channel4e CHAR(1) NULL DEFAULT('')
   'Channel4f CHAR(1) NULL DEFAULT('')
   'Channel4g CHAR(1) NULL DEFAULT('')
   'Channel4h CHAR(1) NULL DEFAULT('')
   
   'Inventory
   'Channel5a CHAR(1) NULL DEFAULT('z')
   'Channel5b CHAR(1) NULL DEFAULT('')
   'Channel5c CHAR(1) NULL DEFAULT('')
   'Channel5d CHAR(1) NULL DEFAULT('y')
   'Channel5e CHAR(1) NULL DEFAULT('')
   'Channel5f CHAR(1) NULL DEFAULT('')
   'Channel5g CHAR(1) NULL DEFAULT('')
   'Channel5h CHAR(1) NULL DEFAULT('')
   
   'Quality
   'Channel6a CHAR(1) NULL DEFAULT('')
   'Channel6b CHAR(1) NULL DEFAULT('')
   'Channel6c CHAR(1) NULL DEFAULT('')
   'Channel6d CHAR(1) NULL DEFAULT('')
   'Channel6e CHAR(1) NULL DEFAULT('')
   'Channel6f CHAR(1) NULL DEFAULT('')
   'Channel6g CHAR(1) NULL DEFAULT('')
   'Channel6h CHAR(1) NULL DEFAULT('')
   
   'Finance
   'Channel7a CHAR(1) NULL DEFAULT('')
   'Channel7b CHAR(1) NULL DEFAULT('')
   'Channel7c CHAR(1) NULL DEFAULT('')
   'Channel7d CHAR(1) NULL DEFAULT('')
   'Channel7e CHAR(1) NULL DEFAULT('')
   'Channel7f CHAR(1) NULL DEFAULT('')
   'Channel7g CHAR(1) NULL DEFAULT('')
   'Channel7h CHAR(1) NULL DEFAULT('')
   'Channel8a CHAR(1) NULL DEFAULT('')
   
   'Spare
   'Channel8b CHAR(1) NULL DEFAULT('')
   'Channel8c CHAR(1) NULL DEFAULT('')
   'Channel8d CHAR(1) NULL DEFAULT('')
   'Channel8e CHAR(1) NULL DEFAULT('')
   'Channel8f CHAR(1) NULL DEFAULT('')
   'Channel8g CHAR(1) NULL DEFAULT('')
   'Channel8h CHAR(1) NULL DEFAULT('')
   
   'Customer (Sets this up)
   'ChannelId CHAR(10)NULL DEFAULT('')
   'ChannelRow TINYINT NULL DEFAULT(1))
   
End Sub


'Gets the msdb and only works on SQL 7.0 and later

Public Sub GetCustomerPermissions()
   Dim RdoPrm As rdoResultset
   Dim bFalse As Byte
   Dim bGroup As Byte
   Dim bSect As Byte
   Dim btest As Byte
   Dim i As Integer
   
   MouseCursor 11
   
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
On Error GoTo ModErr1:
   sSql = "use msdb"
   RdoCon.Execute sSql, rdExecDirect
   sSql = "SELECT ChannelID FROM syschannels"
   bSqlRows = GetDataSet(RdoPrm)
   If bSqlRows Then
      With RdoPrm
         If Trim(!ChannelId) = "" Then bFalse = 1 _
                 Else bFalse = 0
         If bFalse = 0 Then
            sSql = "SELECT "
            For i = 97 To 103
               bGroup = i
               sSql = sSql & "Channel" & bSect & Chr$(bGroup) & ","
            Next
            bGroup = i
            sSql = sSql & "Channel" & bSect & Chr$(bGroup) & " " _
                   & "FROM syschannels WHERE ChannelRow=1"
            bSqlRows = GetDataSet(RdoPrm, ES_FORWARD)
            If bSqlRows Then
               With RdoPrm
                  For i = 0 To 7
                     btest = Asc(.rdoColumns(i))
                     If btest > 109 Then btest = 1 Else btest = 0
                     bCustomerGroups(i + 1) = btest
                  Next
                  .Cancel
               End With
            End If
         End If
         .Cancel
      End With
   Else
      bFalse = 1
   End If
   
   If bFalse = 1 Then
      For i = 1 To 8
         bCustomerGroups(i) = 1
      Next
   End If
   sSql = "use " & sDataBase
   RdoCon.Execute sSql, rdExecDirect
   Set RdoPrm = Nothing
   Exit Sub
   
ModErr1:
   'Either the table doesn't exist yet or there are the
   'user doesn't have permission
   'In either case, we have to let them go
   Resume ModErr2
ModErr2:
   sSql = "use " & sDataBase
   For i = 1 To 8
      bCustomerGroups(i) = 1
   Next
   
End Sub
