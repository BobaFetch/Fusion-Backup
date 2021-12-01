Attribute VB_Name = "TerryTime"
'common code used by EsiPom and EsiTime

'Public gblnSqlRows As Boolean

' Job Datatype

Type Job
   strPart As String
   lngRun As Long
   intOp As Integer
   sngQty As Single
End Type

Type Employee
   intNumber As Long ' Employee number
   strFirstName As String ' Employee first name
   strLastName As String ' Employee last name
   strCurShop As String ' * 12  ' Current shop selected
   strCurWC As String ' * 12  ' Current work center selected
   strAccount As String ' Indirect Labor Field
   sngRate As Single ' Employee Rate
   jobCurMO() As Job ' Jobs logged into      'see MAX_CONCURRENT_LOGINS
   strTimeCard As String ' Time card for the day
End Type

Public mempCurrentEmployee As Employee

Option Explicit

'Sub UpdateTimeDatabase()
'   'update time related tables from the POM or EsiTime modules
'   'call immediately after OpenSqlServer
'
'   On Error Resume Next
'
'   'add TCPRORATE column to TcitTable
'   sSql = "ALTER TABLE TcitTable ADD TCPRORATE decimal(8,3) NULL DEFAULT(0)"
'   RdoCon.Execute sSql, rdExecDirect
'   sSql = "Update TcitTable set TCPRORATE = 0 where TCPRORATE is null"
'   RdoCon.Execute sSql, rdExecDirect
'
'   '    If Err Then
'   '        MsgBox Err.Description
'   '    End If
'End Sub
'


Public Sub UpdateOpFromTimeCharges(PartNo As String, Runno As Long, _
                                   opNo As Integer, complete As Boolean)
   
   'get sums from related operations
   Dim rdo As ADODB.Recordset
   Dim Hours As Single, yield As Single, accept As Single, reject As Single, scrap As Single
   
   On Error Resume Next
   Err.Clear
' MM 9/8    clsADOCon.BeginTrans
   
   sSql = _
          "SELECT ISNULL(SUM(CAST(TCHOURS AS DECIMAL(10,3))),0) AS TCHOURS," & vbCrLf _
          & "ISNULL(CAST(SUM(TCYIELD) AS DECIMAL(10,3)),0) AS TCYIELD," & vbCrLf _
          & "ISNULL(CAST(SUM(TCACCEPT) AS DECIMAL(10,3)),0) AS TCACCEPT," & vbCrLf _
          & "ISNULL(CAST(SUM(TCREJECT) AS DECIMAL(10,3)),0) AS TCREJECT," & vbCrLf _
          & "ISNULL(CAST(SUM(TCSCRAP) AS DECIMAL(10,3)),0) AS TCSCRAP" & vbCrLf _
          & "From TcitTable" & vbCrLf _
          & "WHERE TCPARTREF='" & Compress(PartNo) & "'" & vbCrLf _
          & "AND TCRUNNO=" & Runno & vbCrLf _
          & "AND TCOPNO=" & opNo
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If gblnSqlRows Then
      With rdo
         Hours = !TCHOURS
         yield = !TCYIELD 'this is alays = 0.  Use TCACCEPT for yield
         accept = !TCACCEPT
         reject = !TCREJECT
         scrap = !TCSCRAP
      End With
   End If
   
   sSql = _
          "UPDATE RnopTable" & vbCrLf _
          & "SET OPCHARGED=" & Hours & "," & vbCrLf _
          & "OPYIELD=" & accept & "," & vbCrLf _
          & "OPACCEPT=" & accept & "," & vbCrLf _
          & "OPREJECT=" & reject & "," & vbCrLf _
          & "OPSCRAP=" & scrap & vbCrLf
   
   If complete Then
      sSql = sSql _
             & ",OPCOMPLETE = 1," & vbCrLf _
             & "OPCOMPDATE = convert(datetime,getdate(),101)," & vbCrLf _
             & "OPINSP = '" & mempCurrentEmployee.strFirstName & " " _
             & mempCurrentEmployee.strLastName & "'" & vbCrLf
   End If
   
   sSql = sSql _
          & "From TcitTable" & vbCrLf _
          & "JOIN RnopTable on TCPARTREF=OPREF" & vbCrLf _
          & "AND TCRUNNO=OPRUN" & vbCrLf _
          & "AND  TCOPNO=OPNO" & vbCrLf _
          & "WHERE OPREF='" & Compress(PartNo) & "'" & vbCrLf _
          & "AND OPRUN=" & Runno & vbCrLf _
          & "AND OPNO=" & opNo
   clsADOCon.ExecuteSQL sSql
   
   'if operation is complete, set next op as current
   If complete Then
      Dim nextop As Integer
      sSql = "select min(OPNO) as NEXTOP from RnopTable" & vbCrLf _
             & "WHERE OPREF='" & Compress(PartNo) & "'" & vbCrLf _
             & "AND OPRUN=" & Runno & vbCrLf _
             & "AND OPCOMPLETE = 0"
             
             '& "AND OPNO>" & opNo
      gblnSqlRows = clsADOCon.GetDataSet(sSql, rdo)
      If gblnSqlRows Then
         nextop = rdo!nextop
      Else
         nextop = 0
      End If
      
      sSql = "UPDATE RunsTable" & vbCrLf _
             & "SET RUNOPCUR=" & nextop & vbCrLf _
             & "WHERE RUNREF='" & Compress(PartNo) & "'" & vbCrLf _
             & "AND RUNNO=" & Runno
      clsADOCon.ExecuteSQL sSql
      
   End If
   
   If Err Then
' MM 9/8       clsADOCon.RollbackTrans
      MsgBox "Update failed: " & Err.Description, vbInformation, "UpdateOpFromTimeCharges"
   Else
' MM 9/8       clsADOCon.CommitTrans
   End If
End Sub
