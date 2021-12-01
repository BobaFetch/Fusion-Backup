VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub CreateSplit(poNumber As Long, PoRelease As Integer, _
   poItem As Integer, poRev As String, remainingQty As Currency, _
   CopyOnDockDeliveryInfo As Boolean)
   'create a split item for an existing item
   'PoItem/PoRev is the item being split
   'make sure this happens inside the transaction
   
   Dim sRev As String
   Dim rdoRev As ADODB.Recordset
      
   'get revision letter
   sRev = "A"
'   sSql = "select max(PIREV) from PoitTable" & vbCrLf _
'      & "where PINUMBER=" & poNumber & " and PIRELEASE=" & PoRelease & vbCrLf _
'      & "and PIITEM=" & poItem & " and PIREV='" & poRev & "'"
   sSql = "select max(PIREV) from PoitTable" & vbCrLf _
      & "where PINUMBER=" & poNumber & " and PIRELEASE=" & PoRelease & vbCrLf _
      & "and PIITEM=" & poItem
   If clsADOCon.GetDataSet(sSql, rdoRev, ES_FORWARD) Then
      sRev = Trim(Left(rdoRev.Fields(0), 1))
      If sRev = "" Then
         sRev = "A"
      Else
         sRev = Chr(Asc(sRev) + 1)
      End If
   End If
   Set rdoRev = Nothing
'   sSql = "insert into PoitTable (PINUMBER,PIRELEASE,PIITEM," & vbCrLf _
'      & "PIREV,PITYPE,PIPART,PIPDATE,PIPQTY,PIESTUNIT," & vbCrLf _
'      & "PIRUNPART,PIRUNNO,PIRUNOPNO,PICOMT,PIVENDOR,PIONDOCK," & vbCrLf _
'      & "PIUSER,PIENTERED,PIACCOUNT) " & vbCrLf _
'      & "select " & poNumber & "," & PoRelease & "," & poItem & "," & vbCrLf _
'      & "'" & sRev & "'," & IATYPE_PoOpenItem & ",PIPART,PIPDATE," & vbCrLf _
'      & RemainingQty & "," & vbCrLf _
'      & "PIESTUNIT,PIRUNPART,PIRUNNO,PIRUNOPNO,PICOMT,PIVENDOR,PIONDOCK," & vbCrLf _
'      & "'" & sInitials & "',getdate(),PIACCOUNT" & vbCrLf _
'      & "from PoitTable" & vbCrLf _
'      & "where PINUMBER=" & poNumber & " and PIRELEASE=" & PoRelease & vbCrLf _
'      & "and PIITEM=" & poItem & " and PIREV='" & poRev & "'"
   sSql = "insert into PoitTable (PINUMBER,PIRELEASE,PIITEM," & vbCrLf _
      & "PIREV,PITYPE,PIPART,PIPDATE,PIPQTY,PIESTUNIT," & vbCrLf _
      & "PIRUNPART,PIRUNNO,PIRUNOPNO,PICOMT,PIVENDOR,PIONDOCK," & vbCrLf _
      & "PIUSER,PIENTERED,PIACCOUNT"
      
   If CopyOnDockDeliveryInfo Then
      sSql = sSql & ",PIODDELIVERED,PIODDELDATE,PIODDELQTY,PIODDELPSNUMBER"
   End If
   
   sSql = sSql & ") " & vbCrLf _
      & "select " & poNumber & "," & PoRelease & "," & poItem & "," & vbCrLf _
      & "'" & sRev & "'," & IATYPE_PoOpenItem & ",PIPART,PIPDATE," & vbCrLf _
      & remainingQty & "," & vbCrLf _
      & "PIESTUNIT,PIRUNPART,PIRUNNO,PIRUNOPNO,PICOMT,PIVENDOR,PIONDOCK," & vbCrLf _
      & "'" & sInitials & "',getdate(),PIACCOUNT"
   
   If CopyOnDockDeliveryInfo Then
      sSql = sSql & ",PIODDELIVERED,PIODDELDATE," & remainingQty & ",PIODDELPSNUMBER"
   End If
   
   sSql = sSql & vbCrLf _
      & "from PoitTable" & vbCrLf _
      & "where PINUMBER=" & poNumber & " and PIRELEASE=" & PoRelease & vbCrLf _
      & "and PIITEM=" & poItem & " and PIREV='" & poRev & "'"
   clsADOCon.ExecuteSql sSql
   
   ' point operation records tied to this po item to the new po item
   sSql = "update RnopTable" & vbCrLf _
      & "set OPPOREV = '" & sRev & "'" & vbCrLf _
      & "where OPPONUMBER = " & poNumber & " and OPPOITEM = " & poItem & " and OPPOREV = '" & poRev & "'"
   Dim success As Boolean
   success = clsADOCon.ExecuteSql(sSql)


End Sub


Public Sub CancelPOItem(poNumber As Long, PoRelease As Integer, _
   poItem As Integer, poRev As String, remainingQty As Currency)
   'create a split item for an existing item
   'PoItem/PoRev and cancel the item
   
   Dim sRev As String
   Dim rdoRev As ADODB.Recordset
      
   'get revision letter
   sRev = "A"
   sSql = "select max(PIREV) from PoitTable" & vbCrLf _
      & "where PINUMBER=" & poNumber & " and PIRELEASE=" & PoRelease & vbCrLf _
      & "and PIITEM=" & poItem
   If clsADOCon.GetDataSet(sSql, rdoRev, ES_FORWARD) Then
      sRev = Trim(Left(rdoRev.Fields(0), 1))
      If sRev = "" Then
         sRev = "A"
      Else
         sRev = Chr(Asc(sRev) + 1)
      End If
   End If
      
   sSql = "insert into PoitTable (PINUMBER,PIRELEASE,PIITEM," & vbCrLf _
      & "PIREV,PITYPE,PIPART,PIPDATE,PIPQTY,PIESTUNIT," & vbCrLf _
      & "PIRUNPART,PIRUNNO,PIRUNOPNO,PICOMT,PIVENDOR,PIONDOCK," & vbCrLf _
      & "PIUSER,PIENTERED,PIACCOUNT"
   
   sSql = sSql & ") " & vbCrLf _
      & "select " & poNumber & "," & PoRelease & "," & poItem & "," & vbCrLf _
      & "'" & sRev & "'," & IATYPE_PoCanceledItem & ",PIPART,PIPDATE," & vbCrLf _
      & remainingQty & "," & vbCrLf _
      & "PIESTUNIT,PIRUNPART,PIRUNNO,PIRUNOPNO,PICOMT,PIVENDOR,PIONDOCK," & vbCrLf _
      & "'" & sInitials & "',getdate(),PIACCOUNT"
   
   
   sSql = sSql & vbCrLf _
      & "from PoitTable" & vbCrLf _
      & "where PINUMBER=" & poNumber & " and PIRELEASE=" & PoRelease & vbCrLf _
      & "and PIITEM=" & poItem & " and PIREV='" & poRev & "'"
   clsADOCon.ExecuteSql sSql ', rdExecDirect

End Sub
