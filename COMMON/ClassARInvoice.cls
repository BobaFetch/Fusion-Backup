VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassARInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GetNextInvoiceNumber() As Long
   'gets the next available invoice number from ComnTable.COLASTINVOICENUMBER
   'if the next number is in use, this function searches upward until an available number
   'is found.
   
   Dim rdo As ADODB.Recordset
   Dim nextNumber As Long
   Dim bRet As Boolean
   
   sSql = "select isnull(COLASTINVOICENUMBER,0) as COLASTINVOICENUMBER from ComnTable"
   bRet = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bRet Then
      nextNumber = rdo!COLASTINVOICENUMBER + 1
   Else
      nextNumber = 1
   End If
   rdo.Close
   Set rdo = Nothing
   'search upward until a free number is found
   Do
      sSql = "select INVNO from CihdTable where INVNO = " & nextNumber
      bRet = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
      If Not bRet Then
         Exit Do
      End If
      nextNumber = nextNumber + 1
   Loop While True
   rdo.Close
   Set rdo = Nothing
   
   'return the new invoice number
   GetNextInvoiceNumber = nextNumber
   
End Function

Public Sub SaveLastInvoiceNumber(InvoiceNumber As Long)
   'save the last used invoice number in ComnTable.COLASTINVOICENUMBER
   'this should be called from within a transaction
   
   'Dim rdo As ADODB.Recordset
   sSql = "update ComnTable set COLASTINVOICENUMBER = " & InvoiceNumber
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
End Sub

Public Function DuplicateInvNumber(InvoiceNumber As Long, Optional ByRef iCanceled As Integer = 0)
   Dim rdo As ADODB.Recordset
   Dim bRet As Boolean
   
   sSql = "select INVNO,INVCANCELED from CihdTable where INVNO = " & InvoiceNumber & " AND INVTYPE <> 'TM'"
   bRet = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If Not bRet Then
      DuplicateInvNumber = False
      iCanceled = 0
   Else
      iCanceled = rdo!INVCANCELED
      ClearResultSet rdo
      DuplicateInvNumber = True
   End If
   
   Set rdo = Nothing
End Function

