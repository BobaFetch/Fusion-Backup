VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPackSlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetLastPackSlipNumber() As String
   'gets the most recently used packing slip number from
   'ComnTable.COPSPREFIX and ComnTable.COLASTPSENUMBER
   
   Dim rdo As ADODB.Recordset
   Dim prefix As String, nextNumber As Long
   
   sSql = "select isnull(rtrim(COPSPREFIX),'') as pfx," & vbCrLf _
      & "isnull(COLASTPSNUMBER,0) as lastnum from ComnTable"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      prefix = rdo!pfx
      nextNumber = rdo!lastnum
   Else
      prefix = "PS"
      nextNumber = 0
   End If
   rdo.Close
   
   Dim str As String, psnum As String
   str = Format(nextNumber, "00000000")
   psnum = prefix & Mid(str, Len(prefix) + 1, 8 - Len(prefix))
   GetLastPackSlipNumber = psnum
   
End Function


Public Function GetNextPackSlipNumber() As String
   'gets the next available packing slip number from
   'ComnTable.COPSPREFIX and ComnTable.COLASTPSENUMBER
   'if the next number is in use, this function searches upward until an available number
   'is found.
   
   Dim rdo As ADODB.Recordset
   Dim prefix As String, nextNumber As Long
   
   sSql = "select isnull(rtrim(COPSPREFIX),'') as pfx," & vbCrLf _
      & "isnull(COLASTPSNUMBER,0) as lastnum from ComnTable"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      prefix = rdo!pfx
      nextNumber = rdo!lastnum + 1
   Else
      prefix = "PS"
      nextNumber = 1
   End If
   rdo.Close
   
   'search upward until a free number is found
   Do
      Dim str As String, psnum As String
      str = Format(nextNumber, "00000000")
      psnum = prefix & Mid(str, Len(prefix) + 1, 8 - Len(prefix))
      
      sSql = "select PSNUMBER from PshdTable where PSNUMBER = '" & psnum & "'"
      If Not clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
         Exit Do
      End If
      nextNumber = nextNumber + 1
   Loop While True
   rdo.Close
   
   Set rdo = Nothing
   'return the new invoice number
   GetNextPackSlipNumber = psnum
   
End Function

Public Function GetPackSlipPrefix() As String
   'gets the packing slip prefix from
   'ComnTable.COPSPREFIX
   
   Dim rdo As ADODB.Recordset
   
   sSql = "select isnull(rtrim(COPSPREFIX),'') as pfx" & vbCrLf _
      & "from ComnTable"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      GetPackSlipPrefix = rdo!pfx
   Else
      GetPackSlipPrefix = "PS"
   End If
   rdo.Close
   Set rdo = Nothing
   
End Function

Public Sub SaveLastPSNumber(PsNumber As String)
   'saves the numeric portion of the packing slip number
   'in ComnTable.COLASTPSNUMBER
   'this should be called from within a transaction
   
   Dim rdo As ADODB.Recordset
   sSql = "update ComnTable set COLASTPSNUMBER " & vbCrLf _
      & "= substring('" & PsNumber & "', 1 + len(COPSPREFIX), 8 - len(COPSPREFIX) )" & vbCrLf _
      & "from ComnTable"
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
End Sub

Public Sub FillPSComboAll(cbo As ComboBox)
   sSql = "SELECT PSNUMBER FROM PshdTable" & vbCrLf _
      & "WHERE PSDATE > '" & DateAdd("yyyy", -1, Now) & "'" & vbCrLf _
      & "ORDER BY PSNUMBER DESC"
   LoadComboBoxAndSelect cbo
End Sub

Public Sub FillPSComboUnprinted(cbo As ComboBox)
   sSql = "SELECT PSNUMBER FROM PshdTable" & vbCrLf _
      & "WHERE PSDATE > '" & DateAdd("yyyy", -2, Now) & "'" & vbCrLf _
      & "AND PSTYPE=1 AND PSPRINTED IS NULL" & vbCrLf _
      & "ORDER BY PSNUMBER DESC"
   LoadComboBoxAndSelect cbo
End Sub

Public Sub FillPSComboPrintedNotShipped(cbo As ComboBox)
   sSql = "SELECT PSNUMBER FROM PshdTable" & vbCrLf _
      & "WHERE PSPRINTED IS NOT NULL AND PSINVOICE=0" & vbCrLf _
      & "AND PSSHIPPED=0" & vbCrLf _
      & "ORDER BY PSNUMBER DESC"
   LoadComboBoxAndSelect cbo
End Sub
