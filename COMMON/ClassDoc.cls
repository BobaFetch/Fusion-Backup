VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FillClasses(cbo As ComboBox, ShowAllClasses As Boolean)

   'ShowAllClasses = True to include classes with no current documents
   
   If ShowAllClasses Then
      sSql = "SELECT DCLNAME FROM DclsTable" & vbCrLf _
         & "ORDER BY DCLNAME"
   Else
      sSql = "SELECT DISTINCT DCLNAME FROM DclsTable" & vbCrLf _
         & "JOIN DdocTable on DOCLASS = DCLREF" & vbCrLf _
         & "ORDER BY DCLNAME"
   End If
   
   LoadComboBox cbo, -1
   'If cbo.ListCount = 0 Then
   '   MsgBox "Please Create At Least One Drawing Class.", vbExclamation, "No drawing classes defined"
   '   Unload Me
   'End If
End Sub

Public Sub FillDocuments(cboClass As ComboBox, cboDoc As ComboBox)
   sSql = "SELECT DISTINCT DONUM" & vbCrLf _
      & "FROM DdocTable " & vbCrLf _
      & "WHERE DOCLASS='" & Compress(cboClass) & "'" & vbCrLf _
      & "ORDER BY DONUM"
   LoadComboBox cboDoc, -1
End Sub

Public Function GetClassDesc(cbo As ComboBox) As String
   Dim rdo As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetDocumentClass '" & Compress(cbo) & "' "
   If clsADOCon.GetDataSet(sSql, rdo, ES_KEYSET) Then
      With rdo
         GetClassDesc = "" & Trim(!DCLDESC)
         ClearResultSet rdo
      End With
   End If
   Set rdo = Nothing
   Exit Function

DiaErr1:
   sProcName = "GetClassDesc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Public Function GetDocDesc(cboClass As ComboBox, cboDoc As ComboBox) As String
   Dim rdo As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "select DODESCR from DdocTable" & vbCrLf _
      & "where DOCLASS = '" & Compress(cboClass) & "'" & vbCrLf _
      & "and DOREF = '" & Compress(cboDoc) & "'"
   If clsADOCon.GetDataSet(sSql, rdo, ES_KEYSET) Then
      With rdo
         GetDocDesc = "" & Trim(!DODESCR)
         ClearResultSet rdo
      End With
   End If
   Set rdo = Nothing
   Exit Function

DiaErr1:
   sProcName = "GetDocDesc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Public Function DeleteClass(ClassName As String) As Boolean
   'returns True if successful
   
   On Error GoTo whoops
   sSql = "delete from DclsTable" & vbCrLf _
      & "WHERE DCLREF = '" & Compress(ClassName) & "'" & vbCrLf _
      & "and not exists (select DOCLASS from DDocTable where DOCLASS = '" & Compress(ClassName) & "')"
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   If clsADOCon.RowsAffected > 0 Then
      DeleteClass = True
   End If
   
   Exit Function
   
whoops:
   sProcName = "DeleteDoc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm

End Function

Public Sub FillRevisions(cboClass As ComboBox, cboDoc As ComboBox, cboRev As ComboBox)
   sSql = "SELECT DISTINCT DOREV FROM DdocTable " & vbCrLf _
          & "WHERE DOCLASS = '" & Compress(cboClass) & "' AND DOREF='" & Compress(cboDoc) & "' " & vbCrLf _
          & "ORDER BY DOREV"
   LoadComboBox cboRev, -1
End Sub

Public Sub FillSheets(cboClass As ComboBox, cboDoc As ComboBox, cboRev As ComboBox, cboSheet As ComboBox)
   cboSheet.Clear

   sSql = "SELECT DOSHEET FROM DdocTable" & vbCrLf _
      & "WHERE DOCLASS='" & Compress(cboClass) & "' AND DOREF='" & Compress(cboDoc) & "'" & vbCrLf _
      & "AND DOREV='" & Compress(cboRev) & "'" & vbCrLf _
      & "ORDER BY DOSHEET"
   LoadComboBox cboSheet, -1
End Sub



