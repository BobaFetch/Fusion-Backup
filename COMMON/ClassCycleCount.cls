VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCycleCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub PopulateCycleCountCombo(cbo As ComboBox, IncludeLocked As Integer, IncludeReconciled As Integer)
   'Include... = -1 to include if not locked or reconciled
   '           = 0 don't care
   'Include... = 1 to include if locked or reconciled
   '
   cbo.Clear
   sSql = "select CCREF from CchdTable" & vbCrLf
'   If IncludeLocked <> 0 Or IncludeReconciled <> 0 Then
'      sSql = sSql & "where "
'   End If
   
   Dim whereClause As Boolean
   
   Select Case IncludeLocked
   Case -1
      If Not whereClause Then
         sSql = sSql & "where "
         whereClause = True
      Else
         sSql = sSql & "and "
      End If
      sSql = sSql & "CCCOUNTLOCKED = 0" & vbCrLf
   Case 1
      If Not whereClause Then
         sSql = sSql & "where "
         whereClause = True
      Else
         sSql = sSql & "and "
      End If
      sSql = sSql & "CCCOUNTLOCKED = 1" & vbCrLf
   End Select
   
   Select Case IncludeReconciled
   Case -1
      If Not whereClause Then
         sSql = sSql & "where "
         whereClause = True
      Else
         sSql = sSql & "and "
      End If
      sSql = sSql & "CCUPDATED = 0" & vbCrLf
   Case 1
      If Not whereClause Then
         sSql = sSql & "where "
         whereClause = True
      Else
         sSql = sSql & "and "
      End If
      sSql = sSql & "CCUPDATED = 1" & vbCrLf
   End Select
   
   sSql = sSql & "ORDER BY CCREF"
   
   LoadComboBox cbo, -1
   If cbo.ListCount > 0 Then
      cbo = cbo.List(0)
   End If
End Sub



Public Function AnalyzeCounts(CCID As String, TotalItems As Integer, CountsEntered As Integer, _
   NoLots As Integer, CountsRequired) As Boolean
   'gets count information from the database
   'returns True if ready to reconcile (i.e. CountsRequired = 0)
   
   TotalItems = 0
   CountsEntered = 0
   NoLots = 0
   CountsRequired = 0
   
   Dim Ado As ADODB.Recordset
   sSql = "select count(*) from CcltTable where CLREF = '" & CCID & "'"
   If clsADOCon.GetDataSet(sSql, Ado) Then
      TotalItems = Ado.Fields(0)
   End If
   
   sSql = "select count(*) from CcltTable where CLREF = '" & CCID & "'" & vbCrLf _
      & "and CLENTERED = 1"
   If clsADOCon.GetDataSet(sSql, Ado) Then
      CountsEntered = Ado.Fields(0)
   End If
   
   sSql = "select count(*) from CcltTable" & vbCrLf _
      & "join CcitTable on CIREF = CLREF and CIPARTREF = CLPARTREF" & vbCrLf _
      & "and CILOTTRACK = 1 and rtrim(CLLOTNUMBER) = ''" & vbCrLf _
      & "where CLREF = '" & CCID & "'" & vbCrLf
   If clsADOCon.GetDataSet(sSql, Ado) Then
      NoLots = Ado.Fields(0)
   End If
   
   CountsRequired = TotalItems - CountsEntered - NoLots
   
   If CountsRequired = 0 Then
      AnalyzeCounts = True
   Else
      AnalyzeCounts = False
   End If
   Set Ado = Nothing
End Function

