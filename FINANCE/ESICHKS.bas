Attribute VB_Name = "ESICHKS"
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
'
' ESICHKS - ES/2000 Check Writing Support
'
' Notes:
' - CHKTYPE Documentation
'   1 = External Check
'   2 = Computer Check
'   3 = No Invoice Check
'   4 = Electronic Check (not used yet)
'
' Created: 04/17/02 (nth)
' Revisons:
'   12/20/02 (nth) Added and changed CHKTYPE documentation.
'   12/26/02 (nth) Added muliple checking account logic to check support functions.
'   02/04/03 (nth) Added clear PreChkSetDte.
'   04/01/03 (nth) Enhanced NumberOfChecks to include vendor and account options.
'   05/14/03 (nth) Change NumberOfChecks to exclude voided checks.
'   09/16/03 (nth) Fix error in NumberOfChecks miscounting checks in setup.
'   05/12/04 (nth) Corrected GetNextCheck to account for periods in check numbers.
'   03/17/05 cjs   Added ModErr checking
'*************************************************************************************

Option Explicit


Public Sub SaveLastCheck(sCheck As String, sAccount As String)
   On Error GoTo modErr1
   If IsNumeric(sCheck) Then
      sSql = "UPDATE GlacTable SET GLLASTCHK = '" & Trim(sCheck) _
             & "' WHERE GLACCTREF = '" & Compress(sAccount) & "'"
      clsADOCon.ExecuteSql sSql
   End If
   Exit Sub
modErr1:
   sProcName = "savelast"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Sub

Public Function GetNextCheck(sAccount As String) As Double
   Dim RdoNum As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT GLLASTCHK FROM GlacTable " _
          & "WHERE GLACCTREF = '" & Compress(sAccount) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNum)
   If bSqlRows Then
      With RdoNum
         If Not IsNull(.Fields(0)) Then
            GetNextCheck = CDbl(.Fields(0)) + 1
         Else
            GetNextCheck = 1
         End If
         .Cancel
      End With
   Else
      GetNextCheck = 1
   End If
   Set RdoNum = Nothing
   Exit Function
modErr1:
   sProcName = "getnextcheck"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Public Function DeleteCheckSetup() As Byte
   Err = 0
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "DELETE FROM ChseTable WHERE CHKREPRINTNO = 0"
   clsADOCon.ExecuteSql sSql
   sSql = "UPDATE Preferences SET PreChkSetDte = NULL"
   clsADOCon.ExecuteSql sSql
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      DeleteCheckSetup = True
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      DeleteCheckSetup = False
   End If
End Function

' Validate check number prevents duplicates.

Public Function ValidateCheck(sCheck As String, _
                              Optional sChkAcct As String) As Byte
   Dim RdoChk As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT DISTINCT CHKNUMBER FROM ChksTable " _
          & "WHERE CHKNUMBER = '" & sCheck & "'"
' TODO: Temporarily commented, once the reports are fixed
' We need to remove this comment.
' MM removed the comment below - We allow same check number for different check account.
   If Trim(sChkAcct) <> "" Then
      sSql = sSql & " AND CHKACCT = '" & sChkAcct & "'"
   End If

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      ValidateCheck = False ' Check Exists
   Else
      ValidateCheck = True
   End If
   Set RdoChk = Nothing
   Exit Function
modErr1:
   sProcName = "validcheck"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function

' Returns the number of checks in setup
' or number of checks for a vendor
' or number of checks for a account (pending)

Public Function NumberOfChecks(Optional sVendor As String, Optional bExcludeCleared As Boolean) As Integer
   Dim RdoChk As ADODB.Recordset
   
   On Error GoTo modErr1
   If sVendor <> "" Then
      sSql = "SELECT COUNT(CHKNUMBER) FROM ChksTable " _
             & "WHERE CHKVENDOR = '" & Compress(sVendor) _
             & "' AND CHKVOID = 0"
      
      If bExcludeCleared Then
         sSql = sSql & " AND CHKCLEARDATE IS NULL"
      End If
   
   
   Else
      sSql = "SELECT COUNT(DISTINCT CHKVND) FROM ChseTable " _
             & "WHERE CHKREPRINTNO = 0"
   End If
   
   'If bExcludeCleared Then
   '   sSql = sSql & " AND CHKCLEARDATE IS NULL"
   'End If
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      With RdoChk
         NumberOfChecks = .Fields(0)
      End With
   End If
   
   Set RdoChk = Nothing
   Exit Function
modErr1:
   sProcName = "numberofchecks"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function
