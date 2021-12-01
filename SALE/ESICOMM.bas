Attribute VB_Name = "ESICOMM"
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'1/11/06 CJS Revised conflicts in GetThisSalesPerson (Was GetSalesPerson)
Option Explicit

Public bCash As Byte ' = 0 if invoice
Public Const CURRENCYMASK = "#,###,###,##0.00"

Public Sub FillSalesPersons(frm As Form)
   Dim rdoSlp As ADODB.Recordset
   
   On Error GoTo ModErr1
   sSql = "Qry_FillSalesPersons"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSlp, ES_FORWARD)
   If bSqlRows Then
      With rdoSlp
         Do Until .EOF
            AddComboStr frm.cmbSlp.hWnd, .Fields(0)
            .MoveNext
         Loop
         ClearResultSet rdoSlp
      End With
      If frm.cmbSlp.ListCount > 0 Then frm.cmbSlp.ListIndex = 0
   End If
   Set rdoSlp = Nothing
   Exit Sub
   
ModErr1:
   sProcName = "fillsalep"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
End Sub

Public Function GetThisSalesPerson(frm As Form) As Byte
   Dim RdoNme As ADODB.Recordset
   
   On Error GoTo ModErr1
   sSql = "SELECT SPLAST,SPFIRST FROM SprsTable WHERE SPNUMBER = '" _
          & Trim(frm.cmbSlp) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNme)
   If bSqlRows Then
      With RdoNme
         frm.lblSlp = "" & Trim(.Fields(1)) & " " & Trim(.Fields(0))
      End With
      GetThisSalesPerson = 1
      ClearResultSet RdoNme
   Else
      GetThisSalesPerson = 0
      frm.lblSlp = "*** Sales Person Not Found ***"
   End If
   Set RdoNme = Nothing
   Exit Function
   
ModErr1:
   sProcName = "getspsos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
End Function

Public Function GetSPVendor(sSP As String) As String
   Dim RdoVnd As ADODB.Recordset
   sSql = "SELECT SPVENDOR FROM SprsTable WHERE SPNUMBER = '" _
          & sSP & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd)
   If bSqlRows Then
      With RdoVnd
         GetSPVendor = "" & Trim(.Fields(0))
      End With
      ClearResultSet RdoVnd
   End If
   Set RdoVnd = Nothing
End Function

Public Function GetSPAccount(sSP As String) As String
   Dim RdoAct As ADODB.Recordset
   sSql = "SELECT SPACCOUNT FROM SprsTable WHERE SPNUMBER = '" _
          & sSP & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct)
   If bSqlRows Then
      With RdoAct
         GetSPAccount = "" & Trim(.Fields(0))
      End With
      ClearResultSet RdoAct
   End If
   Set RdoAct = Nothing
End Function

Public Function InvOrCash() As Byte
   ' 0 = Pay commissions on invoice
   ' 1 = Pay commissions on cash receipt
   Dim RdoPay As ADODB.Recordset
   sSql = "SELECT COCOMMISSION FROM ComnTable WHERE COREF = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPay)
   If bSqlRows Then
      With RdoPay
         InvOrCash = .Fields(0)
      End With
      ClearResultSet RdoPay
   End If
   Set RdoPay = Nothing
End Function

Public Sub GetSPSOs(frm As Form)
   Dim RdoSO As ADODB.Recordset
   
   On Error GoTo ModErr1
   
   frm.cmbSon.Clear
   If bCash = 0 Then
      sSql = "SELECT DISTINCT ITSO FROM SoitTable INNER JOIN SpcoTable ON ITSO = SMCOSO " _
             & "AND ITNUMBER = SMCOSOIT AND ITREV = SMCOITREV LEFT OUTER JOIN " _
             & "SpapTable ON SMCOSO = COSO AND SMCOSOIT = COSOIT AND SMCOITREV = COSOITREV " _
             & "WHERE SMCOSM = '" & frm.cmbSlp & "' AND COAPINV IS NULL AND ITINVOICE <> 0"
   Else
      sSql = "SELECT DISTINCT ITSO FROM SoitTable INNER JOIN SpcoTable ON ITSO = SMCOSO AND " _
             & "ITNUMBER = SMCOSOIT AND ITREV = SMCOITREV INNER JOIN CihdTable ON " _
             & "ITINVOICE = INVNO LEFT OUTER JOIN SpapTable ON SMCOSO = COSO AND " _
             & "SMCOSOIT = COSOIT AND SMCOITREV = COSOITREV AND SMCOSM = COSPNUMBER " _
             & "WHERE SMCOSM = '" & frm.cmbSlp & "' AND COAPINV IS NULL AND INVPIF = 1"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSO, ES_FORWARD)
   If bSqlRows Then
      With RdoSO
         Do Until .EOF
            AddComboStr frm.cmbSon.hWnd, Format(.Fields(0), SO_NUM_FORMAT)
            .MoveNext
         Loop
         ClearResultSet RdoSO
      End With
      If frm.cmbSon.ListCount > 0 Then frm.cmbSon.ListIndex = 0
   End If
   Set RdoSO = Nothing
   Exit Sub
   
ModErr1:
   sProcName = "getspsos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
End Sub
