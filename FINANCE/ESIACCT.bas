Attribute VB_Name = "ESIACCT"
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'**************************************************************************************
' ESIACCT.BAS - ES/2000 GL accounts source
'
' Notes:
'
' Created: (nth)
' Revisons:
' 10/02/00 (cjs) Accounting options if Co.GlVerify = 1 then see if there is a journal open in for this date.
' 01/12/04 (nth) Enhanced UpdateActDesc (bNoMult) disable multiple accounts
' 09/28/04 (nth) Added bNoCash option to UpdateActDesc
' 11/17/04 (nth) Added Account Balance
'
'**************************************************************************************

Public sJournalID As String
Public iTotal As Integer
Public iInActive As Integer ' = 1 if include inactive
Public iCurType As Integer
Public iFsLevel As Integer
Public iLevel As Integer
Public DbAct As Recordset
Public bChart As Byte ' = 1 if we are just displying the chart of accounts

Public Function UpdateActDesc( _
                              sAct As ComboBox, _
                              Optional lblLabel As Label, _
                              Optional bNoMult As Byte, _
                              Optional bNoCash As Byte)
   
   
   Dim rdoAct As ADODB.Recordset
   Dim frm As Form
   
   On Error GoTo modErr1
   Set frm = MdiSect.ActiveForm
   
   ' pre 11/3/05 logic: sact = trim(sact) caused click event loop withdropdown list
   Dim sAcct As String
   sAcct = Trim(sAct)
   
   If sAcct = "" Or UCase(sAcct) = "ALL" Then
      If bNoMult Then
         If sAct.ListCount > 0 Then
            sAct.ListIndex = 0
            UpdateActDesc = UpdateActDesc(sAct, lblLabel, bNoMult)
         End If
      Else
         UpdateActDesc = "Multiple Accounts Selected."
      End If
      Exit Function
   End If
   
   sSql = "SELECT GLDESCR FROM GlacTable WHERE GLACCTREF = '" _
          & Compress(sAct) & "'"
   If bNoCash Then
      sSql = sSql & " AND GLCASH = 0"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   On Error Resume Next
   If bSqlRows Then
      lblLabel.ForeColor = frm.ForeColor
      UpdateActDesc = "" & Trim(rdoAct!GLDESCR)
   Else
      lblLabel.ForeColor = ES_RED
      UpdateActDesc = "*** Invalid Account Number ***"
   End If
   Set rdoAct = Nothing
   Exit Function
   
modErr1:
   sProcName = "updateactdesc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

Public Function GetThisAccount(sAccount As String) As String
   Dim RdoGlm As ADODB.Recordset
   
   On Error GoTo modErr1
   sAccount = Compress(sAccount)
   
   sSql = "SELECT GLACCTREF,GLACCTNO FROM GlacTable " _
          & "WHERE GLACCTREF='" & sAccount & "' AND GLINACTIVE=0"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      With RdoGlm
         GetThisAccount = Trim(!GLACCTNO)
         .Cancel
      End With
   Else
      GetThisAccount = ""
   End If
   Set RdoGlm = Nothing
   Exit Function
   
modErr1:
   Resume modErr2
modErr2:
   On Error GoTo 0
End Function

'WIP Labor Accounts 8/28/99
'Use local errors

Public Function GetLaborAcct(sPartNumber As String, sCode As String, bLevel As Byte) As String
   Dim RdoLac As ADODB.Recordset
   
   'Part Number
   sSql = "SELECT PAINVLABACCT FROM PartTable WHERE " _
          & "PARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLac, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoLac.Fields(0)) Then GetLaborAcct = _
                    Trim(RdoLac.Fields(0))
   End If
   
   'Product Code
   If GetLaborAcct = "" Then
      sSql = "SELECT PCINVLABACCT FROM PcodTable WHERE " _
             & "PCREF='" & sCode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoLac.Fields(0)) Then GetLaborAcct = _
                       Trim(RdoLac.Fields(0))
      End If
   End If
   
   'Default
   If GetLaborAcct = "" Then
      sSql = "SELECT COINVLABACCT" & Trim(str(bLevel)) & " FROM ComnTable " _
             & "WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoLac.Fields(0)) Then GetLaborAcct = _
                       Trim(RdoLac.Fields(0))
      End If
   End If
   Set RdoLac = Nothing
   
End Function

'WIP Overhead Accounts 8/28/99
'Use local errors

Public Function GetOverHeadAcct(sPartNumber As String, sCode As String, bLevel As Byte) As String
   Dim RdoOac As ADODB.Recordset
   
   'Part Number
   sSql = "SELECT PAINVOHDACCT FROM PartTable WHERE " _
          & "PARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOac, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoOac.Fields(0)) Then GetOverHeadAcct = _
                    Trim(RdoOac.Fields(0))
   End If
   
   'Product Code
   If GetOverHeadAcct = "" Then
      sSql = "SELECT PCINVOHDACCT FROM PcodTable WHERE " _
             & "PCREF='" & sCode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoOac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoOac.Fields(0)) Then GetOverHeadAcct = _
                       Trim(RdoOac.Fields(0))
      End If
   End If
   
   'Default
   If GetOverHeadAcct = "" Then
      sSql = "SELECT COINVOHDACCT" & Trim(str(bLevel)) & " FROM ComnTable " _
             & "WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoOac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoOac.Fields(0)) Then GetOverHeadAcct = _
                       Trim(RdoOac.Fields(0))
      End If
   End If
   Set RdoOac = Nothing
   
End Function

'WIP Material Accounts 8/28/99
'Use local errors

Public Function GetMaterialAcct(sPartNumber As String, sCode As String, bLevel As Byte) As String
   Dim RdoMac As ADODB.Recordset
   
   'Part Number
   sSql = "SELECT PAINVMATACCT FROM PartTable WHERE " _
          & "PARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMac, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoMac.Fields(0)) Then GetMaterialAcct = _
                    Trim(RdoMac.Fields(0))
   End If
   
   'Product Code
   If GetMaterialAcct = "" Then
      sSql = "SELECT PCINVMATACCT FROM PcodTable WHERE " _
             & "PCREF='" & sCode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoMac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoMac.Fields(0)) Then GetMaterialAcct = _
                       Trim(RdoMac.Fields(0))
      End If
   End If
   
   'Default
   If GetMaterialAcct = "" Then
      sSql = "SELECT COINVMATACCT" & Trim(str(bLevel)) & " FROM ComnTable " _
             & "WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoMac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoMac.Fields(0)) Then GetMaterialAcct = _
                       Trim(RdoMac.Fields(0))
      End If
   End If
   Set RdoMac = Nothing
   
End Function

'WIP Expense Accounts 8/28/99
'Use local errors

Public Function GetExpenseAcct(sPartNumber As String, sCode As String, bLevel As Byte) As String
   Dim RdoEac As ADODB.Recordset
   
   'Part Number
   sSql = "SELECT PAINVEXPACCT FROM PartTable WHERE " _
          & "PARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEac, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoEac.Fields(0)) Then GetExpenseAcct = _
                    Trim(RdoEac.Fields(0))
   End If
   
   'Product Code
   If GetExpenseAcct = "" Then
      sSql = "SELECT PCINVEXPACCT FROM PcodTable WHERE " _
             & "PCREF='" & sCode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoEac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoEac.Fields(0)) Then GetExpenseAcct = _
                       Trim(RdoEac.Fields(0))
      End If
   End If
   
   'Default
   If GetExpenseAcct = "" Then
      sSql = "SELECT COINVEXPACCT" & Trim(str(bLevel)) & " FROM ComnTable " _
             & "WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoEac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoEac.Fields(0)) Then GetExpenseAcct = _
                       Trim(RdoEac.Fields(0))
      End If
   End If
   Set RdoEac = Nothing
   Exit Function
   
modErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function






Public Sub FindAccount(frm As Form)
   'use local Errors
   Dim rdoAct As ADODB.Recordset
   Dim sAccount As String
   sAccount = Compress(frm.cmbAct)
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR FROM " _
          & "GlacTable WHERE GLACCTREF='" & sAccount & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         frm.cmbAct = "" & Trim(!GLACCTNO)
         frm.lblDsc = "" & Trim(!GLDESCR)
         .Cancel
      End With
   Else
      If Len(sAccount) > 0 Then
         frm.lblDsc = "*** Account Wasn't Found ***"
      Else
         frm.lblDsc = "Multiple Accounts Selected."
      End If
   End If
   Set rdoAct = Nothing
End Sub

'*************************************************************************
'   Function: GetPartInvoiceAccounts
'
'   Notes: Retrieve accounts for invoices. Returns to account values
'          to the caller byref.
'
'   Created:
'   Modified:   6/1/01 (nth) To return more than just two accounts
'*************************************************************************
'

Public Function GetPartInvoiceAccounts(SPartRef As String, iLevel As Integer, sCode As String, _
                                       Optional sRevAccount As String, _
                                       Optional sDisAccount As String, _
                                       Optional sCGSMaterialAccount As String, _
                                       Optional sCGSLaborAccount As String, _
                                       Optional sCGSExpAccount As String, _
                                       Optional sCGSOhAccount As String, _
                                       Optional sInvMaterialAccount As String, _
                                       Optional sInvLaborAccount As String, _
                                       Optional sInvExpAccount As String, _
                                       Optional sInvOhAccount As String) As Boolean
   
   Dim rdoAct As ADODB.Recordset
   On Error GoTo modErr1
   
   'Part
   GetPartInvoiceAccounts = True
   SPartRef = Compress(SPartRef)
   
   sSql = "SELECT PACGSMATACCT,PACGSLABACCT,PACGSEXPACCT,PACGSOHDACCT," _
          & "PAINVMATACCT,PAINVLABACCT,PAINVEXPACCT,PAINVOHDACCT," _
          & "PAREVACCT,PADISACCT FROM PartTable WHERE PARTREF='" & SPartRef & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         
         sRevAccount = "" & Trim(!PAREVACCT)
         sDisAccount = "" & Trim(!PADISACCT)
         
         sCGSMaterialAccount = "" & Trim(!PACGSMATACCT)
         sCGSLaborAccount = "" & Trim(!PACGSLABACCT)
         sCGSExpAccount = "" & Trim(!PACGSEXPACCT)
         sCGSOhAccount = "" & Trim(!PACGSOHDACCT)
         
         sInvMaterialAccount = "" & Trim(!PAINVMATACCT)
         sInvLaborAccount = "" & Trim(!PAINVLABACCT)
         sInvExpAccount = "" & Trim(!PAINVEXPACCT)
         sInvOhAccount = "" & Trim(!PAINVOHDACCT)
         
         .Cancel
      End With
   End If
   
   ' Now check the accounts, if any are blank then fill then from the
   ' product code
   sCode = Compress(sCode)
   
   sSql = "SELECT PCCGSMATACCT,PCCGSLABACCT,PCCGSEXPACCT,PCCGSOHDACCT," _
          & "PCINVMATACCT,PCINVLABACCT,PCINVEXPACCT,PCINVOHDACCT," _
          & "PCREVACCT,PCDISCACCT FROM PcodTable WHERE PCREF='" & sCode & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         
         If sRevAccount = "" Then sRevAccount = "" & Trim(!PCREVACCT)
         If sDisAccount = "" Then sDisAccount = "" & Trim(!PCDISCACCT)
         
         If sCGSMaterialAccount = "" Then sCGSMaterialAccount = "" & Trim(!PCCGSMATACCT)
         If sCGSLaborAccount = "" Then sCGSLaborAccount = "" & Trim(!PCCGSLABACCT)
         If sCGSExpAccount = "" Then sCGSExpAccount = "" & Trim(!PCCGSEXPACCT)
         If sCGSOhAccount = "" Then sCGSOhAccount = "" & Trim(!PCCGSOHDACCT)
         
         If sInvMaterialAccount = "" Then sInvMaterialAccount = "" & Trim(!PCINVMATACCT)
         If sInvLaborAccount = "" Then sInvLaborAccount = "" & Trim(!PCINVLABACCT)
         If sInvExpAccount = "" Then sInvExpAccount = "" & Trim(!PCINVEXPACCT)
         If sInvOhAccount = "" Then sInvOhAccount = "" & Trim(!PCINVOHDACCT)
         
         .Cancel
      End With
   End If
   
   ' Last check the company setup and fill any accounts that are still empty.
   sSql = "SELECT COREVACCT" & Trim(str(iLevel)) & "," _
          & "COAPDISCACCT," _
          & "COCGSMATACCT" & Trim(str(iLevel)) & "," _
          & "COCGSLABACCT" & Trim(str(iLevel)) & "," _
          & "COCGSEXPACCT" & Trim(str(iLevel)) & "," _
          & "COCGSOHDACCT" & Trim(str(iLevel)) & "," _
          & "COINVMATACCT" & Trim(str(iLevel)) & "," _
          & "COINVLABACCT" & Trim(str(iLevel)) & "," _
          & "COINVEXPACCT" & Trim(str(iLevel)) & "," _
          & "COINVOHDACCT" & Trim(str(iLevel)) & " FROM " _
          & "ComnTable WHERE COREF=1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         
         If sRevAccount = "" Then sRevAccount = "" & Trim(.Fields(0))
         If sDisAccount = "" Then sDisAccount = "" & Trim(.Fields(1))
         
         If sCGSMaterialAccount = "" Then sCGSMaterialAccount = "" & Trim(.Fields(2))
         If sCGSLaborAccount = "" Then sCGSLaborAccount = "" & Trim(.Fields(3))
         If sCGSExpAccount = "" Then sCGSExpAccount = "" & Trim(.Fields(4))
         If sCGSOhAccount = "" Then sCGSOhAccount = "" & Trim(.Fields(5))
         
         If sInvMaterialAccount = "" Then sInvMaterialAccount = "" & Trim(.Fields(6))
         If sInvLaborAccount = "" Then sInvLaborAccount = "" & Trim(.Fields(7))
         If sInvExpAccount = "" Then sInvExpAccount = "" & Trim(.Fields(8))
         If sInvOhAccount = "" Then sInvOhAccount = "" & Trim(.Fields(9))
         
         .Cancel
      End With
   End If
   
   
   Set rdoAct = Nothing
   Exit Function
   
modErr1:
   sProcName = "GetPartInvoiceAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   MsgBox CurrError.Number & " " & CurrError.Description
End Function

Public Sub FillLevel1(sMaster As String)
   Dim i As Integer
   Dim RdoAct1 As ADODB.Recordset
   Dim iSubAccounts As Integer
   
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct1)
   If bSqlRows Then
      With RdoAct1
         Do Until .EOF
            With DbAct
               .AddNew
               !Act00 = iCurType
               !Act01 = 1
               !Act02 = "" & Trim(RdoAct1!GLACCTREF)
               !Act03 = String(2, Chr(160)) & "" & Trim(RdoAct1!GLACCTNO)
               !Act04 = String(2, Chr(160)) & "" & Trim(RdoAct1!GLDESCR)
               !Act05 = RdoAct1!GLINACTIVE
               
               iFsLevel = GetAcctLevel(Trim(RdoAct1!GLACCTREF))
               !Act06 = iFsLevel
               .Update
            End With
            If iLevel > 1 Then
               iSubAccounts = FillLevel2(Trim(!GLACCTREF))
               If bChart = 0 And iSubAccounts > 0 Then
                  InsertTotalRow RdoAct1, 1
               End If
            End If
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      iSubAccounts = 0
   End If
   Set RdoAct1 = Nothing
   Exit Sub
DiaErr1:
   sProcName = "filllevel1"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Sub

Public Function FillLevel2(sMaster As String) As Integer
   Dim i As Integer
   Dim RdoAct2 As ADODB.Recordset
   Dim iSubAccounts As Integer
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct2)
   If bSqlRows Then
      With RdoAct2
         Do Until .EOF
            With DbAct
               .AddNew
               !Act00 = iCurType
               !Act01 = 2
               !Act02 = "" & Trim(RdoAct2!GLACCTREF)
               !Act03 = String$(4, Chr$(160)) & "" & Trim(RdoAct2!GLACCTNO)
               !Act04 = String$(4, Chr$(160)) & "" & Trim(RdoAct2!GLDESCR)
               !Act05 = RdoAct2!GLINACTIVE
               iFsLevel = GetAcctLevel(Trim(RdoAct2!GLACCTREF))
               !Act06 = iFsLevel
               .Update
            End With
            If iLevel > 2 Then
               iSubAccounts = FillLevel3(Trim(!GLACCTREF))
               If bChart = 0 And iSubAccounts > 0 Then
                  InsertTotalRow RdoAct2, 2
               End If
            End If
            FillLevel2 = FillLevel2 + 1
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoAct2 = Nothing
   Exit Function
DiaErr1:
   sProcName = "filllevel2"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

Public Function FillLevel3(sMaster As String) As Integer
   Dim i As Integer
   Dim RdoAct3 As ADODB.Recordset
   Dim iSubAccounts As Integer
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct3)
   If bSqlRows Then
      With RdoAct3
         Do Until .EOF
            With DbAct
               .AddNew
               !Act00 = iCurType
               !Act01 = 3
               !Act02 = "" & Trim(RdoAct3!GLACCTREF)
               !Act03 = String$(6, Chr$(160)) & "" & Trim(RdoAct3!GLACCTNO)
               !Act04 = String$(6, Chr$(160)) & "" & Trim(RdoAct3!GLDESCR)
               !Act05 = RdoAct3!GLINACTIVE
               iFsLevel = GetAcctLevel(Trim(RdoAct3!GLACCTREF))
               !Act06 = iFsLevel
               .Update
            End With
            If iLevel > 3 Then
               iSubAccounts = FillLevel4(Trim(!GLACCTREF))
               If bChart = 0 And iSubAccounts > 0 Then
                  InsertTotalRow RdoAct3, 3
               End If
            End If
            FillLevel3 = FillLevel3 + 1
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoAct3 = Nothing
   Exit Function
DiaErr1:
   sProcName = "filllevel3"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

Public Function FillLevel4(sMaster As String) As Integer
   Dim i As Integer
   Dim RdoAct4 As ADODB.Recordset
   Dim iSubAccounts As Integer
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct4)
   If bSqlRows Then
      With RdoAct4
         Do Until .EOF
            With DbAct
               .AddNew
               !Act00 = iCurType
               !Act01 = 4
               !Act02 = "" & Trim(RdoAct4!GLACCTREF)
               !Act03 = String$(8, Chr$(160)) & "" & Trim(RdoAct4!GLACCTNO)
               !Act04 = String$(8, Chr$(160)) & "" & Trim(RdoAct4!GLDESCR)
               !Act05 = RdoAct4!GLINACTIVE
               iFsLevel = GetAcctLevel(Trim(RdoAct4!GLACCTREF))
               !Act06 = iFsLevel
               .Update
            End With
            If iLevel > 4 Then
               iSubAccounts = FillLevel5(Trim(!GLACCTREF))
               If bChart = 0 And iSubAccounts > 0 Then
                  InsertTotalRow RdoAct4, 4
               End If
            End If
            FillLevel4 = FillLevel4 + 1
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoAct4 = Nothing
   Exit Function
DiaErr1:
   sProcName = "filllevel4"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

Public Function FillLevel5(sMaster As String) As Integer
   Dim i As Integer
   Dim RdoAct5 As ADODB.Recordset
   Dim iSubAccounts As Integer
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct5)
   If bSqlRows Then
      With RdoAct5
         Do Until .EOF
            With DbAct
               .AddNew
               !Act00 = iCurType
               !Act01 = 5
               !Act02 = "" & Trim(RdoAct5!GLACCTREF)
               !Act03 = String$(10, Chr$(160)) & "" & Trim(RdoAct5!GLACCTNO)
               !Act04 = String$(10, Chr$(160)) & "" & Trim(RdoAct5!GLDESCR)
               !Act05 = RdoAct5!GLINACTIVE
               iFsLevel = GetAcctLevel(Trim(RdoAct5!GLACCTREF))
               !Act06 = iFsLevel
               .Update
            End With
            If iLevel > 5 Then
               iSubAccounts = FillLevel6(Trim(!GLACCTREF))
               If bChart = 0 And iSubAccounts > 0 Then
                  InsertTotalRow RdoAct5, 5
               End If
            End If
            FillLevel5 = FillLevel5 + 1
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoAct5 = Nothing
   Exit Function
DiaErr1:
   sProcName = "filllevel5"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Public Function FillLevel6(sMaster As String) As Integer
   Dim i As Integer
   Dim RdoAct6 As ADODB.Recordset
   Dim iSubAccounts As Integer
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct6)
   If bSqlRows Then
      With RdoAct6
         Do Until .EOF
            With DbAct
               .AddNew
               !Act00 = iCurType
               !Act01 = 6
               !Act02 = "" & Trim(RdoAct6!GLACCTREF)
               !Act03 = String$(12, Chr$(160)) & "" & Trim(RdoAct6!GLACCTNO)
               !Act04 = String$(12, Chr$(160)) & "" & Trim(RdoAct6!GLDESCR)
               !Act05 = RdoAct6!GLINACTIVE
               iFsLevel = GetAcctLevel(Trim(RdoAct6!GLACCTREF))
               !Act06 = iFsLevel
               .Update
            End With
            
            If iLevel > 6 Then
               iSubAccounts = FillLevel7(Trim(!GLACCTREF))
               If bChart = 0 And iSubAccounts > 0 Then
                  InsertTotalRow RdoAct6, 6
               End If
            End If
            FillLevel6 = FillLevel6 + 1
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoAct6 = Nothing
   Exit Function
DiaErr1:
   sProcName = "filllevel6"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Public Function FillLevel7(sMaster As String) As Integer
   Dim i As Integer
   Dim RdoAct7 As ADODB.Recordset
   Dim iSubAccounts As Integer
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct7)
   If bSqlRows Then
      With RdoAct7
         Do Until .EOF
            With DbAct
               .AddNew
               !Act00 = iCurType
               !Act01 = 7
               !Act02 = "" & Trim(RdoAct7!GLACCTREF)
               !Act03 = String$(14, Chr$(160)) & "" & Trim(RdoAct7!GLACCTNO)
               !Act04 = String$(14, Chr$(160)) & "" & Trim(RdoAct7!GLDESCR)
               !Act05 = RdoAct7!GLINACTIVE
               iFsLevel = GetAcctLevel(Trim(RdoAct7!GLACCTREF))
               !Act06 = iFsLevel
               .Update
            End With
            
            If iLevel > 7 Then
               iSubAccounts = FillLevel8(Trim(!GLACCTREF))
               If bChart = 0 And iSubAccounts > 0 Then
                  InsertTotalRow RdoAct7, 7
               End If
            End If
            FillLevel7 = FillLevel7 + 1
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoAct7 = Nothing
   Exit Function
DiaErr1:
   sProcName = "filllevel7"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Public Function FillLevel8(sMaster As String) As Integer
   Dim i As Integer
   Dim RdoAct8 As ADODB.Recordset
   Dim iSubAccounts As Integer
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct8)
   If bSqlRows Then
      With RdoAct8
         Do Until .EOF
            With DbAct
               .AddNew
               !Act00 = iCurType
               !Act01 = 8
               !Act02 = "" & Trim(RdoAct8!GLACCTREF)
               !Act03 = String$(16, Chr$(160)) & "" & Trim(RdoAct8!GLACCTNO)
               !Act04 = String$(16, Chr$(160)) & "" & Trim(RdoAct8!GLDESCR)
               !Act05 = RdoAct8!GLINACTIVE
               iFsLevel = GetAcctLevel(Trim(RdoAct8!GLACCTREF))
               !Act06 = iFsLevel
               .Update
            End With
            
            If iLevel > 8 Then
               iSubAccounts = FillLevel9(Trim(!GLACCTREF))
               If bChart = 0 And iSubAccounts > 0 Then
                  InsertTotalRow RdoAct8, 8
               End If
            End If
            FillLevel8 = FillLevel8 + 1
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      iSubAccounts = 0
   End If
   Set RdoAct8 = Nothing
   Exit Function
DiaErr1:
   sProcName = "filllevel8"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

Public Function FillLevel9(sMaster As String) As Integer
   Dim i As Integer
   Dim RdoAct9 As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct9)
   If bSqlRows Then
      With RdoAct9
         Do Until .EOF
            With DbAct
               .AddNew
               !Act00 = iCurType
               !Act01 = 9
               !Act02 = "" & Trim(RdoAct9!GLACCTREF)
               !Act03 = String$(18, Chr$(160)) & "" & Trim(RdoAct9!GLACCTNO)
               !Act04 = String$(18, Chr$(160)) & "" & Trim(RdoAct9!GLDESCR)
               !Act05 = RdoAct9!GLINACTIVE
               iFsLevel = GetAcctLevel(Trim(RdoAct9!GLACCTREF))
               !Act06 = iFsLevel
               .Update
            End With
            FillLevel9 = FillLevel9 + 1
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoAct9 = Nothing
   Exit Function
DiaErr1:
   sProcName = "filllevel9"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

Public Function GetAcctLevel(sAcctRef As String) As Integer
   Dim RdoLvl As ADODB.Recordset
   sSql = "SELECT GLACCTREF,GLFSLEVEL FROM GlacTable " _
          & "WHERE GLACCTREF='" & sAcctRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLvl, ES_FORWARD)
   If bSqlRows Then
      GetAcctLevel = RdoLvl!GLFSLEVEL
   Else
      GetAcctLevel = 0
   End If
   Set RdoLvl = Nothing
End Function

Public Sub InsertTotalRow(rdoAct As ADODB.Recordset, iLev As Integer)
   With DbAct
      .AddNew
      !Act00 = iCurType
      !Act01 = iLev
      !Act02 = "" & Trim(rdoAct!GLACCTREF)
      !Act03 = String(iLev * 2, Chr(160)) & "" & Trim(rdoAct!GLACCTNO)
      !Act04 = String(iLev * 2, Chr(160)) & "TOTAL " & Trim(rdoAct!GLDESCR)
      !Act05 = rdoAct!GLINACTIVE
      iFsLevel = GetAcctLevel(Trim(rdoAct!GLACCTREF))
      !Act06 = iFsLevel
      .Update
   End With
End Sub

Public Function CashAccountBalance( _
                                   sAccount As String, _
                                   sDate As String) As Currency
   
   Dim rdoGL As ADODB.Recordset
   Dim rdoCsh As ADODB.Recordset
   Dim rdoCks As ADODB.Recordset
   
   'On Error GoTo modErr1
   
   sAccount = Compress(sAccount)
   
   ' GL (no sub journal XC CC CR ect...)
   sSql = "SELECT SUM(GjitTable.JIDEB - GjitTable.JICRD) FROM GjitTable " _
          & "INNER JOIN GjhdTable ON JINAME = GJNAME LEFT OUTER JOIN " _
          & "JrhdTable ON GJNAME = MJGLJRNL WHERE GJPOST <= '" & sDate & "' " _
          & "AND JIACCOUNT = '" & sAccount & "' AND MJGLJRNL IS NULL"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoGL)
   If bSqlRows Then
      With rdoGL
         If Not IsNull(.Fields(0)) Then
            CashAccountBalance = CCur(.Fields(0))
         End If
         .Cancel
      End With
   End If
   Set rdoGL = Nothing
   
   ' Cash Receipts
   sSql = "SELECT SUM(CACKAMT) FROM GjhdTable RIGHT OUTER JOIN " _
          & "CashTable INNER JOIN JritTable ON CACHECKNO = DCCHECKNO " _
          & "AND CACASHACCT = DCACCTNO AND CACUST = DCCUST ON GJNAME " _
          & "= DCHEAD WHERE CARCDATE <= '" & sDate & "' AND CACASHACCT ='" & sAccount & "'" _
          & " AND GJNAME IS NULL "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCsh)
   If bSqlRows Then
      With rdoCsh
         If Not IsNull(.Fields(0)) Then
            CashAccountBalance = CashAccountBalance _
                                 + CCur(.Fields(0))
         End If
         .Cancel
      End With
   End If
   Set rdoCsh = Nothing
   
   ' Checks
   sSql = "SELECT SUM(ChksTable.CHKAMOUNT) FROM ChksTable INNER JOIN " _
          & "JritTable ON CHKACCT = DCACCTNO AND CHKNUMBER = DCCHECKNO " _
          & "LEFT OUTER JOIN GjhdTable ON DCHEAD = GJNAME WHERE (CHKPOSTDATE " _
          & "<='" & sDate & "' AND CHKACCT = '" & sAccount & "' AND CHKVOIDDATE > '" _
          & sDate & "') OR (CHKPOSTDATE <='" & sDate & "' AND CHKACCT = '" _
          & sAccount & "' AND CHKVOIDDATE IS NULL)"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCks)
   If bSqlRows Then
      With rdoCks
         If Not IsNull(.Fields(0)) Then
            CashAccountBalance = CashAccountBalance _
                                 - CCur(.Fields(0))
         End If
         .Cancel
      End With
   End If
   Set rdoCks = Nothing
   Exit Function
modErr1:
   sProcName = "CashAcco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function
