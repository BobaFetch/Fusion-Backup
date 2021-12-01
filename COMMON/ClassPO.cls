VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type VendorTerms
   cNetDays As Currency
   cDDays As Currency
   cDiscount As Currency
   cProxDt As Currency
   cProxdue As Currency
   cbTaxable As Currency
   strFob As String
   strBcontact As String
   strShipTo As String
   strBuyer As String
End Type

Private Function GetNewPo(ByRef strNewPO As String)
   Dim RdoCmn As ADODB.Recordset
   Dim lOldPo As Long
   Static sOldLast As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT (COLASTPURCHASEORDER + 1) COLASTPURCHASEORDER From ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmn, ES_FORWARD)
   If bSqlRows Then
      If RdoCmn!COLASTPURCHASEORDER > 0 Then lOldPo = RdoCmn!COLASTPURCHASEORDER
   End If
   
   If lOldPo = 0 Then
      sSql = "SELECT MAX(PONUMBER) + 1 AS NEWPO FROM PohdTable "
      Set RdoCmn = clsADOCon.GetRecordSet(sSql, adOpenDynamic)
      If Not RdoCmn!NEWPO Then
         If RdoCmn!NEWPO > 0 Then lOldPo = RdoCmn!NEWPO
      Else
         strNewPO = "000000"
      End If
   End If
   
   If lOldPo = 0 Then
      strNewPO = "000000"
   Else
      strNewPO = Format(lOldPo, "000000")
   End If
   Set RdoCmn = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetNewPo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function CheckPo(lPoNum As Long) As Byte
   Dim RdoPon As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PONUMBER FROM PohdTable WHERE " _
          & "PONUMBER=" & Val(lPoNum) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPon)
   If bSqlRows Then
      CheckPo = True
   Else
      CheckPo = False
   End If
   Set RdoPon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "CheckPo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub GetVendorTerms(ByVal strVendRef As String, ByRef udtVenTerms As VendorTerms)
   Dim RdoTrm As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT VEREF,VENETDAYS,VEDDAYS,VEDISCOUNT," _
          & "VEPROXDT,VEPROXDUE,VEFOB,VEBNAME,VEBADR,VEBCITY," _
          & "VEBSTATE,VEBZIP,VEBCONTACT,VEBUYER FROM VndrTable WHERE " _
          & "VEREF='" & strVendRef & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTrm)
   If bSqlRows Then
      With RdoTrm
         On Error Resume Next
         udtVenTerms.cNetDays = 0 + !VENETDAYS
         udtVenTerms.cDDays = 0 + !VEDDAYS
         udtVenTerms.cDiscount = 0 + !VEDISCOUNT
         udtVenTerms.cProxDt = 0 + !VEPROXDT
         udtVenTerms.cProxdue = 0 + !VEPROXDUE
         udtVenTerms.strFob = "" & Trim(!VEFOB)
         udtVenTerms.strBcontact = "" & Trim(!VEBCONTACT)
         udtVenTerms.strBuyer = "" & Trim(!VEBUYER)
      End With
   End If
   Set RdoTrm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getvendorterms"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function AddNewPo(strVendor As String) As String
   Dim RdoNpo As ADODB.Recordset
   Dim strVendRef As String
   Dim strShipTo As String
   Dim strDate As String
   Dim strReqBy As String
   Dim udtVenTerms As VendorTerms
   Dim cbTaxable As Currency
   Dim strPoNum As String
   Dim bNewPo As Boolean
   Dim strService As String
   
   GetNewPo strPoNum
   If (strPoNum <> "") Then
      bNewPo = CheckPo(CLng(strPoNum))
      If bNewPo = True Then
         MsgBox "That Po Number Has Been Used.", vbInformation
         AddNewPo = ""
         Exit Function
      End If
   Else
      MsgBox "Couln't create new PO Number.", vbInformation
      AddNewPo = ""
      Exit Function
   End If
   
   GetShipTo strShipTo
   udtVenTerms.strShipTo = strShipTo
   strVendRef = Compress(strVendor)
   GetVendorTerms strVendRef, udtVenTerms
   strDate = Format(ES_SYSDATE, "mm/dd/yy")
   On Error GoTo DiaErr1
   
   strReqBy = sInitials
   strService = "0"
   cbTaxable = 1
   
   sSql = "INSERT INTO PohdTable(PONUMBER, POVENDOR, PODATE, POREQBY," _
            & "POSERVICE, PONETDAYS, PODDAYS, PODISCOUNT, POPROXDT, " _
            & "POPROXDUE, POFOB, POBCONTACT, POBUYER, POSHIPTO," _
            & "POTAXABLE) VALUES(" _
            & "'" & strPoNum & "','" & strVendRef & "','" & strDate & "'," _
            & "'" & strReqBy & "','" & strService & "','" & CStr(udtVenTerms.cNetDays) & "'," _
            & "'" & CStr(udtVenTerms.cDDays) & "','" & CStr(udtVenTerms.cDiscount) & "','" _
            & CStr(udtVenTerms.cProxDt) & "','" & CStr(udtVenTerms.cProxdue) & "','" _
            & CStr(udtVenTerms.strFob) & "','" & CStr(udtVenTerms.strBcontact) & "','" _
            & CStr(udtVenTerms.strBuyer) & "','" & CStr(udtVenTerms.strShipTo) & "'," _
            & CStr(cbTaxable) & ")"
   
   Debug.Print sSql
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "UPDATE ComnTable SET COLASTPURCHASEORDER=" & Val(strPoNum) & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "UPDATE ComnTable SET CURPONUMBER=" & Val(strPoNum) & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   Set RdoNpo = Nothing
   
   AddNewPo = strPoNum
   
   Exit Function
   
DiaErr1:
   sProcName = "addnewpo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Function AddPOItem(strPoNum As String, strPartNum As String, strQty As String, _
         strUnitPrice As String, strPartRqd As String, strActDate As String, strVendor As String)
   
   On Error GoTo DiaErr1
   
   Dim iNewItem As Integer
   Dim iNextItem As Integer
   Dim clsPart As ClassPart
   Dim cPrice As Currency
   Dim cQuantity As Currency
   Dim strDebitAcc As String
   Dim strCreditAcc As String
   
   GetLastItem strPoNum, iNextItem
   iNewItem = iNextItem
   
   ' get Part Unit cost
   Set clsPart = New ClassPart
   If (Not clsPart Is Nothing) Then
      'clsPart.PartNumber = strPartNum
      'strUnitPrice = clsPart.GetPartUnitPrice
      clsPart.GetPartAccounts strPartNum, False, strDebitAcc, strCreditAcc
   Else
      strDebitAcc = ""
   End If
   
   cPrice = Format(Val(strUnitPrice), ES_PurchasedDataFormat)
   cQuantity = Format(Val(strQty), ES_QuantityDataFormat)

   
   sSql = "INSERT INTO PoitTable (PINUMBER,PIITEM,PITYPE,PIPART,PIPDATE,PIPORIGDATE,PIVENDOR," _
          & "PIPQTY, PILOT, PIESTUNIT, PIONDOCK, PIUSER,PIENTERED,PIACCOUNT) " & vbCrLf _
          & "VALUES(" & strPoNum & "," _
          & iNewItem & ",14,'" & strPartNum & "','" _
          & strActDate & "','" & strPartRqd & "','" & strVendor & "','" _
          & CStr(cQuantity) & "',0,'" & CStr(cPrice) & "',0," _
          & "'" & sInitials & "','" & Format(ES_SYSDATE, "mm/dd/yy") & "','" & strDebitAcc & "')"
   
   Debug.Print sSql
   
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   
   Set clsPart = Nothing
   
   Exit Function
DiaErr1:
   sProcName = "AddPOItem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub GetShipTo(ByRef strShipTo As String)
   On Error GoTo DiaErr1
   GetCompany True
   strShipTo = Co.Name & vbCrLf _
             & Co.Addr(1) & vbCrLf _
             & Co.Addr(2) & vbCrLf _
             & Co.Addr(3) & vbCrLf _
             & Co.Addr(4)
   Exit Sub
   
DiaErr1:
   sProcName = "getshipto"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


'Private Sub FillReqBy()
'   On Error GoTo DiaErr1
'   sSql = "SELECT DISTINCT POREQBY FROM PohdTable WHERE PODATE> '" _
'          & Format(ES_SYSDATE - 360, "mm/dd/yy") & "' ORDER BY POREQBY"
'   LoadComboBox cmbReq, -1
'   On Error Resume Next
'   optReq.Value = Val(GetSetting("Esi2000", "EsiProd", "PRe01aRb", optReq))
'   If optReq.Value = vbChecked Then cmbReq = GetSetting("Esi2000", "EsiProd", "PRe01a", cmbReq)
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillreqby"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'

'Private Sub LoadVendorAddress()
'    Dim RdoVndr As ADODB.Recordset
'    Dim sAddress As String
'    sAddress = ""
'
'    sSql = "SELECT VEBADR, VEBCITY, VEBSTATE, VEBZIP,VEBCONTACT, VEBPHONE, VEBEXT FROM VndrTable WHERE VEREF = '" & Compress(cmbVnd) & "'"
'    bSqlRows = clsADOCon.GetRecordSet(sSql,RdoVndr, ES_FORWARD)
'    If bSqlRows Then
'        If Len(Trim("" & RdoVndr!VEBCONTACT)) > 0 Then sAddress = sAddress & Trim(RdoVndr!VEBCONTACT) & vbCrLf
'        If Len(Trim("" & RdoVndr!VEBADR)) > 0 Then sAddress = sAddress & Trim(RdoVndr!VEBADR) & vbCrLf
'        If Len(Trim("" & RdoVndr!VEBCITY)) > 0 Then sAddress = sAddress & Trim(RdoVndr!VEBCITY) & " ," & Trim(RdoVndr!VEBSTATE) & " " & Trim(RdoVndr!VEBZIP) & vbCrLf
'        If Len(Trim("" & RdoVndr!VEBPHONE)) > 0 And Trim("" & RdoVndr!VEBPHONE) <> "___-___-____" Then sAddress = sAddress & "" & Trim(RdoVndr!VEBPHONE)
'        If Len(Trim("" & Trim(RdoVndr!VEBEXT))) > 0 And Val("" & Trim(RdoVndr!VEBEXT)) > 0 Then sAddress = sAddress & " Ext: " & Trim(RdoVndr!VEBEXT)
'    End If
'    Set RdoVndr = Nothing
'    If Len(sAddress) > 0 Then
'        cmbVnd.ToolTipText = sAddress
''        FusionToolTip.ToolText(cmbVnd) = sAddress
'    Else
'        cmbVnd.ToolTipText = "Enter the Vendor"
'    End If
'End Sub
'
'
Public Sub GetLastItem(ByVal strPoNum As String, ByRef iNextItem As Integer)
   Dim RdoNxt As ADODB.Recordset
   'On Error Resume Next
   sSql = "SELECT MAX(PIITEM) FROM PoitTable " _
          & "WHERE PINUMBER=" & strPoNum & " "
   Set RdoNxt = clsADOCon.GetRecordSet(sSql, adOpenDynamic)
   If Not IsNull(RdoNxt.Fields(0)) Then
      If Not IsNull(RdoNxt.Fields(0)) Then
         If Val(RdoNxt.Fields(0)) > 0 Then iNextItem = RdoNxt.Fields(0) + 1
      Else
         iNextItem = 1
      End If
   Else
      iNextItem = 1
   End If
   Set RdoNxt = Nothing
   
End Sub


