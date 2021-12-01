Attribute VB_Name = "EDIOutFiles"
Option Explicit

Dim arrValue() As Variant
Dim arrFieldName() As Variant

Public Function CreateASNOut(strFilePath As String, strFileName As String, Optional strInDate As String = "")

   Dim strFullpath As String
   Dim strDate As String
   Dim strCust As String
   
   'strFileName = txtEdiFilePath.Text

   'strFilePath = "C:\Development\FusionCode\EDIFiles\Testing\"
   'strFileName = "ASNOUT.EDI"
   strFullpath = strFilePath & strFileName
   
   Dim nFileNum As Integer, lLineCount As Long
   Dim strBlank As String
   
   If (strInDate = "") Then
      strDate = Format(Now, "mm/dd/yy")
   Else
      strDate = strInDate
   End If
   strCust = ""
   

   If (Trim(strFileName) <> "") Then
      ' Open the file
      nFileNum = FreeFile
      Open strFullpath For Output As nFileNum
      
      If EOF(nFileNum) Then
         ASNOutFile nFileNum, strDate
      End If
      ' Close the file
      Close nFileNum
   End If


End Function



Function ASNOutFile(nFileNum As Integer, strDate As String) As Integer
   
   Dim rdoPS As rdoResultset
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
   Dim strFileName As String
   Dim strCust As String
   Dim strPONumber As String
   Dim strPartNum As String
   Dim strPiPartRef As String
   Dim strPSNum As String
   Dim strQty As String
   Dim strCarton As String
   Dim strContainer As String
   Dim strPrevContainer As String
   Dim strLoadNum As String
   Dim strPSVia As String
   Dim bPartFound As Boolean
   Dim bIncRow As Boolean
   Dim strPullNum As String
   Dim strBinNum As String
   Dim strShipNo As String
   Dim strGrossWt As String
   Dim strCarrierNum As String
   Dim iItem As Integer
   Dim bSelected As Boolean
   
   sSql = "SELECT DISTINCT PSNUMBER, PSCONTAINER, PSCUST, PSSHIPNO, PSNUMBER, ISNULL(PSCARTON, '') PSCARTON," _
            & "ISNULL(PSGROSSLBS, '0.00') PSGROSSLBS,ISNULL(PSCARRIERNUM, '') PSCARRIERNUM, " _
            & " PSLOADNO, PSVIA, SOPO,PIQTY , PIPART, PARTNUM, ISNULL(PULLNUM, '') PULLNUM, ISNULL(BINNUM, '') BINNUM " _
         & " From PshdTable, psitTable, sohdTable, SoitTable, Parttable " _
         & " WHERE PshdTable.PSDATE = '" & strDate & "'" _
          & " AND PSNUMBER = PIPACKSLIP" _
          & " AND SONUMBER = ITSO" _
          & " AND ITPSNUMBER = ITPSNUMBER" _
          & " AND SoitTable.ITSO = PsitTable.PISONUMBER" _
          & " AND SoitTable.ITNUMBER = PsitTable.PISOITEM" _
          & " AND SoitTable.ITREV = PsitTable.PISOREV" _
          & " AND PARTREF = PIPART" _
          & " AND PshdTable.PSCUST IN " _
          & " (SELECT DISTINCT a.CUREF " _
          & "     FROM ASNInfoTable a, custtable b WHERE " _
          & "        A.CUREF = b.CUREF AND TRUCKPLANT = 1)" _
          & " ORDER BY PSSHIPNO"

          '& " AND PshdTable.PSINVOICE = 0 "
          'PshdTable.PSCUST LIKE '" & strCust & "%' AND
          
   Debug.Print sSql
   
   bSqlRows = GetDataSet(rdoPS, rdOpenStatic)
   
   strPrevContainer = ""
   bSelected = False
   If bSqlRows Then
      
      With rdoPS
      While Not .EOF
         
         strPSNum = Trim(!PsNumber)
         strContainer = Trim(!PSCONTAINER)
         strCust = Trim(!PSCUST)
         strShipNo = "" & Trim(!PSSHIPNO)
         strPONumber = Trim(!PsNumber)
         strPartNum = Trim(!PARTNUM)
         strPiPartRef = Trim(!PIPART)
         strCarton = Trim(!PSCARTON)
         strLoadNum = Trim(!PSLOADNO)
         strPSVia = Trim(!PSVIA)
         strPONumber = Trim(!SOPO)
         strQty = Trim(!PIQTY)
         strPullNum = Trim(!PULLNUM)
         strBinNum = Trim(!BINNUM)
         strGrossWt = Trim(!PSGROSSLBS)
         strCarrierNum = Trim(!PSCARRIERNUM)
         
            
         If (strContainer <> "") Then
   
            If ((strPrevContainer = "") Or (strPrevContainer <> strContainer)) Then
               
               Dim strBusPartner As String
               Dim strBusDetail As String
               Dim strBuyerCode As String
               GetBuyerInfo strCust, strBusPartner, strBusDetail, strBuyerCode
                              
               ' Add Header detail
               Dim strHeader As String
               CreateASNHeader strCust, strContainer, strCarton, strGrossWt, _
                     strLoadNum, strCarrierNum, strBusPartner, strBusDetail, strDate, strHeader
               
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strHeader
                  Debug.Print strHeader
               End If
               
               ' Add CD
               Dim strHeadCD As String
               Dim iTotItems As Integer
               iTotItems = TotalPs(strContainer, strCust, strDate)
               
               CreateCD strContainer, strLoadNum, iTotItems, strHeadCD
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strHeadCD
                  Debug.Print strHeadCD
               End If
               
               ' Add H2
               Dim strHeader2 As String
               CreateHeader2 strContainer, strPSVia, strBuyerCode, strHeader2
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strHeader2
                  Debug.Print strHeader2
               End If
               
               
               ' Add
               Dim strHeadR1 As String
               CreateR1 strContainer, strCarrierNum, strHeadR1
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strHeadR1
                  Debug.Print strHeadR1
               End If
               
               ' Add Shipping info
               Dim strN1 As String
               Dim strN2 As String
               CreateShipInfo strCust, strContainer, strBuyerCode, strN1, strN2
               
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strN1
                  Print #nFileNum, strN2
                  Debug.Print strN1
                  Debug.Print strN2
               End If
               
               strPrevContainer = strContainer
            End If
            
            ' Not add the Details
            Dim strDT As String
            CreateASNDetail strContainer, strPartNum, strQty, strPONumber, _
               strPSNum, strPullNum, strDT
            
            If EOF(nFileNum) Then
               Print #nFileNum, strDT
               Debug.Print strDT
            End If
            ' If any selcted set the dirty flag to true
            bSelected = True
            
         End If
         
         .MoveNext
      Wend
      .Close
      End With
   
      If (bSelected = True) Then
         'MsgBox "ASN File created.", vbExclamation, Caption
      End If
      
   End If

   MouseCursor ccArrow
   Set rdoPS = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GenerateASNFile"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Function

Private Function GetBuyerInfo(ByVal strCust As String, _
                  ByRef strBusPartner As String, ByRef strBusDetail As String, _
                  ByRef strBuyerCode As String)

   On Error GoTo ModErr1
   Dim RdoBuy As rdoResultset
   If Trim(strCust) <> "" Then
      
      sSql = "SELECT SHPTOIDCODE, SHPTOCODEQUAL, SHPFRMIDCODE, " _
               & " SHPFRMCODEQUAL , SHPREF, SHPDETAIL, " _
               & " SHPADDRS, BUYERCODE FROM ASNInfoTable " _
               & "WHERE CUREF = '" & strCust & "'"

      bSqlRows = GetDataSet(RdoBuy, ES_FORWARD)
      If bSqlRows Then
         With RdoBuy
            strBusPartner = "" & Trim(!SHPREF)
            strBusDetail = "" & Trim(!SHPDETAIL)
            strBuyerCode = "" & Trim(!BUYERCODE)
            ClearResultSet RdoBuy
            
         End With
      End If
   End If
   Set RdoBuy = Nothing
   Exit Function
   
ModErr1:
   sProcName = "GetBuyerInfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Function

Public Function CreateInvoiceEDIFile(strFilePath As String, strFileName As String, Optional strInDate As String = "")

   Dim strFullpath As String
   
   Dim strDate As String
   Dim strCust As String
   
   
   'strFilePath = "C:\Development\FusionCode\EDIFiles\Testing\"
   'strFileName = "INVOUT.EDI"
   strFullpath = strFilePath & strFileName
   
   Dim nFileNum As Integer, lLineCount As Long
   Dim strBlank As String
   
   
   If (strInDate = "") Then
      strDate = Format(Now, "mm/dd/yy")
   Else
      strDate = strInDate
   End If

   If (Trim(strFileName) <> "") Then
      ' Open the file
      nFileNum = FreeFile
      Open strFullpath For Output As nFileNum
      
      If EOF(nFileNum) Then
         GenerateINVFile nFileNum, strDate
      End If
      ' Close the file
      Close nFileNum
   End If
End Function



Function GenerateINVFile(ByVal nFileNum As Integer, ByVal strDate As String) As Integer

    Dim rdoPS As rdoResultset

    MouseCursor (ccHourglass)
    On Error GoTo DiaErr1

    Dim strPONumber As String
    Dim strPartNum As String
    Dim strCust As String
    Dim strPSNum As String
    Dim strInvNum  As String
    Dim strQty As String
    Dim strInvDate As String
    Dim strSODate As String
    Dim strContainer As String
    Dim strUnitPrice As String
    Dim strPAUnit As String
    Dim bPartFound, bIncRow As Boolean
    Dim strInvTot As String
    Dim strBook As String
    Dim iList As Integer

   sSql = "SELECT PIPART, PARTNUM, INVNO, PSNUMBER, INVTOTAL, INVDATE, INVCUST, SOPO, " _
             & "SODATE, PISELLPRICE,PIQTY , PSCONTAINER, PSSHIPNO, ITDOLLARS,PAPRICE, PAUNITS" _
             & "  FROM cihdTable, sohdtable, soitTable, pshdTable, psittable, Parttable" _
             & "  WHERE INVDATE = '" & strDate & "'" _
             & "    AND ITPSNUMBER = PSNUMBER" _
             & "    AND ITSO = SONUMBER" _
             & "    AND PSINVOICE = INVNO" _
             & "    AND PSNUMBER = PIPACKSLIP" _
             & "    AND PARTREF = PIPART" _
             & "    AND INVCUST IN " _
             & "           (SELECT DISTINCT a.CUREF FROM " _
             & "                 ASNInfoTable a, custtable b WHERE " _
             & "              A.CUREF = b.CUREF AND PACCARDPART = 1)"

             '& "    AND INVCUST LIKE '" & strCust & "%'"
    Debug.Print (sSql)

    bSqlRows = GetDataSet(rdoPS, rdOpenStatic)

    If bSqlRows Then
        With rdoPS
            While Not .EOF

                strInvNum = Trim(!invno)
                strPSNum = Trim(!PsNumber)
                strContainer = Trim(!PSCONTAINER)
                strPONumber = Trim(!SOPO)
                strPartNum = Trim(!PARTNUM)
                strQty = Trim(!PIQTY)
                strInvTot = Format(Trim(!INVTOTAL), ES_MoneyFormat)
                strInvDate = Trim(!INVDATE)
                strSODate = Trim(!SODATE)
                strCust = Trim(!INVCUST)
                'strUnitPrice = Trim(!PAPRICE)
                strUnitPrice = Format(Trim(!ITDOLLARS), ES_MoneyFormat)
                'strBook = "PACPARTS"
                'GetBookPrice strPartNum, strBook, strUnitPrice
                
                strPAUnit = Trim(!PAUNITS)

               ' Add Header detail
               Dim strHeader As String
               CreateInvHeader strCust, strInvNum, strPONumber, strSODate, _
                     strInvDate, strContainer, strHeader
               
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strHeader
                  Debug.Print (strHeader)
               End If
               
               ' Add Detail
               Dim strInvDetail As String
               CreateInvDetail strInvNum, strQty, strPAUnit, strUnitPrice, strPartNum, strInvDetail
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strInvDetail
                  Debug.Print (strInvDetail)
               End If
               
               ' Add So
               Dim strSODetail As String
               CreateSODetail strInvNum, strQty, strInvTot, strSODetail
               ' Read the contents of the file
               If EOF(nFileNum) Then
                  Print #nFileNum, strSODetail
                  Debug.Print (strSODetail)
               End If

                .MoveNext
            Wend
            .Close
         End With
         ClearResultSet rdoPS
    End If

   MouseCursor (ccArrow)
   
   Set rdoPS = Nothing
   Exit Function

DiaErr1:
    sProcName = "GenerateINVFile"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description

End Function


Function AddEDIFieldsLength(ByVal strType As String, ImpType As String)
   Dim iIndex As Integer
   Dim j As Integer
   Dim RdoEdi As rdoResultset
   Dim strValue As String
   Dim iTotLen As Integer
   Dim iTotalItems As Integer
   Dim iNumChar As Integer
   Dim strFields As String
   Dim strFldVal As String
   
   On Error GoTo DiaErr1
   
   If (strType <> "") Then
      sSql = "SELECT FIELDNAME,NUMCHARS FROM ProEdiFormat WHERE " _
             & "HEADER = '" & strType & "' AND IMPORTTYPE = '" & ImpType & "' ORDER BY FORATORDER"
      
      bSqlRows = GetDataSet(RdoEdi, rdOpenStatic)
      ReDim arrValue(0 To RdoEdi.RowCount + 1)
      ReDim arrFieldName(0 To RdoEdi.RowCount + 1)
      If bSqlRows Then
         With RdoEdi
         iTotalItems = 0
         While Not .EOF
            iNumChar = !NUMCHARS
            arrValue(iTotalItems) = CStr(iNumChar)
            arrFieldName(iTotalItems) = !FieldName
            iTotalItems = iTotalItems + 1
            .MoveNext
         Wend
         .Close
         End With
         ClearResultSet RdoEdi
      End If
      
   End If
   Set RdoEdi = Nothing
   
   Exit Function
DiaErr1:
   sProcName = "DecodeEdiFormat"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description

End Function


Private Function CreateASNHeader(strCust As String, strContainer As String, strCarton As String, _
               strGrossWt As String, strLoadNum As String, strCarrierNum As String, _
               strBusPartner As String, strBusDetail As String, strDate As String, _
               ByRef strHeader As String)
   On Error GoTo DiaErr1
      
   Dim strHeader1 As String
   Dim strBlank As String
   Dim strUnit As String
   Dim strTime As String
   Dim strDateConv As String
   
   strHeader = "H1"
   strUnit = "LB"
   ' Get Fields Chars
   'strContainer = "8028"
   'strGrossWt = "1987"
   'strCarton = "1234"
   strBlank = ""
   'strBusPartner = "PACCAR"
   'strBusDetail = "CH"
   
   ' get the Field lenght
   AddEDIFieldsLength "H1", "ASN"
   
   strContainer = FormatEDIString(strContainer, arrValue(0), "0")
   strBusPartner = FormatEDIString(strBusPartner, arrValue(1), "@")
   strBusDetail = strBusDetail & FormatEDIString(" ", (arrValue(2) - Len(strBusDetail)), "@")
   strHeader = strHeader & strContainer & strBusPartner & strBusDetail
   
   strConverDate strDate, strDateConv
   
   strDateConv = FormatEDIString(strDateConv, arrValue(3), "0")
   strTime = "170000"
   strTime = FormatEDIString(strTime, arrValue(4), "0")
   strGrossWt = FormatEDIString(strGrossWt, arrValue(5), "0")
   strUnit = FormatEDIString(strUnit, arrValue(6), "0")
   
   strHeader = strHeader & strDateConv & strTime & strGrossWt & strUnit
   
   
   Exit Function
   
DiaErr1:
   sProcName = "SetPartHeader"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description

End Function

Private Function CreateCD(strContainer As String, strLoadNum As String, iTotItems As Integer, _
                  ByRef strHeadCD As String)
   On Error GoTo DiaErr1
      
   Dim strTotItem As String
   
   AddEDIFieldsLength "CD", "ASN"
   ' Get total Items
   strLoadNum = FormatEDIString(CStr(strLoadNum), arrValue(1), "@")
   strTotItem = FormatEDIString(CStr(iTotItems), arrValue(2), "0")
   
   strHeadCD = "CD" & strContainer & strLoadNum & strTotItem
   
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateCD"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description

End Function


Private Function CreateHeader2(strContainer As String, strPSVia As String, _
         strBuyerCode As String, ByRef strHeader2 As String)
   On Error GoTo DiaErr1
      
   Dim strTransMethod As String
   Dim strEquipDesc As String
   
   AddEDIFieldsLength "H2", "ASN"
   ' Get total Items
   strTransMethod = "M"
   strEquipDesc = "TL"
   
   strPSVia = strPSVia & FormatEDIString(" ", (arrValue(2) - Len(strPSVia)), "@")
   strTransMethod = strTransMethod & FormatEDIString(" ", (arrValue(3) - Len(strTransMethod)), "@")
   strEquipDesc = strEquipDesc & FormatEDIString(" ", (arrValue(4) - Len(strTransMethod)), "@")
   strHeader2 = "H2" & strContainer & strBuyerCode & strPSVia & strTransMethod & strEquipDesc
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateHeader2"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description

End Function

Private Function CreateR1(strContainer As String, strCarrierNum As String, _
                  ByRef strHeadR1 As String)
   On Error GoTo DiaErr1
      
   Dim strCarrType As String
   Dim strBuyerCode As String
   
   AddEDIFieldsLength "R1", "ASN"
   
   strCarrType = FormatEDIString("CN", arrValue(1), "@")
   strCarrierNum = FormatEDIString(strCarrierNum, arrValue(2), "0")
   strHeadR1 = "R1" & strContainer & strCarrType & strCarrierNum
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateR1"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Private Function CreateShipInfo(strCust As String, strContainer As String, strBuyerCode As String, _
               ByRef strN1 As String, ByRef strN2 As String)
   On Error GoTo DiaErr1
      
   Dim strFromAddrs As String
   Dim strFrmVndID As String
   Dim strFrmVnd As String
   Dim strShpFrom As String
   Dim strToAddrs As String
   Dim strToVndID As String
   Dim strToVnd As String
   Dim strShpTo As String
   
      
   GetShipInfo strCust, strFrmVnd, strFrmVndID, strFromAddrs, _
               strToVnd, strToVndID, strToAddrs
      
   AddEDIFieldsLength "N1", "ASN"
   ' Get total Items
   strFromAddrs = strFromAddrs & FormatEDIString(" ", (arrValue(2) - Len(strFromAddrs)), "@")
   strFrmVndID = FormatEDIString(strFrmVndID, arrValue(3), "0")
   'strFrmVnd = Format(strFrmVnd, String(arrValue(4), "0"))
   strN1 = "N1" & strContainer & "SF" & strFromAddrs & strFrmVndID & strFrmVnd
   
   AddEDIFieldsLength "N1", "ASN"
   ' Get total Items
   strToAddrs = strToAddrs & FormatEDIString(" ", (arrValue(2) - Len(strToAddrs)), "@")
   strToVndID = FormatEDIString(strToVndID, arrValue(3), "0")
   strN2 = "N1" & strContainer & "ST" & strToAddrs & strToVndID & strToVnd
   
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateShipInfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Private Function CreateInvHeader(strCust As String, strInvNum As String, strPONumber As String, _
               strSODate As String, strInvDate As String, strContainer As String, ByRef strHeader As String)
   On Error GoTo DiaErr1
      
   Dim strBlank As String
   Dim strUnit As String
   Dim strTime As String
   Dim strInvDtConv As String
   Dim strSODateConv As String
   Dim strBusPartner As String
   Dim strBusDetail As String
   Dim strBuyerCode As String
   
   strHeader = "H"
   strUnit = "EA"
   ' Get Fields Chars
   'strContainer = "8028"
   'strGrossWt = "1987"
   'strCarton = "1234"
   strBlank = ""
   'strBusPartner = "PACCAR"
   'strBusDetail = "DE"
   
   GetBuyerInfo strCust, strBusPartner, strBusDetail, strBuyerCode
   
   ' get the Field length
   AddEDIFieldsLength "H", "INV"
   
   strInvNum = FormatEDIString(strInvNum, arrValue(0), "0")
   strConverDate strInvDate, strInvDtConv
   strInvDtConv = FormatEDIString(strInvDtConv, arrValue(1), "0")
   strBusPartner = FormatEDIString(strBusPartner, arrValue(2), "@")
   strBusDetail = strBusDetail & FormatEDIString(" ", (arrValue(3) - Len(strBusDetail)), "@")
   strPONumber = strPONumber & FormatEDIString(" ", (arrValue(4) - Len(strPONumber)), "@")
   
   strConverDate strInvDate, strInvDtConv
   strInvDtConv = FormatEDIString(strInvDtConv, arrValue(5), "0")
   
   strConverDate strSODate, strSODateConv
   strSODateConv = FormatEDIString(strSODateConv, arrValue(6), "0")
   
   strContainer = FormatEDIString(strContainer, arrValue(7), "0")
   
   ' MM not using the SODATE 3/29/2012
   strHeader = strHeader & strInvNum & strInvDtConv & strBusPartner & strBusDetail _
                  & strPONumber & strInvDtConv & strInvDtConv & strContainer

'   strHeader = strHeader & strInvNum & strInvDtConv & strBusPartner & strBusDetail _
'                  & strPONumber & strInvDtConv & strSODateConv & strContainer
   Exit Function
   
DiaErr1:
   sProcName = "SetPartHeader"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function


Private Function CreateASNDetail(strContainer As String, strPartNum As String, _
               strQty As String, strPONumber As String, strPS As String, _
                  strPullNum As String, ByRef strDT As String)
   On Error GoTo DiaErr1
      
         Dim strVendPartNum As String
         
         AddEDIFieldsLength "DT", "ASN"
         ' Get total Items
         strVendPartNum = Mid(strPartNum, 1, arrValue(2))
         strPartNum = strPartNum & FormatEDIString(" ", (arrValue(1) - Len(strPartNum)), "@")
         strQty = FormatEDIString(strQty, arrValue(3), "0")
         strPONumber = strPONumber & FormatEDIString(" ", (arrValue(4) - Len(strPONumber)), "@")
         strPS = FormatEDIString(Mid(strPS, 3, (Len(strPS) - 2)), arrValue(5), "0")
         strPullNum = strPullNum ' Shows as 9 characters 'MM & FormatEDIString(strPullNum, arrValue(6), "0")
   
         strDT = "DT" & strContainer & strPartNum & strVendPartNum & strQty & strPONumber & strPS & strPullNum

   Exit Function
   
DiaErr1:
   sProcName = "SetPartHeader"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description

End Function

Private Function GetShipInfo(ByVal strCust As String, ByRef strFrmVnd As String, _
                  ByRef strFrmVndID As String, ByRef strFromAddrs As String, _
                  ByRef strToVnd As String, ByRef strToVndID As String, ByRef strToAddrs As String)

   On Error GoTo ModErr1
   Dim RdoBuy As rdoResultset
   If Trim(strCust) <> "" Then
      
      sSql = "SELECT SHPTOIDCODE, SHPTOCODEQUAL, SHPFRMIDCODE, " _
               & " SHPFRMCODEQUAL , SHPREF, SHPDETAIL, " _
               & " SHPADDRS FROM ASNInfoTable " _
               & "WHERE CUREF = '" & strCust & "'"

      bSqlRows = GetDataSet(RdoBuy, ES_FORWARD)
      If bSqlRows Then
         With RdoBuy
            strFrmVnd = "" & Trim(!SHPFRMIDCODE)
            strFrmVndID = "" & Trim(!SHPFRMCODEQUAL)
            strFromAddrs = "U.S. CASTINGS LLC."
            strToVnd = "" & Trim(!SHPTOIDCODE)
            strToVndID = "" & Trim(!SHPTOCODEQUAL)
            strToAddrs = "" & Trim(!SHPADDRS)
            
            ClearResultSet RdoBuy
            
         End With
      End If
   End If
   Set RdoBuy = Nothing
   Exit Function
   
ModErr1:
   sProcName = "GetBuyerInfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Function

Private Function strConverDate(strDate As String, ByRef strDateConv As String)
   strDateConv = Format(CDate(strDate), "yymmdd")
End Function

Private Function TotalPs(strContainer As String, strCust As String, strDate As String) As Integer
   On Error GoTo ModErr1
   
   Dim iMaxPSSel As Integer
   Dim rdoPS As rdoResultset
   iMaxPSSel = 0
   If Trim(strContainer) <> "" Then
      
      sSql = "SELECT MAX(PSCONTAINER) maxPS" _
         & " From PshdTable, psitTable, sohdTable, SoitTable " _
         & " WHERE PshdTable.PSCUST = '" & strCust & "' AND PshdTable.PSDATE = '" & strDate & "'" _
          & " AND PSCONTAINER = '" & strContainer & "'" _
          & " AND PSNUMBER = PIPACKSLIP" _
          & " AND SONUMBER = ITSO" _
          & " AND ITPSNUMBER = ITPSNUMBER" _
          & " AND SoitTable.ITSO = PsitTable.PISONUMBER" _
          & " AND SoitTable.ITNUMBER = PsitTable.PISOITEM" _
          & " AND SoitTable.ITREV = PsitTable.PISOREV" _
          & " AND PshdTable.PSINVOICE = 0 " _

      bSqlRows = GetDataSet(rdoPS, ES_FORWARD)
      If bSqlRows Then
         With rdoPS
            iMaxPSSel = "" & Trim(!maxPS)
            ClearResultSet rdoPS
         End With
      End If
      Set rdoPS = Nothing
   End If
   
   TotalPs = iMaxPSSel
   Exit Function
   
ModErr1:
   sProcName = "GetBuyerInfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function
Private Function CreateInvDetail(strInvNum As String, strQty As String, strPAUnits As String, _
                  strUnitPrice As String, strPartNum As String, ByRef strInvDetail As String)
   On Error GoTo DiaErr1
      
   Dim strTotItem As String
   Dim strBlank As String
   Dim strBlank1 As String
   Dim strVendPartNum As String
   strBlank = ""
   strBlank1 = ""
   
   AddEDIFieldsLength "D", "INV"
   ' Get total Items
   strInvNum = FormatEDIString(CStr(strInvNum), arrValue(0), "0")
   strQty = FormatEDIString(strQty, arrValue(1), "0")
   strPAUnits = FormatEDIString(strPAUnits, arrValue(2), "@")
   strBlank = FormatEDIString(" ", arrValue(3), "@")
   strUnitPrice = FormatEDIString(Replace(strUnitPrice, ".", ""), arrValue(4), "0")
   
   strPartNum = strPartNum & FormatEDIString(" ", (arrValue(5) - Len(strPartNum)), "@")
   strBlank1 = FormatEDIString(" ", arrValue(6), "@")
   strVendPartNum = Mid(strPartNum, 1, arrValue(7))

   strInvDetail = "D" & strInvNum & strQty & strPAUnits & strBlank & _
            strUnitPrice & strPartNum & strBlank1 & strVendPartNum
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateCD"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function



Private Function CreateSODetail(strInvNum As String, strQty As String, _
         strInvTot As String, ByRef strSODetail As String)
   On Error GoTo DiaErr1
      
   Dim strInvTot1 As String
   strInvTot1 = strInvTot
   
   AddEDIFieldsLength "S", "INV"
   ' Get total Items
   
   strInvNum = FormatEDIString(CStr(strInvNum), arrValue(0), "0")
   strQty = FormatEDIString(strQty, arrValue(1), "0")
   strInvTot = FormatEDIString(Replace(strInvTot, ".", ""), arrValue(2), "0")
   strInvTot1 = FormatEDIString(Replace(strInvTot1, ".", ""), arrValue(3), "0")
   
   strSODetail = "S" & strInvNum & strQty & strInvTot & strInvTot1
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateSODetail"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function


Private Function FormatEDIString(strInput As String, iLen As Variant, strPad As String) As String
   
   If (iLen > 0) Then
      If (strPad = "0") Then
         strInput = Format(strInput, String(iLen, "0"))
      ElseIf (strPad = "@") Then
         strInput = Format(strInput, String(iLen, "@"))
      End If
   Else
      strInput = ""
   End If

   FormatEDIString = strInput
   
End Function


Private Function GetBookPrice(strPart As String, strBook As String, ByRef strPrice As String)
   Dim RdoBok As rdoResultset
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT PARTREF,PARTNUM,PBIREF,PBIPARTREF,PBIPRICE " _
          & "FROM PartTable,PbitTable WHERE (PARTREF=PBIPARTREF) AND " _
          & "(PBIREF='" & strBook & "') AND (PARTREF='" & Compress(strPart) & "')"
   bSqlRows = GetDataSet(RdoBok, ES_FORWARD)
   If bSqlRows Then
      With RdoBok
         strPrice = Format(!PBIPRICE, ES_MoneyFormat)
         ClearResultSet RdoBok
      End With
   Else
      strPrice = Format(0, ES_MoneyFormat)
   End If
   
   Set RdoBok = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getbookpr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function


