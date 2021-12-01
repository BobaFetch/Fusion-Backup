Attribute VB_Name = "ImportEDISO"
Option Explicit

Dim bCutOff As Byte
Dim bOnLoad As Byte
Dim bUnload As Boolean
Dim bOptionSel As Boolean

Dim sLastPrefix As String
Dim sNewsonumber As String
Dim sCust As String
Dim sStName As String
Dim sStAdr As String
Dim sContact As String
Dim sConIntPhone As String
Dim sConPhone As String
Dim sConIntFax As String
Dim sConFax As String
Dim sConExt As String
Dim sDivision As String
Dim sOldSoNumber As String
Dim sRegion As String
Dim sSterms As String
Dim sVia As String
Dim sFob As String
Dim sSlsMan As String
Dim sTaxExempt As String
Dim strFilePath As String
Dim strEDIFormat  As String

Dim iDays As Integer
Dim iFrtDays As Integer
Dim iNetDays As Integer

Dim cDiscount As Currency

Dim arrValue() As Variant
Dim arrFieldName() As Variant

Dim strPartNum As String
Dim strPartCnt As String
Dim strPAUnit As String
Dim strPartInfo As String
Dim strPullNum As String
Dim strBinNum As String
Dim strECNum As String
Dim strBldStation As String

Dim strPartNumFld As String
Dim strPartCntFld As String
Dim strPAUnitFld As String
Dim strPartInfoFld As String
Dim strPullNumFld As String
Dim strBinNumFld As String
Dim strECNumFld As String
Dim strBldStationFld As String

Public Function ImpEDISalesOrder(ByVal strFilePath As String, ByVal strFileName As String) As Boolean
   
   Dim strEDIDataType As String
   Dim strFullpath As String
   
   On Error GoTo DiaErr1
   
   If ((Trim(strFilePath) <> "") And (Trim(strFileName) <> "")) Then

      strFullpath = strFilePath & strFileName
      strEDIDataType = CheckImportType(strFileName, strFullpath)
      
      
      If (strEDIDataType = "850_EDI") Then
         
         DeleteOldData ("Inhd830_EDI")
         DeleteOldData ("Init830_EDI")
         DeleteOldData ("Inhd830_850EDI")
         DeleteOldData ("Init830_850EDI")
         
         ImportEDIFile strFullpath, strEDIDataType
         Create850SO
      
      ElseIf (strEDIDataType = "862_EDI") Then
         
         DeleteOldData ("Inhd862_EDI")
         DeleteOldData ("Init862_EDI")
         
         ImportEDIFile strFullpath, strEDIDataType
         Create862SalesOrder
         'Fill862Grid (CStr(cmbCst))
      
      End If
   End If
   
   Exit Function
   
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
End Function


Public Function ImportEDIFile(ByVal strFilePath As String, strEDIDataType As String) As Integer
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
   ' Read the content if the text file.
   Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
   Dim lngPos As Integer
   Dim bFound As Boolean
   Dim gstrSenderCode As String
' Get a free file number
   nFileNum = FreeFile
   
   Open strFilePath For Input As nFileNum
   ' Read the contents of the file
   bFound = False
   Do While Not EOF(nFileNum)
      Line Input #nFileNum, sNextLine
      Debug.Print sNextLine
      
      If (strEDIDataType = "850_EDI") Then
         Decode850EdiFormat sNextLine
      ElseIf (strEDIDataType = "862_EDI") Then
         Decode862EdiFormat sNextLine
      End If
   Loop
   Close nFileNum

   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   If (nFileNum > 0) Then
      Close nFileNum
   End If
   sProcName = "ImportEDIFile"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Public Function Decode850EdiFormat(ByVal strEdiData As String)
   Dim iIndex As Integer
   Dim j As Integer
   Dim RdoEdi As adodb.Recordset
   Dim strValue As String
   Dim strType As String
   Dim iTotLen As Integer
   Dim iTotalItems As Integer
   Dim iNumChar As Integer
   Dim strFields As String
   Dim strFldVal As String
   Dim strTabName As String
   
   On Error GoTo DiaErr1
   
   If (strEdiData <> "") Then
      iIndex = 2
      iTotLen = Len(strEdiData)
      strType = Mid(strEdiData, 1, iIndex)
      iIndex = iIndex + 1
      sSql = "SELECT FIELDNAME,NUMCHARS FROM ProEdiFormat WHERE " _
             & "HEADER = '" & strType & "' AND IMPORTTYPE = 'PO' ORDER BY FORATORDER"
      
      bSqlRows = clsAdoCon.GetDataSet(sSql, RdoEdi, rdOpenStatic)
      ReDim arrValue(0 To RdoEdi.RecordCount + 1)
      ReDim arrFieldName(0 To RdoEdi.RecordCount + 1)
      If bSqlRows Then
         With RdoEdi
         iTotalItems = 0
         While Not .EOF
            iNumChar = !NUMCHARS
            
            If (iNumChar > 0) Then
               strValue = Mid(strEdiData, iIndex, iNumChar)
            Else
               strValue = Mid(strEdiData, iIndex, (iTotLen - iIndex))
            End If
            
            arrValue(iTotalItems) = RemoveSQLString(Trim(strValue))
            arrFieldName(iTotalItems) = !FieldName
            iIndex = iIndex + iNumChar
            iTotalItems = iTotalItems + 1
            .MoveNext
         Wend
         .Close
         End With
      End If
      
      For j = 0 To iTotalItems - 1
         If (strFields = "") Then
            strFields = arrFieldName(j)
            strFldVal = "'" & arrValue(j) & "'"
         Else
            strFields = strFields + "," + arrFieldName(j)
            strFldVal = strFldVal + "," + "'" + arrValue(j) + "'"
         End If
      Next
      
      If (strFldVal <> "") Then
         If (strType = "H0") Then
            strTabName = "Inhd830_EDI"
         Else
            strTabName = "Init830_EDI"
         End If
         
         sSql = "INSERT INTO " & strTabName & " (" & strFields & ") " _
                & " VALUES (" & strFldVal & ")"
         
         Debug.Print sSql
         
         clsAdoCon.ExecuteSql sSql
      End If
      
   End If

   Exit Function
DiaErr1:
   sProcName = "Decode850EdiFormat"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function


Function Decode862EdiFormat(ByVal strEdiData As String)
   Dim iIndex As Integer
   Dim j As Integer
   Dim RdoEdi As adodb.Recordset
   Dim strValue As String
   Dim strType As String
   Dim iTotLen As Integer
   Dim iTotalItems As Integer
   Dim iNumChar As Integer
   Dim strFields As String
   Dim strFldVal As String
   Dim strTabName As String
   
   On Error GoTo DiaErr1
   
   If (strEdiData <> "") Then
      iIndex = 2
      iTotLen = Len(strEdiData)
      strType = Mid(strEdiData, 1, iIndex)
      iIndex = iIndex + 1
      sSql = "SELECT FIELDNAME,NUMCHARS FROM ProEdiFormat WHERE " _
             & "HEADER = '" & strType & "' AND IMPORTTYPE = 'SHP' ORDER BY FORATORDER"
      
      bSqlRows = clsAdoCon.GetDataSet(sSql, RdoEdi, rdOpenStatic)
      ReDim arrValue(0 To RdoEdi.RecordCount + 1)
      ReDim arrFieldName(0 To RdoEdi.RecordCount + 1)
      If bSqlRows Then
         With RdoEdi
         iTotalItems = 0
         While Not .EOF
            iNumChar = !NUMCHARS
            
            If (iNumChar > 0) Then
               strValue = Mid(strEdiData, iIndex, iNumChar)
            Else
               strValue = Mid(strEdiData, iIndex, ((iTotLen - iIndex) + 1))
            End If
            
            arrValue(iTotalItems) = RemoveSQLString(Trim(strValue))
            arrFieldName(iTotalItems) = !FieldName
            iIndex = iIndex + iNumChar
            iTotalItems = iTotalItems + 1
            .MoveNext
         Wend
         .Close
         End With
      End If
      
      If (strType = "H1") Then
      
         For j = 0 To iTotalItems - 1
            If (strFields = "") Then
               strFields = arrFieldName(j)
               strFldVal = "'" & arrValue(j) & "'"
            Else
               strFields = strFields + "," + arrFieldName(j)
               strFldVal = strFldVal + "," + "'" + arrValue(j) + "'"
            End If
         
            strPartNum = ""
            strPartCnt = ""
            strPAUnit = ""
            strPartInfo = ""
            strBldStation = ""
            strECNum = ""
            
            strPartNumFld = ""
            strPartCntFld = ""
            strPAUnitFld = ""
            strPartInfoFld = ""
            strBldStationFld = ""
            strECNumFld = ""
         
         Next
         strTabName = "Inhd862_EDI"
         
         sSql = "INSERT INTO " & strTabName & " (" & strFields & ") " _
                & " VALUES (" & strFldVal & ")"
      
         Debug.Print sSql
         clsAdoCon.ExecuteSql sSql
      Else
         If (strType = "D1") Then
            ' Partnum
            strPartCntFld = arrFieldName(2)
            strPartCnt = arrValue(2)
            ' Partnum
            strPartNumFld = arrFieldName(3)
            strPartNum = arrValue(3)
            ' EC Number
            strECNumFld = arrFieldName(5)
            strECNum = arrValue(5)
            ' Pull#
            strPullNumFld = arrFieldName(6)
            strPullNum = arrValue(6)
            ' Partnum
            strPAUnitFld = arrFieldName(7)
            strPAUnit = arrValue(7)
            ' BinNum
            strBinNumFld = arrFieldName(8)
            strBinNum = arrValue(8)
            ' Build Station
            strBldStationFld = arrFieldName(9)
            strBldStation = arrValue(9)
            
         Else
            For j = 0 To iTotalItems - 1
               If (strFields = "") Then
                  strFields = arrFieldName(j)
                  strFldVal = "'" & arrValue(j) & "'"
               Else
                  strFields = strFields + "," + arrFieldName(j)
                  strFldVal = strFldVal + "," + "'" + arrValue(j) + "'"
               End If
            Next
         
            strTabName = "Init862_EDI"
            
            strPartInfoFld = "," + strPartNumFld + "," + strPartCntFld + "," + _
                           strPAUnitFld + "," + strPullNumFld + "," + _
                           strBinNumFld + "," + strECNumFld + "," + _
                           strBldStationFld + ","
            strPartInfo = ",'" + strPartNum + "','" + strPartCnt + "','" + _
                        strPAUnit + "','" + strPullNum + "','" + strBinNum + _
                           "','" + strECNum + "','" + strBldStation + "',"
            
            If (Trim(strPartNumFld) <> "") Then
               sSql = "INSERT INTO " & strTabName & " (EDI_ELEMENTTYPE" & strPartInfoFld & strFields & ") " _
                      & " VALUES ('" & strType & "'" & strPartInfo & strFldVal & ")"
               
               Debug.Print sSql
               clsAdoCon.ExecuteSql sSql
            End If
            
         End If
      End If
      
      
   End If

   Exit Function
DiaErr1:
   sProcName = "Decode862EdiFormat"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function


Private Function CheckImportType(ByVal strFileName As String, ByVal strFilePath As String) As String
   
   Dim strType As String
   Dim strEDIFormat  As String
   
   strType = Mid(strFileName, 1, 5) ' first 5 char as the filename
   
   If (strType = "in830") Then
      strEDIFormat = CheckFileFormat(strFilePath)
   ElseIf (strType = "in862") Then
      strEDIFormat = "862_EDI"
   Else
      strEDIFormat = ""
   End If
   
   CheckImportType = strEDIFormat
End Function

Private Function CheckFileFormat(ByVal strFilePath As String) As String

   On Error GoTo DiaErr1

   Dim strRecipientsCode As String
   Dim strEDIFormat  As String
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
   ' Read the content if the text file.
   Dim nFileNum As Integer
   Dim strLine As String
' Get a free file number
   nFileNum = FreeFile
   
   Open strFilePath For Input As nFileNum
   ' Read the contents of the file
   If Not EOF(nFileNum) Then
      Line Input #nFileNum, strLine
      Debug.Print strLine
      
      If (strLine <> "") Then
         strRecipientsCode = Mid(strLine, 16, 15)
         
         If (Trim(strRecipientsCode) = "11555AA") Then
            strEDIFormat = "850_EDI"
         Else
            strEDIFormat = "830_PlanSchedule"
         End If
      End If
   End If
   Close nFileNum

   CheckFileFormat = strEDIFormat
   
   Exit Function

DiaErr1:
   sProcName = "getbookpr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function


Function Create850SO() As Integer
   
   Dim strSenderCode As String
   Dim strCust As String
   Dim strPONumber As String
   
   Dim RdoHdEdi As adodb.Recordset
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
   
   strSenderCode = "097248199"

   sSql = "SELECT DISTINCT PONUMBER " _
            & " FROM Inhd830_EDI " _
            & "WHERE EDISENDERCODE = '" & strSenderCode & "'"

   Debug.Print sSql
   
   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoHdEdi, rdOpenStatic)
   
   If bSqlRows Then
      With RdoHdEdi
      While Not .EOF
         
         strPONumber = Trim(!PONUMBER)
               
         ' get the customer name
         If (GetCustFromPOPrefix(strPONumber, strCust)) Then
              
            ' Get new Sales Order number
            Dim strSoType As String
            Dim strItem As String
            Dim strNewSO As String
            strSoType = "S"
            
            GetNewSO strNewSO, strSoType
            
            CreateSOFromEDIData strSenderCode, strPONumber, strNewSO, strSoType, strCust
              
         End If
         .MoveNext
      Wend
      .Close
      End With
   End If

   Set RdoHdEdi = Nothing
   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   sProcName = "Fill850Grid "
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Private Function Create862SalesOrder()
   Dim bByte As Byte
   Dim lNewSoNum As Long
   Dim iList As Long
   
   ' Get new Sales Order number
   Dim strSoType As String
   Dim strItem As String
   strSoType = "S"
   
   Dim strSenderCode As String
   Dim strCust As String
   
   
   strSenderCode = "097248199"
   
   Dim RdoEdi As adodb.Recordset
   Dim strCustCont As String
   Dim strShpToName As String
   Dim strShpAddr1 As String
   Dim strShpAddr2 As String
   Dim strShipTo4 As String
   Dim strShipTo5 As String
   Dim strShpToAddress As String
   Dim strPOItem As String
   Dim strUOM As String
   Dim strDueDt As String
   Dim strShpDt As String
   Dim bPartFound, bIncRow As Boolean
   Dim strSORemark As String
   Dim bSoExists As Boolean
   Dim iItem As Integer
   Dim strBook As String
   Dim strBldStation As String
   Dim strPONumber As String
   
   Dim strPartID As String
   Dim strPart
   Dim strQty As String
   Dim strRefDesc As String
   Dim strUnitPrice As String
   Dim strReqDt As String
   Dim strSoNum As String
   Dim strContactNum As String
   
   Dim strECNum As String
   
   bSoExists = False
   
   sSql = "select DISTINCT a.PONUMBER, b.PARTNUM, b.PAUNITS, PORELEASE," _
            & " SHPQTY, DUEDATE, SHPDATE,a.SHIPCODE," _
            & " SHIPNAME, SHIPADDRESS1, SHIPADDRESS2," _
            & " PULLNUM, BINNUM, BUILDSTATION, ECNUMBER " _
         & " FROM Inhd862_EDI a, Init862_EDI b" _
         & " WHERE a.PONUMBER = b.PONUMBER AND EDI_ELEMENTTYPE = 'D2'"
    
   Debug.Print sSql
   
   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoEdi, rdOpenStatic)
   
   If bSqlRows Then
      With RdoEdi
      While Not .EOF
         
         strPONumber = Trim(!PONUMBER)
         strPOItem = Trim(!PORELEASE)
         strPartID = Trim(!PartNum)
         strQty = Trim(!SHPQTY)
         strDueDt = ConvertToDate(Trim(!DUEDATE))
         strShpDt = ConvertToDate(Trim(!SHPDATE))
         strShpToName = Trim(!SHIPNAME)
         strShpAddr1 = Trim(!SHIPADDRESS1)
         strShpAddr2 = Trim(!SHIPADDRESS2)
         strUOM = Trim(!PAUNITS)
         
         strPullNum = Trim(!PULLNUM)
         strBinNum = Trim(!BINNUM)
         
         strBldStation = Trim(!BUILDSTATION)
         strECNum = Trim(!ECNUMBER)
         
         If (GetCustFromPOPrefix(strPONumber, strCust)) Then
         
            GetPartPrice strPartID, strUnitPrice
   '               strBook = "PACPARTS"
   '               GetBookPrice strPartID, strBook, strUnitPrice
            
            strCustCont = "" 'Trim(!SHIPPERSON)
            strContactNum = ""
            strSORemark = ""
            
            MakeAddress strShpToName, strShpAddr1, strShpAddr2, _
                     "", "", strShpToAddress
            
            ' if the SO header is alrady added don't add the PO again
            bSoExists = CheckOfExistingSO(strPONumber, strPartID, strSoNum)
            If (bSoExists = False) Then
               GetNewSO strSoNum, strSoType
               AddSalesOrder strSoNum, strPONumber, strCustCont, strContactNum, _
                                 strShpToName, strCust, strShpToAddress, strSoType, strSORemark
         Else
            ' Get the customer inforamtion
            Dim bGoodCust As Byte
            bGoodCust = GetCustomerData(strCust)
            If bCutOff = 1 Then
               bGoodCust = 0
            End If
            If Not bGoodCust Then Exit Function
         End If
            
            ' Add So items
            AddSoItem strSoNum, CStr(strPOItem), strPONumber, strPOItem, _
               strPartID, strQty, strUnitPrice, strDueDt, strPullNum, _
               strBinNum, strBldStation, strECNum, strShpDt
         End If
         
         .MoveNext
      Wend
      .Close
      End With
   End If
   

   Set RdoEdi = Nothing
   
   
   
End Function



Function CreateSOFromEDIData(ByVal strSenderCode As String, ByVal strInputPONum As String, _
               ByVal strNewSO As String, ByVal strSoType As String, ByVal strCusName As String) As Integer
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
   Dim RdoEdi As adodb.Recordset
   Dim strCustCont As String
   Dim strShpToName As String
   Dim strShipTo2 As String
   Dim strShipTo3 As String
   Dim strShipTo4 As String
   Dim strShipTo5 As String
   Dim strShpToAddress As String
   Dim strPOItem As String
   Dim strPartID As String
   Dim strUOM As String
   Dim strUnitPrice As String
   Dim strReqDt As String
   Dim bPartFound, bIncRow As Boolean
   Dim strQty As String
   Dim strContactNum As String
   Dim strSORemark As String
   Dim bSOHdAdded As Boolean
   Dim iItem As Integer
   Dim strBook As String
   
   bSOHdAdded = False
   
   sSql = "SELECT EDISENDERCODE, SHIPTO1, SHIPTO2, SHIPTO3," _
             & " SHIPTO4 , SHIPTO5, CUSTCONTACT, Inhd830_EDI.PONUMBER AS PONUMBER1," _
             & "POITEM,POPART ,POPAUNIT, POQTY, POREQDT, POAMT " _
            & " FROM Inhd830_EDI, Init830_EDI " _
            & "WHERE Inhd830_EDI.PONUMBER = Init830_EDI.PONUMBER AND " _
            & " EDISENDERCODE = '" & strSenderCode & "' AND " _
            & "Inhd830_EDI.PONUMBER = '" & strInputPONum & "'"

   Debug.Print sSql
   
   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoEdi, rdOpenStatic)
   
   If bSqlRows Then
      With RdoEdi
      While Not .EOF
         
         'strPONumber = Trim(!PONUMBER1)
         strPartID = Trim(!POPART)
         strSenderCode = Trim(!EDISENDERCODE)
         strShpToName = Trim(!SHIPTO1)
         strShipTo2 = Trim(!SHIPTO2)
         strShipTo3 = Trim(!SHIPTO3)
         strShipTo4 = Trim(!SHIPTO4)
         strShipTo5 = Trim(!SHIPTO5)
         strCustCont = Trim(!CUSTCONTACT)
         strQty = Trim(!POQTY)
         strPOItem = Trim(!POITEM)
         strUOM = Trim(!POPAUNIT)
         strReqDt = ConvertToDate(Trim(!POREQDT))
         'Ignore the proice from EDI..get the price from ProceBook
         'strUnitPrice = Trim(!POAMT)
         
         strBook = "PACPARTS"
         GetBookPrice strPartID, strBook, strUnitPrice

         
         strContactNum = ""
         strSORemark = ""
         
         MakeAddress strShpToName, strShipTo2, strShipTo3, _
                  strShipTo4, strShipTo5, strShpToAddress
         
         ' if the SO header is alrady added don't add the PO again
         If (bSOHdAdded = False) Then
            AddSalesOrder strNewSO, strInputPONum, strCustCont, strContactNum, _
                              strShpToName, strCusName, strShpToAddress, strSoType, strSORemark
            bSOHdAdded = True
         Else
            ' Get the customer inforamtion
            Dim bGoodCust As Byte
            bGoodCust = GetCustomerData(strCusName)
            If bCutOff = 1 Then
               bGoodCust = 0
            End If
            If Not bGoodCust Then Exit Function
         End If
         
         ' Add So items
         Dim strPullNum As String, strBinNum As String
         Dim strBldStation As String, strECNum As String
         strPullNum = ""
         strBinNum = ""
         strBldStation = ""
         strECNum = ""
         
         
         AddSoItem strNewSO, CStr(strPOItem), strInputPONum, strPOItem, _
            strPartID, strQty, strUnitPrice, strReqDt, strPullNum, _
            strBinNum, strBldStation, strECNum

         .MoveNext
      Wend
      .Close
      End With
   End If
   
   Set RdoEdi = Nothing
   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateSOFromXMLData"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Private Function MakeAddress(strShpTo1 As String, strShpTo2 As String, strStreet As String, _
                  strRegionCode As String, strPostalCode As String, ByRef strShpToAddress As String)

   Dim strNewAddress As String
   
   strShpToAddress = ""
   
   If (strShpTo1 <> "") Then strNewAddress = strNewAddress & strShpTo1 & vbCrLf
   If (strShpTo2 <> "") Then strNewAddress = strNewAddress & strShpTo2 & vbCrLf
   If (strStreet <> "") Then strNewAddress = strNewAddress & strStreet & vbCrLf
   
   
   If (strPostalCode <> "") Then
      If (strRegionCode <> "") Then
         strNewAddress = strNewAddress & ", " & IIf((strRegionCode <> ""), strRegionCode, "") & " - " & strPostalCode
      Else
         strNewAddress = strNewAddress & " - " & strPostalCode
      End If
   End If
   
   strShpToAddress = strNewAddress

End Function

Private Function CheckRecordExits(sSql As String)
    
   Dim RdoCon As adodb.Recordset
   
   On Error GoTo ERR1
      
   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoCon, ES_FORWARD)
   If bSqlRows Then
       CheckRecordExits = True
   Else
       CheckRecordExits = False
   End If
   Set RdoCon = Nothing
   Exit Function
   
ERR1:
    CheckRecordExits = False

End Function

Private Function DeleteOldData(strTableName As String)

   If (strTableName <> "") Then
      sSql = "DELETE FROM " & strTableName
      clsAdoCon.ExecuteSql sSql
   End If

End Function

Private Function CheckOfExistingSO(strPONumber As String, strPartID As String, ByRef strSoNum As String) As Boolean
   Dim RdoSO As adodb.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT ISNULL(MAX(SONUMBER),0) SONUMBER  FROM sohdTable,SoitTable WHERE " _
             & " SONUMBER = ITSO AND SOPO = '" & strPONumber & "'" _
             & "  AND ITPART = '" & Compress(strPartID) & "'"

   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoSO, ES_FORWARD)
   If bSqlRows Then
      With RdoSO
         If (Trim(!SoNumber) = 0) Then
            strSoNum = ""
            CheckOfExistingSO = False
         Else
            strSoNum = Trim(!SoNumber)
            CheckOfExistingSO = True
         End If
         ClearResultSet RdoSO
      End With
   Else
      strSoNum = ""
      CheckOfExistingSO = False
      
   End If
   
   Set RdoSO = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "CheckOfExistingSO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Private Function ConvertToDate(strDate As String) As String
   
   Dim strDateConv As String
   If (Trim(strDate) <> "" And Not IsNull(strDate)) Then
      strDateConv = Mid(strDate, 3, 2) & "/" & Mid(strDate, 5, 2) & "/" & Mid(strDate, 1, 2)
   Else
      strDateConv = ""
   End If
   
   ConvertToDate = strDateConv
End Function


Private Function GetBookPrice(strPart As String, strBook As String, ByRef strPrice As String)
   Dim RdoBok As adodb.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT PARTREF,PARTNUM,PBIREF,PBIPARTREF,PBIPRICE " _
          & "FROM PartTable,PbitTable WHERE (PARTREF=PBIPARTREF) AND " _
          & "(PBIREF='" & strBook & "') AND (PARTREF='" & Compress(strPart) & "')"
   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoBok, ES_FORWARD)
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
 
 Private Function GetPartPrice(strPart As String, ByRef strPrice As String)
   Dim RdoPrice As adodb.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT PAPRICE FROM PartTable WHERE PARTREF='" & Compress(strPart) & "'"
   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoPrice, ES_FORWARD)
   If bSqlRows Then
      With RdoPrice
         strPrice = Format(!PAPRICE, ES_MoneyFormat)
         ClearResultSet RdoPrice
      End With
   Else
      strPrice = Format(0, ES_MoneyFormat)
   End If
   
   Set RdoPrice = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getbookpr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Private Function RemoveSQLString(varString As Variant) As String
   Dim PartNo As String
   Dim NewPart As String
   
   On Error GoTo modErr1
   PartNo = Trim$(varString)
   If Len(PartNo) > 0 Then
      NewPart = Replace(PartNo, Chr$(39), "")    'single quote
      NewPart = Replace(NewPart, Chr$(44), "")   ' comma
   End If
   RemoveSQLString = NewPart
   Exit Function
   
modErr1:
   Resume modErr2
modErr2:
   On Error Resume Next
   RemoveSQLString = varString
   
End Function


Private Sub AddSalesOrder(strNewSO As String, strBuyerOrderNumber As String, _
                     strContactName As String, strContactNum As String _
                     , strShipName As String, strCusName As String, strNewAddress As String, _
                     strSoType As String, strSORemark As String)
                     
   Dim sNewDate As Variant
   Dim bGoodCust As Byte
   
   bGoodCust = GetCustomerData(strCusName)
   If bCutOff = 1 Then
      'MsgBox "This Customer's Credit Is On Hold.", _
         vbInformation, Caption
      bGoodCust = 0
   End If
   If Not bGoodCust Then Exit Sub
   On Error GoTo DiaErr1
   
   sNewDate = Format(ES_SYSDATE, "mm/dd/yy")
   sSql = "INSERT SohdTable (SONUMBER,SOTYPE,SOCUST,SODATE," _
          & "SOSALESMAN,SOSTNAME,SOSTADR,SODIVISION,SOREGION,SOSTERMS," _
          & "SOVIA,SOFOB,SOARDISC,SODAYS,SONETDAYS,SOFREIGHTDAYS," _
          & "SOTEXT,SOTAXEXEMPT,SOPO, SOREMARKS) " _
          & "VALUES(" & Val(strNewSO) & ",'" & strSoType & "','" _
          & strCusName & "','" & sNewDate & "','" & sSlsMan & "','" _
          & strShipName & "','" & strNewAddress & "','" & sDivision & "','" _
          & sRegion & "','" & sSterms & "','" & sVia & "','" _
          & sFob & "'," & cDiscount & "," & iDays & "," & iNetDays _
          & "," & iFrtDays & ",'" & strNewSO & "','" & sTaxExempt & "','" _
          & Trim(strBuyerOrderNumber) & "','" & strSORemark & "')"
   
   Debug.Print sSql
   
   clsAdoCon.ExecuteSql sSql, rdExecDirect
   If clsAdoCon.RowsAffected Then
      On Error Resume Next
'      MsgBox "Sales Order Added.", vbInformation, Caption
      sSql = "UPDATE SohdTable SET SOCCONTACT='" & strContactName & "'," _
             & "SOCPHONE='" & strContactNum & "',SOCINTFAX='" & sConIntFax _
             & "',SOCFAX='" & sConFax & "',SOCEXT=" & sConExt _
             & " WHERE SONUMBER=" & Val(strNewSO) & ""
      Debug.Print sSql
      
      clsAdoCon.ExecuteSql sSql, rdExecDirect
      
      sSql = "UPDATE ComnTable SET COLASTSALESORDER='" & Trim(strSoType) _
             & Trim(strNewSO) & "' WHERE COREF=1"
      clsAdoCon.ExecuteSql sSql, rdExecDirect
   
   Else
      'MsgBox "Couldn't Add Sales Order.", vbExclamation, Caption
   End If
   Exit Sub
   
DiaErr1:
   MsgBox Err.Description
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Sub


Private Sub AddSoItem(strNewSO As String, strItem As String, _
                     strBuyerOrderNumber As String, strBuyerItemLine As String, _
                     strPartID As String, strQty As String, _
                     strUnitPrice As String, strReqDt As String, _
                     strPullNum As String, strBinNum As String, _
                     strBldStation As String, strECNum As String, _
                     Optional strShippedDt As String = "")

   On Error GoTo DiaErr1
   
   Dim strShedDt As String
   ' Create the ShipDate
   If (strShippedDt <> "") Then
      strShedDt = strShippedDt
   Else
      strShedDt = strReqDt
   End If
   
   If (iFrtDays > 0) Then
      strShedDt = Format(DateAdd("d", -iFrtDays, strShedDt), "mm/dd/yy")
   End If

   
   Dim RdoSoit As adodb.Recordset
   sSql = "SELECT DISTINCT ITSO FROM SoitTable WHERE " _
             & " ITSO = '" & strNewSO & "'" _
             & "  AND ITNUMBER = '" & Val(strItem) & "'"

   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoSoit, ES_FORWARD)
   If bSqlRows Then
      ClearResultSet RdoSoit
      Set RdoSoit = Nothing
      
      If (Val(strQty) = 0) Then
         sSql = "UPDATE SoitTable SET ITQTY = " & Val(strQty) & ", ITSCHED = '" & strShedDt & "'," _
                  & " ITCUSTREQ = '" & strReqDt & "', PULLNUM = '" & strPullNum & "', BINNUM = '" & strBinNum & "', " _
                  & "ITACTUAL=NULL, ITCANCELED=1, ITCANCELDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "' " _
                  & " WHERE ITSO = '" & strNewSO & "' AND ITNUMBER = '" & Val(strItem) & "'"
      
      Else
         sSql = "UPDATE SoitTable SET ITQTY = " & Val(strQty) & ", ITSCHED = '" & strShedDt & "'," _
                  & " ITCUSTREQ = '" & strReqDt & "', PULLNUM = '" & strPullNum & "', BINNUM = '" & strBinNum & "', " _
                  & " BUILDSTATION = '" & strBldStation & "', ECNUMBER = '" & strECNum & "' " _
                  & " WHERE ITSO = '" & strNewSO & "' AND ITNUMBER = '" & Val(strItem) & "'"
      End If
             
         
      Debug.Print sSql
      
      clsAdoCon.ExecuteSql sSql, rdExecDirect
      
      Exit Sub
   End If
      
   clsAdoCon.BeginTrans
      
   sSql = "INSERT SoitTable (ITSO,ITNUMBER,ITCUSTITEMNO, ITPART,ITQTY,ITCUSTREQ, ITSCHED,ITBOOKDATE," _
          & "ITDOLLORIG, ITDOLLARS, ITUSER, PULLNUM, BINNUM, BUILDSTATION, ECNUMBER) " _
          & "VALUES(" & strNewSO & "," & strItem & ",'" & strBuyerItemLine & "','" _
          & Compress(strPartID) & "'," & Val(strQty) & ",'" & strReqDt & "','" & strShedDt & "','" _
          & Format(ES_SYSDATE, "mm/dd/yy") & "','" & CCur(strUnitPrice) & "','" _
          & CCur(strUnitPrice) & "','" & sInitials & "','" & strPullNum & "','" _
          & strBinNum & "','" & strBldStation & "','" & strECNum & "')"
   
   Debug.Print sSql
   
   clsAdoCon.ExecuteSql sSql, rdExecDirect
   
   'Add commission if applicable.
'   If cmdCom.Enabled Then
     Dim Item As New ClassSoItem
     Dim bUserMsg As Boolean
     bUserMsg = False
     Item.InsertCommission CLng(strNewSO), CLng(strItem), "", ""
     Item.UpdateCommissions CLng(strNewSO), CLng(strItem), "", bUserMsg
 '  End If
   
   clsAdoCon.CommitTrans
   Exit Sub
   
DiaErr1:
   sProcName = "addsoitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub

Public Sub GetNewSO(ByRef sNewSo As String, ByVal sSoType As String)
   Dim RdoSon As adodb.Recordset
   Dim lSales As Long
   On Error GoTo DiaErr1
   
   sSql = "SELECT (MAX(SONUMBER)+ 1)AS SalesOrder FROM SohdTable"
   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         If Not IsNull(.Fields(0)) Then
            sNewSo = "" & Format$(!SalesOrder, "00000")
         Else
            sNewSo = "00000"
         End If
         ClearResultSet RdoSon
      End With
   End If
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "GetNewSO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub

Private Function GetPartComm(ByVal strGetPart As String, _
            ByRef strPartNum As String, ByRef bComm As Boolean) As Byte
   Dim RdoPrt As adodb.Recordset
   
   On Error GoTo DiaErr1
   bComm = False
   strGetPart = Compress(strGetPart)
   If Len(strGetPart) > 0 Then
      sSql = "SELECT PARTNUM,PADESC,PAEXTDESC,PAPRICE,PAQOH," _
             & "PACOMMISSION FROM PartTable WHERE PARTREF='" & strGetPart & "'"
      bSqlRows = clsAdoCon.GetDataSet(sSql, RdoPrt, ES_STATIC)
      If bSqlRows Then
         With RdoPrt
            strPartNum = "" & Trim(!PartNum)
            If !PACOMMISSION = 1 Then bComm = True _
                               Else bComm = False
            GetPartComm = 1
            ClearResultSet RdoPrt
         End With
      Else
         GetPartComm = 0
      End If
      'On Error Resume Next
      Set RdoPrt = Nothing
   Else
      GetPartComm = 0
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetPartComm"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Function

Private Function GetCustomerData(strCusName As String) As Byte
   Dim RdoCst As adodb.Recordset
   sCust = Compress(strCusName)
   On Error GoTo DiaErr1
   sSql = "SELECT CUREF,CUSTNAME,CUSTNAME,CUSTADR,CUARDISC," _
          & "CUDAYS,CUNETDAYS,CUDIVISION,CUREGION,CUSTERMS," _
          & "CUVIA,CUFOB,CUSALESMAN,CUDISCOUNT,CUSTSTATE," _
          & "CUSTCITY,CUSTZIP,CUCCONTACT,CUCPHONE,CUCEXT,CUCINTPHONE," _
          & "CUFRTDAYS,CUINTFAX,CUFAX,CUTAXEXEMPT,CUCUTOFF " _
          & "FROM CustTable WHERE CUREF='" & strCusName & "'"
   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoCst)
   If bSqlRows Then
      With RdoCst
         bCutOff = !CUCUTOFF
         sStName = "" & Trim(!CUSTNAME)
         sStAdr = "" & Trim(!CUSTADR) & vbCrLf _
                  & "" & Trim(!CUSTCITY) & " " & Trim(!CUSTSTATE) _
                  & "  " & Trim(!CUSTZIP)
         sDivision = "" & Trim(!CUDIVISION)
         sRegion = "" & Trim(!CUREGION)
         sSterms = "" & Trim(!CUSTERMS)
         sVia = "" & Trim(!CUVIA)
         sFob = "" & Trim(!CUFOB)
         sSlsMan = "" & Trim(!CUSALESMAN)
         sContact = "" & Trim(!CUCCONTACT)
         sConIntPhone = "" & Trim(!CUCINTPHONE)
         sConPhone = "" & Trim(!CUCPHONE)
         sConIntFax = "" & Trim(!CUINTFAX)
         sConFax = "" & Trim(!CUFAX)
         sConExt = "" & Trim(str$(!CUCEXT))
         cDiscount = Format(0 + !CUARDISC, "##0.000")
         iDays = Format(!CUDAYS, "###0")
         iNetDays = Format(!CUNETDAYS, "###0")
         iFrtDays = Format(!CUFRTDAYS, "##0")
         sTaxExempt = "" & Trim(!CUTAXEXEMPT)
         ClearResultSet RdoCst
      End With
      GetCustomerData = True
   Else
      sStName = ""
      sStAdr = ""
      sDivision = ""
      sRegion = ""
      sSterms = ""
      sVia = ""
      sFob = ""
      sSlsMan = ""
      iFrtDays = 0
      'MsgBox "Couldn't Retrieve Customer.", vbExclamation, Caption
      GetCustomerData = False
   End If
   Set RdoCst = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcustda"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description

End Function

Private Sub GetCustomerRef(ByRef strCusFullName As String, ByRef strCusName As String)

   Dim RdoCus As adodb.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT CUREF FROM CustTable WHERE CUNAME = '" & strCusFullName & "'"
   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoCus)
   If bSqlRows Then
      With RdoCus
         strCusName = Trim(!CUREF)
         ClearResultSet RdoCus
      End With
   Else
      strCusName = ""
   End If
   Set RdoCus = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "GetCustomerRef"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   'GetCustomerRef = False

End Sub

Private Function GetCustFromPOPrefix(strPO As String, ByRef strCust As String) As Boolean
   
   Dim strPrefix As String
   On Error GoTo modErr1
   
   ' ST-S or EX-S
   If IsNumeric(strPO) Then
       strPrefix = "#"
       GetCustomer strPrefix, strCust
       GetCustFromPOPrefix = True
       Exit Function
   End If
      
   ' Check for the 4 characters
   strPrefix = Mid(strPO, 1, 4)
   If (GetCustomer(strPrefix, strCust)) Then
      GetCustFromPOPrefix = True
      Exit Function
   End If
   ' Check for the 4 characters
   strPrefix = Mid(strPO, 1, 3)
   If (GetCustomer(strPrefix, strCust)) Then
      GetCustFromPOPrefix = True
      Exit Function
   End If
   
   '
   strPrefix = Mid(strPO, 4, 1)
   If (GetCustomer(strPrefix, strCust)) Then
      GetCustFromPOPrefix = True
      Exit Function
   End If
   
   GetCustFromPOPrefix = False
   Exit Function
   
modErr1:
   sProcName = "CheckPOLetter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Function

Private Function GetCustomer(strPrefix As String, ByRef strCust As String) As Boolean

   Dim RdoCst As adodb.Recordset
   sSql = "SELECT DISTINCT CUREF FROM " _
            & " ASNInfoTable WHERE POLETTERREF = '" & strPrefix & "'"

   bSqlRows = clsAdoCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         strCust = Trim(!CUREF)
         ClearResultSet RdoCst
      End With
      GetCustomer = True
   Else
      strCust = ""
      GetCustomer = False
   End If
   
   Set RdoCst = Nothing
End Function

