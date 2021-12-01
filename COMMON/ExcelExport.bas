Attribute VB_Name = "ExcelExport"
'*** ES2000 is the property of                                ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***

Option Explicit

Public Sub SaveAsExcel(ByVal rs As ADODB.Recordset, ByRef aFieldsToExport() As String, Optional ByVal FileName As String = "", Optional OpenWhenDone As Boolean = False, Optional bHeaders As Boolean = True, Optional bUseDescriptiveFieldNames As Boolean = True, Optional bShowBooleanAsYN As Boolean = True, Optional ByRef pg As Variant)

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim i, fd As Integer
    Dim iSheetsPerBook As Integer
    

    'Cell count, the cells we can use
    Dim iCell As Integer

    Screen.MousePointer = vbHourglass
    ' Assign object references to the variables. Use
    ' Add methods to create new workbook and worksheet
    ' objects.
    
      
'If a session of Excel is already running then grab it.
    On Error Resume Next
    
    
    
    
    If Not IsMissing(pg) Then
        pg.Visible = True
        pg.Min = 0
        pg.max = 100
        pg.progress = 0
        DoEvents
    End If
    
    
    
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo SaveToExcelError
'Otherwise instantiate a new instance.
    If xlApp Is Nothing Then Set xlApp = New Excel.Application

    iSheetsPerBook = xlApp.SheetsInNewWorkbook
    xlApp.SheetsInNewWorkbook = 1
    Set xlBook = xlApp.Workbooks.Add
    xlApp.SheetsInNewWorkbook = iSheetsPerBook
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set xlSheet = xlBook.Worksheets(1)
    
    'Get the field names
    If bHeaders Then
        'CellCnt = 1
        For fd = 0 To rs.Fields.count - 1
            iCell = IsInArray(rs.Fields(fd).Name, aFieldsToExport)
            If iCell >= 0 Then
                Select Case rs.Fields(fd).Type
                Case adBinary, adLongVarChar, adLongVarBinary, adVarBinary
                ' This type of data can't export to excel
                Case Else
                    If bUseDescriptiveFieldNames Then
                        Dim strTblName As String
                        Dim strFldName As String
                        
                        If Not IsNull(rs.Fields(fd).Properties(3).Value) Then
                           strTblName = rs.Fields(fd).Properties(3).Value
                        Else
                           strTblName = ""
                        End If
                        
                        strFldName = rs.Fields(fd).Name
                        If (strTblName <> "") Then
                           xlSheet.Cells(1, iCell + 1).Value = FriendlyFieldName(strTblName, strFldName)
                        Else
                           xlSheet.Cells(1, iCell + 1).Value = strFldName
                        End If
                     Else
                        xlSheet.Cells(1, iCell + 1).Value = rs.Fields(fd).Name
                     End If
                     
                    xlSheet.Cells(1, iCell + 1).Interior.ColorIndex = 33
                    xlSheet.Cells(1, iCell + 1).Font.Bold = True
                    xlSheet.Cells(1, iCell + 1).BorderAround xlContinuous
                    'CellCnt = CellCnt + 1
                End Select
            End If
        Next
        i = 2
    Else
        i = 1
    End If


'Rewind the rescordset
    rs.MoveFirst

'    Do While Not rs.EOF()
'        For fd = 0 To rs.rdoColumns.Count - 1
'
'            iCell = IsInArray(rs.rdoColumns(fd).Name, aFieldsToExport)
'            If iCell >= 0 Then
'
'            Select Case rs.rdoColumns(fd).Type
'            Case rdTypeBINARY, rdTypeLONGVARCHAR, rdTypeLONGVARBINARY, rdTypeBINARY
'            ' This type of data can't export to excel
'            Case Else
'                xlSheet.Cells(i, iCell + 1).Value = rs.rdoColumns(fd).Value
'                xlSheet.Columns().AutoFit
'            End Select
'            End If
'        Next
'        rs.MoveNext
'        i = i + 1
'    Loop
    If Not IsMissing(pg) Then pg.max = rs.RecordCount
On Error Resume Next

    Do While Not rs.EOF()
        iCell = 0
        For fd = LBound(aFieldsToExport) To UBound(aFieldsToExport)
            Select Case rs.Fields(aFieldsToExport(fd)).Type
            Case adBinary, adLongVarChar, adLongVarBinary, adVarBinary
            Case Else
                xlSheet.Cells(i, iCell + 1).Value = rs.Fields(aFieldsToExport(fd)).Value
                'xlSheet.Columns().AutoFit
            End Select
            iCell = iCell + 1
        Next fd
        rs.MoveNext
        i = i + 1
   Debug.Print i
        If Not IsMissing(pg) Then
           ' pg.Visible = True
           ' pg.Min = 0
           If i <= pg.max Then pg.Value = i Else pg.Value = pg.max
            'pg.Value = i - 1
            'DoEvents
        End If
        
    Loop
    
'xlSheet.Columns.AutoFit
xlSheet.rows.RowHeight = 20

'Fit all columns
'    CellCnt = 1
'    For fd = 0 To rs.rdoColumns.Count - 1
'            If IsInArray(rs.rdoColumns(fd).Name, aFieldsToExport) Then
'
'        Select Case rs.rdoColumns(fd).Type
'        Case rdTypeBINARY, rdTypeLONGVARCHAR, rdTypeLONGVARBINARY, rdTypeBINARY
'        ' This type of data can't export to excel
'        Case Else
'            xlSheet.Columns(CellCnt).AutoFit
'            CellCnt = CellCnt + 1
'        End Select
'        End If
'    Next

' Save the Worksheet.
If Len(Trim(FileName)) > 0 Then
    If InStr(1, FileName, ".") = 0 Then FileName = FileName + ".xlsx"
    xlBook.SaveAs FileName
End If

'xlSheet.SaveAs filename
', xlWorkbookNormal

' Close the Workbook
'xlBook.Close
' Close Microsoft Excel with the Quit method.
'xlApp.Quit
'
'Set xlSheet = Nothing
'xlBook.Close False
'Set xlBook = Nothing

If OpenWhenDone Then
    Screen.MousePointer = vbDefault
    'xlApp.Workbooks.Open filename
    xlApp.Visible = True
    xlApp.DisplayAlerts = True
Else
'    If InStr(1, filename, ".") = 0 Then filename = filename + ".xlsx"
'    xlBook.SaveAs filename
    Set xlSheet = Nothing
    xlBook.Close False
    Set xlBook = Nothing
    xlApp.Visible = True
    xlApp.DisplayAlerts = True
    xlApp.Quit
    Set xlApp = Nothing
End If


' Release the objects.
'Set xlApp = Nothing
'Set xlBook = Nothing
'Set xlSheet = Nothing

Screen.MousePointer = vbDefault
If Not IsMissing(pg) Then pg.Visible = False
'Shell App.Path & "\test.xlsx"
Exit Sub

SaveToExcelError:
MsgBox Err.Description & " Row = " & str(i) & " Column = " & str(iCell)


End Sub


Public Sub SaveAsExcelSupDup(ByVal rs As ADODB.Recordset, ByRef aFieldsToExport() As String, Optional ByVal FileName As String = "", _
            Optional bSupDup As Boolean = False, Optional iDupCol As Integer, Optional iSupCol As Integer, _
            Optional OpenWhenDone As Boolean = False, Optional bHeaders As Boolean = True, Optional bUseDescriptiveFieldNames As Boolean = True, _
            Optional bShowBooleanAsYN As Boolean = True, _
            Optional ByRef pg As Variant)

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim i, fd As Integer
    Dim iSheetsPerBook As Integer
    

    'Cell count, the cells we can use
    Dim iCell As Integer

    Screen.MousePointer = vbHourglass
    ' Assign object references to the variables. Use
    ' Add methods to create new workbook and worksheet
    ' objects.
    
      
'If a session of Excel is already running then grab it.
    On Error Resume Next
    
    
    
    
    If Not IsMissing(pg) Then
        pg.Visible = True
        pg.Min = 0
        pg.max = 100
        pg.progress = 0
        DoEvents
    End If
    
    
    
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo SaveToExcelError
    Err.Clear
'Otherwise instantiate a new instance.
    If xlApp Is Nothing Then Set xlApp = New Excel.Application

    iSheetsPerBook = xlApp.SheetsInNewWorkbook
    xlApp.SheetsInNewWorkbook = 1
    Set xlBook = xlApp.Workbooks.Add
    xlApp.SheetsInNewWorkbook = iSheetsPerBook
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set xlSheet = xlBook.Worksheets(1)
    
    'Get the field names
    If bHeaders Then
        'CellCnt = 1
        For fd = 0 To rs.Fields.count - 1
            iCell = IsInArray(rs.Fields(fd).Name, aFieldsToExport)
            If iCell >= 0 Then
                Select Case rs.Fields(fd).Type
                Case adBinary, adLongVarChar, adLongVarBinary, adVarBinary
                ' This type of data can't export to excel
                Case Else
                    If bUseDescriptiveFieldNames Then
                        Dim strTblName As String
                        Dim strFldName As String
                        
                        If Not IsNull(rs.Fields(fd).Properties(3).Value) Then
                           strTblName = rs.Fields(fd).Properties(3).Value
                        Else
                           strTblName = ""
                        End If
                        
                        strFldName = rs.Fields(fd).Name
                        If (strTblName <> "") Then
                           xlSheet.Cells(1, iCell + 1).Value = FriendlyFieldName(strTblName, strFldName)
                        Else
                           xlSheet.Cells(1, iCell + 1).Value = strFldName
                        End If
                     Else
                        xlSheet.Cells(1, iCell + 1).Value = rs.Fields(fd).Name
                     End If
                     
                    xlSheet.Cells(1, iCell + 1).Interior.ColorIndex = 33
                    xlSheet.Cells(1, iCell + 1).Font.Bold = True
                    xlSheet.Cells(1, iCell + 1).BorderAround xlContinuous
                    'CellCnt = CellCnt + 1
                End Select
            End If
        Next
        i = 2
    Else
        i = 1
    End If


'Rewind the rescordset
    rs.MoveFirst

'    Do While Not rs.EOF()
'        For fd = 0 To rs.rdoColumns.Count - 1
'
'            iCell = IsInArray(rs.rdoColumns(fd).Name, aFieldsToExport)
'            If iCell >= 0 Then
'
'            Select Case rs.rdoColumns(fd).Type
'            Case rdTypeBINARY, rdTypeLONGVARCHAR, rdTypeLONGVARBINARY, rdTypeBINARY
'            ' This type of data can't export to excel
'            Case Else
'                xlSheet.Cells(i, iCell + 1).Value = rs.rdoColumns(fd).Value
'                xlSheet.Columns().AutoFit
'            End Select
'            End If
'        Next
'        rs.MoveNext
'        i = i + 1
'    Loop
    If Not IsMissing(pg) Then pg.max = rs.RecordCount
'On Error Resume Next

   Dim strPrevVal As String
   Dim strCurVal As String
   Dim bSupress As Boolean
   strPrevVal = ""
   strCurVal = ""
   bSupress = False
   
    Do While Not rs.EOF()
        iCell = 0
        For fd = LBound(aFieldsToExport) To UBound(aFieldsToExport)
            Select Case rs.Fields(aFieldsToExport(fd)).Type
            Case adBinary, adLongVarChar, adLongVarBinary, adVarBinary
            Case Else
               If (bSupDup) And iSupCol = (iCell + 1) Then
                  If (bSupress = False) Then
                     xlSheet.Cells(i, iCell + 1).Value = Replace(rs.Fields(aFieldsToExport(fd)).Value, "/", "-")
                  Else
                     xlSheet.Cells(i, iCell + 1).Value = ""
                  End If
               Else
                  xlSheet.Cells(i, iCell + 1).Value = Replace(rs.Fields(aFieldsToExport(fd)).Value, "/", "-")
               End If
               
               If (iDupCol = (iCell + 1)) Then
                  strCurVal = rs.Fields(aFieldsToExport(fd)).Value
                  If (strPrevVal = "" And strCurVal = "") Then
                     strPrevVal = strCurVal
                     bSupress = False
                  ElseIf (strPrevVal <> strCurVal) Then
                     strPrevVal = strCurVal
                     bSupress = False
                  Else
                     bSupress = True
                  End If
               End If
               
               xlSheet.Columns().AutoFit
            End Select
            iCell = iCell + 1
        Next fd
        rs.MoveNext
        i = i + 1
        
        If Not IsMissing(pg) Then
           ' pg.Visible = True
           ' pg.Min = 0
           If i <= pg.max Then pg.Value = i Else pg.Value = pg.max
            'pg.Value = i - 1
            'DoEvents
        End If
        
    Loop
    
    Debug.Print Err.Number & Err.Description
    
'xlSheet.Columns.AutoFit
xlSheet.rows.RowHeight = 20

'Fit all columns
'    CellCnt = 1
'    For fd = 0 To rs.rdoColumns.Count - 1
'            If IsInArray(rs.rdoColumns(fd).Name, aFieldsToExport) Then
'
'        Select Case rs.rdoColumns(fd).Type
'        Case rdTypeBINARY, rdTypeLONGVARCHAR, rdTypeLONGVARBINARY, rdTypeBINARY
'        ' This type of data can't export to excel
'        Case Else
'            xlSheet.Columns(CellCnt).AutoFit
'            CellCnt = CellCnt + 1
'        End Select
'        End If
'    Next

' Save the Worksheet.
If Len(Trim(FileName)) > 0 Then
    If InStr(1, FileName, ".") = 0 Then FileName = FileName + ".xlsx"
    xlBook.SaveAs FileName
End If

'xlSheet.SaveAs filename
', xlWorkbookNormal

' Close the Workbook
'xlBook.Close
' Close Microsoft Excel with the Quit method.
'xlApp.Quit
'
'Set xlSheet = Nothing
'xlBook.Close False
'Set xlBook = Nothing

If OpenWhenDone Then
    Screen.MousePointer = vbDefault
    'xlApp.Workbooks.Open filename
    xlApp.Visible = True
    xlApp.DisplayAlerts = True
Else
'    If InStr(1, filename, ".") = 0 Then filename = filename + ".xlsx"
'    xlBook.SaveAs filename
    Set xlSheet = Nothing
    xlBook.Close False
    Set xlBook = Nothing
    xlApp.Visible = True
    xlApp.DisplayAlerts = True
    xlApp.Quit
    Set xlApp = Nothing
End If


' Release the objects.
'Set xlApp = Nothing
'Set xlBook = Nothing
'Set xlSheet = Nothing

Screen.MousePointer = vbDefault
If Not IsMissing(pg) Then pg.Visible = False
'Shell App.Path & "\test.xlsx"
Exit Sub

SaveToExcelError:
MsgBox Err.Description & " Row = " & str(i) & " Column = " & str(iCell)


End Sub


Public Function IsInArray(FindValue As String, arrSearch() As String) As Integer
    Dim i As Integer
    
    On Error GoTo LocalError
    IsInArray = -1
    If Not IsArray(arrSearch) Then Exit Function
    For i = LBound(arrSearch) To UBound(arrSearch)
        If UCase(Trim(FindValue)) = UCase(Trim(arrSearch(i))) Then
            IsInArray = i
            Exit For
        End If
    Next i
    
'    IsInArray = InStr(1, vbNullChar & Join(arrSearch, _
'     vbNullChar) & vbNullChar, vbNullChar & FindValue & _
'     vbNullChar) > 0

Exit Function
LocalError:
End Function



Public Function FriendlyFieldName(sTbl As String, sFld As String) As String
    Select Case UCase(sTbl)
    Case "PARTTABLE"
        Select Case UCase(sFld)
        Case "PARTNUM": FriendlyFieldName = "Part Number"
        Case "PADESC": FriendlyFieldName = "Part Description"
        Case "PACLASS": FriendlyFieldName = "Part Class"
        Case "PAPRODCODE": FriendlyFieldName = "Product Code"
        Case "PALEVEL": FriendlyFieldName = "Part Type"
        Case "PAMAKEBUY": FriendlyFieldName = "Part MakeBuy"
        Case "PAQOH": FriendlyFieldName = "Part QOH"
        Case "PALOCATION": FriendlyFieldName = "Part Location"
        
        Case Else
            FriendlyFieldName = sFld
        End Select
    Case "POHDTABLE"
        Select Case UCase(sFld)
        Case "PODATE": FriendlyFieldName = "PO Date"
        Case Else
            FriendlyFieldName = sFld
        End Select
    Case "POITTABLE"
        Select Case UCase(sFld)
        Case "PINUMBER": FriendlyFieldName = "PO Number"
        Case "PIRELEASE": FriendlyFieldName = "PO Released"
        Case "PIITEM": FriendlyFieldName = "PO Item"
        Case "PIREV": FriendlyFieldName = "PO Revision"
        Case "PITYPE": FriendlyFieldName = "Part Type"
        Case "PIPART": FriendlyFieldName = "Part Number"
        Case "PIPDATE": FriendlyFieldName = "Delivery Date"
        Case "PIADATE": FriendlyFieldName = "Date Received"
        Case "PIPQTY": FriendlyFieldName = "PO Item Qty"
        Case "PIAQTY": FriendlyFieldName = "Quantity Received"
        Case "PIAMT": FriendlyFieldName = "PO Item Amount"
        Case "PIESTUNIT": FriendlyFieldName = "PO Item Unit Price"
        Case "PIADDERS": FriendlyFieldName = "<Unused>"
        Case "PILOT": FriendlyFieldName = "Lot Number"
        Case "PIRUNPART": FriendlyFieldName = "MO Number"
        Case "PIRUNNO": FriendlyFieldName = "Run Number"
        Case "PIRUNOPNO": FriendlyFieldName = "Run Operation"
        Case "PISN": FriendlyFieldName = "<Unused>"
        Case "PISNNO": FriendlyFieldName = "<Unused>"
        Case "PIFRTADDERS": FriendlyFieldName = "<Unused>"
        Case "PIWIP": FriendlyFieldName = "<Unused>"
        Case "PIONDOC": FriendlyFieldName = "PO Item On Dock?"
        Case "PIREJECTED": FriendlyFieldName = "PO Item Rejected?"
        Case "PIWASTE": FriendlyFieldName = "<Unused>"
        Case "PIINSBY": FriendlyFieldName = "<Unused>"
        Case "PIINSDATE": FriendlyFieldName = "PIINSDATE"
        Case "PIUSER": FriendlyFieldName = "User Entered"
        Case "PIENTERED": FriendlyFieldName = "Date Item Entered"
        Case "PIODATE": FriendlyFieldName = "<Unused>"
        Case "PIRECEIVED": FriendlyFieldName = "<Unused>"
        Case "PIORIGSCHEDQTY": FriendlyFieldName = "<Unused>"
        Case "PICOMT": FriendlyFieldName = "Item Comment"
        Case "PILOTNUMBER": FriendlyFieldName = "Lot Number"
        Case "PIONDOCK": FriendlyFieldName = "Item On Dock?"
        Case "PIONDOCKINSPECTED": FriendlyFieldName = "On Dock Inspected?"
        Case "PIONDOCKINSPDATE": FriendlyFieldName = "Date On Dock Inspected"
        Case "PIONDOCKREJTAG": FriendlyFieldName = "<Unused>"
        Case "PIONDOCKQTYACC": FriendlyFieldName = "On Dock Qty Accepted"
        Case "PIONDOCKQTYREJ": FriendlyFieldName = "On Dock Qty Rejected"
        Case "PIONDOCKINSPECTOR": FriendlyFieldName = "On Dock Inspector"
        Case "PIONDOCKCOMMENT": FriendlyFieldName = "On Dock Comment"
        Case "PIODDELIVERED": FriendlyFieldName = "On Dock Delivered?"
        Case "PIODDELDATE": FriendlyFieldName = "Date Delivered On Dock"
        Case "PIODDELPSNUMBER": FriendlyFieldName = "<Unused>"
        Case "PIODDELQTY": FriendlyFieldName = "On Dock Delivered Qty"
        Case "PIACCOUNT": FriendlyFieldName = "PO Item GL Account"
        Case "PIVENDOR": FriendlyFieldName = "Vendor"
        Case "PIPICKRECORD": FriendlyFieldName = "PO Item Picked?"
        Case "PIPRESPLITFROM": FriendlyFieldName = "PO Item Split From"
        Case "PIONDOCKQTYWASTE": FriendlyFieldName = "On Dock Qty Waste"
        Case "PIPORIGDATE": FriendlyFieldName = "Original Due Date"
        Case Else
            FriendlyFieldName = sFld
        End Select
    Case "RUNSTABLE"
        Select Case UCase(sFld)
        Case "RUNREF": FriendlyFieldName = "Run PartRef"
        Case "RUNNO": FriendlyFieldName = "Run Number"
        Case "RUNFROZEN": FriendlyFieldName = "Frozen"
        Case "RUNSCHED": FriendlyFieldName = "Scheduled"
        Case "RUNCOMPLETE": FriendlyFieldName = "Complete"
        Case "RUNDIVISION": FriendlyFieldName = "Division"
        Case "RUNPKSTART": FriendlyFieldName = "Pick Start"
        Case "RUNMATL": FriendlyFieldName = "Material"
        Case "RUNTYPE": FriendlyFieldName = "Type"
        Case "RUNCOMMENTS": FriendlyFieldName = "Comments"
        Case "RUNLABOR": FriendlyFieldName = "Labor"
        Case "RUNSTDCOST": FriendlyFieldName = "Standard Cost"
        Case "RUNPURGED": FriendlyFieldName = "Purged"
        Case "RUNQTY": FriendlyFieldName = "Quantity"
        Case "RUNYIELD": FriendlyFieldName = "Yield"
        Case "RUNSTART": FriendlyFieldName = "Start"
        Case "RUNENGREV": FriendlyFieldName = "<Unused>"
        Case "RUNPLDATE": FriendlyFieldName = "Date Printed"
        Case "RUNPKQTY": FriendlyFieldName = "Pick Quantity"
        Case "RUNSTATUS": FriendlyFieldName = "Run Status"
        Case "RUNOPCUR": FriendlyFieldName = "Current Operation"
        Case "RUNPRIORITY": FriendlyFieldName = "Priority"
        Case "RUNREV": FriendlyFieldName = "Revision"
        Case "RUNEXP": FriendlyFieldName = "Expired"
        Case "RUNAPPURGED": FriendlyFieldName = "<Unused>"
        Case "RUNPKPURGED": FriendlyFieldName = "Run Pick Purged"
        Case "RUNBUDLAB": FriendlyFieldName = "<Unused>"
        Case "RUNBUDMAT": FriendlyFieldName = "<Unused>"
        Case "RUNBUDEXP": FriendlyFieldName = "<Unused>"
        Case "RUNBUDOH": FriendlyFieldName = "<Unused>"
        Case "RUNBUDHRS": FriendlyFieldName = "<Unused>"
        Case "RUNCHARGED": FriendlyFieldName = "Run Charged"
        Case "RUNCOST": FriendlyFieldName = "Run Cost"
        Case "RUNOHCOST": FriendlyFieldName = "Overhead Cost"
        Case "RUNCMATL": FriendlyFieldName = "Material Cost"
        Case "RUNCEXP": FriendlyFieldName = "Expense Cost"
        Case "RUNCHRS": FriendlyFieldName = "Hours Cost"
        Case "RUNCLAB": FriendlyFieldName = "Labor Cost"
        Case "RUNAPPDT": FriendlyFieldName = "<Unused>"
        Case "RUNAPPBY": FriendlyFieldName = "<Unused>"
        Case "RUNFINCOMP": FriendlyFieldName = "<Unused>"
        Case "RUNCLOSED": FriendlyFieldName = "Closed"
        Case "RUNREVBY": FriendlyFieldName = "Revised By"
        Case "RUNREVDT": FriendlyFieldName = "Revision Date"
        Case "RUNCREATE": FriendlyFieldName = "Created"
        Case "RUNPRINTED": FriendlyFieldName = "Printed"
        Case "RUNPDATE": FriendlyFieldName = "Print Date"
        Case "RUNRELEASED": FriendlyFieldName = "Released"
        Case "RUNLOTNUMBER": FriendlyFieldName = "Lot Number"
        Case "RUNPARTIALQTY": FriendlyFieldName = "Partial Quantity"
        Case "RUNPARTIALDATE": FriendlyFieldName = "Partial Date"
        Case "RUNSCRAP": FriendlyFieldName = "Scrap"
        Case "RUNREWORK": FriendlyFieldName = "Rework"
        Case "RUNREMAININGQTY": FriendlyFieldName = "Remaining Quantity"
        Case "RUNCANCELED": FriendlyFieldName = "Cancelled"
        Case "RUNCANCELEDBY": FriendlyFieldName = "Cancelled By"
        Case "RUNLASTSPLITREF": FriendlyFieldName = "Last Split Reference"
        Case "RUNLASTSPLITRUNNO": FriendlyFieldName = "Last Split Run Number"
        Case "RUNSPLITFROMREF": FriendlyFieldName = "<Unused>"
        Case "RUNSPLITFROMRUNNO": FriendlyFieldName = "Split From Run Number"
        Case "RUNRTNUM": FriendlyFieldName = "Run Routing Number"
        Case "RUNRTDESC": FriendlyFieldName = "Run Routing Description"
        Case "RUNRTBY": FriendlyFieldName = "Routed By"
        Case "RUNRTAPPBY": FriendlyFieldName = "Routing App By"
        Case "RUNRTAPPDATE": FriendlyFieldName = "Routing App Date"
        Case "RUNMAINTCOSTED": FriendlyFieldName = "<Unused>"
        Case Else
            FriendlyFieldName = sFld
        End Select
    Case "SOHDTABLE"
        Select Case UCase(sFld)
        Case "SONUMBER": FriendlyFieldName = "Sales Order Number"
        Case "SOTYPE": FriendlyFieldName = "Sales Order Type"
        Case "SOCUST": FriendlyFieldName = "Customer"
        Case "SODATE": FriendlyFieldName = "Sales Order Date"
        Case "SOSALESMAN": FriendlyFieldName = "Salesman"
        Case "SOREP": FriendlyFieldName = "Sales Order Rep"
        Case "SOPO": FriendlyFieldName = "PO Number"
        Case "SOSTNAME": FriendlyFieldName = "Ship To Street"
        Case "SOSTADR": FriendlyFieldName = "Ship To Address"
        Case "SOCCONTACT": FriendlyFieldName = "Ship To Contact"
        Case "SOCPHONE": FriendlyFieldName = "Ship To Phone"
        Case "SOCEXT": FriendlyFieldName = "Ship To Extension"
        Case "SOJOBNO": FriendlyFieldName = "Job Number"
        Case "SODIVISION": FriendlyFieldName = "Division"
        Case "SOREGION": FriendlyFieldName = "Region"
        Case "SOBUSUNIT": FriendlyFieldName = "Business Unit"
        Case "SODELDATE": FriendlyFieldName = "SO Date Delivered"
        Case "SOCREATED": FriendlyFieldName = "SO Date Created"
        Case "SOREVISED": FriendlyFieldName = "SO Date Revised"
        Case "SOREMARKS": FriendlyFieldName = "SO Remarks"
        Case "SOSHIPDATE": FriendlyFieldName = "SO Ship Date"
        Case "SOFREIGHTDAYS": FriendlyFieldName = "Freight Days"
        Case "SOCANDATE": FriendlyFieldName = "Date Cancelled"
        Case "SOCANCELED": FriendlyFieldName = "SO Cancelled?"
        Case Else
            FriendlyFieldName = sFld
        End Select
    Case "SOITTABLE"
        Select Case UCase(sFld)
        Case "ITSO": FriendlyFieldName = "Sales Order Number"
        Case "ITNUMBER": FriendlyFieldName = "SO Item Number"
        Case "ITREV": FriendlyFieldName = "Item Revision"
        Case "ITPART": FriendlyFieldName = "Part Number"
        Case "ITQTY": FriendlyFieldName = "SO Qty"
        Case "ITCUSTREQ": FriendlyFieldName = "Date Requested"
        Case "ITCOMMISSION": FriendlyFieldName = "Item Commission"
        Case "ITCOST": FriendlyFieldName = "Item Cost"
        Case "ITINV": FriendlyFieldName = "<Unused>"
        Case "ITDOLLARS": FriendlyFieldName = "Unit Price"
        Case "ITSCHED": FriendlyFieldName = "Scheduled Ship Date"
        Case "ITACTUAL": FriendlyFieldName = "Actual Delivery Date"
        Case "ITINTSPAR1": FriendlyFieldName = "<Unused>"
        Case "ITADJUST": FriendlyFieldName = "Item Adjustment"
        Case "ITORIGQTY": FriendlyFieldName = "<Unused>"
        Case "ITDISCRATE": FriendlyFieldName = "Discount Rate"
        Case "ITDISCAMOUNT": FriendlyFieldName = "Discount Amount"
        Case "ITCGSACCT": FriendlyFieldName = "<Unused>"
        Case "ITINVACCT": FriendlyFieldName = "<Unused>"
        Case "ITDISACCT": FriendlyFieldName = "<Unused>"
        Case "ITSLSTXACCT": FriendlyFieldName = "<Unused>"
        Case "ITTFRPRICE": FriendlyFieldName = "<Unused>"
        Case "ITTAXRATE": FriendlyFieldName = "Tax Rate"
        Case "ITTAXAMT": FriendlyFieldName = "Tax Amount"
        Case "ITSTATE": FriendlyFieldName = "State"
        Case "ITTAXCODE": FriendlyFieldName = "Tax Code"
        Case "ITQTYORIG": FriendlyFieldName = "Original Qty"
        Case "ITDOLLORIG": FriendlyFieldName = "<Unused>"
        Case "ITDISCRORIG": FriendlyFieldName = "<Unused>"
        Case "ITDISCAORIG": FriendlyFieldName = "<Unused>"
        Case "ITTAXRORIG": FriendlyFieldName = "<Unused>"
        Case "ITTAXAORIG": FriendlyFieldName = "<Unused>"
        Case "ITBOSTATIC": FriendlyFieldName = "<Unused>"
        Case "ITBOSTATE": FriendlyFieldName = "<Unused>"
        Case "ITBOCODE": FriendlyFieldName = "<Unused>"
        Case "ITCOSTED": FriendlyFieldName = "Item Costed?"
        Case "ITQUOTEPROB": FriendlyFieldName = "Problem with Quote?"
        Case "ITPSNUMBER": FriendlyFieldName = "Packing Slip Number"
        Case "ITPSITEM": FriendlyFieldName = "Pack Slip Item"
        Case "ITPSCARTON": FriendlyFieldName = "Pack Slip Carton"
        Case "ITPSSHIPNO": FriendlyFieldName = "Pack Slip Shipping Number"
        Case "ITFRTALLOW": FriendlyFieldName = "Freight Allowance"
        Case "ITSCHEDDEL": FriendlyFieldName = "Scheduled Delivery Date"
        Case "ITBOOKDATE": FriendlyFieldName = "Date Booked"
        Case "ITCREATED": FriendlyFieldName = "Item Created Date"
        Case "ITREVISED": FriendlyFieldName = "Item Revised Date"
        Case "ITCANCELDATE": FriendlyFieldName = "Date Item Cancelled"
        Case "ITCANCELED": FriendlyFieldName = "Item Cancelled?"
        Case "ITUSER": FriendlyFieldName = "User"
        Case "ITCANCELEDBY": FriendlyFieldName = "Cancelled By"
        Case "ITCOMMENTS": FriendlyFieldName = "Item Comments"
        Case "ITINVOICE": FriendlyFieldName = "Invoice Number"
        Case "ITSERNO": FriendlyFieldName = "<Unused>"
        Case "ITREVACCT": FriendlyFieldName = "<Unused>"
        Case "ITLOTNUMBER": FriendlyFieldName = "Lot Number"
        Case "ITCUSTITEMNO": FriendlyFieldName = "Customer Item No"
        Case "ITPSSHIPPED": FriendlyFieldName = "Pack Slipped Shipped?"
        Case "ITMOCREATED": FriendlyFieldName = "MO Created?"
        Case "ITFEDTAXRATE": FriendlyFieldName = "Federal Tax Rate"
        Case "ITFEDTAXAMT": FriendlyFieldName = "Federal Tax Amt"
        Case "ITFEDTAXACCT": FriendlyFieldName = "Federal Tax Account"
        Case "ITFEDTAXCODE": FriendlyFieldName = "Federal Tax Code"
        Case "SORD_POINTER": FriendlyFieldName = "<Unused>"
        Case Else
            FriendlyFieldName = sFld
        End Select
        
    Case "MRPLTABLE"
        Select Case UCase(sFld)
        Case "MRP_PARTREF": FriendlyFieldName = "MO PartNumber"
        Case "MRP_PARTQTYRQD": FriendlyFieldName = "Part Qty Rqd"
        Case "MRP_PARTDATERQD": FriendlyFieldName = "Part Date Rqd"
        Case "MRP_SOCUST": FriendlyFieldName = "Sales Order Cust"
        Case "MRP_SONUM": FriendlyFieldName = "Sales Order Number"
        Case "MRP_SOITEM": FriendlyFieldName = "SO Item Number"
        Case "MRP_COMMENT": FriendlyFieldName = "Comment"
        Case Else
            FriendlyFieldName = sFld
        End Select
        
    Case "VW_RNOPPIVOT"
        Select Case UCase(sFld)
        Case "OPNO1": FriendlyFieldName = "OPCenter1"
        Case "OPNO2": FriendlyFieldName = "OPCenter2"
        Case "OPNO3": FriendlyFieldName = "OPCenter3"
        Case "OPNO4": FriendlyFieldName = "OPCenter4"
        Case "OPNO5": FriendlyFieldName = "OPCenter5"
        Case "OPNO6": FriendlyFieldName = "OPCenter6"
        Case "OPNO7": FriendlyFieldName = "OPCenter7"
        Case "OPNO8": FriendlyFieldName = "OPCenter8"
        Case "OPNO9": FriendlyFieldName = "OPCenter9"
        Case "OPNO10": FriendlyFieldName = "OPCenter10"
        Case "OPNO11": FriendlyFieldName = "OPCenter11"
        Case "OPNO12": FriendlyFieldName = "OPCenter12"
        Case "OPNO13": FriendlyFieldName = "OPCenter13"
        Case "OPNO14": FriendlyFieldName = "OPCenter14"
        Case "OPNO15": FriendlyFieldName = "OPCenter15"
        Case "OPNO16": FriendlyFieldName = "OPCenter16"
        Case "OPNO17": FriendlyFieldName = "OPCenter17"
        Case "OPNO18": FriendlyFieldName = "OPCenter18"
        Case "OPNO19": FriendlyFieldName = "OPCenter19"
        Case "OPNO20": FriendlyFieldName = "OPCenter20"
        Case "OPNO21": FriendlyFieldName = "OPCenter21"
        Case "OPNO22": FriendlyFieldName = "OPCenter22"
        Case "OPNO23": FriendlyFieldName = "OPCenter23"
        Case "OPNO24": FriendlyFieldName = "OPCenter24"
        Case "OPNO25": FriendlyFieldName = "OPCenter25"
        Case "OPNO26": FriendlyFieldName = "OPCenter26"
        Case "OPNO27": FriendlyFieldName = "OPCenter27"
        Case Else
            FriendlyFieldName = sFld
        End Select


    Case Else
        FriendlyFieldName = sFld
    End Select
End Function



Public Function ParseFieldName(sItem As String) As String
    Dim iStart, iStop As Integer
    
    ParseFieldName = ""
    If Len(Trim(sItem)) = 0 Then Exit Function
    
    iStart = InStr(1, sItem, "[")
    iStop = InStr(iStart, sItem, "]")
    If iStart = 0 Or iStop = 0 Then
        ParseFieldName = sItem
        Exit Function
    End If
    ParseFieldName = Trim(Mid(sItem, iStart + 1, (iStop - iStart) - 1))
End Function




Public Sub LoadTableColumns(ByVal sTableNme As String, ByRef arrColumns() As String)
    Dim i As Integer
    Dim rdoFlds As ADODB.Recordset
    Erase arrColumns
    ReDim arrColumns(0 To 0) As String
    
    On Error GoTo LTC1
    sSql = "select column_name from information_schema.columns where table_name = '" & sTableNme & "'"
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoFlds, ES_STATIC)
    If bSqlRows Then
        ReDim arrColumns(0 To rdoFlds.RecordCount - 1) As String
        i = 0
        While Not rdoFlds.EOF
            arrColumns(i) = UCase(Trim("" & rdoFlds!column_name))
            i = i + 1
            rdoFlds.MoveNext
        Wend
    End If
    Set rdoFlds = Nothing
    Exit Sub
LTC1:
End Sub


Public Sub SaveLoadListbox(lstLB As ListBox, ByVal sModule As String, ByVal intSaveOrLoad As Integer)
    Dim i As Long
    Dim sListItems, sLBContent As String
    
    Select Case intSaveOrLoad
    Case 1
        sLBContent = ""
        For i = 0 To lstLB.ListCount - 1
            lstLB.Selected(i) = True
            sLBContent = sLBContent & lstLB.List(lstLB.ListIndex) & vbTab
        Next i
        SaveSetting "Esi2000", sModule, lstLB.Name, sLBContent
   Case 2
    lstLB.Clear
    sLBContent = GetSetting("Esi2000", sModule, lstLB.Name, sLBContent)
    sListItems = ""
    For i = 1 To Len(sLBContent)
        If Mid(sLBContent, i, 1) = vbTab Then
            If Len(Trim(sListItems)) > 0 Then lstLB.AddItem sListItems
            sListItems = ""
        Else
            sListItems = sListItems & Mid(sLBContent, i, 1)
        End If
    Next i
    End Select
End Sub




