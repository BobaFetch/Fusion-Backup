Attribute VB_Name = "PrintReports"
Public Sub PrintReportSalesOrderAllocations(frm As Form, PartNumber As String, Runno As Long)
   
   On Error GoTo whoops
   'SetMdiReportsize MDISect
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
  
   
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'   MDISect.Crw.Formulas(2) = "PartNumber='" & Trim(PartNumber) & "'"
'   MDISect.Crw.Formulas(3) = "RunNumber='" & Runno & "'"

   sCustomReport = GetCustomReport("prdsh17")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "PartNumber"
    aFormulaName.Add "RunNumber"
    aFormulaName.Add "Run"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & Trim(PartNumber) & "'")
    aFormulaValue.Add CStr("'" & Runno & "'")
    aFormulaValue.Add Runno
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

'   sSql = "{RunsTable.RUNREF}='" & Compress(cmbPrt) & "' " _
'          & "AND {RunsTable.RUNNO}=" & Val(cmbRun) & " "
'   If Val(txtSon) > 0 Then
'      sSql = sSql & "AND {RnalTable.RASO}=" & Val(txtSon) & " "
'   End If
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction frm
   
   sSql = "{RunsTable.RUNREF} = '" & Trim(Compress(PartNumber)) & "' and {RunsTable.RUNNO} = " & Val(Runno)
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject frm, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

   Exit Sub
   
whoops:
   sProcName = "PrintSalesOrderAllocations"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
End Sub


