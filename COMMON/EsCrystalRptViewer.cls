VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EsCrystalRptViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private sReportFileName As String 'local copy
Private sReportTitle As String 'local copy

Public bBold As Byte
Public sReportPath As String
Public bNoCrystal As Boolean
Public iZoomLevel As Integer
Public bUserAction As Boolean
Public iBarOnTop As Byte

Public crxApplication As CRAXDRT.Application
Public crxReport As CRAXDRT.report
Dim crxSections As CRAXDRT.Sections
Dim crxSection As CRAXDRT.Section
Dim crViwerFrm As CRViewerFrm
Dim crxSubreportObject As SubreportObject
Dim CRXReportObject As Object

Public Function AddQtyDetail()

End Function


Public Property Let ReportFile(ByVal rptFileName As String)
    sReportFileName = rptFileName
    
End Property


Public Property Get ReportFile() As String
    ReportFile = sReportFileName
End Property

Public Property Let SetReportTitle(ByVal rptTitle As String)
    If (Not crxReport Is Nothing) Then
        crxReport.ReportTitle = rptTitle
        sReportTitle = rptTitle
    End If
End Property


Public Property Get GetReportTitle() As String
    GetReportTitle = sReportTitle
End Property

Public Sub Init()
    'Open the report
    Set crxApplication = New CRAXDRT.Application
    
    Set crViwerFrm = New CRViewerFrm
    Set crxReport = Nothing
End Sub

Public Function SetReportFileName(ByVal sRptFile As String, ByVal sRptPath As String)
    
On Error GoTo modErr1
    Dim strFullPath As String
    strFullPath = sRptPath & sRptFile
    If (CheckPath(strFullPath)) Then
        Set crxReport = crxApplication.OpenReport(sRptPath & sRptFile, 1)
        sReportFileName = sRptFile
        crxReport.DiscardSavedData
    Else
        MsgBox ("Crystal report File was not found - " & strFullPath)
        Set crxReport = Nothing
    End If
    
    Exit Function
modErr1:
    sProcName = "SetReportFileName"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
    DoModuleErrors Me
    
End Function

    
Public Function SetReportDBParameters(aRptPara As Collection, aRptParaType As Collection)
    Dim j, i As Integer
    If (Not crxReport Is Nothing) Then
        For j = 1 To aRptPara.count
            i = crxReport.ParameterFields.Item(j).ParameterType
            crxReport.ParameterFields.Item(j).ClearCurrentValueAndRange
            If (aRptParaType.Item(j) = "Int") Then
                crxReport.ParameterFields.Item(j).AddCurrentValue CInt(aRptPara.Item(j))
            ElseIf (aRptParaType.Item(j) = "Bool") Then
                crxReport.ParameterFields.Item(j).AddCurrentValue CBool(aRptPara.Item(j))
            Else
                crxReport.ParameterFields.Item(j).AddCurrentValue CStr(aRptPara.Item(j))
            End If
        Next
    End If
End Function

Public Function SetSubRptDBParameters(sSecName As String, aRptPara As Collection, aRptParaType As Collection)
    Dim j, i As Integer
    If (Not crxReport Is Nothing) Then
        Dim crxSubreport As CRAXDRT.report
        
        ' Set subreport table links
         Set crxSections = crxReport.Sections
         For Each crxSection In crxSections ' for each section
              For Each CRXReportObject In crxSection.ReportObjects ' for each reports
                 If CRXReportObject.Kind = crSubreportObject Then ' only if sub report
                    If (CRXReportObject.SubreportName = sSecName) Then
                        Set crxSubreportObject = CRXReportObject
                        Set crxSubreport = crxSubreportObject.OpenSubreport
                        
                        For j = 1 To aRptPara.count
                            i = crxSubreport.ParameterFields.Item(j).ParameterType
                            crxSubreport.ParameterFields.Item(j).ClearCurrentValueAndRange
                            If (aRptParaType.Item(j) = "Int") Then
                                crxSubreport.ParameterFields.Item(j).AddCurrentValue CInt(aRptPara.Item(j))
                            Else
                                crxSubreport.ParameterFields.Item(j).AddCurrentValue CStr(aRptPara.Item(j))
                            End If
                        Next
                        Set crxSubreport = Nothing
                    End If
                 End If
              Next
        Next
    End If
End Function


Public Function SetSortFields(aSortFields As Collection)
    If (Not crxReport Is Nothing) Then
        Dim iTable As Integer
        Dim jField As Integer
        Dim AllSortfld As Integer
        Dim strFieldName As String
        
        For AllSortfld = 1 To aSortFields.count
            strFieldName = CStr(aSortFields(AllSortfld))
            For iTable = 1 To crxReport.Database.Tables.count
               For jField = 1 To crxReport.Database.Tables(iTable).Fields.count
                  If (UCase$(crxReport.Database.Tables(iTable).Fields(jField).Name) = UCase$(strFieldName)) Then
                     crxReport.RecordSortFields.Add crxReport.Database.Tables(iTable).Fields(jField), crAscendingOrder
                     'crxReport.RecordSortFields(jField).Field = crxReport.Database.Tables(iTable).Fields(jField)
                     'crxReport.RecordSortFields(jField).SortDirection = crAscendingOrder
                  End If
               Next ' next field
            Next ' next Table
        Next ' All sort fileds
    
    End If
End Function

Public Function SetReportFormulaFields(aFormulaName As Collection, aFormulaValue As Collection)
    Dim j As Integer
    Dim sFieldName As String
    Dim sFieldValue As String
        
        
    On Error Resume Next    ' We do not want to stop setting other formulas if there is an error in one
    
    If (Not crxReport Is Nothing) Then
        For j = 1 To aFormulaName.count
            sFieldName = aFormulaName.Item(j)
            sFieldValue = aFormulaValue.Item(j)
            crxReport.FormulaFields.GetItemByName(sFieldName).Text = sFieldValue
        Next
    End If
    If Err.Number > 0 Then MsgBox "There are formulas missing from your report" & vbCrLf & "You may have a custom report that requires updating", vbOKOnly
    On Error GoTo 0 'Resume normal error handling
        

End Function

Public Function SetReportFormula(ByVal Index As Integer, ByVal sSql As String)
    
    If (Not crxReport Is Nothing) Then
        crxReport.FormulaFields(Index).Text = sSql
    End If

End Function


Public Function SetReportSection(ByVal strSec As String, ByVal bShow As Boolean)
    
    If (Not crxReport Is Nothing) Then
        Set crxSection = crxReport.Sections(strSec)
        crxSection.Suppress = bShow
    End If

End Function

Public Function SetReportSelectionFormula(sSql As String)
    If (Not crxReport Is Nothing) Then
        Dim crxSubreport As CRAXDRT.report
        
        crxReport.FormulaSyntax = crCrystalSyntaxFormula
        crxReport.RecordSelectionFormula = sSql
    End If
End Function

Public Function GetReportSelectionFormula()
    If (Not crxReport Is Nothing) Then
        Dim crxSubreport As CRAXDRT.report
        crxReport.FormulaSyntax = crCrystalSyntaxFormula
        GetReportSelectionFormula = crxReport.RecordSelectionFormula
    End If
End Function

Public Function SetSubRptSelFormula(sSecName As String, sSubSql As String)
    If (Not crxReport Is Nothing) Then
        Dim crxSubreport As CRAXDRT.report
        
        ' Set subreport table links
         Set crxSections = crxReport.Sections
         For Each crxSection In crxSections ' for each section
              For Each CRXReportObject In crxSection.ReportObjects ' for each reports
                 If CRXReportObject.Kind = crSubreportObject Then ' only if sub report
                    If (CRXReportObject.SubreportName = sSecName) Then
                        Set crxSubreportObject = CRXReportObject
                        Set crxSubreport = crxSubreportObject.OpenSubreport
                        ' sub report selection
                        crxSubreport.FormulaSyntax = crCrystalSyntaxFormula
                        crxSubreport.RecordSelectionFormula = sSubSql
                        Set crxSubreport = Nothing
                    End If
                 End If
              Next
        Next
    End If
End Function


Public Function SetReportDistinctRecords(ByVal bFlag As Boolean)
    If (Not crxReport Is Nothing) Then
        crxReport.EnableSelectDistinctRecords = bFlag
    End If
End Function


Public Function CRViewerSize(frm As Form)
    Dim bWindowSize As Byte
    Dim A As Integer
    Dim b As Integer
        
    On Error Resume Next
    sProcName = "printreport"
    frm.optPrn.enabled = False
    frm.optDis.enabled = False

    A = Screen.TwipsPerPixelX
    b = Screen.TwipsPerPixelY
    
    ' Set the drill down option of the report
    frm.report.DiscardSavedData
    bWindowSize = GetSetting("Esi2000", "System", "ReportMax", bWindowSize)
    If bWindowSize = 0 Then
        If iBarOnTop = False Then
            crViwerFrm.WindowState = 0
            crViwerFrm.Top = 1280 / b
            crViwerFrm.Height = MdiSect.Height - 1750
            crViwerFrm.Left = 220 / A
            crViwerFrm.Width = MdiSect.Width - 3330
        Else
            crViwerFrm.Top = 1280 / b
            crViwerFrm.Height = (MdiSect.Height / b) - (1750 / b)
            crViwerFrm.Left = 220 / A
            crViwerFrm.Width = (MdiSect.Width / A) - (750 / A)
        End If
    Else
        crViwerFrm.WindowState = 2
        crViwerFrm.Top = 0
        crViwerFrm.Height = Screen.Height
        crViwerFrm.Left = 0
        crViwerFrm.Width = Screen.Width
    End If
    

End Function

Public Sub OpenCrystalReportObject(frm As Form, aFormulaName As Collection, _
               Optional ByVal nCopies As Integer = 1, _
               Optional sDefaultPrinter As String = "", _
               Optional ByVal bNoView As Boolean = False, _
               Optional bModal As Boolean = False)
    Dim b As Byte
    Dim bInstr As Byte
    Dim FormDriver As String
    Dim FormPort As String
    Dim FormPrinter As String
    
    'MouseCursor ccHourglass
    
    On Error GoTo modErr1
    
    ' if debugging, show info
    If Debugging() Then
        Select Case MsgBox(DisplayCRParameterString(frm, aFormulaName) & vbCrLf & "Proceed?", vbYesNo + vbQuestion)
        Case vbNo
            On Error Resume Next          'there are not always buttons with these names
            frm.optPrn.enabled = True
            frm.optDis.enabled = True
            MouseCursor ccDefault
            Exit Sub
        End Select
    End If
    SetReportTitle = frm.Caption
    
    'MsgBox DisplayCRParameterString(frm, aFormulaName) & vbCrLf & "Proceed?", vbYesNo + vbQuestion
    
    On Error Resume Next
    'SetTableConnection
   
'        If (Not crxReport Is Nothing) Then
'            ' Set the Viewer form caption
'            crViwerFrm.Caption = frm.Caption
'            crViwerFrm.ShowReport crxReport
'            crViwerFrm.Show
'        End If


    If frm.optPrn.Value = True _
            Or frm.optPrn.Value = vbChecked _
            Or bNoView = True Then

       On Error Resume Next
       FormPrinter = Trim(frm.lblPrinter)
       If Err > 0 Then FormPrinter = ""
       If FormPrinter = "Default Printer" Then FormPrinter = ""
       If Len(sDefaultPrinter) > 0 Then FormPrinter = sDefaultPrinter
'       If Not bBold Then
'          MdiSect.crw.SectionFont(0) = "ALL;;;;N"
'       Else
'          MdiSect.crw.SectionFont(0) = "ALL;;;;Y"
'       End If
       If Len(Trim(FormPrinter)) > 0 Then
          b = GetPrinterPort(FormPrinter, FormDriver, FormPort)
       Else
          FormPrinter = ""
          FormDriver = ""
          FormPort = ""
       End If
         Dim oldPageOrien As CRPaperOrientation
         Dim oldPageSize As CRPaperSize
         Dim oldPageSource As CRPaperSource
         Dim bByte As Boolean
        If (Not crxReport Is Nothing) Then
            
            If (FormPrinter <> "") Then
               
               
               oldPageOrien = crxReport.PaperOrientation
               oldPageSize = crxReport.PaperSize
               'oldPageSource = crxReport.PaperSource
               
               crxReport.SelectPrinter FormDriver, FormPrinter, FormPort

               If (oldPageSize <> crDefaultPaperSize) Then
                  crxReport.PaperOrientation = oldPageOrien
                  crxReport.PaperSize = oldPageSize
               End If
               
               
'               oldPageOrien = crxReport.PaperOrientation
'               oldPageSize = crxReport.PaperSize
'               'oldPageSource = crxReport.PaperSource
'
'               ' MM Casgas Label Issue
'               bByte = InStr(1, FormPrinter, "GK420d")
'
'               If bByte = False Then
'                  crxReport.SelectPrinter FormDriver, FormPrinter, FormPort
'               End If
'
'               If (oldPageSize <> crDefaultPaperSize) Then
'                  crxReport.PaperOrientation = oldPageOrien
'                  crxReport.PaperSize = oldPageSize
'               End If
            
            End If
            
            'crxReport.PaperSource = oldPageSource
            crxReport.PrintOut False, nCopies
            ' unload the form
            Unload crViwerFrm
        End If

    Else

        If (Not crxReport Is Nothing) Then
            ' Set the Viewer form caption
            crViwerFrm.Caption = frm.Caption
            crViwerFrm.ShowReport crxReport
            If (bModal = False) Then
               crViwerFrm.Show
            Else
               crViwerFrm.Show vbModal, frm
            End If
            
            
        End If
    End If
    
    frm.optPrn.enabled = True
    frm.optDis.enabled = True
    MouseCursor ccDefault
    Exit Sub
    
modErr1:
    sProcName = "OpenCrystalReportObj"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
modErr2:
    DoModuleErrors frm
    On Error Resume Next 'following buttons may not exist
    frm.optPrn.enabled = True
    frm.optDis.enabled = True
    MouseCursor ccDefault
End Sub

Public Sub SetDbTableConnection(Optional ByVal bStorProc As Boolean = False)

   'when the rpt calls a stored procedure for its dataset, DO NOT SET bStorProc = True
   'this results in the parameters being set blank
   ' 3/13/19: The above does not seem to be true any longer
   'but now the parameter is no longer required as stored procedures are automatically detected below
    
   Dim bInstr As Byte
   Dim crxDatabaseTable As CRAXDRT.DatabaseTable
   Dim crxSubreport As CRAXDRT.report
        
   On Error GoTo modErr1
   
   Err.Clear
   'DNS sDsn = "FusionDB" '"ESI2000" 'is for me 'TODO: RegisterSqlDsn ("ESI2000")
   sDsn = GetConfUserSetting(USERSETTING_SqlDsn)

   
   Dim strServer As String
   ' DNS strServer = UCase(GetUserSetting(USERSETTING_ServerName))
   strServer = UCase(GetConfUserSetting(USERSETTING_ServerName))
  
   If (Not crxReport Is Nothing) Then
      crxReport.Database.LogOnServer "crdb_odbc.dll", sDsn, sDataBase, sSaAdmin, sSaPassword
      
      'MsgBox "LogOnServer:" & CStr(Err.Number) & ":" & Err.Description
      For Each crxDatabaseTable In crxReport.Database.Tables
    
         'crxDatabaseTable.ConnectBufferString = sConStr ' sDsn, sDataBase, sSaAdmin, sSaPassword
         crxDatabaseTable.SetLogOnInfo sDsn, sDataBase, sSaAdmin, sSaPassword
         
         ' 3/13/2019: if the "table" is in fact a stored procedure, repoint the catalog = the database from the ODBC connection
         'bStorProc is redundant and can be removed as a parameter at a future date.  Any table suffixed with ;1 is a stored procedure
         'note: when you do this, you wipe out parameters that have been passed before this call.  Code has been changed to make
         'the parameter call after the SetDbTableConnection call, but just for safety's sake, we are only going to make this call
         'when the current database name is not already correct
         If bStorProc Or InStr(1, crxDatabaseTable.Name, ";1") > 1 Then
            If InStr(1, crxDatabaseTable.ConnectBufferString, ";Database=" & sDataBase & ";") = 0 Then
                crxDatabaseTable.Location = sDataBase & ".dbo." & crxDatabaseTable.Name
            End If
         End If
      Next crxDatabaseTable
    
    ' Set subreport table links
     Set crxSections = crxReport.Sections
    
     For Each crxSection In crxSections
          For Each CRXReportObject In crxSection.ReportObjects
             If CRXReportObject.Kind = crSubreportObject Then
               Set crxSubreportObject = CRXReportObject
               Set crxSubreport = crxSubreportObject.OpenSubreport
    
               For Each crxDatabaseTable In crxSubreport.Database.Tables
                  'crxDatabaseTable.ConnectBufferString = sConStr ' sDsn, sDataBase, sSaAdmin, sSaPassword
                  crxDatabaseTable.SetLogOnInfo sDsn, sDataBase, sSaAdmin, sSaPassword
               Next
               Set crxSubreport = Nothing
             End If
          Next
      Next
   End If
Exit Sub
   
modErr1:
    sProcName = "SetDbTableConnection"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
   
End Sub


Public Sub ClearFieldCollection(ByRef aForFields As Collection)
    Dim arrayCount, j As Integer
    arrayCount = aForFields.count
    For j = 1 To arrayCount
        aForFields.Remove (1) 'first in the item
    Next
End Sub

Public Sub ShowDrillDown(bDrillDown As Boolean)
    crViwerFrm.CRViewerObj.EnableDrillDown = bDrillDown
    
End Sub

Public Sub ShowGroupTree(bGrpTree As Boolean)
    crViwerFrm.CRViewerObj.DisplayGroupTree = bGrpTree
End Sub

Public Function DisplayCRParameterString(frm As Form, aForFields As Collection)
    If (Not crxReport Is Nothing) Then
    
        Dim S As String
        S = "Crystal Reports parameters for report " & crxReport.ReportTitle & " called from " & frm.Name & " form" & vbCrLf
        Dim j As Integer
        
        On Error Resume Next
        If Not aForFields Is Nothing Then
            For j = 1 To aForFields.count
               If j > 13 Then Exit For    ' error for imainc mo
                S = S & "(" & j & ") " & aForFields.Item(j) & " : " & crxReport.FormulaFields.GetItemByName(aForFields.Item(j)).Text & vbCrLf
            Next
         End If
        
        'now add StoredProcParam
        ' this sometimes blows up with sp parameters
        Dim spParamsFound As Boolean
        For j = 1 To crxReport.ParameterFields.count
          'If crxReport.ParameterFields.Item(j).GetNthCurrentValue(1) <> "" Then
            If j > 13 Then Exit For    'error for imainc mo
             If Not spParamsFound Then
                spParamsFound = True
                S = S & "Parameters :" & vbCrLf
             End If
             'S = S & "(" & j & ") " & crxReport.ParameterFields.Item(j).Name & " : " & crxReport.ParameterFields.Item(j).GetNthCurrentValue(1) & vbCrLf
             S = S & "(" & j & ") " & crxReport.ParameterFields.Item(j).Name & " : "
             S = S & crxReport.ParameterFields.Item(j).GetNthCurrentValue(1) & vbCrLf
          'End If
        Next
        On Error GoTo 0
        
        S = S & "SQL: " & crxReport.RecordSelectionFormula & vbCrLf
        DisplayCRParameterString = S
    
    
        Dim strError As String
        Dim blnHasNoError As Boolean
        Dim i As Integer
        
        For i = 1 To crxReport.FormulaFields.count
             crxReport.FormulaFields(i).Check blnHasNoError, strError
           
             If blnHasNoError = True Then
             'fine, there are no errors
             Else
                 MsgBox (crxReport.FormulaFields(i).FormulaFieldName & " --  " & strError)
                 'strText = CrRep.FormulaFields(i).Text
             End If
         Next i
    
    End If
End Function


Public Function GetPrinterPort(devPrinter As String, devDriver As String, devPort As String) As Byte
   Dim SysPrinter As Printer
   For Each SysPrinter In Printers
      If Trim(SysPrinter.DeviceName) = devPrinter Then
         devDriver = SysPrinter.DriverName
         devPort = SysPrinter.Port
         Exit For
      End If
   Next
   
End Function

Public Function GroupingByField(ByVal sFieldName As String) As Boolean
    On Error Resume Next
    Dim CRXGroupField As CRAXDRT.GroupNameFieldDefinition
    
    GroupingByField = False
    Set CRXGroupField = crxReport.GroupNameFields
    For Each CRXGroupField In crxReport.GroupNameFields
        If InStr(1, UCase(CRXGroupField.Name), UCase(sFieldName)) > 0 Then
            GroupingByField = True
            Exit For
        End If
    Next
    
    Set CRXGroupField = Nothing
End Function

