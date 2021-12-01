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

Public crxApplication As CRAXDRT.Application
Public crxReport As CRAXDRT.report



Public Property Let ReportFile(ByVal rptFileName As String)
    sReportFileName = rptFileName
End Property


Public Property Get ReportFile() As String
    ReportFile = sReportFileName
End Property

Public Property Let SetReportTitle(ByVal rptTitle As String)
    If (IsObject(crxReport)) Then
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
    Set crxReport = Nothing
End Sub

Public Function SetReportFileName(ByVal sRptFile As String, ByVal sRptPath As String)
    
    Set crxReport = crxApplication.OpenReport(sRptPath & sRptFile, 1)
    sReportFileName = sRptFile
    
End Function

    
Public Function SetReportParameters(aRptPara As Collection)
    Dim j, i As Integer
    If (IsObject(crxReport)) Then
        For j = 1 To aRptPara.count
            i = crxReport.ParameterFields.Item(j).ParameterType
            crxReport.ParameterFields.Item(j).ClearCurrentValueAndRange
            crxReport.ParameterFields.Item(j).AddCurrentValue CStr(aRptPara.Item(j))
        Next
    End If

End Function

Public Function SetReportFormulaFields(aFormulaName As Collection, aFormulaValue As Collection)
    Dim j As Integer
    Dim sFieldName As String
    Dim sFieldValue As String
    
    If (IsObject(crxReport)) Then
        For j = 1 To aFormulaName.count
            sFieldName = aFormulaName.Item(j)
            sFieldValue = aFormulaValue.Item(j)
            'crxReport.FormulaFields.GetItemByName(sFieldName).Text = CInt(sFieldValue)
            crxReport.FormulaFields.GetItemByName(sFieldName).Text = sFieldValue
        Next
    End If

End Function


Public Function SetReportSelectionFormula(sSql As String)
    If (IsObject(crxReport)) Then
        crxReport.RecordSelectionFormula = sSql
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
            CRViewerFrm.WindowState = 0
            CRViewerFrm.top = 1280 / b
            CRViewerFrm.Height = MdiSect.Height - 1750
            CRViewerFrm.Left = 220 / A
            CRViewerFrm.Width = MdiSect.Width - 3330
        Else
            CRViewerFrm.top = 1280 / b
            CRViewerFrm.Height = (MdiSect.Height / b) - (1750 / b)
            CRViewerFrm.Left = 220 / A
            CRViewerFrm.Width = (MdiSect.Width / A) - (750 / A)
        End If
    Else
        CRViewerFrm.WindowState = 2
        CRViewerFrm.top = 0
        CRViewerFrm.Height = Screen.Height
        CRViewerFrm.Left = 0
        CRViewerFrm.Width = Screen.Width
    End If
    

End Function

Public Sub OpenCrystalReportObject(frm As Form, aFormulaName As Collection)
    'fires crystal and sets zoom level
    'if user has selected one
    Dim b As Byte
    Dim bInstr As Byte
    Dim FormDriver As String
    Dim FormPort As String
    Dim FormPrinter As String
    
    MouseCursor ccHourglass
    
    On Error GoTo ModErr1
    
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
    
    On Error Resume Next
    SetTableConnection
   
    If frm.optPrn.Value = True Or frm.optPrn.Value = vbChecked Then
       On Error Resume Next
       FormPrinter = Trim(frm.lblPrinter)
       If Err > 0 Then FormPrinter = ""
       If FormPrinter = "Default Printer" Then FormPrinter = ""
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
       crxReport.SelectPrinter FormDriver, FormPrinter, FormPort
       crxReport.PrintOut False
       ' unload the form
       Unload CRViewerFrm

    Else
        ' Set the Viewer form caption
        CRViewerFrm.Caption = frm.Caption
        
        CRViewerFrm.ShowReport crxReport
        CRViewerFrm.Show
    End If
    
    frm.optPrn.enabled = True
    frm.optDis.enabled = True
    MouseCursor ccDefault
    Exit Sub
    
ModErr1:
    sProcName = "SetCrystalAction"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
modErr2:
    DoModuleErrors frm
    On Error Resume Next 'following buttons may not exist
    frm.optPrn.enabled = True
    frm.optDis.enabled = True
    MouseCursor ccDefault
End Sub

Public Sub SetTableConnection()
    
    Dim bInstr As Byte
    Dim crxDatabaseTable As CRAXDRT.DatabaseTable
    
    Err.Clear
    sDsn = RegisterSqlDsn("ESI2000")
    If Err = 20599 Then
       SaveUserSetting USERSETTING_SqlDsn, sDsn
    Else
       If Err > 0 Then
          'any other errors
          '8/19/05 find the report and do it twice
          CurrError.Number = Err.Number
          CurrError.Description = Err.Description
          sProcName = MdiSect.crw.ReportFileName
          sProcName = Left(sProcName, Len(sProcName) - 4)
          bInstr = InStr(4, sProcName, "\")
          sProcName = Right$(sProcName, Len(sProcName) - bInstr)
          bInstr = InStr(4, sProcName, "\")
          sProcName = Right$(sProcName, Len(sProcName) - bInstr)
          GoTo ModErr1
       End If
    End If
    
    
    For Each crxDatabaseTable In crxReport.Database.Tables
        crxDatabaseTable.SetLogOnInfo sDsn, sDataBase, sSaAdmin, sSaPassword
    Next crxDatabaseTable
   
   Exit Sub
   
ModErr1:
   On Error GoTo 0
   
End Sub


Public Sub ClearFieldCollection(ByRef aForFields As Collection)
    Dim arrayCount, j As Integer
    arrayCount = aForFields.count
    For j = 1 To arrayCount
        aForFields.Remove (1) 'first in the item
    Next
End Sub

Public Sub ShowDrillDown(bDrillDown As Boolean)
    CRViewerFrm.CRViewerObj.EnableDrillDown = bDrillDown
    
End Sub

Public Sub ShowGroupTree(bGrpTree As Boolean)
    CRViewerFrm.CRViewerObj.DisplayGroupTree = bGrpTree
End Sub

Public Function DisplayCRParameterString(frm As Form, aForFields As Collection)
    If (IsObject(crxReport)) Then
    
        Dim s As String
        s = "Crystal Reports parameters for report " & crxReport.ReportTitle & " called from " & frm.Name & " form" & vbCrLf
        Dim j As Integer
        
        For j = 1 To aForFields.count
            On Error Resume Next
            s = s & "(" & j & ") " & aForFields.Item(j) & " : " & crxReport.FormulaFields.GetItemByName(aForFields.Item(j)).Text & vbCrLf
        Next
        
        'now add StoredProcParam
        Dim spParamsFound As Boolean
        For j = 1 To crxReport.ParameterFields.count
          If crxReport.ParameterFields.Item(j).GetNthCurrentValue(1) <> "" Then
             If Not spParamsFound Then
                spParamsFound = True
                s = s & "Parameters :" & vbCrLf
             End If
             s = s & "(" & j & ") " & crxReport.ParameterFields.Item(j).Name & " : " & crxReport.ParameterFields.Item(j).GetNthCurrentValue(1) & vbCrLf
          End If
        Next
        
        s = s & "SQL: " & crxReport.RecordSelectionFormula & vbCrLf
        DisplayCRParameterString = s
    End If
End Function


