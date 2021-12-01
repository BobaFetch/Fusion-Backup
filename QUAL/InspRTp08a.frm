VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTp08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inspections By Inspector"
   ClientHeight    =   3075
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3075
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTp08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox optGrp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      ToolTipText     =   "Chart Results"
      Top             =   2280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "InspRTp08a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "InspRTp08a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbIns 
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Leading Characters Or Select From List"
      Top             =   1080
      Width           =   1665
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3075
      FormDesignWidth =   7275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   4
      Left            =   5760
      TabIndex        =   14
      Top             =   1800
      Width           =   1308
   End
   Begin VB.Label lblIns 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   13
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Chart"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Chart Results"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   1
      Left            =   5760
      TabIndex        =   11
      Top             =   1440
      Width           =   1308
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspection Report Dates"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   9
      Top             =   1800
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   360
      Picture         =   "InspRTp08a.frx":0AB6
      ToolTipText     =   "Chart Results"
      Top             =   0
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspector(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1665
   End
End
Attribute VB_Name = "InspRTp08a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/1/05 Changed date handling
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbIns_Click()
   GetInspector
   
End Sub

Private Sub cmbIns_LostFocus()
   cmbIns = CheckLen(cmbIns, 12)
   If Trim(cmbIns) = "" Then cmbIns = "ALL"
   GetInspector
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillInspectorsAll"
   LoadComboBox cmbIns, -1
   If cmbIns.ListCount > 0 Then cmbIns = cmbIns.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set InspRTp08a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sInsp As String
   Dim sBegDate As String
   Dim sEnddate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   
   If Not IsDate(txtBeg) Then
      sBegDate = "1995,01,01"
   Else
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEnddate = "2024,12,31"
   Else
      sEnddate = Format(txtEnd, "yyyy,mm,dd")
   End If
   On Error GoTo DiaErr1
   If cmbIns <> "ALL" Then sInsp = Compress(cmbIns) Else sInsp = ""
   sCustomReport = GetCustomReport("quarj08")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowGroup"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbIns & "." _
                        & "And Dates From " & txtBeg & " To " & txtEnd) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add optGrp.value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   sSql = "{RinsTable.INSID} LIKE '" & sInsp & "*' AND " _
          & "{RjhdTable.REJDATE} In Date(" & sBegDate & ") " _
          & "To Date(" & sEnddate & ")"
   cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport1()
   Dim sInsp As String
   Dim sBegDate As String
   Dim sEnddate As String
   MouseCursor 13
   
   If Not IsDate(txtBeg) Then
      sBegDate = "1995,01,01"
   Else
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEnddate = "2024,12,31"
   Else
      sEnddate = Format(txtEnd, "yyyy,mm,dd")
   End If
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   If cmbIns <> "ALL" Then sInsp = Compress(cmbIns) Else sInsp = ""
   sCustomReport = GetCustomReport("quarj08")
   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.Crw.Formulas(1) = "Includes='" & cmbIns & "." _
                        & "And Dates From " & txtBeg & " To " & txtEnd & "'"
   MdiSect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   sSql = "{RinsTable.INSID} LIKE '" & sInsp & "*' AND " _
          & "{RjhdTable.REJDATE} In Date(" & sBegDate & ") " _
          & "To Date(" & sEnddate & ")"
   If optGrp.value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(0) = "REPORTFTR.0.0;F;;;"
   Else
      MdiSect.Crw.SectionFormat(0) = "REPORTFTR.0.0;T;;;"
   End If
   MdiSect.Crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sIns As String * 12
   sIns = cmbIns
   sOptions = sIns & Trim(txtBeg) & Trim(txtEnd) _
              & Trim(str(optGrp.value))
   SaveSetting "Esi2000", "EsiQual", "rj08", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiQual", "rj08", Trim(sOptions))
   If Len(Trim(sOptions)) > 8 Then
      cmbIns = Trim(Left(sOptions, 12))
      optGrp = Val(Right(sOptions, 1))
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 3)
   
End Sub

Private Sub Image1_Click()
   If optGrp = vbUnchecked Then optGrp = vbChecked Else _
               optGrp = vbUnchecked
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub GetInspector()
   Dim RdoIns As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT INSID,INSFIRST,INSMIDD,INSLAST FROM " _
          & "RinsTable WHERE INSID='" & Compress(cmbIns) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoIns, ES_FORWARD)
   If bSqlRows Then
      With RdoIns
         cmbIns = "" & Trim(!INSID)
         lblIns = "" & Trim(!INSFIRST) & " " _
                  & Trim(!INSMIDD) & " " & Trim(!INSLAST)
      End With
   Else
      If cmbIns = "ALL" Then
         lblIns = "All Inspectors."
      Else
         lblIns = "Inspectors Starting With " & Trim(cmbIns) & "."
      End If
   End If
   Set RdoIns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getinspect"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDate(txtBeg)
   
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDate(txtEnd)
   
End Sub
