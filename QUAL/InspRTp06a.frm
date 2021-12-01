VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTp06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inspection Reports By Vendor"
   ClientHeight    =   3555
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3555
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTp06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   288
      Left            =   4200
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   288
      Left            =   2040
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame z2 
      Height          =   540
      Left            =   2040
      TabIndex        =   14
      Top             =   2280
      Width           =   3855
      Begin VB.OptionButton optCom 
         Caption         =   "Both"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optCom 
         Caption         =   "Incomplete"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optCom 
         Caption         =   "Complete"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox optGrp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "Chart Results"
      Top             =   2880
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "InspRTp06a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "InspRTp06a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select, Leading Char(s) Or Blank For ALL"
      Top             =   1080
      Width           =   1555
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3555
      FormDesignWidth =   7260
   End
   Begin VB.Label lblVEName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2040
      TabIndex        =   21
      Top             =   1440
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   3
      Left            =   5520
      TabIndex        =   19
      Top             =   1800
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   5520
      TabIndex        =   18
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspect Report Dates"
      Height          =   288
      Index           =   5
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   288
      Index           =   1
      Left            =   3240
      TabIndex        =   16
      Top             =   1800
      Width           =   912
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Corrective Action"
      Height          =   288
      Index           =   4
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "Chart Results"
      Top             =   2520
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   2
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "Chart Results"
      Top             =   2160
      Width           =   1692
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   360
      Picture         =   "InspRTp06a.frx":0AB6
      ToolTipText     =   "Chart Results"
      Top             =   0
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chart"
      Height          =   288
      Index           =   11
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Chart Results"
      Top             =   2880
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "InspRTp06a"
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

Private Sub cmbVnd_Click()
   GetThisVendor
   
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 12)
   If cmbVnd = "" Then cmbVnd = "ALL"
   GetThisVendor
   
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
   sSql = "Qry_FillVendors"
   LoadComboBox cmbVnd
   If cmbVnd.ListCount > 0 Then cmbVnd = cmbVnd.List(0)
   GetThisVendor
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
   Set InspRTp06a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sCode As String
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
   If Trim(cmbVnd) <> "ALL" Then sCode = Compress(cmbVnd)
   
   On Error GoTo DiaErr1
   sCustomReport = GetCustomReport("quarj06")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowGroup"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Vendor(s) " & CStr(cmbVnd & "... And " _
                        & "Dates From " & txtBeg & " To " & txtEnd) & "...'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add optGrp.value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   
   sSql = "{RjhdTable.REJVENDOR} LIKE '" & sCode & "*' AND " _
          & "{RjhdTable.REJDATE} In Date(" & sBegDate & ") " _
          & "To Date(" & sEnddate & ")"
   
   If optCom(0).value = True Then
      sSql = sSql & " AND {RjitTable.RITACT}=1"
   Else
      If optCom(1).value = True Then sSql = sSql & " AND {RjitTable.RITACT}=0"
   End If
   sSql = sSql & " and {RjhdTable.REJTYPE} = 'V'"
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
   Dim sCode As String
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
   If Trim(cmbVnd) <> "ALL" Then sCode = Compress(cmbVnd)
   
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("quarj06")
   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.Crw.Formulas(1) = "Includes='Vendor(s) " & cmbVnd & "... And " _
                        & "Dates From " & txtBeg & " To " & txtEnd & "...'"
   MdiSect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   sSql = "{RjhdTable.REJVENDOR} LIKE '" & sCode & "*' AND " _
          & "{RjhdTable.REJDATE} In Date(" & sBegDate & ") " _
          & "To Date(" & sEnddate & ")"
   
   If optCom(0).value = True Then
      sSql = sSql & " AND {RjitTable.RITACT}=1"
   Else
      If optCom(1).value = True Then sSql = sSql & " AND {RjitTable.RITACT}=0"
   End If
   If optGrp.value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(0) = "REPORTFTR.0.1;F;;;"
   Else
      MdiSect.Crw.SectionFormat(0) = "REPORTFTR.0.1;T;;;"
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
   Dim sCode As String * 10
   Dim sCom As String
   
   If optCom(0).value = True Then sCom = 0
   If optCom(1).value = True Then sCom = 1
   If optCom(2).value = True Then sCom = 2
   
   sCode = cmbVnd
   sOptions = sCode & txtBeg & txtEnd & sCom _
              & Trim(str(optGrp.value))
   SaveSetting "Esi2000", "EsiQual", "rj06", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim sCom As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiQual", "rj06", Trim(sOptions))
   If Len(Trim(sOptions)) > 0 Then
      cmbVnd = Trim(Mid(sOptions, 1, 10))
      sCom = Mid(sOptions, 27, 1)
      optGrp = Val(Mid(sOptions, 30, 1))
      optCom(Val(sCom)).value = True
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 3)
   
End Sub

Private Sub Image1_Click()
   If optGrp = vbUnchecked Then optGrp = vbChecked Else _
               optGrp = vbUnchecked
   
End Sub

Private Sub optCom_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optGrp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
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
