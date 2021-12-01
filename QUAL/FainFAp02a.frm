VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form FainFAp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Of First Article Reports"
   ClientHeight    =   3645
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
   ScaleHeight     =   3645
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "FainFAp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optCmp 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   1920
      TabIndex        =   7
      Top             =   2640
      Width           =   735
   End
   Begin VB.Frame Z2 
      Height          =   520
      Left            =   1920
      TabIndex        =   18
      Top             =   1800
      Width           =   3735
      Begin VB.OptionButton optBoth 
         Caption         =   "Both"
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   200
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optInc 
         Caption         =   "Incomplete"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   200
         Width           =   1215
      End
      Begin VB.OptionButton OptCom 
         Caption         =   "Complete"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   200
         Width           =   1215
      End
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Reports"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   1920
      TabIndex        =   6
      Top             =   2400
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
         Picture         =   "FainFAp02a.frx":07AE
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
         Picture         =   "FainFAp02a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3645
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   7
      Left            =   5640
      TabIndex        =   20
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   0
      Left            =   3120
      TabIndex        =   17
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Created From"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Number(s)"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   13
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Tag             =   " "
      Top             =   1920
      Width           =   1425
   End
End
Attribute VB_Name = "FainFAp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillFirstArticles "
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbPrt_Validate(Cancel As Boolean)
   cmbPrt = CheckLen(cmbPrt, 30)
   If Trim(cmbPrt) = "" Then cmbPrt = "ALL"
   
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

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
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
   Set FainFAp02a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sReport As String
   Dim sBegDate As String
   Dim sEnddate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   
   If IsDate(txtBeg) Then
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   Else
      sBegDate = "1995,01,01"
   End If
   If IsDate(txtEnd) Then
      sEnddate = Format(txtEnd, "yyyy,mm,dd")
   Else
      sEnddate = "2020,12,31"
   End If
   If Trim(cmbPrt) <> "ALL" Then sReport = Compress(cmbPrt)
   On Error GoTo DiaErr1
   sCustomReport = GetCustomReport("quafa02")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDetails"
   aFormulaName.Add "ShowComp"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbPrt & ", From " _
                        & txtBeg & " Through " & txtEnd) & "...'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add optDet.value
   aFormulaValue.Add optCmp.value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   
   sSql = "{FahdTable.FA_REF} LIKE '" & sReport & "*' " _
          & "AND {FahdTable.FA_CREATED} In Date(" & sBegDate & ") " _
          & "To Date(" & sEnddate & ") "
   If optCom.value = True Then
      sSql = sSql & "AND {FahdTable.FA_COMPLETE}=1 "
   Else
      sSql = sSql & "AND {FahdTable.FA_COMPLETE}=0 "
   End If
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
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport1()
   Dim sReport As String
   Dim sBegDate As String
   Dim sEnddate As String
   MouseCursor 13
   
   If Not IsDate(txtBeg) Then
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   Else
      sBegDate = "1995,01,01"
   End If
   If Not IsDate(txtEnd) Then
      sEnddate = Format(txtEnd, "yyyy,mm,dd")
   Else
      sEnddate = "2020,12,31"
   End If
   If Trim(cmbPrt) <> "ALL" Then sReport = Compress(cmbPrt)
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("quafa02")
   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.Crw.Formulas(1) = "Includes='" & cmbPrt & ", From " _
                        & txtBeg & " Through " & txtEnd & "...'"
   MdiSect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   sSql = "{FahdTable.FA_REF} LIKE '" & sReport & "*' " _
          & "AND {FahdTable.FA_CREATED} In Date(" & sBegDate & ") " _
          & "To Date(" & sEnddate & ") "
   If optCom.value = True Then
      sSql = sSql & "AND {FahdTable.FA_COMPLETE}=1 "
   Else
      sSql = sSql & "AND {FahdTable.FA_COMPLETE}=0 "
   End If
   If optDet.value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(0) = "GROUPHDR.1.0;F;;;"
      MdiSect.Crw.SectionFormat(1) = "DETAIL.0.0;F;;;"
   Else
      MdiSect.Crw.SectionFormat(0) = "GROUPHDR.1.0;T;;;"
      MdiSect.Crw.SectionFormat(1) = "DETAIL.1.0;T;;;"
   End If
   If optCmp.value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.0.0;F;;;"
   Else
      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.0.0;T;;;"
   End If
   MdiSect.Crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtBeg = "01/01/" & Right(txtEnd, 2)
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(str(optDet.value)) & Trim(str(optCmp.value))
   SaveSetting "Esi2000", "EsiQual", "fa02", Trim(sOptions)
   
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiQual", "fa02", sOptions)
   If Len(sOptions) > 0 Then
      optDet.value = Val(Left(sOptions, 1))
      optCmp.value = Val(Right(sOptions, 1))
   Else
      optDet.value = vbChecked
      optCmp.value = vbChecked
   End If
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
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
