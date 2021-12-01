VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Document Used On"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboClass 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "DocuDCp05a.frx":0000
      Left            =   2040
      List            =   "DocuDCp05a.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "9"
      ToolTipText     =   "Select Class From List"
      Top             =   600
      Width           =   2000
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCp05a.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cboRev 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "9"
      ToolTipText     =   "Document Revision"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboSheet 
      Height          =   315
      Left            =   4560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "9"
      ToolTipText     =   "Sheet (If Marked In Class)"
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox OptExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   645
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   645
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5760
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5760
      TabIndex        =   10
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "DocuDCp05a.frx":07B2
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
         Picture         =   "DocuDCp05a.frx":0930
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cboDoc 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "9"
      ToolTipText     =   "Contains Parts With A Document List"
      Top             =   1020
      Width           =   3285
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   6855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document Class"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   1725
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sheet"
      Height          =   255
      Index           =   3
      Left            =   3780
      TabIndex        =   14
      Top             =   1500
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1020
      Width           =   1725
   End
End
Attribute VB_Name = "DocuDCp05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private doc As ClassDoc
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboClass_Click()
   doc.FillDocuments cboClass, cboDoc
End Sub

Private Sub cboClass_LostFocus()
   doc.FillDocuments cboClass, cboDoc
End Sub

Private Sub cboDoc_Click()
   'GetSheets
   doc.FillRevisions cboClass, cboDoc, cboRev
End Sub


Private Sub cboDoc_LostFocus()
   'cboDoc = CheckLen(cboDoc, 30)
   'If Len(cboDoc) Then GetSheets
   doc.FillRevisions cboClass, cboDoc, cboRev
End Sub


'Private Sub cmbRev_Click()
'   doc.FillSheets cboClass, cboDoc, cboRev, cboSheet
'End Sub
'
'Private Sub cmbRev_LostFocus()
'   'cmbRev = CheckLen(cmbRev, 6)
'   doc.FillSheets cboClass, cboDoc, cboRev, cboSheet
'End Sub
'
'
'Private Sub cmbSht_Click()
''   GetRevisions
''
'End Sub
'
'Private Sub cmbSht_LostFocus()
''   cmbSht = CheckLen(cmbSht, 6)
''   GetRevisions
''
'End Sub
'

Private Sub cboRev_Click()
   doc.FillSheets cboClass, cboDoc, cboRev, cboSheet
End Sub

Private Sub cboRev_LostFocus()
   doc.FillSheets cboClass, cboDoc, cboRev, cboSheet
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' cboDoc = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


'Private Sub FillCombo()
'   On Error GoTo DiaErr1
'   sSql = "Qry_FillDocuments"
'   LoadComboBox cbodoc
'   If cbodoc.ListCount > 0 Then cbodoc = cbodoc.List(0)
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillcombo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub

Private Sub Form_Activate()
   'On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      Set doc = New ClassDoc
      doc.FillClasses Me.cboClass, False
      'FillCombo
      'doc.FillDocuments cboClass, Me.cboDoc
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   bOnLoad = 1
   FormLoad Me
   FormatControls
   GetOptions
   'FillCombo
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set DocuDCp05a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sClass As String
   Dim sDoc As String
   Dim sRev As String
   Dim sSheet As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   sClass = Compress(cboClass)
   sDoc = Compress(cboDoc)
   sRev = Compress(cboRev)
   sSheet = Compress(cboSheet)
   
   sCustomReport = GetCustomReport("engdc05")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "PartDescription"
   aFormulaName.Add "ExtendedDescription"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(optDsc.value) & "'")
   aFormulaValue.Add CStr("'" & CStr(OptExt.value) & "'")
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{DdocTable.DOREF} = '" & sDoc & "' " _
          & "AND {DdocTable.DOSHEET} = '" & sSheet & "' " _
          & "AND {DdocTable.DOREV} = '" & sRev & "' " _
          & "AND {DdocTable.DOCLASS} = '" & sClass & "'"
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
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




Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(optDsc.value) _
              & RTrim(OptExt.value)
   SaveSetting "Esi2000", "EsiEngr", "dc05", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "dc05", sOptions)
   If Len(sOptions) > 0 Then
      optDsc.value = Val(Left(sOptions, 1))
      OptExt.value = Val(Mid(sOptions, 2, 1))
   End If
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

''
''
''Private Sub GetSheets()
''   On Error GoTo DiaErr1
''   cmbSht.Clear
''   sSql = "SELECT DISTINCT DOSHEET FROM DdocTable " _
''          & "WHERE DOREF='" & Compress(cboDoc) & "' "
''   LoadComboBox cmbSht, -1
''   If cmbSht.ListCount > 0 Then cmbSht = cmbSht.List(0)
''   Exit Sub
''
''DiaErr1:
''   sProcName = "getsheets"
''   CurrError.Number = Err.Number
''   CurrError.Description = Err.Description
''   DoModuleErrors Me
''
''
''End Sub
''
''Private Sub GetRevisions()
''   Dim sDocument As String
''   Dim sSheet As String
''
''   On Error GoTo DiaErr1
''   cmbRev.Clear
''   sDocument = Compress(cboDoc)
''   sSheet = Compress(cmbSht)
''   sSql = "SELECT DISTINCT DOREV FROM DdocTable " _
''          & "WHERE DOREF='" & sDocument & "' " _
''          & "AND DOSHEET='" & sSheet & "' "
''   LoadComboBox cmbRev, -1
''   If cmbRev.ListCount > 0 Then cmbRev = cmbRev.List(0)
''   Exit Sub
''
''DiaErr1:
''   sProcName = "getrevisi"
''   CurrError.Number = Err.Number
''   CurrError.Description = Err.Description
''   DoModuleErrors Me
''
''End Sub
