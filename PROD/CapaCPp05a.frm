VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Of Work Center Calendars"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPp05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5280
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "CapaCPp05a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "CapaCPp05a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbYer 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Year"
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox cmbWcn 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Shop From List"
      Top             =   960
      Width           =   1815
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   6420
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Shops"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1830
      TabIndex        =   6
      Top             =   1680
      Width           =   105
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "CapaCPp05a"
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

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbShp_Click()
   FillWorkCenters
   
End Sub

Private Sub cmbShp_LostFocus()
   cmbShp = CheckLen(cmbShp, 12)
   
End Sub

Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 12)
   
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
   On Error GoTo Pca031
   sSql = "Qry_FillShops "
   LoadComboBox cmbShp
   cmbShp = "ALL"
   Exit Sub
   
Pca031:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillWorkCenters()
   'Dim RdoCmb As ADODB.Recordset
   cmbWcn.Clear
   If cmbShp = "ALL" Then
      cmbWcn = "ALL"
      Exit Sub
   End If
   On Error GoTo DiaErr1
   sSql = "Qry_FillWorkCenters '" & Compress(cmbShp) & "'"
   LoadComboBox cmbWcn
   cmbWcn = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillworkcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillCombo
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim iList As Integer
   FormatControls
   FormLoad Me
   cmbWcn = "ALL"
   cmbYer = "ALL"
   AddComboStr cmbYer.hwnd, "ALL"
   For iList = 1995 To 2024
      AddComboStr cmbYer.hwnd, Format$(iList)
   Next
   AddComboStr cmbYer.hwnd, Format$(iList)
   bOnLoad = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set CapaCPp05a = Nothing
   
End Sub
Private Sub PrintReport()
   Dim sRout As String
   Dim sShop As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   sRout = Compress(cmbWcn)
   sShop = Compress(cmbShp)
   If Len(sRout) = 0 Then cmbWcn = "ALL"
   If sRout = "ALL" Then sRout = ""
   
   MouseCursor 13
   On Error GoTo DiaErr1
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaValue.Add CStr("'Includes" & CStr(cmbWcn & "... And Shop " & cmbShp) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("prdca12")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   'MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{WcntTable.WCNREF} LIKE '" & sRout & "*'"
   If sShop <> "ALL" Then
      sSql = sSql & " AND {WcntTable.WCNSHOP}='" & sShop & "'"
   End If
   If cmbYer <> "ALL" Then
      sSql = sSql & " AND '" & cmbYer & "' IN {WcclTable.WCCREF}"
   End If
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
  ' MDISect.Crw.SelectionFormula = sSql
  ' SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optDis_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   MouseCursor 13
   PrintReport
   
End Sub
