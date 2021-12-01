VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ToolTLp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tools By Description"
   ClientHeight    =   3630
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3630
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "ALL"
      Top             =   1440
      Width           =   3225
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   3120
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.Frame z2 
      Height          =   735
      Left            =   1800
      TabIndex        =   18
      Top             =   1800
      Width           =   4335
      Begin VB.OptionButton optCap 
         Caption         =   "Non Expense"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   260
         Width           =   1335
      End
      Begin VB.OptionButton optExp 
         Caption         =   "Expendable"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   260
         Width           =   1335
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   260
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbCls 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "2"
      ToolTipText     =   "Select - Leading Chars"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1800
      TabIndex        =   5
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
         Picture         =   "ToolTLp02a.frx":07AE
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
         Picture         =   "ToolTLp02a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3630
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Comments"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   3120
      Width           =   1785
   End
   Begin VB.Label T 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Types"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   0
      Left            =   5640
      TabIndex        =   16
      Top             =   1440
      Width           =   1428
   End
   Begin VB.Label T 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Descriptions"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Detail"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   2880
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5640
      TabIndex        =   13
      Top             =   1080
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Tag             =   " "
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Label T 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Class(es)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "ToolTLp02a"
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
   sSql = "Qry_FillToolClasses"
   LoadComboBox cmbCls, -1
   If cmbCls.ListCount > 0 Then cmbCls = cmbCls.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbCls_LostFocus()
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
   
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
   MDISect.lblBotPanel = Caption
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
   Set ToolTLp02a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sClass As String
   Dim sPart As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   
   If cmbCls <> "ALL" Then sClass = UCase$(cmbCls)
   If txtDsc <> "ALL" Then sPart = UCase$(txtDsc)
   On Error GoTo DiaErr1
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDetails"
   aFormulaName.Add "ShowComments"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Includes Class(es) " & CStr(cmbCls _
                        & ", Tool(s) " & txtDsc) & "...'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add optDet.value
   aFormulaValue.Add optCmt.value
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("engtl02")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "UpperCase({TohdTable.TOOL_CLASS}) LIKE '" & sClass & "*' " _
          & "AND UpperCase({TohdTable.TOOL_DESC}) LIKE '" & sPart & "*' "
   If optExp.value = True Then _
                     sSql = sSql & " AND {TohdTable.TOOL_EXPENDABLE}=1"
   If optCap.value = True Then _
                     sSql = sSql & " AND {TohdTable.TOOL_EXPENDABLE}=0"
   
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
   txtDsc = "ALL"
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim$(optDet.value) & Trim$(optCmt.value)
   SaveSetting "Esi2000", "EsiEngr", "tl02", sOptions
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   sOptions = Trim(GetSetting("Esi2000", "EsiEngr", "tl02", sOptions))
   On Error Resume Next
   If Trim(sOptions) <> "" Then
      optDet.value = Left(sOptions, 1)
      optCmt.value = Right(sOptions, 1)
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

Private Sub txtDsc_LostFocus()
   txtDsc = Trim(txtDsc)
   If txtDsc = "" Then txtDsc = "ALL"
   If txtDsc <> "ALL" Then txtDsc = StrCase(txtDsc)
   
End Sub