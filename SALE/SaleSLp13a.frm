VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLp13a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Price Book Discrepancy Report"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox cbShowPartDesc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ComboBox txtBok 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   2200
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   840
      Top             =   120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2040
      FormDesignWidth =   5955
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   480
      Width           =   1215
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "SaleSLp13a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Print The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "SaleSLp13a.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Display The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Part Desc"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price Book(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   3960
      TabIndex        =   5
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "SaleSLp13a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'08/25/11 New
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillPriceBooks"
   LoadComboBox txtBok
   If txtBok.ListCount > 0 Then txtBok = txtBok.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub Form_Activate()
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
   bOnLoad = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set SaleSLp13a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   'txtEnd = Format(Now, "mm/dd/yy")
   'txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 3)
   'txtCust = "ALL"
   
End Sub






Private Sub optDis_Click()
    PrintReport
End Sub


Private Sub PrintReport()
   Dim sBook As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
  
    
   MouseCursor 13

   On Error GoTo DiaErr1
   If txtBok = "" Then txtBok = "ALL"
   If txtBok <> "ALL" Then sBook = Compress(txtBok)
     
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowPartDesc"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("' " & CStr(txtBok) & "...'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add cbShowPartDesc.Value

   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("slepb04.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{PbhdTable.PBHREF} LIKE '" & sBook & "*' AND " & _
          "{SoitTable.ITCANCELED}=0 AND {SoitTable.ITPSSHIPPED} = 0 AND " & _
          "{SohdTable.SOCREATED}>={PbhdTable.PBHSTARTDATE} AND {SohdTable.SOCREATED}<={PbhdTable.PBHENDDATE} AND " & _
          "{PbitTable.PBIPRICE}<>{SoitTable.ITDOLLARS} "

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

Private Sub optPrn_Click()
    PrintReport
End Sub


Private Sub txtBok_Validate(Cancel As Boolean)
   txtBok = CheckLen(txtBok, 12)
   If Trim(txtBok) = "" Then txtBok = "ALL"
   
End Sub
