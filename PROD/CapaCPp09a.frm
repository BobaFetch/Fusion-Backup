VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPp09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Center Used On"
   ClientHeight    =   3105
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
   ScaleHeight     =   3105
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPp09a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optCmp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   4
      Top             =   2520
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   2280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   2280
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Shop From List"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox cmbWcn 
      ForeColor       =   &H00000000&
      Height          =   288
      Left            =   2280
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   2040
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
         Picture         =   "CapaCPp09a.frx":07AE
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
         Picture         =   "CapaCPp09a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3105
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Completed Operations"
      Height          =   288
      Index           =   7
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   2004
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Comments"
      Height          =   288
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   2004
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   14
      Top             =   1320
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center(s)"
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   13
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Shops"
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Work Centers"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   2004
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   10
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Tag             =   " "
      Top             =   1800
      Width           =   1428
   End
End
Attribute VB_Name = "CapaCPp09a"
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

Private Sub FillShops()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_FillShops"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         AddComboStr cmbShp.hwnd, "ALL"
         Do Until .EOF
            AddComboStr cmbShp.hwnd, "" & Trim(!SHPNUM)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   cmbShp = cmbShp.List(0)
   FillCenters
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbShp_Click()
   FillCenters
   
End Sub

Private Sub cmbShp_LostFocus()
   If Trim(cmbShp) = "" Then cmbShp = "ALL"
   FillCenters
   
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
   If bOnLoad Then
      FillShops
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
   Set CapaCPp09a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sShop As String
   Dim sCenter As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   If cmbShp <> "ALL" Then sShop = Compress(cmbShp)
   If cmbWcn <> "ALL" Then sCenter = Compress(cmbWcn)
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Includes"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'Work Center(s)" & CStr(cmbWcn & " And " _
                        & "Work Center(s) " & cmbWcn) & "...'")
    
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("prdca16")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   sSql = "{WcntTable.WCNREF} LIKE '" & sCenter & "*' AND " _
          & "{ShopTable.SHPREF} LIKE '" & sShop & "*' "

'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
    aFormulaName.Add "ServiceCenters"
    'aFormulaValue.Add optDet.value
    
   If optDet.Value = vbUnchecked Then
      sSql = sSql & "AND {WcntTable.WCNSERVICE}=0 "
      'MDISect.Crw.Formulas(3) = "ServiceCenters='N'"
     aFormulaValue.Add CStr("'N'")
   Else
     ' MDISect.Crw.Formulas(3) = "ServiceCenters='Y'"
     aFormulaValue.Add CStr("'Y'")
   End If
   
    aFormulaName.Add "CompletedOps"
    'aFormulaValue.Add optCmp.value
       
   If optCmp.Value = vbUnchecked Then
      sSql = sSql & "AND {RnopTable.OPCOMPLETE}=0 "
   '   MDISect.Crw.Formulas(4) = "CompletedOps='N'"
      aFormulaValue.Add CStr("'N'")
   Else
    '  MDISect.Crw.Formulas(4) = "CompletedOps='Y'"
      aFormulaValue.Add CStr("'Y'")
   End If
     aFormulaName.Add "ShowComments"
     aFormulaValue.Add OptCmt.Value
'   If optCmt.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.2.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.2.0;T;;;"
'   End If
   sSql = sSql & " AND {RunsTable.RUNSTATUS} <> 'CA' "
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
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
   sOptions = Trim(str(optDet.Value)) _
              & Trim(str(OptCmt.Value)) _
              & Trim(str(optCmp.Value))
   SaveSetting "Esi2000", "EsiProd", "ca16", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = Trim(GetSetting("Esi2000", "EsiProd", "ca16", sOptions))
   If Len(sOptions) > 0 Then
      optDet.Value = Val(Left(sOptions, 1))
      OptCmt.Value = Val(Mid(sOptions, 2, 1))
      optCmp.Value = Val(Right(sOptions, 1))
   End If
   
End Sub

Private Sub optCmp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
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

Private Sub FillCenters()
   Dim RdoWcn As ADODB.Recordset
   cmbWcn.Clear
   If cmbShp = "ALL" Then
      sSql = "SELECT DISTINCT WCNREF,WCNNUM FROM WcntTable " _
             & "ORDER BY WCNREF"
   Else
      sSql = "SELECT WCNREF,WCNNUM FROM WcntTable " _
             & "WHERE WCNSHOP='" & Compress(cmbShp) & "' ORDER BY WCNREF"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWcn, ES_FORWARD)
   If bSqlRows Then
      With RdoWcn
         AddComboStr cmbWcn.hwnd, "ALL"
         Do Until .EOF
            AddComboStr cmbWcn.hwnd, "" & Trim(!WCNNUM)
            .MoveNext
         Loop
         ClearResultSet RdoWcn
      End With
   Else
      AddComboStr cmbWcn.hwnd, "ALL"
   End If
   cmbWcn = cmbWcn.List(0)
   Set RdoWcn = Nothing
End Sub
