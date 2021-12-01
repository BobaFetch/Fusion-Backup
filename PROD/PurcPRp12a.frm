VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRp12a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Numbers By Buyer"
   ClientHeight    =   3045
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3045
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRp12a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Or Leave Blank"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox cmbByr 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Leave Blank"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2400
      TabIndex        =   3
      Top             =   2400
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Leading Chars Or Blank For All"
      Top             =   1440
      Width           =   3255
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
         Picture         =   "PurcPRp12a.frx":07AE
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
         Picture         =   "PurcPRp12a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3045
      FormDesignWidth =   7215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   7
      Left            =   5880
      TabIndex        =   15
      Top             =   1800
      Width           =   1400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code(s)"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   5880
      TabIndex        =   13
      Top             =   1440
      Width           =   1400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1788
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5880
      TabIndex        =   10
      Top             =   1080
      Width           =   1400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Tag             =   " "
      Top             =   2160
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "PurcPRp12a"
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
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   '    sSql = "SELECT DISTINCT CUREF,CUNICKNAME,SOCUST FROM " _
   '        & "CustTable,SohdTable WHERE CUREF=SOCUST"
   '    bsqlrows = clsadocon.getdataset(ssql, RdoCmb, ES_FORWARD)
   '        If bSqlRows Then
   '            With RdoCmb
   '                Do Until .EOF
   '                    AddComboStr cmbCst.hWnd, "" & Trim(!CUNICKNAME)
   '                    .MoveNext
   '                Loop
   '                .Cancel
   '            End With
   '        Else
   '            lblNme = "*** No Customers With SO's Found ***"
   '        End If
   '    Set RdoCmb = Nothing
   '    cmbCst = "ALL"
   '    GetCustomer
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbByr_LostFocus()
   If cmbByr = "" Then cmbByr = "ALL"
   
End Sub


Private Sub cmbCde_LostFocus()
   If Trim(cmbCde) = "" Then cmbCde = "ALL"
   
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
      cmbByr.AddItem "Unassigned"
      FillBuyers
      FillProductCodes
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
   Set PurcPRp12a = Nothing
   
End Sub
Private Sub PrintReport()
    Dim bShow As Byte
    Dim sBuyer As String
    Dim sCode As String
    Dim sPart As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   If cmbByr <> "ALL" Then sBuyer = Compress(cmbByr)
   If cmbByr = cmbByr.List(0) Then
      sBuyer = ""
      bShow = 1
   End If
   If txtPrt <> "ALL" Then sPart = Compress(txtPrt)
   If cmbCde <> "ALL" Then sCode = Compress(cmbCde)
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Includes"
    aFormulaName.Add "ShowDetails"


    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & "Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbByr) & "...'")
    aFormulaValue.Add optDet.Value
     
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("prdpr15")
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue '   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   
   If bShow = 1 Then
      sSql = "{PartTable.PABUYER} = '' "
   Else
      sSql = "{PartTable.PABUYER} LIKE '" & sBuyer & "*' "
   End If
   sSql = sSql & "AND {PartTable.PARTREF} LIKE '" & sPart & "*' "
   sSql = sSql & "AND {PartTable.PAPRODCODE} LIKE '" & sCode & "*' "
   
'   If optDet.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.1.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.1.0;T;;;"
'   End If
    
    cCRViewer.SetReportSelectionFormula (sSql)
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
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
   cmbByr = "ALL"
   txtPrt = "ALL"
   cmbCde = "ALL"
   
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiProd", "pr15", optDet.Value
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "pr15", sOptions)
   If Trim(sOptions) <> "" Then optDet.Value = Val(sOptions)
   
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

Private Sub txtPrt_LostFocus()
   If Trim(txtPrt) = "" Then txtPrt = "ALL"
   
End Sub
