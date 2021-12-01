VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRp16a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturers By Part Number"
   ClientHeight    =   3105
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
   ScaleHeight     =   3105
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optCon 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   288
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers Assigned To A Manufacturer"
      Top             =   960
      Width           =   3360
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRp16a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optAdr 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PurcPRp16a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PurcPRp16a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   3120
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
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   288
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contacts"
      Height          =   288
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1680
      TabIndex        =   14
      Top             =   1320
      Width           =   3372
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   288
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   1428
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   1428
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ext Descriptions"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   1788
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5400
      TabIndex        =   9
      Top             =   1080
      Width           =   1428
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Tag             =   " "
      Top             =   1800
      Width           =   1428
   End
End
Attribute VB_Name = "PurcPRp16a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'11/20/06 New
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE PAMANUFACTURER<>'' "
   LoadComboBox cmbPrt
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbPrt_Click()
   FindPart cmbPrt, 1
   
End Sub


Private Sub cmbPrt_LostFocus()
   If Len(cmbPrt) Then
      FindPart cmbPrt, 1
   Else
      cmbPrt = "ALL"
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      If cmdHlp Then
         MouseCursor 13
         OpenHelpContext 907
         MouseCursor 0
         cmdHlp = False
      End If
   End If
   
End Sub

Private Sub Form_Activate()
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
   Set PurcPRp16a = Nothing
   
End Sub

Private Sub PrintReport()
    Dim sIncludes As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   If cmbPrt = "" Then cmbPrt = "ALL"
   If cmbPrt <> "ALL" Then sIncludes = Compress(cmbPrt)
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowAddress"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbPrt) & "...'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add optAdr.Value
    
    sCustomReport = GetCustomReport("prdpr19")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{PartTable.PARTREF} LIKE '" & sIncludes & "*' "
'   If optAdr.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.1;T;;;"
'   End If
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
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
   On Error Resume Next
   sOptions = Trim(str(optAdr.Value))
   SaveSetting "Esi2000", "EsiProd", "prp16a", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "prp16a", sOptions)
   If Trim(sOptions) <> "" Then _
           optAdr.Value = Val(Mid(sOptions, 1, 1))
   
End Sub

Private Sub lblDsc_Change()
   If lblDsc = "" Or lblDsc = "*** Part Number Wasn't Found ***" Then lblDsc = _
               "*** Range Of Part Numbers ***"
   
End Sub

Private Sub optAdr_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optCon_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub
