VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MO Status By Customer"
   ClientHeight    =   3060
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
   ScaleHeight     =   3060
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   2040
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1440
      Width           =   1250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1440
      Width           =   1250
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
         Picture         =   "ShopSHp06a.frx":07AE
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
         Picture         =   "ShopSHp06a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Contains Customers With Allocations"
      Top             =   1080
      Width           =   1575
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   9
      Left            =   5400
      TabIndex        =   13
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Tag             =   " "
      Top             =   2040
      Width           =   1545
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   7
      Left            =   3240
      TabIndex        =   10
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Orders Booked"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Tag             =   " "
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "ShopSHp06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/28/05 Changed date handling
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCst_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbCst = CheckLen(cmbCst, 10)
   For iList = 0 To cmbCst.ListCount - 1
      If cmbCst = cmbCst.List(iList) Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbCst = cmbCst.List(0)
   End If
   
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
   sSql = "SELECT DISTINCT SOCUST,CUREF,CUNICKNAME " _
          & "FROM RnalTable,SohdTable,CustTable WHERE " _
          & "(RASO=SONUMBER AND SOCUST=CUREF) "
   LoadComboBox cmbCst, 1
   If cmbCst.ListCount > 0 Then
      cmbCst = cmbCst.List(0)
   Else
      cmbCst = "No Allocations"
      cmbCst.ForeColor = ES_RED
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
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
   Set ShopSHp06a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sCust As String
   Dim sBegDate As String
   Dim sEndDate As String
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   MouseCursor 13
   On Error Resume Next
   sCust = Compress(cmbCst)
   If Not IsDate(txtBeg) Then
      sBegDate = "1995,01,01"
   Else
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEndDate = "2024,12,31"
   Else
      sEndDate = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   MouseCursor 13
   On Error GoTo DiaErr1
   

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowDescription"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & txtBeg & " To " & txtEnd & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add optDsc.Value

   sCustomReport = GetCustomReport("prdsh06")
 '  MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

'   If optDsc.value = vbUnchecked Then
''      MDISect.Crw.SectionFormat(0) = "DETAIL.0.1;F;;;"
''      MDISect.Crw.Formulas(3) = "Desc='0'"
'      aFormulaValue.Add optDsc.value
'   Else
''      MDISect.Crw.SectionFormat(0) = "DETAIL.0.1;T;;;"
''      MDISect.Crw.Formulas(3) = "Desc='1'"
'      aFormulaValue.Add optDsc.value
'   End If
   'MDISect.Crw.SelectionFormula = sSql
   'SetCrystalAction Me
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{SohdTable.SOCUST} LIKE '" & sCust & "*' AND " _
          & "{SohdTable.SODATE} in Date(" & sBegDate _
          & ") to Date(" & sEndDate & ")"
   
   cCRViewer.SetReportSelectionFormula sSql
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
'   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
'   txtBeg = "01/01/" & Right(txtEnd, 2)
   txtEnd = ""
   txtBeg = ""
   
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiProd", "sh06", Trim(str(optDsc.Value))
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh06", sOptions)
   If Len(sOptions) Then
      optDsc.Value = Val(sOptions)
   Else
      optDsc.Value = vbChecked
   End If
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) > 3 Then
      txtBeg = CheckDateEx(txtBeg)
   Else
      txtBeg = "ALL"
   End If
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) > 3 Then
      txtEnd = CheckDateEx(txtEnd)
   Else
      txtEnd = "ALL"
   End If
   
End Sub
