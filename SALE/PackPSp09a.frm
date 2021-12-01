VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PackPSp09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Packing Slips Printed Not Shipped"
   ClientHeight    =   2895
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2895
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbDateField 
      Height          =   315
      ItemData        =   "PackPSp09a.frx":0000
      Left            =   2040
      List            =   "PackPSp09a.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSp09a.frx":0036
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Tag             =   "4"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   3
      Tag             =   "4"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Packing Slips"
      Top             =   1080
      Width           =   1555
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
         Picture         =   "PackPSp09a.frx":07E4
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
         Picture         =   "PackPSp09a.frx":0962
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7080
      Top             =   3840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2895
      FormDesignWidth =   7545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Report by"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   1665
   End
   Begin VB.Label lblCUName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2040
      TabIndex        =   15
      Top             =   1440
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   13
      Top             =   2280
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   12
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Dates "
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Report Is Useful For Prepackaged Goods Only"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   5625
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5640
      TabIndex        =   9
      Top             =   1200
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Select A Customer Or Leave Blank"
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "PackPSp09a"
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
   sSql = "Qry_FillPSNotPrinted"
   LoadComboBox cmbCst, 2
   cmbCst = "ALL"
   GetThisCustomer
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbCst_Click()
   GetThisCustomer
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Trim(cmbCst) = "" Then cmbCst = "ALL"
   GetThisCustomer
   
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
   If bOnLoad = 1 Then FillCombo
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
   Set PackPSp09a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sBegDate As String
   Dim sEndDate As String
   Dim sCust As String
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
      sEndDate = Format(txtEnd, "yyyy,mm,dd")
   Else
      sEndDate = "2024,12,31"
   End If
   If cmbCst <> "ALL" Then sCust = Compress(cmbCst)
   On Error GoTo DiaErr1
'   SetMdiReportsize MdiSect
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "Includes='Customer(s) " & cmbCst & " And " _
'                        & "Printed From " & txtBeg & " To " & txtEnd & "'"
'   MdiSect.Crw.Formulas(2) = "RequestBy='Requested By:" & sInitials & "'"
'   MdiSect.Crw.ReportFileName = sReportPath & "sleps09.rpt"
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Customer(s) " & CStr(cmbCst & " And " _
                        & cmbDateField & " From " & txtBeg & " To " & txtEnd) & "...'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("sleps09.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = ""
   sSql = cCRViewer.GetReportSelectionFormula
   
   If (sSql <> "") Then
      sSql = sSql & " AND "
   End If
   
   sSql = "{PshdTable.PSCUST} LIKE '" & sCust & "*' "
   
   If Me.cmbDateField.ListIndex = 0 Then sSql = sSql & "AND {PshdTable.PSPRINTED} in " _
      Else sSql = sSql & "AND {SoitTable.ITSCHED} in "
      
    sSql = sSql & "Date(" & sBegDate & ") to Date(" & sEndDate & ")" _
          & " AND ({SoitTable.ITPSNUMBER} = {PshdTable.PSNUMBER}) and {SoitTable.ITPSSHIPPED} = 0.00 and " _
          & "{PshdTable.PSSHIPPRINT} = 1.00"
          
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
          
'   MdiSect.Crw.SelectionFormula = sSql
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
   b = AllowPsPrepackaging()
   If b = 1 Then z1(4).Visible = False
   GetPsDates
   
End Sub

Private Sub SaveOptions()
    SaveSetting "Esi2000", "EsiSale", "PackPSp09a", LTrim(str(cmbDateField.ListIndex))
End Sub

Private Sub GetOptions()
    cmbDateField.ListIndex = Val(GetSetting("Esi2000", "EsiSale", "PackPSp09a", "0"))
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Trim(txtEnd) <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub


Private Sub GetPsDates()
   Dim RdoGdt As ADODB.Recordset
   sSql = "SELECT MIN(PSPRINTED) FROM PshdTable WHERE PSPRINTED " _
          & "IS NOT NULL"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGdt, ES_FORWARD)
   If bSqlRows Then
      txtBeg = Format(RdoGdt.Fields(0), "mm/dd/yyyy")
   Else
      txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   
   Set RdoGdt = Nothing
End Sub

