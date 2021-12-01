VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSp13a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Transfers (Report)"
   ClientHeight    =   2925
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   HelpContextID   =   934
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2925
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSp13a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optTransfers 
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cmbCst 
      Height          =   288
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Pack Slips"
      Top             =   960
      Width           =   1555
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   2400
      Width           =   735
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PackPSp13a.frx":07AE
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
         Picture         =   "PackPSp13a.frx":092C
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
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2925
      FormDesignWidth =   7095
   End
   Begin VB.Label lblCUName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1920
      TabIndex        =   20
      Top             =   1320
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfers Are Active"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   18
      ToolTipText     =   "Company Setting"
      Top             =   480
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   7
      Left            =   5520
      TabIndex        =   16
      Top             =   1680
      Width           =   2268
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   6
      Left            =   3120
      TabIndex        =   15
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   14
      Top             =   960
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer(s) "
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   1548
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Slips From"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1908
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   288
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1788
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   1788
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   5
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1668
   End
End
Attribute VB_Name = "PackPSp13a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'8/2/05 New
Option Explicit
Dim bOnLoad As Byte
Dim lTransfer As Long

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub CheckTransferInvoice()
   Dim RdoInv As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT TransferInvoice,AllowTransfers FROM Preferences " _
          & "WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   If bSqlRows Then
      With RdoInv
         lTransfer = !TransferInvoice
         optTransfers.Value = !AllowTransfers
         ClearResultSet RdoInv
      End With
   End If
   Set RdoInv = Nothing
   Exit Sub
   
DiaErr1:
   lTransfer = 0
   
End Sub

Private Sub cmbCst_Click()
   GetThisCustomer
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(Trim(cmbCst)) = 0 Then cmbCst = "ALL"
   GetThisCustomer
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 934
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PSCUST,CUREF,CUNICKNAME FROM " _
          & "PshdTable,CustTable WHERE (PSCUST=CUREF AND PSINVOICE=" _
          & lTransfer & ") "
   LoadComboBox cmbCst, 1
   cmbCst = "ALL"
   GetThisCustomer
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CheckTransferInvoice
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
   Set PackPSp13a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sBeg As String
   Dim sEnd As String
   Dim sCust As String
   Dim sCustomReport As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   MouseCursor 13
   On Error GoTo DiaErr1
   If Trim(cmbCst) <> "ALL" Then sCust = Compress(cmbCst)
   If IsDate(txtBeg) Then
      sBeg = Format(txtBeg, "yyyy,mm,dd")
   Else
      sBeg = "1995,01,01"
   End If

   If IsDate(txtEnd) Then
      sEnd = Format(txtEnd, "yyyy,mm,dd")
   Else
      sEnd = "2024,12,31"
   End If
'   SetMdiReportsize MdiSect
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "Includes='Customer(s) " & cmbCst & ", " _
'                        & "From " & txtBeg & " Through " & txtBeg & "'"
'   MdiSect.Crw.Formulas(2) = "RequestBy='Requested By:" & sInitials & "'"
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDetails"
   aFormulaName.Add "ShowDescription"

   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Customer(s)" & CStr(cmbCst & ", " _
                        & "From " & txtBeg & " Through " & txtBeg) & "...'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optDet.Value
   aFormulaValue.Add optDsc.Value

   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("sleps19")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{PshdTable.PSCUST} LIKE '" & sCust & "*' " _
          & "AND {PshdTable.PSDATE} in Date(" & sBeg _
          & ") to Date(" & sEnd & ") AND {PshdTable.PSINVOICE}=" & lTransfer & " "
'   If optDet.value = vbUnchecked Then
'      MdiSect.Crw.SectionFormat(1) = "GROUPHDR.0.1;F;;;"
'      MdiSect.Crw.SectionFormat(2) = "DETAIL.0.0;F;;;"
'      MdiSect.Crw.SectionFormat(3) = "DETAIL.0.1;F;;;"
'      MdiSect.Crw.SectionFormat(4) = "DETAIL.0.2;F;;;"
'      MdiSect.Crw.SectionFormat(5) = "GROUPFTR.0.0;F;;;"
'   Else
'      MdiSect.Crw.SectionFormat(1) = "GROUPHDR.0.1;T;;;"
'      MdiSect.Crw.SectionFormat(2) = "DETAIL.0.0;T;;;"
'      If optDsc.value = vbChecked Then
'         MdiSect.Crw.SectionFormat(3) = "DETAIL.0.1;T;;;"
'      Else
'         MdiSect.Crw.SectionFormat(3) = "DETAIL.0.1;F;;;"
'      End If
'      MdiSect.Crw.SectionFormat(4) = "DETAIL.0.2;T;;;"
'      MdiSect.Crw.SectionFormat(5) = "GROUPFTR.0.0;T;;;"
'   End If
'   MdiSect.Crw.SelectionFormula = sSql
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
   sOptions = Trim(str(optDet.Value)) & Trim(str(optDsc.Value))
   SaveSetting "Esi2000", "EsiSale", "sh19", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "sh19", sOptions)
   If Len(Trim(sOptions)) Then
      optDet.Value = Val(Left(sOptions, 1))
      optDsc.Value = Val(Right(sOptions, 1))
   Else
      optDet.Value = vbChecked
      optDsc.Value = vbChecked
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 5)
   cmbCst = "ALL"
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   If lTransfer = 0 Then
      MsgBox "Transfers Have Not Been Setup.", _
         vbInformation, Caption
   Else
      PrintReport
   End If
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   If lTransfer = 0 Then
      MsgBox "Transfers Have Not Been Setup.", _
         vbInformation, Caption
   Else
      PrintReport
   End If
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If txtBeg = "" Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If txtEnd = "" Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtBeg)
   
End Sub
