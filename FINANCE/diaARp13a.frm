VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARp13a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tax Codes (Report)"
   ClientHeight    =   2790
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkCountry 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1020
      Width           =   855
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      TabIndex        =   19
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaARp13a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "diaARp13a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox chkComments 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   2280
      Width           =   855
   End
   Begin VB.ComboBox cmbCode 
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox cmbSte 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CheckBox chkBO 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   420
      Width           =   855
   End
   Begin VB.CheckBox chkSales 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6000
      Top             =   2340
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2790
      FormDesignWidth =   6705
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   9
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaARp13a.frx":0308
      PictureDn       =   "diaARp13a.frx":044E
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaARp13a.frx":05A0
      PictureDn       =   "diaARp13a.frx":06E6
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Country Codes?"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1020
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Comments?"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Include B && O Tax Codes?"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   420
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Sales Tax Codes?"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "State / Country"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   1395
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaARp13a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaARp13a - Tax Codes Reports
'
' Created: 09/12/02 (JH)
' Revisions:
'
'
'*********************************************************************************

Dim bOnLoad As Byte
'Dim bBOType     As Byte
'Dim bSalesType  As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub chkBO_Click()
   'bBOType = chkBO.Value
   FillStates
   If "" & Trim(cmbSte) <> "" Then
      FillCodes
   End If
End Sub

Private Sub chkBO_LostFocus()
   'bBOType = chkBO.Value
   FillStates
   If "" & Trim(cmbSte) <> "" Then
      FillCodes
   End If
End Sub

Private Sub chkSales_Click()
   'bSalesType = chkSales.Value
   FillStates
   If "" & Trim(cmbSte) <> "" Then
      FillCodes
   End If
End Sub

Private Sub chkSales_LostFocus()
   'bSalesType = chkSales.Value
   FillStates
   If "" & Trim(cmbSte) <> "" Then
      FillCodes
   End If
End Sub

Private Sub cmbSte_Click()
   If Not bOnLoad Then
      FillCodes
   End If
End Sub


Private Sub cmbSte_LostFocus()
   If Not bOnLoad Then
      FillCodes
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillStates
      'bBOType = 0
      'bSalesType = 0
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaARp13a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub PrintReport()
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   If chkSales.Value = 0 And chkBO.Value = 0 And chkCountry.Value = 0 Then
      MsgBox "You Must Select A Tax Code Type.", vbInformation, Caption
      chkSales.SetFocus
      Exit Sub
   End If
   
   MouseCursor 13
   On Error GoTo DiaErr1
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finar13a.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   
   sSql = ""
   If "" & Trim(cmbSte) = "" And "" & Trim(cmbCode) = "" Then
       aFormulaName.Add "Selections"
       aFormulaValue.Add CStr("'For States: ALL & Codes ALL'")
      'no sql required
   ElseIf "" & Trim(cmbSte) <> "" And "" & Trim(cmbCode) = "" Then
       aFormulaName.Add "Selections"
       aFormulaValue.Add CStr("'For: " & CStr(Trim(cmbSte)) & " Codes ALL'")
       sSql = "{TxcdTable.TAXSTATE} = '" & Compress(cmbSte) & "'"
   Else
       aFormulaName.Add "Selections"
       aFormulaValue.Add CStr("'For: " & CStr(Trim(cmbSte) & " Codes " & Trim(cmbCode)) & "'")
       sSql = "{TxcdTable.TAXREF} = '" & Compress(cmbSte) & Compress(cmbCode) & "'"
   End If
   
       aFormulaName.Add "IncludeComments"
       aFormulaValue.Add chkComments.Value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   'tax type selection
   Dim codes As String
   codes = TaxTypeCrystal()
   If codes <> "" Then
      If sSql <> "" Then
         sSql = sSql & " and "
      End If
      sSql = sSql & codes
   End If
   
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
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   
   'If bSalesType = False And bBOType = False Then
   If chkSales.Value = 0 And chkBO.Value = 0 And chkCountry.Value = 0 Then
      MsgBox "You Must Select A Tax Code Type.", vbInformation, Caption
      chkSales.SetFocus
      Exit Sub
'   ElseIf "" & Trim(cmbSte) <> "" And "" & Trim(cmbCode) = "" Then
'      MsgBox "You Must Select A Tax Code.", vbInformation, Caption
'      cmbCode.SetFocus
'      Exit Sub
   End If
      
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   MdiSect.crw.ReportFileName = sReportPath & "finar13a.rpt"
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   
   sSql = ""
   If "" & Trim(cmbSte) = "" And "" & Trim(cmbCode) = "" Then
      MdiSect.crw.Formulas(2) = "Selections='For States: ALL & Codes ALL'"
      'no sql required
   ElseIf "" & Trim(cmbSte) <> "" And "" & Trim(cmbCode) = "" Then
      MdiSect.crw.Formulas(2) = "Selections='For: " & Trim(cmbSte) & " Codes ALL'"
      'sSql = "{TxcdTable.TAXREF} = '" & Compress(cmbSte) & Compress(cmbCode) & "'"
      sSql = "{TxcdTable.TAXSTATE} = '" & Compress(cmbSte) & "'"
   Else
      MdiSect.crw.Formulas(2) = "Selections='For: " & Trim(cmbSte) & " Codes " & Trim(cmbCode) & "'"
      sSql = "{TxcdTable.TAXREF} = '" & Compress(cmbSte) & Compress(cmbCode) & "'"
   End If
   
   MdiSect.crw.Formulas(3) = "IncludeComments=" & chkComments.Value
   
'   sSql = ""
'   If Len(Trim(cmbSte)) <> 0 And Len(Trim(cmbCode)) <> 0 Then
'      sSql = "{TxcdTable.TAXREF} = '" & Compress(cmbSte) & Compress(cmbCode) & "'"
'      MdiSect.crw.SelectionFormula = sSql
'   ElseIf chkSales.Value And chkBO.Value And chkCountry.Value Then
'      'don't do anything
'   Else
'      sSql = TaxTypeCrystal()
'   End If
   
   'tax type selection
   Dim codes As String
   codes = TaxTypeCrystal()
   If codes <> "" Then
      If sSql <> "" Then
         sSql = sSql & " and "
      End If
      sSql = sSql & codes
   End If
   
   MdiSect.crw.SelectionFormula = sSql
   
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function TaxTypeCrystal() As String
   ' returns {TxcdTable.TAXTYPE} IN [...]
   ' or blank if no types selected
   
   Dim insql As String
   insql = ""
   
   If chkBO.Value Then
      insql = insql & 0
   End If
   
   If chkSales.Value Then
      If Len(insql) > 0 Then
         insql = insql & ","
      End If
      insql = insql & 1
   End If
   
   If chkCountry.Value Then
      If Len(insql) > 0 Then
         insql = insql & ","
      End If
      insql = insql & 2
   End If
   
   If Len(insql) = 0 Then
      TaxTypeCrystal = ""
   Else
      TaxTypeCrystal = " {TxcdTable.TAXTYPE} IN [" & insql & "] "
   End If
   
End Function

Private Function TaxTypeSQL() As String
   ' returns TAXTYPE in (...)
   ' or blank if no types selected
   
   Dim insql As String
   insql = ""
   If chkBO.Value Then
      insql = insql & 0
   End If
   
   If chkSales.Value Then
      If Len(insql) > 0 Then
         insql = insql & ","
      End If
      insql = insql & 1
   ElseIf chkCountry.Value Then
      If Len(insql) > 0 Then
         insql = insql & ","
      End If
      insql = insql & 2
   End If
   If Len(insql) = 0 Then
      TaxTypeSQL = ""
   Else
      TaxTypeSQL = " TAXTYPE in (" & insql & ") "
   End If
   
End Function

Public Sub FillCodes()
   Dim RdoCmb As ADODB.Recordset
   
   On Error GoTo DiaErr1
   cmbCode.Clear
   'If bSalesType <> 0 And bBOType <> 0 Then
   If chkSales.Value And chkBO.Value And chkCountry.Value Then
      sSql = "SELECT TAXCODE FROM TxcdTable WHERE TAXSTATE = '" & Trim(cmbSte) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
      If bSqlRows Then
         With RdoCmb
            While Not .EOF
               AddComboStr cmbCode.hwnd, "" & Trim(!taxCode)
               .MoveNext
            Wend
         End With
         cmbCode.ListIndex = 0
      End If
      Set RdoCmb = Nothing
      'ElseIf bSalesType = 0 And bBOType = 0 Then
   ElseIf Not chkSales.Value And Not chkBO.Value And Not chkCountry.Value Then
      Exit Sub
   Else
      'sSql = "SELECT TAXCODE FROM TxcdTable WHERE TAXTYPE = " & bSalesType
      sSql = "SELECT TAXCODE FROM TxcdTable WHERE " & TaxTypeSQL()
      If Trim(cmbSte) <> "" Then
         sSql = sSql & " AND TAXSTATE = '" & Trim(cmbSte) & "'"
      End If
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
      If bSqlRows Then
         With RdoCmb
            While Not .EOF
               AddComboStr cmbCode.hwnd, "" & Trim(!taxCode)
               .MoveNext
            Wend
         End With
         cmbCode.ListIndex = 0
      End If
      Set RdoCmb = Nothing
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "FillCodes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillStates()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   cmbSte.Clear
   
   'If (bSalesType = 0 And bBOType = 0) Or (bSalesType <> 0 And bBOType <> 0) Then
   If Len(TaxTypeSQL()) = 0 Or Len(TaxTypeSQL()) >= 20 Then
      sSql = "SELECT DISTINCT TAXSTATE FROM TxcdTable"
   Else
      sSql = "SELECT DISTINCT TAXSTATE FROM TxcdTable WHERE " & TaxTypeSQL()
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_KEYSET)
   If bSqlRows Then
      With RdoCmb
         
         Do Until .EOF
            cmbSte.AddItem "" & Trim(!taxState)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "Fill States"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SaveOptions()
   Dim sbuf As String
   sbuf = chkComments.Value
   SaveSetting "Esi2000", "EsiFina", Me.Name, sbuf
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(sOptions) Then
      chkComments.Value = Mid(sOptions, 1, 1)
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub
