VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSp10a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipped Items By Sales Person"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6405
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSp10a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbcde 
      Height          =   288
      Left            =   1680
      TabIndex        =   3
      Tag             =   "8"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CheckBox optDol 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   3240
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   3000
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   2760
      Width           =   735
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   10
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PackPSp10a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   490
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PackPSp10a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5280
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.ComboBox cmbSls 
      Height          =   288
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Contains Customers With Invoices"
      Top             =   960
      Width           =   1555
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   3960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3735
      FormDesignWidth =   6405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5280
      TabIndex        =   22
      Top             =   1680
      Width           =   1164
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   252
      Index           =   12
      Left            =   240
      TabIndex        =   21
      Top             =   2520
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code(s)"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   20
      Top             =   2205
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   11
      Left            =   5280
      TabIndex        =   19
      Top             =   2208
      Width           =   1188
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Sales Person By Highest Dollar"
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   18
      Top             =   3240
      Width           =   2892
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipments From"
      Height          =   288
      Index           =   6
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   1548
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ext. Descriptions"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   3000
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Width           =   1452
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   5
      Left            =   2880
      TabIndex        =   14
      Top             =   1725
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   0
      Left            =   5280
      TabIndex        =   13
      Top             =   960
      Width           =   1188
   End
   Begin VB.Label lblSls 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1680
      TabIndex        =   12
      Top             =   1320
      Width           =   3000
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1428
   End
End
Attribute VB_Name = "PackPSp10a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of                     ***
'*** ESI Software Engineering Inc, Seattle, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***
'See the UpdateTables prodecure for database revisions
' PackPSp10a - Shipped By Sales Person
'
' Created: 1/05/04 (JCW)
' Revisions:
'   01/22/04 (JCW) Fixed Spelling/Layout/naming struct./report names
'3/23/06 (CJS) Corrected formulae and added Grouping/Dates "ALL"
'5/10/06 See PrintReport
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress(3) As New EsiKeyBd
Private txtGotFocus(3) As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbSls_LostFocus()
   cmbSls = CheckLen(cmbSls, 10)
   If Trim(cmbSls) <> "" And Trim(UCase(cmbSls)) <> "ALL" Then
      FindSalesPerson
   Else
      cmbSls = "ALL"
      lblSls = "***All Sales People Selected.***"
   End If
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      FillProductCodes
      GetOptions
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   txtBeg = Format(Now, "mm/01/yyyy")
   txtEnd = Format(Now, "mm/dd/yyyy")
   bOnLoad = 1
   'GetOptions
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PackPSp10a = Nothing
End Sub



Private Sub FindSalesPerson()
   Dim rdoFind As ADODB.Recordset
   sSql = "SELECT SPFIRST,SPLAST FROM SprsTable WHERE SPNUMBER = '" & cmbSls & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoFind)
   If bSqlRows Then
      With rdoFind
         lblSls = Trim(!SPFIRST) & " " & Trim(!SPLAST)
      End With
   Else
      lblSls = "Range Of Sales Persons Selected."
   End If
End Sub

Private Sub cmbSls_Click()
   FindSalesPerson
   
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
   sSql = "Qry_FillSalesPersons"
   LoadComboBox cmbSls, -1
   cmbSls = "ALL"
   FindSalesPerson
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub



Private Sub PrintReport()
   Dim sBegDate As String
   Dim sEndDate As String
   Dim sCode As String
   Dim sCustomReport As String
   Dim cCRViewer As EsCrystalRptViewer
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
   
   On Error GoTo DiaErr1
   If cmbSls = "" Then cmbSls = "ALL"
   If cmbcde = "" Then cmbcde = "ALL"
   If cmbcde <> "ALL" Then sCode = Trim(cmbcde)
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   aFormulaName.Add "Person"
   aFormulaName.Add "ProductClass"
   aFormulaName.Add "ShowDollars"
   aFormulaName.Add "Desc"
   aFormulaName.Add "ExtDesc"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'Shipped Items By Sales Person'")
   aFormulaValue.Add CStr("'Shipped From " & CStr(txtBeg & "  Through " & txtEnd) & "'")
   aFormulaValue.Add CStr("'Sales Person(s): " & CStr(cmbSls) & "'")
   aFormulaValue.Add CStr("'Product Code(s):" & CStr(cmbcde) & "'")
   aFormulaValue.Add CStr("'" & CStr(optDol.Value) & "'")
   aFormulaValue.Add CStr("'" & CStr(optDsc.Value) & "'")
   aFormulaValue.Add CStr("'" & CStr(optExt.Value) & "'")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   If optDol.Value = vbChecked Then
      sCustomReport = GetCustomReport("sleSh03b.rpt")
   Else
      sCustomReport = GetCustomReport("sleSh03a.rpt")
   End If
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{SoitTable.ITCANCELED} = 0 and " _
          & "({CihdTable.INVNO} <> 0 or trim(cstr({PsitTable.PIPACKSLIP})) <> '') " _
          & "and trim(cstr({@Shipped})) <> '' AND {@shipped} In Date('" _
          & sBegDate & "') To Date('" & sEndDate & "')"
   
   If Trim(cmbSls) <> "ALL" Then _
           sSql = sSql & " AND {SprsTable.SPNUMBER} Like '" & cmbSls & "*' "
   sSql = sSql & " AND {PartTable.PAPRODCODE} Like '" & Trim(sCode) & "*'"
   
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
   
   Set txtGotFocus(0).esCmbGotfocus = txtBeg
   Set txtGotFocus(1).esCmbGotfocus = txtEnd
   Set txtGotFocus(2).esCmbGotfocus = cmbSls
   
   Set txtKeyPress(0).esCmbKeyDate = txtBeg
   Set txtKeyPress(1).esCmbKeyDate = txtEnd
   Set txtKeyPress(2).esCmbKeyCase = cmbSls
   cmbSls = "ALL"
   cmbcde = "ALL"
   
End Sub



Private Sub lblSls_Change()
   If Left(lblSls, 8) = "***Sales" Then
      lblSls = "Range Of Sales Persons Selected."
   Else
      lblSls.ForeColor = vbBlack
   End If
   
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

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(optExt.Value) _
              & RTrim(optDsc.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   
End Sub
