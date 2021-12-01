VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PackPSp05a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipped Items By Shipping Date"
   ClientHeight    =   3495
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
   ScaleHeight     =   3495
   ScaleWidth      =   6405
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSp05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Tag             =   "8"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Contains Customers With Invoices"
      Top             =   960
      Width           =   1555
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
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PackPSp05a.frx":07AE
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
         Picture         =   "PackPSp05a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   490
      End
   End
   Begin VB.ComboBox cmbEndDate 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cmbStartDate 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   735
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6000
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3495
      FormDesignWidth =   6405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   252
      Index           =   12
      Left            =   240
      TabIndex        =   20
      Top             =   2520
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5280
      TabIndex        =   19
      Top             =   1680
      Width           =   1164
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   18
      Top             =   2205
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   11
      Left            =   5280
      TabIndex        =   17
      Top             =   2208
      Width           =   1164
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer "
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   1320
      Width           =   3120
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   0
      Left            =   5280
      TabIndex        =   14
      Top             =   960
      Width           =   1164
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   5
      Left            =   3120
      TabIndex        =   13
      Top             =   1725
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ext. Descriptions"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipments From"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   10
      Top             =   1725
      Width           =   1305
   End
End
Attribute VB_Name = "PackPSp05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of                     ***
'*** ESI Software Engineering Inc, Seattle, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***
'See the UpdateTables prodecure for database revisions
' PackPSp05a - Shipped Items By Date
'
' Created: 1/05/04 (JCW)
' Revisions:
' 01/22/04 (JCW) Fixed Spelling/Layout/naming struct./report names
'3/23/06 Fixed Formulae and added Groupings
'5/10/06 See Print Report
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress(3) As New EsiKeyBd
Private txtGotFocus(3) As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   
   'FindCustomer does not return not found if invalid
   'So i will just do it manually
   lblCst = "***Customer Not Found.***"
   'then if it returned rows overwrite the message
   
   If Trim(cmbCst) = "" Or UCase(Trim(cmbCst)) = "ALL" Then
      cmbCst = "ALL"
      lblCst = "***All Customers Selected.***"
   Else
      FindCustomer Me, cmbCst
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
   sSql = "Qry_FillCustomers"
   LoadComboBox cmbCst
   cmbCst = "ALL"
   lblCst = "Range Of Customers Selected."
   
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
   sCurrForm = Caption
   cmbStartDate = Format(Now, "mm/01/yyyy")
   cmbEndDate = Format(Now, "mm/dd/yyyy")
   bOnLoad = 1
   'GetOptions
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   If Len(Trim(cmbCst)) Then cUR.CurrentCustomer = cmbCst
   SaveCurrentSelections
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PackPSp05a = Nothing
   
End Sub

Private Sub PrintReport()
    Dim sCustomReport As String
    Dim sBegDate As String
    Dim sEndDate As String
    Dim sCust As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    aFormulaName.Add "Customer"
    aFormulaName.Add "ProductClass"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowExDescription"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'Shipped Items By Shipping Date'")
    aFormulaValue.Add CStr("'Shipped From " & CStr(cmbStartDate & "  Through " & cmbEndDate) & "'")
    aFormulaValue.Add CStr("'Customer(s):" & CStr(cmbCst) & "'")
   
   'CUSTOM REPORT
   
   sCustomReport = GetCustomReport("sleSh02a")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.ShowGroupTree False

   If IsDate(cmbStartDate) Then
      sBegDate = Format(cmbStartDate, "yyyy,mm,dd")
   Else
      sBegDate = "1995,01,01"
   End If
   
   If IsDate(cmbEndDate) Then
      sEndDate = Format(cmbEndDate, "yyyy,mm,dd")
   Else
      sEndDate = "2024,12,31"
   End If
   If cmbCde = "" Then cmbCde = "ALL"
   If cmbCst <> "ALL" Then sCust = Compress(cmbCst)
''   sSql = "{SoitTable.ITCANCELED} = 0 and " _
''          & "({CihdTable.INVNO} <> 0 or trim(cstr({PsitTable.PIPACKSLIP})) <> '') " _
''          & "and trim(cstr({@Shipped})) <> ''" _
''          & " and {@Shipped} IN {@Start} to {@End}"
   
   
   
'   sSql = "{SoitTable.ITCANCELED} = 0" & vbCrLf _
'      & "and ({PshdTable.PSSHIPPEDDATE} IN Date('" & sBegDate & "') to Date('" & sEndDate & "')" & vbCrLf _
'      & "or (isnull({PshdTable.PSSHIPPEDDATE}) " _
'      & "and {CihdTable.INVSHIPDATE} IN Date('" & sBegDate & "') to Date('" & sEndDate & "')) & vbcrlf"
   
   
   
   sSql = "{SoitTable.ITCANCELED} = 0" & vbCrLf _
      & "and ({PshdTable.PSSHIPPEDDATE} IN Date('" & sBegDate & "') to Date('" & sEndDate & "')" & vbCrLf _
      & "or( isnull({PshdTable.PSSHIPPEDDATE})" & vbCrLf _
      & "and {CihdTable.INVSHIPDATE} IN Date('" & sBegDate & "') to Date('" & sEndDate & "')))"
   If Trim(cmbCst) <> "ALL" Then sSql = sSql & " and {CustTable.CUREF} Like '" _
           & sCust & "*'"
   
   If Trim(cmbCde) <> "ALL" Then
      sSql = sSql & " AND {PartTable.PAPRODCODE} Like '" & Trim(cmbCde) & "*'"
      aFormulaValue.Add CStr("'Product Code: " & CStr(cmbCde) & "'")
   Else
      aFormulaValue.Add CStr("'Product Code: ALL'")
   End If
      aFormulaValue.Add optDsc.Value
      aFormulaValue.Add optExt.Value
   
    aFormulaName.Add "StartDate"
    aFormulaName.Add "EndDate"
    aFormulaValue.Add CStr("'" & CStr(cmbStartDate) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbEndDate) & "'")
    
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
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
   
   Set txtGotFocus(0).esCmbGotfocus = cmbStartDate
   Set txtGotFocus(1).esCmbGotfocus = cmbEndDate
   Set txtGotFocus(2).esCmbGotfocus = cmbCst
   
   Set txtKeyPress(0).esCmbKeyDate = cmbStartDate
   Set txtKeyPress(1).esCmbKeyDate = cmbEndDate
   Set txtKeyPress(2).esCmbKeyCase = cmbCst
   cmbCde = "ALL"
   cmbCst = "ALL"
   
End Sub

Private Sub lblcst_Change()
   If Left(lblCst, 8) = "***Custo" Then
      lblCst = "Range Of Customers Selected."
   Else
      lblCst.ForeColor = vbBlack
   End If
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub



Private Sub optPrn_Click()
   PrintReport
   
End Sub
Private Sub cmbStartDate_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub cmbStartDate_LostFocus()
   If Len(Trim(cmbStartDate)) = 0 Then cmbStartDate = "ALL"
   If cmbStartDate <> "ALL" Then cmbStartDate = CheckDateEx(cmbStartDate)
   
End Sub


Private Sub cmbEndDate_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub cmbEndDate_LostFocus()
   If Len(Trim(cmbEndDate)) = 0 Then cmbEndDate = "ALL"
   If Trim(cmbEndDate) <> "ALL" Then cmbEndDate = CheckDateEx(cmbEndDate)
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(optExt.Value) _
              & RTrim(optDsc.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optExt.Value = Val(Mid(sOptions, 1, 1))
      optDsc.Value = Val(Mid(sOptions, 2, 1))
   Else
      optExt.Value = vbUnchecked
      optDsc.Value = vbUnchecked
   End If
   
End Sub
