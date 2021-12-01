VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form BookBKp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bookings By Salesperson"
   ClientHeight    =   3975
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
   ScaleHeight     =   3975
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BookBKp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtCst 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Sales Orders"
      Top             =   1680
      Width           =   1555
   End
   Begin VB.ComboBox cmbSlp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Salesperson From List"
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   3360
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox cmbDiv 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "8"
      ToolTipText     =   "Select Division From List"
      Top             =   2400
      Width           =   860
   End
   Begin VB.ComboBox cmbReg 
      Height          =   288
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "8"
      ToolTipText     =   "Select Region From List"
      Top             =   2760
      Width           =   780
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox optGrp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   3612
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "BookBKp03a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "BookBKp03a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3975
      FormDesignWidth =   7260
   End
   Begin VB.Label lblCUName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2040
      TabIndex        =   27
      Top             =   2040
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   11
      Left            =   5520
      TabIndex        =   25
      Tag             =   " "
      Top             =   1320
      Width           =   2385
   End
   Begin VB.Label lblSlp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3240
      TabIndex        =   24
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Salesperson"
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   23
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   10
      Left            =   5520
      TabIndex        =   22
      Tag             =   " "
      Top             =   2760
      Width           =   2388
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   9
      Left            =   5520
      TabIndex        =   21
      Tag             =   " "
      Top             =   2400
      Width           =   2388
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   8
      Left            =   5520
      TabIndex        =   20
      Tag             =   " "
      Top             =   1680
      Width           =   2385
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   6
      Left            =   360
      TabIndex        =   19
      Top             =   3120
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   288
      Index           =   5
      Left            =   360
      TabIndex        =   18
      Tag             =   " "
      Top             =   3360
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      Height          =   288
      Index           =   4
      Left            =   360
      TabIndex        =   17
      Tag             =   " "
      Top             =   2760
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   16
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Start Date"
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   15
      Top             =   1320
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chart"
      Height          =   288
      Index           =   7
      Left            =   360
      TabIndex        =   14
      Top             =   3612
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   360
      Picture         =   "BookBKp03a.frx":0AB6
      ToolTipText     =   "Chart Results"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   288
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Tag             =   " "
      Top             =   2400
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer(s)"
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   1425
   End
End
Attribute VB_Name = "BookBKp03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/25/05 Changed dates and Options
'8/30/06 Removed Chart
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbDiv_LostFocus()
   cmbDiv = CheckLen(cmbDiv, 4)
   If Len(cmbDiv) = 0 Then cmbDiv = "ALL"
   
End Sub

Private Sub cmbReg_LostFocus()
   cmbReg = CheckLen(cmbReg, 3)
   If Len(cmbReg) = 0 Then cmbReg = "ALL"
   
End Sub

Private Sub cmbSlp_Click()
   GetSalesPerson
   
End Sub

Private Sub cmbSlp_LostFocus()
   cmbSlp = CheckLen(cmbSlp, 4)
   GetSalesPerson
   
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
   sSql = "Qry_GetCustomerSalesOrder"
   LoadComboBox txtCst
   txtCst = "ALL"
   GetThisCustomer 1
   
   sSql = "Qry_FillSalesPersons"
   LoadComboBox cmbSlp, -1
   If cmbSlp.ListCount > 0 Then
      If Trim(cmbSlp) = "" Then cmbSlp = cmbSlp.List(0)
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
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillDivisions
      FillRegions
      If cmbDiv = "" Then cmbDiv = "ALL"
      If cmbReg = "" Then cmbReg = "ALL"
      If txtCst = "" Then txtCst = "ALL"
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
   Set BookBKp03a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sCust As String
   Dim sDiv As String
   Dim sReg As String
   Dim sBegDate As String
   Dim sEndDate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
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
   If Trim(txtCst) = "ALL" Then sCust = "" Else sCust = Compress(txtCst)
   If Trim(cmbDiv) = "ALL" Then sDiv = "" Else sDiv = cmbDiv
   If Trim(cmbReg) = "ALL" Then sReg = "" Else sReg = cmbReg
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDescription"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Customer(s) " & CStr(txtCst & ", Division " _
                        & cmbDiv & ", Region " & cmbReg & " From " & txtBeg _
                        & " To " & txtEnd) & "...'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optDsc.Value
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("slebk03")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{SohdTable.SOSALESMAN}='" & cmbSlp & "' "
   sSql = sSql & "AND {CustTable.CUREF} LIKE '" & sCust & "*' " _
          & "AND {SohdTable.SODIVISION} LIKE '" & sDiv & "*' " _
          & "AND {SohdTable.SOREGION} LIKE '" & sReg & "*' " _
          & "AND {SoitTable.ITBOOKDATE} In Date(" & sBegDate & ") " _
          & "To Date(" & sEndDate & ")"
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
   txtBeg = Format(ES_SYSDATE, "mm/01/yyyy")
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   cmbDiv.AddItem "ALL"
   cmbReg.AddItem "ALL"
   
   
End Sub

Private Sub SaveOptions()
   Dim sDiv As String * 4
   Dim sReg As String * 3
   Dim sCust As String * 10
   Dim sPers As String * 4
   Dim sOptions As String
   'Save by Menu Option
   sDiv = cmbDiv
   sReg = cmbReg
   sCust = txtCst
   sPers = cmbSlp
   sOptions = sDiv & sReg _
              & Trim(str(optDsc.Value)) _
              & Trim(str(optGrp.Value)) _
              & sCust & sPers
   SaveSetting "Esi2000", "EsiSale", "bk03a", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "bk03a", Trim(sOptions))
   If Len(sOptions) Then
      cmbDiv = Mid(sOptions, 1, 4)
      cmbReg = Mid(sOptions, 5, 3)
      optDsc = Mid(sOptions, 8, 1)
      optGrp = Mid(sOptions, 9, 1)
      txtCst = Mid(sOptions, 10, 10)
      cmbSlp = Mid(sOptions, 20, 4)
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 5)
   
End Sub

Private Sub Image1_Click()
   If optGrp.Value = vbChecked Then
      optGrp.Value = vbUnchecked
   Else
      optGrp.Value = vbChecked
   End If
   
End Sub

Private Sub lblSlp_Change()
   If Left(lblSlp, 9) = "*** No Va" Then
      lblSlp.ForeColor = ES_RED
   Else
      lblSlp.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optGrp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
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


Private Sub txtCst_Click()
   GetThisCustomer 1
   
End Sub

Private Sub txtCst_LostFocus()
   txtCst = CheckLen(txtCst, 10)
   If Len(txtCst) = 0 Then txtCst = "ALL"
   GetThisCustomer 1
   
End Sub



Private Sub GetSalesPerson()
   Dim rdoSlp As ADODB.Recordset
   sSql = "SELECT SPNUMBER,SPFIRST,SPMIDD,SPLAST " _
          & "FROM SprsTable WHERE SPNUMBER='" & cmbSlp & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSlp, ES_FORWARD)
   If bSqlRows Then
      With rdoSlp
         cmbSlp = "" & Trim(!SPNumber)
         lblSlp = "" & Trim(!SPFIRST) & " "
         If Len(Trim(!SPMIDD)) Then lblSlp = lblSlp & Trim(!SPMIDD) & ". "
         lblSlp = lblSlp & Trim(!SPLAST)
         ClearResultSet rdoSlp
      End With
   Else
      lblSlp = "*** No Valid Salesperson ***"
   End If
   Set rdoSlp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Trim(txtEnd) <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub

