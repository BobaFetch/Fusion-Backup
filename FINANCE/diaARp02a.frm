VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaARp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Statements (Report)"
   ClientHeight    =   4245
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4245
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPONumber 
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox cbByPO 
      Caption         =   "By PO"
      Height          =   375
      Left            =   3720
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkIncludeZeroBalance 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   3180
      Width           =   735
   End
   Begin VB.CheckBox chkCrDm 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   3780
      Width           =   735
   End
   Begin VB.CheckBox chkskip 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   2880
      Width           =   735
   End
   Begin VB.ComboBox txtSDte 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtdays 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Tag             =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkIncludePaid 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   3480
      Width           =   735
   End
   Begin VB.ComboBox txtEdte 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Contains Customers With Unpaid Invoices"
      Top             =   480
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
      TabIndex        =   14
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaARp02a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaARp02a.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   12
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
      PictureUp       =   "diaARp02a.frx":0308
      PictureDn       =   "diaARp02a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6000
      Top             =   3840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4245
      FormDesignWidth =   6480
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   225
      Left            =   360
      TabIndex        =   13
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaARp02a.frx":0594
      PictureDn       =   "diaARp02a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
      Height          =   255
      Index           =   13
      Left            =   3600
      TabIndex        =   31
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Zero Balance Statements"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   29
      Top             =   3180
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Apply Credit and Debit Memos To Invoices"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   28
      Top             =   3780
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Or"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   27
      Top             =   2040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   10
      Left            =   3360
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   9
      Left            =   3360
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices From"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Skip Payments After ""Through"" Date"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   23
      Top             =   2880
      Width           =   2625
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Or More Days Old"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   22
      Top             =   2400
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Only Invoices "
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   20
      Top             =   480
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Paid Invoices And Applied Credits"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   675
      TabIndex        =   18
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Top             =   840
      Width           =   3000
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer "
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   1425
   End
End
Attribute VB_Name = "diaARp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'************************************************************************************
' diaARp02a - Customer Statement (Report)
'
' Created: (cjs)
' Revisions:
' 11/29/01 (nth) To function like MCS
' 08/13/02 (nth) Fixed runtime error with report selection formula.
' 10/23/03 (nth) Updated report linking
' 10/23/03 (nth) Added custom report
' 02/25/04 (nth) Fixed error with exclude payments after, was excluding all.
' 11/16/04 (nth) Fixed show paid invoices and applied credits option.
' 03/23/05  cjs  Added PrintThemAll (only listed customers)
'************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

Dim iUserLogo As Integer

Dim sTableDef As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


'************************************************************************************

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Trim(cmbCst) = "" Then
      cmbCst = "ALL"
      lblNme = "Multiple Customers Selected."
   Else
      FindCustomer Me, cmbCst
   End If
   FillPOCombo
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillCombo()
   Dim rdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT CUREF,CUNICKNAME,INVCUST FROM " _
        & "CustTable, CihdTable WHERE CUREF=INVCUST"

'          & "CustTable,CihdTable WHERE (CUREF=INVCUST AND INVCANCELED=0 " _
'          & "AND INVPAY<>INVTOTAL)"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst, ES_FORWARD)
   If bSqlRows Then
      With rdoCst
         Do Until .EOF
            AddComboStr cmbCst.hWnd, "" & Trim(!CUNICKNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
  
   If cmbCst.ListCount > 0 Then
      cmbCst = cmbCst.List(0)
      FindCustomer Me, cmbCst
   End If
   Set rdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CreateTempTable
      FillCombo
      
      ' Check to see if we nned to use company logo
      GetUseLogo
      If cbByPO.Value = vbChecked Then
        z1(13).Left = 240
        z1(13).Top = 1680
        cmbPONumber.Left = 1440
        cmbPONumber.Top = 1680
        z1(13).Visible = True
        cmbPONumber.Visible = True
        z1(0).Visible = False
        z1(4).Visible = False
        z1(5).Visible = False
        z1(8).Visible = False
        z1(9).Visible = False
        z1(10).Visible = False
        z1(11).Visible = False
        txtsDte.Visible = False
        txteDte.Visible = False
        txtDays.Visible = False
      Else
        z1(13).Visible = False
        cmbPONumber.Visible = False
        z1(0).Visible = True
        z1(4).Visible = True
        z1(5).Visible = True
        z1(8).Visible = True
        z1(9).Visible = True
        z1(10).Visible = True
        z1(11).Visible = True
        txtsDte.Visible = True
        txteDte.Visible = True
        txtDays.Visible = True
      End If
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   txtDays = "0"
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
   On Error Resume Next
   'Redundant because the table is dropped with the
   'connect, but reduces clutter
   sSql = "DROP TABLE " & sTableDef & " "
   clsADOCon.ExecuteSQL sSql
   
   FormUnload
   Set diaARp02a = Nothing
End Sub

Private Sub PrintReport()
   Dim sCust As String
   Dim sOld As String
   Dim sCustomReport As String
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   sCust = Compress(cmbCst)
   
   SetCrystalParameters sCust
'   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCust As String
   Dim sOld As String
   'Dim sStart As String
   'Dim sEnd   As String
   Dim sCustomReport As String
   
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   sCust = Compress(cmbCst)
   '    sStart = Trim(txtSDte)
   '    sEnd = Trim(txtEdte)
   '
   '    If sEnd = "" Then sEnd = "ALL"
   '    If sStart = "" Then sStart = "ALL"
   '
   '    SetMdiReportsize MdiSect
   '    sCustomReport = GetCustomReport("finar02.rpt")
   '    MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   '
   '    MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   '    MdiSect.crw.Formulas(1) = "InvoicesThrough='" & txtEdte & "'"
   '
   '    MdiSect.crw.Formulas(2) = "Title1='Includes Only Invoices From " _
   '        & sStart & " Through " & sEnd & "'"
   '    MdiSect.crw.Formulas(3) = "Title2='nothing'"
   '
   '    sSql = "{CihdTable.INVCANCELED} = 0 "
   '
   '    If UCase(Trim(cmbCst)) <> "ALL" Then
   '        sSql = sSql & "AND {CihdTable.INVCUST}='" & sCust & "'"
   '    End If
   '
   '    If IsDate(txtEdte) Then
   '        sSql = sSql & " AND {CihdTable.INVDATE} <= #" & txtEdte & "#"
   '    End If
   '    If IsDate(txtSDte) Then
   '        sSql = sSql & " AND {CihdTable.INVDATE} >= #" & txtSDte & "#"
   '    End If
   '    MdiSect.crw.Formulas(4) = "SkipPay=" & Val(chkskip)
   '    MdiSect.crw.Formulas(5) = chkIncludePaid.Name & "=" & chkIncludePaid.Value
   '    MdiSect.crw.Formulas(6) = "ShowZeroBalance='" & Val(chkIncludeZeroBalance) & "'"
   '    MdiSect.crw.SelectionFormula = sSql
   
   SetCrystalParameters sCust
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Sub SetCrystalParameters(sCust As String)
   
   Dim sStart As String
   Dim sEnd As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   sStart = Trim(txtsDte)
   sEnd = Trim(txteDte)
   
   If sEnd = "" Then sEnd = "ALL"
   If sStart = "" Then sStart = "ALL"
   
'   SetMdiReportsize MdiSect
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finar02.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "InvoicesThrough"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(txteDte) & "'")
   aFormulaValue.Add CStr("'Includes Only Invoices From " & CStr(sStart & " Through " & sEnd) & "'")
   aFormulaValue.Add CStr("'Nothing'")

'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "InvoicesThrough='" & txtEdte & "'"
'   MdiSect.Crw.Formulas(2) = "Title1='Includes Only Invoices From " _
'                        & sStart & " Through " & sEnd & "'"
'   MdiSect.Crw.Formulas(3) = "Title2='nothing'"
   
   
   sSql = "{CihdTable.INVCANCELED} = 0 "
   
   'If UCase(Trim(cmbCst)) <> "ALL" Then
   sSql = sSql & "AND {CihdTable.INVCUST}='" & sCust & "'"
   'End If
   
   If cbByPO.Value = vbChecked Then
        sSql = sSql & BuildPOWhereClause(cmbPONumber)
   Else
        If IsDate(txteDte) Then
          sSql = sSql & " AND {CihdTable.INVDATE} <= #" & txteDte & "#"
        End If
        If IsDate(txtsDte) Then
            sSql = sSql & " AND {CihdTable.INVDATE} >= #" & txtsDte & "#"
        End If
   End If
   aFormulaName.Add "SkipPay"
   aFormulaName.Add "ShowZeroBalance"
   aFormulaName.Add "ShowPaid"
   aFormulaName.Add "ShowMemos"
   aFormulaName.Add "ShowOurLogo"

   aFormulaValue.Add CStr(Val(chkskip))
   aFormulaValue.Add CStr(Val(chkIncludeZeroBalance))
   aFormulaValue.Add CStr(chkIncludePaid.Value)
   aFormulaValue.Add CStr(Val(chkCrDm))
   aFormulaValue.Add CStr("'" & CStr(iUserLogo) & "'")

'   MdiSect.Crw.Formulas(4) = "SkipPay=" & Val(chkskip)
'   MdiSect.Crw.Formulas(5) = "ShowZeroBalance=" & Val(chkIncludeZeroBalance)
'   MdiSect.Crw.Formulas(6) = "ShowPaid=" & chkIncludePaid.Value
'   MdiSect.Crw.Formulas(7) = "ShowMemos=" & Val(chkCrDm)
'    ' Set the logo field
'   MdiSect.Crw.Formulas(8) = "ShowOurLogo='" & iUserLogo & "'"
   
'   MdiSect.Crw.SelectionFormula = sSql
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

End Sub

Sub SetCrystalParameters1(sCust As String)
   
   Dim sStart As String
   Dim sEnd As String
   
   'sCust = Compress(cmbCst)
   sStart = Trim(txtsDte)
   sEnd = Trim(txteDte)
   
   If sEnd = "" Then sEnd = "ALL"
   If sStart = "" Then sStart = "ALL"
   
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("finar02.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "InvoicesThrough='" & txteDte & "'"
   MdiSect.crw.Formulas(2) = "Title1='Includes Only Invoices From " _
                        & sStart & " Through " & sEnd & "'"
   MdiSect.crw.Formulas(3) = "Title2='nothing'"
   
   
   sSql = "{CihdTable.INVCANCELED} = 0 "
   
   'If UCase(Trim(cmbCst)) <> "ALL" Then
   sSql = sSql & "AND {CihdTable.INVCUST}='" & sCust & "'"
   'End If
   
   If cbByPO.Value = vbChecked Then
        sSql = sSql & BuildPOWhereClause(cmbPONumber)
   Else
       If IsDate(txteDte) Then
          sSql = sSql & " AND {CihdTable.INVDATE} <= #" & txteDte & "#"
       End If
       If IsDate(txtsDte) Then
          sSql = sSql & " AND {CihdTable.INVDATE} >= #" & txtsDte & "#"
       End If
   End If
   MdiSect.crw.Formulas(4) = "SkipPay=" & Val(chkskip)
   MdiSect.crw.Formulas(5) = "ShowZeroBalance=" & Val(chkIncludeZeroBalance)
   MdiSect.crw.Formulas(6) = "ShowPaid=" & chkIncludePaid.Value
   MdiSect.crw.Formulas(7) = "ShowMemos=" & Val(chkCrDm)
    ' Set the logo field
   MdiSect.crw.Formulas(8) = "ShowOurLogo='" & iUserLogo & "'"
   
   MdiSect.crw.SelectionFormula = sSql
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txteDte = Format(Now, "mm/dd/yy")
   txtsDte = Left(txteDte, 3) & "01" & Right(txteDte, 3)
   
End Sub

Public Sub SaveOptions()
   Dim sbuf As String
   sbuf = chkIncludePaid.Value & chkskip.Value & chkCrDm.Value & chkIncludeZeroBalance
   SaveSetting "Esi2000", "EsiFina", Me.Name, sbuf
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, "000000")
   If Len(sOptions) Then
      chkIncludePaid.Value = Mid(sOptions, 1, 1)
      chkskip.Value = Mid(sOptions, 2, 1)
      chkCrDm.Value = Mid(sOptions, 3, 1)
      If Len(sOptions) > 3 Then
         chkIncludeZeroBalance.Value = Mid(sOptions, 4, 1)
      End If
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub optDis_Click()
   If Trim(cmbCst) = "" Or Trim(cmbCst) = "ALL" Then
      MsgBox "You can only display statements for one customer at a time." _
         & "'ALL' is only an option for printing.  Please select an individual customer to display.", _
         vbInformation, Caption
   Else
      PrintReport
   End If
   
End Sub

Private Sub optPrn_Click()
   If Trim(cmbCst) = "" Then cmbCst = "ALL"
   If Trim(cmbCst) = "ALL" Then
      GetInvCustomers
      PrintAllStatements
   Else
      PrintReport
   End If
   
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub


Private Sub txtDays_LostFocus()
   If Trim(txtDays) = "" Then
      txtDays = "0"
      txteDte = Format(Now, "mm/dd/yy")
   Else
      txteDte = Format(DateAdd("d", (-1 * Val(txtDays)), Now), "mm/dd/yy")
   End If
End Sub

Private Sub txtEDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEdte_LostFocus()
   If Trim(txteDte) = "" Or txteDte = "ALL" Then
      txteDte = "ALL"
      txtDays = "0"
   Else
      txteDte = CheckDate(txteDte)
      If CDate(txteDte) <= Now Then
         txtDays = DateDiff("d", txteDte, Now)
      Else
         txtDays = "0"
      End If
   End If
End Sub

Private Sub txtSDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtSDte_LostFocus()
   If Trim(txtsDte) = "" Or txtsDte = "ALL" Then
      txtsDte = "ALL"
   Else
      txtsDte = CheckDate(txtsDte)
   End If
End Sub

Public Sub PrintAllStatements()
   Dim RdoInv As ADODB.Recordset
   Dim sCust As String
   Dim sOld As String
   'Dim sStart As String
   'Dim sEnd   As String
   Dim sCustomReport As String
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   'sStart = Trim(txtSDte)
   'sEnd = Trim(txtEdte)
   
   'If sEnd = "" Then sEnd = "ALL"
   'If sStart = "" Then sStart = "ALL"
   
   'sCustomReport = GetCustomReport("finar02.rpt")
   'SetMdiReportsize MdiSect
   
   sSql = "SELECT CUREF FROM " & sTableDef & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   If bSqlRows Then
      With RdoInv
         Do Until .EOF
            optPrn.enabled = False
            optDis.enabled = False
            sCust = "" & Trim(!CUREF)
            'MdiSect.crw.ReportFileName = sReportPath & sCustomReport
            'MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
            'MdiSect.crw.Formulas(1) = "InvoicesThrough='" & txtEdte & "'"
            'MdiSect.crw.Formulas(2) = "Title1='Includes Only Invoices From " _
            '    & sStart & " Through " & sEnd & "'"
            
            'sSql = "{CihdTable.INVCANCELED} = 0 "
            'sSql = sSql & "AND {CihdTable.INVCUST}='" & sCust & "'"
            
            'If IsDate(txtEdte) Then
            '    sSql = sSql & " AND {CihdTable.INVDATE} <= #" & txtEdte & "#"
            'End If
            
            'If IsDate(txtSDte) Then
            '    sSql = sSql & " AND {CihdTable.INVDATE} >= #" & txtSDte & "#"
            'End If
            'MdiSect.crw.Formulas(4) = "SkipPay='" & Val(chkskip) & "'"
            'MdiSect.crw.Formulas(5) = chkIncludePaid.Name & "=" & chkIncludePaid.Value
            'MdiSect.crw.Formulas(6) = "ShowZeroBalance='" & Val(chkIncludeZeroBalance) & "'"
            'MdiSect.crw.SelectionFormula = sSql
            
            Dim bPrint As Boolean
            Dim bContinue As Boolean
            bPrint = True
            If Debugging Then
               Select Case MsgBox("Next Statement for " & sCust & ". Print?", vbYesNoCancel)
                  Case vbYes
                  Case vbNo
                     bPrint = False
                  Case vbCancel
                     Exit Do
               End Select
            End If
            
            If bPrint Then
               SetCrystalParameters sCust
               'SetCrystalAction Me
               optPrn.enabled = False
               optDis.enabled = False
               Sleep 1000
            End If
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      MsgBox "No Qualifying Customers Were Selected.", _
         vbInformation, Caption
   End If
   optPrn.enabled = True
   optDis.enabled = True
   MouseCursor 0
   Set RdoInv = Nothing
   Exit Sub
   
DiaErr1:
   optPrn.enabled = True
   optDis.enabled = True
   sProcName = "printallstate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetInvCustomers()
   Dim rdoCst As ADODB.Recordset
   Dim sStartDate As String
   Dim sEnddate As String
   
   On Error GoTo DiaErr1
   If txtsDte = "ALL" Then
      sStartDate = "01/01/1990"
   Else
      sStartDate = txtsDte
   End If
   If txteDte = "ALL" Then
      sEnddate = "01/01/2024"
   Else
      sEnddate = txteDte
   End If
   
   sSql = "TRUNCATE TABLE " & sTableDef & " "
   sSql = "SELECT DISTINCT CUREF,CUNICKNAME,INVCUST FROM " _
          & "CustTable,CihdTable WHERE (CUREF=INVCUST AND INVCANCELED=0 " _
          & "AND INVPAY<>INVTOTAL AND INVDATE BETWEEN '" & sStartDate _
          & "' AND '" & sEnddate & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst, ES_FORWARD)
   If bSqlRows Then
      With rdoCst
         Do Until .EOF
            sSql = "INSERT INTO " & sTableDef & " " _
                   & "(CUREF) VALUES('" & Trim(!CUREF) & "')"
            clsADOCon.ExecuteSQL sSql
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set rdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getinvcust"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'3/23/05 Temp Table to gather customers

Private Sub CreateTempTable()
   On Error Resume Next
   sTableDef = "##" & sInitials & "Statements"
   sSql = "CREATE TABLE " & sTableDef & " " _
          & "(CUREF char(10) NULL DEFAULT(''))"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "CREATE UNIQUE CLUSTERED INDEX CustomerIdx ON " _
          & sTableDef & "(CUREF) WITH FILLFACTOR = 80"
   clsADOCon.ExecuteSQL sSql
   
End Sub
Private Sub GetUseLogo()
    Dim RdoLogo As ADODB.Recordset
    Dim bRows As Boolean
    ' Assumed that COMREF is 1 all the time
    sSql = "SELECT ISNULL(COLUSELOGO, 0) as COLUSELOGO FROM ComnTable WHERE COREF = 1"
    bRows = clsADOCon.GetDataSet(sSql, RdoLogo, ES_FORWARD)

    If bRows Then
        With RdoLogo
            iUserLogo = !COLUSELOGO
        End With
        'RdoLogo.Close
        ClearResultSet RdoLogo
    End If
    Set RdoLogo = Nothing
End Sub


Private Function BuildPOWhereClause(ByVal PONumber As String) As String
    Dim rdoInvoices As ADODB.Recordset
    Dim bRows As Boolean
    Dim sTemp As String
    
    sTemp = sSql
    BuildPOWhereClause = ""

    sSql = "SELECT DISTINCT ITINVOICE FROM SohdTable " & _
      " INNER JOIN SoitTable ON SONUMBER=ITSO " & _
      " WHERE SOPO = '" & PONumber & "' AND SOCUST='" & cmbCst & "' AND ITINVOICE>0 "
    bRows = clsADOCon.GetDataSet(sSql, rdoInvoices, ES_FORWARD)
    If bRows Then
      With rdoInvoices
        Do Until .EOF
          BuildPOWhereClause = BuildPOWhereClause & LTrim(str(!ITINVOICE)) & ","
          .MoveNext
        Loop
        .Cancel
      End With
      If Len(BuildPOWhereClause) > 0 Then
        BuildPOWhereClause = Left(BuildPOWhereClause, Len(BuildPOWhereClause) - 1)
        BuildPOWhereClause = " AND {CihdTable.INVNO} IN [" & BuildPOWhereClause & "]"
      End If
    End If
    Set rdoInvoices = Nothing
    sSql = sTemp
    'BuildPOWhereClause = sWhere
End Function



Private Sub FillPOCombo()
    Dim sCustomer As String
    
    If cbByPO.Value = vbUnchecked Then Exit Sub
    If UCase(Trim(cmbCst)) = "ALL" Then sCustomer = "" Else sCustomer = cmbCst
    cmbPONumber.Clear
    sSql = "SELECT DISTINCT TOP 30000 SOPO FROM SohdTable WHERE SOPO IS NOT NULL AND SOCUST LIKE '" & sCustomer & "%' ORDER BY SOPO DESC "
   
    LoadComboBox cmbPONumber, -1
    
    
End Sub
