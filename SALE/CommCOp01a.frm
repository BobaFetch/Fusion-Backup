VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CommCOp01a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commission Status (Report)"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6915
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "CommCOp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CommCOp01a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox txtThr 
      Height          =   288
      Left            =   3840
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ComboBox txtFrm 
      Height          =   288
      Left            =   1560
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CheckBox optPay 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.ComboBox cmbSon 
      Height          =   288
      Left            =   1800
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Qualifying Sales Orders"
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cmbSlp 
      Height          =   288
      Left            =   1560
      TabIndex        =   0
      Tag             =   "3"
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   5640
      TabIndex        =   9
      Top             =   480
      Width           =   1335
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "CommCOp01a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "CommCOp01a.frx":0AC2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   6915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   23
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   21
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   252
      Index           =   6
      Left            =   5400
      TabIndex        =   18
      Top             =   1920
      Width           =   1332
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   252
      Index           =   4
      Left            =   2760
      TabIndex        =   17
      Top             =   1920
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Paid"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   1332
   End
   Begin VB.Label lblPre 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1560
      TabIndex        =   14
      Top             =   1560
      Width           =   252
   End
   Begin VB.Label lblSlp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1560
      TabIndex        =   13
      Top             =   1200
      Width           =   2772
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   1572
   End
End
Attribute VB_Name = "CommCOp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
' Created: 08/27/03 (jcw)
'09/04/03 (nth) Revised and updated.
'09/29/04 (nth) Added sales order from and through dates.
'11/14/05 (cjs) Reformatted entire dialog and criteria
'3/30/06 (cjs) added Item Comments
Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodSP As Byte
Dim bGoodSO As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbSlp_Click()
   'bGoodSP = ESICOMM.GetThisSalesPerson(Me)
   bGoodSP = GetSalesPerson
End Sub

Private Sub cmbSlp_LostFocus()
   If Not bCancel Then
      If Trim(cmbSlp) = "" Then
         cmbSlp = "ALL"
         lblSlp = "All salesmen selected"
      End If
   
      'bGoodSP = ESICOMM.GetThisSalesPerson(Me)
      bGoodSP = GetSalesPerson
      If bGoodSP Then
         FillSon
      Else
         lblPre = ""
         cmbSon.Clear
      End If
   End If
   
End Sub

Private Sub cmbSon_Click()
   bGoodSO = GetSalesOrder()
   
End Sub
   
Private Sub cmbSon_LostFocus()
   If Trim(cmbSon) = "" Then
      cmbSon = "ALL"
   Else
      cmbSon = Format(Abs(Val(cmbSon)), SO_NUM_FORMAT)
      bGoodSO = GetSalesOrder()
   End If
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bCash = InvOrCash()
   txtFrm = Format(ES_SYSDATE, "mm/01/yy")
   txtThr = Format(ES_SYSDATE, "mm/dd/yy")
   GetOptions
   bOnLoad = 1
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillSalesPersons Me
      FillSon
   End If
   bOnLoad = 0
   MouseCursor 0
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set CommCOp01a = Nothing
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub lblPre_Change()
   If Trim(lblPre) = "*" Then lblPre.ForeColor = ES_RED _
           Else lblPre.ForeColor = vbBlack
   
End Sub

Private Sub lblSlp_Change()
   If Left(lblSlp, 6) = "*** Sa" Then lblSlp.ForeColor = ES_RED _
           Else lblSlp.ForeColor = vbBlack
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   If bGoodSP Then PrintReport _
      Else MsgBox "Requires A Valid Salesperson.", _
      vbInformation, Caption
   
End Sub

Private Sub PrintReport()
   Dim sBegDate As String
   Dim sEndDate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If IsDate(txtFrm) Then
      sBegDate = Format(txtFrm, "yyyy,mm,dd")
   Else
      sBegDate = "1995,01,01"
   End If
   
   If IsDate(txtThr) Then
      sEndDate = Format(txtThr, "yyyy,mm,dd")
   Else
      sEndDate = "2024,12,31"
   End If
   
'   SetMdiReportsize MdiSect
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("slecm01.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   
   If Trim(cmbSlp) <> "" And cmbSlp <> "ALL" Then
      sSql = "{SpcoTable.SMCOSM} = '" & Trim(cmbSlp) & "'" & vbCrLf
   Else
      sSql = ""
   End If
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'For Sales Person  " & CStr(Trim(cmbSlp) & " - " _
                        & Trim(lblSlp)) & "'")
   aFormulaName.Add "Title2"
   If Trim(cmbSon) <> "" And cmbSon <> "ALL" Then
      'MdiSect.Crw.Formulas(3) = "Title2 = 'Sales Order:  " & lblPre & cmbSon & "'"
      aFormulaValue.Add CStr("'Sales Order: " & CStr(lblPre & cmbSon) & "'")
      If sSql <> "" Then
         sSql = sSql & "AND "
      End If
      sSql = sSql & "{SohdTable.SONUMBER} = " & Val(cmbSon) & " " & vbCrLf
   Else
      aFormulaValue.Add CStr("'Sales Orders From " & CStr(txtFrm & " Through " & txtThr) & "'")
   End If
   aFormulaName.Add "Inv/CR"
   aFormulaValue.Add CStr(bCash)
   
   If sSql <> "" Then
      sSql = sSql & "AND "
   End If
   sSql = sSql & "{SohdTable.SODATE} In Date(" & sBegDate _
          & ") to Date(" & sEndDate & ") " & vbCrLf
   
   'include paid commissions?
   '(these will appear as paid AP invoices)
   If optPay Then
      aFormulaName.Add "Title3"
      aFormulaValue.Add CStr("'*** Excludes Paid Commissions ***'")
      If sSql <> "" Then
         sSql = sSql & "AND "
      End If
      sSql = sSql & "(ISNULL({VihdTable.VIPIF})"
      sSql = sSql & " OR {VihdTable.VIPIF} = 0)" & vbCrLf
   End If
   
   sSql = sSql & " AND {SohdTable.SOCANCELED} <> 1"
   aFormulaName.Add "ShowItemComments"
   aFormulaName.Add "ExcludePaidCommissions"
   
   aFormulaValue.Add optCmt.Value
   aFormulaValue.Add CStr("'" & CStr(optPay) & "'")
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
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
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function GetSalesOrder() As Byte
   Dim RdoSon As ADODB.Recordset
   If cmbSon = "ALL" Then
      lblPre = ""
   Else
      sSql = "SELECT SOTYPE FROM SohdTable WHERE SONUMBER = " & Val(cmbSon) & " "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
      If bSqlRows Then
         With RdoSon
            lblPre = "" & Trim(!SOTYPE)
            lblPre.ToolTipText = ""
            GetSalesOrder = 1
         End With
         ClearResultSet RdoSon
      Else
         GetSalesOrder = 0
         cmbSon = "ALL"
         lblPre = " *"
         lblPre.ToolTipText = "Sales Order Wasn't Found"
      End If
   End If
   Set RdoSon = Nothing
   
End Function

Private Sub FillSon()
   Dim RdoSon As ADODB.Recordset
   cmbSon.Clear
   sSql = "SELECT DISTINCT SONUMBER FROM SohdTable,SpcoTable WHERE " _
          & "SMCOSO = SONUMBER ORDER BY SONUMBER"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         Do Until .EOF
            AddComboStr cmbSon.hWnd, Format(.Fields(0), SO_NUM_FORMAT)
            .MoveNext
         Loop
         ClearResultSet RdoSon
      End With
   End If
   If cmbSon.ListCount > 0 Then
      cmbSon = cmbSon.List(0)
      bGoodSO = GetSalesOrder()
   End If
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillson"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiSale", Me.Name & "_Printer", lblPrinter
   sOptions = optPay.Value
   SaveSetting "Esi2000", "EsiSale", Me.Name, sOptions
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optPay.Value = Mid(sOptions, 1, 1)
   End If
   lblPrinter = GetSetting("Esi2000", "EsiSale", Me.Name & "_Printer", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub optPay_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   If bGoodSP Then PrintReport _
      Else MsgBox "Requires A Valid Salesperson.", _
      vbInformation, Caption
   
End Sub

Private Sub txtFrm_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtFrm_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtFrm_LostFocus()
   If txtFrm = "" Or txtFrm = "ALL" Then
      txtFrm = "ALL"
   Else
      txtFrm = CheckDate(txtFrm)
   End If
   
End Sub

Private Sub txtThr_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtThr_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtThr_LostFocus()
   If txtThr = "" Or txtThr = "ALL" Then
      txtThr = "ALL"
   Else
      txtThr = CheckDate(txtThr)
   End If
   
End Sub

Private Function GetSalesPerson() As Byte
   Dim RdoNme As ADODB.Recordset
   
   If cmbSlp = "ALL" Then
      GetSalesPerson = True
      Exit Function
   End If
   
   On Error GoTo ModErr1
   sSql = "SELECT SPLAST,SPFIRST FROM SprsTable WHERE SPNUMBER = '" _
          & Trim(cmbSlp) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNme)
   If bSqlRows Then
      With RdoNme
         lblSlp = "" & Trim(.Fields(1)) & " " & Trim(.Fields(0))
      End With
      GetSalesPerson = 1
      ClearResultSet RdoNme
   Else
      GetSalesPerson = 0
      lblSlp = "*** Sales Person Not Found ***"
   End If
   Set RdoNme = Nothing
   Exit Function
   
ModErr1:
   sProcName = "getspsos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function



