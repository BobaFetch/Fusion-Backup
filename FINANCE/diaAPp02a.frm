VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Statements (Old)"
   ClientHeight    =   3225
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3225
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optPad 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With Invoices"
      Top             =   720
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5760
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   5
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
      PictureUp       =   "diaAPp02a.frx":0000
      PictureDn       =   "diaAPp02a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3600
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3225
      FormDesignWidth =   6930
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   17
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
      PictureUp       =   "diaAPp02a.frx":028C
      PictureDn       =   "diaAPp02a.frx":03D2
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Invoices"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Beginning Date"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   13
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Name"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   4
      Top             =   720
      Width           =   2025
   End
End
Attribute VB_Name = "diaAPp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*************************************************************************************
'
' diaAPe02a - Vendor Statments (Report)
'
' Notes:
'
' Created: (nth)
' Revisions:
'   10/22/03 (nth) Added custom report
'
'*************************************************************************************

Dim bOnLoad As Boolean
Dim bCancel As Boolean
Dim bGoodVendor As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmbVnd_Click()
   If cmbVnd <> "ALL" Or Trim(cmbVnd) <> "" Then
      lblNme.ForeColor = Me.ForeColor
      bGoodVendor = FindVendor(Me)
   Else
      lblNme.ForeColor = ES_BLUE
      lblNme = "*** Multiple Vendors Selected ***"
   End If
   
End Sub


Private Sub cmbVnd_LostFocus()
   Dim i As Integer
   Dim bByte As Boolean
   
   cmbVnd = CheckLen(cmbVnd, 10)
   If Len(cmbVnd) = 0 Then cmbVnd = "ALL"
   For i = 0 To cmbVnd.ListCount - 1
      If cmbVnd = cmbVnd.List(i) Then bByte = True
   Next
   If Not bByte Then
      cmbVnd = "ALL"
   End If
   
   If cmbVnd <> "ALL" Then
      lblNme.ForeColor = Me.ForeColor
      bGoodVendor = FindVendor(Me)
   Else
      lblNme.ForeColor = ES_BLUE
      lblNme = "*** Multiple Vendors Selected ***"
   End If
   
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
   Dim RdoVed As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT VIVENDOR,VEREF,VENICKNAME " _
          & "FROM VihdTable,VndrTable WHERE VIVENDOR=VEREF ORDER BY VIVENDOR"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      With RdoVed
         Do Until .EOF
            AddComboStr cmbVnd.hWnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoVed = Nothing
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
      cmbVnd = cUR.CurrentVendor
      FindVendor Me
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtBeg = Format(txtEnd, "mm/01/yy")
   GetOptions
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   bOnLoad = True
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bGoodVendor Then
      cUR.CurrentVendor = cmbVnd
      SaveCurrentSelections
   End If
   FormUnload
   On Error Resume Next
   Set diaAPp02a = Nothing
End Sub
Private Sub PrintReport()
   Dim sBeg As String
   Dim sEnd As String
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "IncDates"
   aFormulaName.Add "RequestBy"

   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbVnd) & "'")
   aFormulaValue.Add CStr("'From " & CStr(txtBeg) & " Through " & CStr(txtEnd) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   
   sCustomReport = GetCustomReport("finap02a.rpt")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   'sSql = "{EsReportVendorStmt.InvNumber} <> ''"
   sSql = ""
   cCRViewer.SetReportSelectionFormula sSql
   
'   sSql = "{VihdTable.VIDATE} >= cDate('" & Trim(txtBeg) & "') AND " _
'          & "{VihdTable.VIDATE} <= cDate('" & Trim(txtEnd) & "')"
'
'   If UCase(cmbVnd) <> "ALL" Then
'      sSql = sSql & " AND {VihdTable.VIVENDOR}='" & Compress(cmbVnd) & "'"
'   End If
'
'   If optPad.Value = vbUnchecked Then
'      sSql = sSql & " AND {VihdTable.VIPIF} = 0"
'   End If
'   cCRViewer.SetReportSelectionFormula sSql
   
   PrepareReportTable
   PopulateReportData
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sBeg As String
   Dim sEnd As String
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   'SetMdiReportsize MdiSect
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Includes='" & cmbVnd & "'"
   MdiSect.crw.Formulas(2) = "IncDates='From " & txtBeg & " Through " & txtEnd & "'"
   MdiSect.crw.Formulas(3) = "RequestBy = 'Requested By: " & sInitials & "'"
   
   sCustomReport = GetCustomReport("finap02a.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   'sSql = "{VihdTable.VIDATE} >= cDate('" & Trim(txtBeg) & "') AND " _
   '       & "{VihdTable.VIDATE} <= cDate('" & Trim(txtEnd) & "')"
   '
   'If UCase(cmbVnd) <> "ALL" Then
   '   sSql = sSql & " AND {VihdTable.VIVENDOR}='" & Compress(cmbVnd) & "'"
   'End If
   '
   'If optPad.Value = vbUnchecked Then
   '   sSql = sSql & " AND {VihdTable.VIPIF} = 0"
   'End If
   
   '
   'MdiSect.Crw.SelectionFormula = sSql
   PrepareReportTable
   PopulateReportData
   
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub SaveOptions()
   Dim sOptions As String
   ' Save by Menu Option
   sOptions = RTrim(optPad.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optPad.Value = Val(Left(sOptions, 1))
   Else
      optPad.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPad_KeyPress(KeyAscii As Integer)
'   KeyLock KeyAscii
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub




Private Sub PrepareReportTable()
    Dim SQL As String
    If TableExists("EsReportVendorStmt") Then
        SQL = "TRUNCATE TABLE EsReportVendorStmt"
    Else
        SQL = "CREATE TABLE [dbo].[EsReportVendorStmt]([VndrNick] [char](10) NULL, [VndrName] [char](40) NULL, [InvNumber] [char](20) NULL, " & vbCrLf _
           & "[InvDate] [datetime] NULL, [InvTotal] [decimal](12, 4) NULL, [CheckNo] [char](12) NULL, [CheckDate] [datetime] NULL, " & vbCrLf _
           & "[ChkVoid] [tinyint] NULL, [DCRef] [decimal](13, 2) NULL, [DCDebit] [decimal] (13,2) NULL, [DCCredit] [decimal](13, 2) NULL," & vbCrLf _
           & "CHKACCT varchar(12) NULL,[Journal] [varchar](12) NULL)"
    End If
    
    clsADOCon.ExecuteSql SQL
    
End Sub


Private Sub PopulateReportData()
    'Dim sWhereClause As String
    'Dim rdoTemp As ADODB.RecordSet
    'Dim sInsertQuery As String
    
    ' The first query is to get all the vendors that have invoices. I also will need to grab the ones that are just manual
    ' checks (with no invoice associated)
    
    MouseCursor vbHourglass
    
    sSql = "INSERT INTO EsReportVendorStmt (VndrNick,VndrName,InvNumber,InvDate," & vbCrLf _
        & "InvTotal, CheckNo, CheckDate, ChkVoid, DCRef, DCDebit, DCCredit, CHKACCT, Journal)" & vbCrLf _
        & "SELECT VEREF, VEBNAME, VINO, VIDATE, VIDUE, CHKNUMBER, CHKPOSTDATE, CHKVOID, DCREF, DCDEBIT, DCCREDIT, DCCHKACCT, DCHEAD" & vbCrLf _
        & "FROM VndrTable " & vbCrLf _
        & "INNER JOIN VihdTable on VndrTable.VEREF = VihdTable.VIVENDOR " & vbCrLf _
        & "LEFT OUTER JOIN JritTable on VihdTable.VIVENDOR = JritTable.DCVENDOR" & vbCrLf _
        & "AND VihdTable.VINO = JritTable.DCVENDORINV AND DCCHECKNO <> ''" & vbCrLf _
        & "LEFT OUTER JOIN ChksTable on JritTable.DCCHECKNO = ChksTable.CHKNUMBER AND DCCHKACCT = CHKACCT " & vbCrLf _
        & "WHERE VihdTable.VIDATE >= '" & Trim(txtBeg) & "' AND VihdTable.VIDATE <= '" & Trim(txtEnd) & "' " & vbCrLf
    If UCase(cmbVnd) <> "ALL" Then sSql = sSql & " AND VndrTable.VEREF='" & Compress(cmbVnd) & "' "
    If optPad.Value = vbUnchecked Then sSql = sSql & "AND VihdTable.VIPIF = 0 "
    clsADOCon.ExecuteSql sSql
    
    'Now insert manual checks info
'    sSql = "INSERT INTO EsReportVendorStmt (VndrNick, VndrName, InvNumber, InvDate, InvTotal, CheckNo, CheckDate, ChkVoid, DCRef, DCDebit, DCCredit, CHKACCT) " & _
'           " Select VEREF, VEBNAME, '', '', CHKAMOUNT, CHKNUMBER, CHKPOSTDATE, CHKVOID, 0, 0, 0, CHKACCT From ChksTable " & _
'           " INNER JOIN VndrTable ON ChksTable.CHKVENDOR=VndrTable.VEREF " & _
'           " WHERE ChksTable.CHKTYPE = 3 AND ChksTable.CHKPOSTDATE >= '" & Trim(txtBeg) & "' AND ChksTable.CHKPOSTDATE <= '" & Trim(txtEnd) & "' "
'    If UCase(cmbVnd) <> "ALL" Then sSql = sSql & "AND ChksTable.CHKVENDOR = '" & Compress(cmbVnd) & "' "
    
    'and JritTable.DCACCTNO = ChksTable.CHKACCT" & vbCrLf
    
'manual checks are being picked up above with the left join
'     sSql = "INSERT INTO EsReportVendorStmt (VndrNick, VndrName, InvNumber, InvDate," & vbCrLf _
'        & "InvTotal, CheckNo, CheckDate, ChkVoid, DCRef, DCDebit, DCCredit, CHKACCT, Journal) " & vbCrLf _
'        & "Select VEREF, VEBNAME, CHKMEMO as InvNumber, CHKPOSTDATE as InvDate, CHKAMOUNT, CHKNUMBER," & vbCrLf _
'        & "CHKPOSTDATE, CHKVOID, CHKTYPE as DCRef, 0 as DCDebit, DCCredit, CHKACCT, DCHEAD" & vbCrLf _
'        & "From ChksTable INNER JOIN VndrTable ON ChksTable.CHKVENDOR=VndrTable.VEREF" & vbCrLf _
'        & "INNER JOIN JritTable on JritTable.DCCHECKNO = ChksTable.CHKNUMBER" & vbCrLf _
'        & "WHERE ChksTable.CHKTYPE = 3 AND ChksTable.CHKPOSTDATE >= '" & Trim(txtBeg) & "'" & vbCrLf _
'        & "AND ChksTable.CHKPOSTDATE <= '" & Trim(txtEnd) & "'" & vbCrLf _
'        & "AND DCCREDIT > 0" & vbCrLf _
'        & "AND ChksTable.CHKVENDOR = '" & Compress(cmbVnd) & "'" & vbCrLf
'    clsADOCon.ExecuteSql sSql
   
    MouseCursor vbNormal

End Sub



