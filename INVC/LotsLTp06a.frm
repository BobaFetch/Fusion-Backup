VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LotsLTp06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mismatched Lots"
   ClientHeight    =   3855
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7365
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTp06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.OptionButton optShow 
      Alignment       =   1  'Right Justify
      Caption         =   "Both"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   27
      ToolTipText     =   "Include Lot Tracked And Not Lot Tracked Items"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.OptionButton optShow 
      Alignment       =   1  'Right Justify
      Caption         =   "Parts Not Tracked"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   26
      ToolTipText     =   "Included Only Items Not Lot Tracked "
      Top             =   2280
      Width           =   1815
   End
   Begin VB.OptionButton optShow 
      Alignment       =   1  'Right Justify
      Caption         =   "Parts Tracked"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   25
      ToolTipText     =   "Include Only Lot Tracked Items"
      Top             =   1920
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   6400
      TabIndex        =   2
      ToolTipText     =   "Build The Report"
      Top             =   1080
      Width           =   875
   End
   Begin VB.TextBox lblDsc 
      Height          =   285
      Left            =   1900
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Leading Char Search  (*  In Front Is A Legal Wild Card)"
      Top             =   1440
      Width           =   3075
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1900
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Leading Char Search  (*  In Front Is A Legal Wild Card)"
      Top             =   1080
      Width           =   3075
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2520
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "LotsLTp06a.frx":07AE
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
         Picture         =   "LotsLTp06a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3855
      FormDesignWidth =   7365
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1920
      TabIndex        =   29
      Top             =   3360
      Width           =   4452
      _ExtentX        =   7858
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Compares Lots By Part Number Where The Lot Activity Doesn't Match The Lot Remaining Quantity"
      Height          =   405
      Index           =   11
      Left            =   240
      TabIndex        =   24
      Tag             =   "P"
      ToolTipText     =   "Total Number Of Parts In The Table"
      Top             =   480
      Width           =   5265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Processed"
      Height          =   285
      Index           =   10
      Left            =   2760
      TabIndex        =   23
      Tag             =   "P"
      ToolTipText     =   "Rows Processed"
      Top             =   3000
      Width           =   1665
   End
   Begin VB.Label lblRows 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   22
      ToolTipText     =   "Rows Processed"
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblPrg 
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "All Lot Detail"
      Height          =   285
      Index           =   9
      Left            =   5520
      TabIndex        =   20
      Tag             =   "P"
      ToolTipText     =   "Total Number Of Lot Transactions In The Table"
      Top             =   2640
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "All Lots"
      Height          =   285
      Index           =   8
      Left            =   5520
      TabIndex        =   19
      Tag             =   "P"
      ToolTipText     =   "Total Number Of Lots  In The Table"
      Top             =   2280
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "All Parts"
      Height          =   285
      Index           =   7
      Left            =   5520
      TabIndex        =   18
      Tag             =   "P"
      ToolTipText     =   "Total Number Of Parts In The Table"
      Top             =   1920
      Width           =   1545
   End
   Begin VB.Label lblRows 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   17
      ToolTipText     =   "Total Number Of Lot Transactions In The Table"
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblRows 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   16
      ToolTipText     =   "Total Number Of Lots  In The Table"
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Detail Rows"
      Height          =   285
      Index           =   6
      Left            =   2760
      TabIndex        =   15
      Tag             =   "P"
      ToolTipText     =   "Total Number Of Lot Transactions In The Table"
      Top             =   2640
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Rows"
      Height          =   285
      Index           =   5
      Left            =   2760
      TabIndex        =   14
      Tag             =   "P"
      ToolTipText     =   "Total Number Of Lots  In The Table"
      Top             =   2280
      Width           =   1545
   End
   Begin VB.Label lblRows 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   13
      ToolTipText     =   "Total Number Of Parts In The Table"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Table Rows"
      Height          =   285
      Index           =   4
      Left            =   2760
      TabIndex        =   12
      Tag             =   "P"
      ToolTipText     =   "Total Number Of Parts In The Table"
      Top             =   1920
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5160
      TabIndex        =   10
      Top             =   1120
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   9
      Tag             =   " "
      Top             =   4200
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1120
      Width           =   1425
   End
End
Attribute VB_Name = "LotsLTp06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'5/13/05 New
'5/16/05 Added Lot type selection criteria
Option Explicit
Dim bOnLoad As Byte
Dim lPartCount As Long
Dim lLotCount As Long
Dim lItemCount As Long

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT COUNT(PARTREF) FROM PartTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then lPartCount = RdoCmb.Fields(0)
   lblRows(0) = lPartCount
   
   sSql = "SELECT COUNT(LOTNUMBER) FROM LohdTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then lLotCount = RdoCmb.Fields(0)
   lblRows(1) = lLotCount
   
   sSql = "SELECT COUNT(LOINUMBER) FROM LoitTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then lItemCount = RdoCmb.Fields(0)
   lblRows(2) = lItemCount
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5580"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSel_Click()
   Dim bResponse As Byte
   txtPrt = Trim(txtPrt)
   If txtPrt = "" Then txtPrt = "ALL"
   If Len(txtPrt) > 0 Then lblDsc = "*** Range Of Parts Selected ***"
   If lblDsc = "" Or txtPrt = "ALL" Then _
               lblDsc = "*** All Parts Selected ***"
   
   bResponse = MsgBox("Run The Report Data?...          ", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      If lPartCount > 10000 And txtPrt = "ALL" Then
         bResponse = MsgBox("This Report May Take 5 Minutes Or More To Run." & vbCr _
                     & "Are You Certain That You Want All Part Numbers?", _
                     ES_NOQUESTION, Caption)
         If bResponse = vbNo Then
            CancelTrans
         End If
      End If
      If bResponse = vbYes Then FillTable
   Else
      CancelTrans
   End If
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      BuildTable
   End If
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
   Set LotsLTp06a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sShow As String
   MouseCursor 13
   If optShow(0).Value = True Then
      sShow = ", Items Lot Tracked"
   ElseIf optShow(1).Value = True Then sShow = ", Items Not Lot Tracked"
      
   Else
      If optShow(2).Value = True Then sShow = ", Items Lot Tracked And Not Lot Tracked"
   End If
   On Error GoTo DiaErr1
'   SetMdiReportsize MdiSect
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
'   MdiSect.Crw.Formulas(2) = "Includes='Parts " & txtPrt & sShow & "...'"
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Includes"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'Parts " & CStr(txtPrt & sShow) & "...'")
   sCustomReport = GetCustomReport("invlt05")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = ""
'   MdiSect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   cCRViewer.SetReportSelectionFormula (sSql)
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
   lblDsc.BackColor = Me.BackColor
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   On Error Resume Next
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = Trim(txtPrt)
   If txtPrt = "" Then txtPrt = "ALL"
   If Len(txtPrt) > 0 Then lblDsc = "*** Range Of Parts Selected ***"
   If lblDsc = "" Or txtPrt = "ALL" Then _
               lblDsc = "*** All Parts Selected ***"
   
End Sub

Private Sub BuildTable()
   'Dim RdoTable As ADODB.Recordset
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT LotNum FROM EsTestLots"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum > 0 Then
      Err.Clear
      sSql = "CREATE TABLE EsTestLots (" _
             & "LotNum char(15) NULL DEFAULT('')," _
             & "LotPart char (30) NULL DEFAULT('')," _
             & "LotOrig smallmoney NULL DEFAULT(0)," _
             & "LotRemains smallmoney NULL DEFAULT(0)," _
             & "LotActivity smallmoney NULL DEFAULT(0))"
      clsADOCon.ExecuteSQL sSql
      
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "CREATE UNIQUE CLUSTERED INDEX LotRef ON EsTestLots ([LotNum]) WITH  FILLFACTOR = 80"
         clsADOCon.ExecuteSQL sSql
      End If
   End If
   
End Sub

Private Sub FillTable()
   Dim RdoLots As ADODB.Recordset
   Dim rdoAct As ADODB.Recordset
   Dim iProg As Integer
   Dim iBar As Integer
   Dim lCOUNTER As Long
   Dim lAllRows As Long
   Dim lDeleted As Long
   Dim nAdder As Currency
   Dim nCount As Currency
   Dim sPartSearch As String
   
   sSql = "truncate table EsTestLots"
   clsADOCon.ExecuteSQL sSql
   fraPrn.Enabled = False
   cmdSel.Enabled = False
   If lLotCount = 0 Then lLotCount = 1
   nAdder = 95 / lLotCount
   MouseCursor 13
   If txtPrt <> "ALL" Then sPartSearch = Compress(txtPrt)
   prg1.Value = 0
   prg1.Visible = True
   z1(10).Caption = "Getting Lots"
   z1(10).Refresh
   If optShow(2) Then
      'All
      sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTORIGINALQTY,LOTREMAININGQTY " _
             & "FROM LohdTable WHERE LOTPARTREF LIKE '" & sPartSearch & "%' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   ElseIf optShow(1) Then
      'Without
      sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTORIGINALQTY,LOTREMAININGQTY," _
             & "PARTREF,PALOTTRACK FROM LohdTable,PartTable WHERE " _
             & "(PARTREF=LOTPARTREF AND LOTPARTREF LIKE '" & sPartSearch & "%' " _
             & "AND PALOTTRACK=0) "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   Else
      'With
      sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTORIGINALQTY,LOTREMAININGQTY," _
             & "PARTREF,PALOTTRACK FROM LohdTable,PartTable WHERE " _
             & "(PARTREF=LOTPARTREF AND LOTPARTREF LIKE '" & sPartSearch & "%' " _
             & "AND PALOTTRACK=1) "
   End If
   sSql = sSql & "ORDER BY LOTNUMBER"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   If bSqlRows Then
      With RdoLots
         On Error Resume Next
         prg1.Value = 5
         iBar = 5
         Do Until .EOF
            lCOUNTER = lCOUNTER + 1
            nCount = nCount + nAdder
            If nCount > 5 Then
               iBar = iBar + 5
               If iBar > 95 Then iBar = 95
               prg1.Value = iBar
               nCount = 0
            End If
            iProg = iProg + 1
            If iProg > 100 Then
               lblRows(3) = lCOUNTER
               iProg = 0
               lblRows(3).Refresh
            End If
            sSql = "INSERT INTO EsTestLots (LotNum,LotPart,LotOrig,LotRemains) " _
                   & "values('" & Trim(!lotNumber) & "','" & Trim(!LotPartRef) & "'," _
                   & !LOTORIGINALQTY & "," & !LOTREMAININGQTY & ")"
            clsADOCon.ExecuteSQL sSql
            If Err > 0 Then Err.Clear
            .MoveNext
         Loop
         ClearResultSet RdoLots
      End With
      Set RdoLots = Nothing
      prg1.Value = 100
      Sleep 2000
      iProg = 0
      prg1.Value = 1
      iBar = 1
      lAllRows = lCOUNTER
      If lCOUNTER = 0 Then lCOUNTER = 1
      nAdder = 95 / lCOUNTER
      lCOUNTER = 0
      z1(10).Caption = "Adding Item Detail"
      z1(10).Refresh
      sSql = "SELECT LotNum FROM EsTestLots"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
      If bSqlRows Then
         With RdoLots
            Do Until .EOF
               lCOUNTER = lCOUNTER + 1
               nCount = nCount + nAdder
               If nCount > 5 Then
                  iBar = iBar + 5
                  If iBar > 95 Then iBar = 95
                  prg1.Value = iBar
                  nCount = 0
               End If
               iProg = iProg + 1
               If iProg > 100 Then
                  lblRows(3) = lCOUNTER
                  iProg = 0
                  lblRows(3).Refresh
               End If
               sSql = "SELECT SUM(LOIQUANTITY) FROM LoitTable " _
                      & "WHERE LOINUMBER='" & !lotnum & "'"
               bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
               If bSqlRows Then
                  sSql = "UPDATE EsTestLots SET LotActivity=" & rdoAct.Fields(0) & " " _
                         & "WHERE LotNum='" & !lotnum & "'"
                  clsADOCon.ExecuteSQL sSql
               End If
               Set rdoAct = Nothing
               .MoveNext
            Loop
            ClearResultSet RdoLots
         End With
      End If
      prg1.Value = 95
      lAllRows = lAllRows + lCOUNTER
      z1(10).Caption = "Deleting Good Rows"
      z1(10).Refresh
      sSql = "SELECT COUNT(LotNum) FROM EsTestLots WHERE LotRemains=LotActivity"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoLots.Fields(0)) Then lDeleted = RdoLots.Fields(0)
      End If
      sSql = "DELETE FROM EsTestLots WHERE LotRemains=LotActivity"
      clsADOCon.ExecuteSQL sSql
      prg1.Value = 100
      Sleep 3000
   End If
   MouseCursor 0
   MsgBox lDeleted & " Rows Were Found And Removed Where The Tables Matched.", _
      vbInformation, Caption
   Set RdoLots = Nothing
   z1(10).Caption = "Total Row Count"
   lblRows(3) = lAllRows
   lblRows(3).Refresh
   z1(10).Refresh
   MsgBox "The Report Table Is Prepared.", vbInformation, Caption
   prg1.Visible = False
   fraPrn.Enabled = True
   cmdSel.Enabled = True
   
End Sub
