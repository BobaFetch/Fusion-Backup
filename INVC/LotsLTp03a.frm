VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LotsLTp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Uncosted Lots"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   8010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2400
      TabIndex        =   20
      Top             =   960
      Width           =   3075
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Leading Char Search  (*  In Front Is A Legal Wild Card)"
      Top             =   960
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "LotsLTp03a.frx":07AE
      Height          =   315
      Left            =   5520
      Picture         =   "LotsLTp03a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   960
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   2160
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6840
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "LotsLTp03a.frx":0E32
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
         Picture         =   "LotsLTp03a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7320
      Top             =   2760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   8010
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   2400
      TabIndex        =   19
      Top             =   2520
      Width           =   4692
      _ExtentX        =   8281
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Greater Or Equal"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Actual Transaction Date)"
      Height          =   288
      Index           =   1
      Left            =   6000
      TabIndex        =   16
      Top             =   1440
      Width           =   1932
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   15
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Lots From"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "( Blank For ALL)"
      Height          =   288
      Index           =   6
      Left            =   6000
      TabIndex        =   11
      Top             =   960
      Width           =   1908
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Detail"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   1785
   End
End
Attribute VB_Name = "LotsLTp03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/1/05 Changed date handling
'5/16/05 corrected group show/hide
'9/15/05 Added Inventory Transfer to report table (32)
Option Explicit
Dim bOnLoad As Byte

Dim iProg As Integer
Dim iTotalLots As Integer
Dim sLots(1000) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   optVew.Value = vbChecked
   ViewParts.Show
   
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
   MouseCursor 0
   FillCombo
   cmbPrt = "ALL"
   
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
   Set LotsLTp01a = Nothing
   
End Sub

Private Sub FillCombo()
   sSql = "Qry_FillSortedParts"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub PrintReport()
   Dim sBook As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   
   On Error GoTo DiaErr1
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDet"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbPrt _
                        & "... From " & txtBeg & " Through  " & txtEnd) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add optDet.Value
   sCustomReport = GetCustomReport("invlt03")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
'   cCRViewer.SetReportSelectionFormula (sSql)
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
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 3)
   txtPrt = "ALL"
   txtQty = "0.000"
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(optDet.Value)
   SaveSetting "Esi2000", "EsiInvc", "lt03", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiInvc", "lt03", sOptions)
   If Len(sOptions) > 0 Then optDet.Value = Val(Mid(sOptions, 2, 1))
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrepareReport
   
End Sub

Private Sub optPrn_Click()
   PrepareReport
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDate(txtBeg)
   
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDate(txtEnd)
   
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Trim(txtPrt) = "" Then txtPrt = "ALL"
   
End Sub
Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Trim(cmbPrt) = "" Then cmbPrt = "ALL"
   
End Sub

Private Sub FillReportTableh()
   Dim RdoLot As ADODB.Recordset
   Dim RdoRpt As ADODB.Recordset
   
   Dim b As Byte
   Dim sLotType As String
   Dim sPartNo As String
   Dim sPartDs As String
   Dim sVnick As String
   Dim sVname As String
   
   Dim sBegDate As String
   Dim sEndDate As String
   
   If txtBeg = "ALL" Then
      sBegDate = "01/01/1995"
   Else
      sBegDate = Format(txtBeg, "mm/dd/yy")
   End If
   If txtEnd = "ALL" Then
      sEndDate = "12/31/24"
   Else
      sEndDate = Format(txtEnd, "mm/dd/yy")
   End If
   
   Erase sLots
   iTotalLots = 0
   sProcName = "FillReporth"
   sSql = "truncate table EsReportLots01h"
   clsADOCon.ExecuteSQL sSql
   sSql = "truncate table EsReportLots01d"
   clsADOCon.ExecuteSQL sSql
   
   If cmbPrt = "ALL" Then sPartNo = "" Else sPartNo = cmbPrt
   sSql = "SELECT *,LOINUMBER,LOITYPE FROM LohdTable,LoitTable " _
          & "WHERE (LOTNUMBER=LOINUMBER AND LOIRECORD=1) AND (" _
          & "LOIPARTREF LIKE '" & sPartNo & "%' AND " _
          & "LOIADATE BETWEEN '" & sBegDate & " 00:00' AND '" _
          & sEndDate & " 23:59' AND LOTREMAININGQTY >=" & Val(txtQty) _
          & " AND LOTDATECOSTED IS NULL)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         b = 1
         sSql = "SELECT * FROM EsReportLots01h WHERE " _
                & "LotSysNumber=''"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_KEYSET)
         Do Until .EOF
            iProg = iProg + 5
            If iProg > 40 Then iProg = 40
            prg1.Value = iProg
            iTotalLots = iTotalLots + 1
            sLots(iTotalLots) = "" & Trim(!lotNumber)
            RdoRpt.AddNew
            Select Case !LOITYPE
               Case 15
                  sLotType = "Purchase Order Receipt"
               Case 6
                  sLotType = "MO Receipt"
               Case 19
                  sLotType = "Manual Inventory Adjustment"
               Case Else
                  sLotType = "Misc Incoming"
            End Select
            sPartNo = "" & Trim(!LotPartRef)
            sPartNo = GetThisPartNumber(sPartNo, sPartDs)
            RdoRpt!LotSysNumber = "" & Trim(!lotNumber)
            RdoRpt!LotUsrNumber = "" & Trim(!LOTUSERLOTID)
            RdoRpt!LotType = sLotType
            RdoRpt!LotPartNumber = sPartNo
            RdoRpt!LotPartDesc = sPartDs
            RdoRpt!LotComment = "" & Trim(!LOTCOMMENTS)
            RdoRpt!LotPlannedDate = Format(!LOTPDATE, "mm/dd/yy")
            RdoRpt!LotSysDate = Format(!LotADate, "mm/dd/yy")
            RdoRpt!LotBeginQty = Format(!LOTORIGINALQTY, ES_QuantityDataFormat)
            RdoRpt!LOTRemainQty = Format(!LOTREMAININGQTY, ES_QuantityDataFormat)
            RdoRpt!LotUnitCost = Format(!LotUnitCost, ES_QuantityDataFormat)
            RdoRpt!LOTLOCATION = "" & Trim(!LOTLOCATION)
            'PO?
            If !LOTPO > 0 Then
               sVnick = GetPOInformation(!LOTPO)
               sVnick = GetThisVendor(sVnick, sVname)
               RdoRpt!LotVendorNick = sVnick
               RdoRpt!LotVendorName = sVname
               RdoRpt!LotPONum = "PO " & Format(!LOTPO, "00000") _
                                 & "-" & Format$(!LOTPOITEM, "###0") & Trim(!LOTPOITEM)
            Else
               'MO?
               If !LOTMORUNNO > 0 Then
                  sPartNo = "" & Trim(!LOTMOPARTREF)
                  sPartNo = GetThisPartNumber(sPartNo, sPartDs)
                  RdoRpt!LotMONumber = "MO " & Trim(sPartNo) & " Run " _
                                       & Format$(!LOTMORUNNO, "00000")
                  RdoRpt!LotMOPartDesc = "" & sPartDs
               End If
            End If
            RdoRpt.Update
            sProcName = "FillReporth"
            .MoveNext
         Loop
         ClearResultSet RdoLot
      End With
   End If
   iProg = 40
   prg1.Value = iProg
   On Error Resume Next
   'RdoRpt.Close
   Set RdoLot = Nothing
   Set RdoRpt = Nothing
   If b = 1 Then FillReportTabled
   
End Sub

Private Function GetThisPartNumber(sPartNum, sDesc) As String
   Dim RdoPrt As ADODB.Recordset
   sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable " _
          & "WHERE PARTREF='" & sPartNum & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         GetThisPartNumber = "" & Trim(!PartNum)
         sDesc = "" & Trim(!PADESC)
         ClearResultSet RdoPrt
      End With
   End If
   Set RdoPrt = Nothing
End Function


Private Function GetPOInformation(lPO As Long) As String
   Dim RdoPoi As ADODB.Recordset
   sSql = "SELECT PONUMBER,POVENDOR FROM PohdTable " _
          & "WHERE PONUMBER=" & lPO & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPoi, ES_FORWARD)
   If bSqlRows Then
      With RdoPoi
         GetPOInformation = "" & Trim(!POVENDOR)
         ClearResultSet RdoPoi
      End With
   End If
   Set RdoPoi = Nothing
   
End Function

Private Sub FillReportTabled()
   Dim RdoLot As ADODB.Recordset
   Dim RdoRpt As ADODB.Recordset
   Dim iList As Integer
   Dim sPackSlip As String
   Dim sActivity As String
   
   sProcName = "FillReportd"
   sSql = "SELECT * FROM EsReportLots01d WHERE " _
          & "LoiSysNumber=''"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_KEYSET)
   For iList = 1 To iTotalLots
      sSql = "SELECT * FROM LoitTable WHERE " _
             & "LOINUMBER='" & sLots(iList) & "' " _
             & "ORDER BY LOIRECORD"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
      If bSqlRows Then
         With RdoLot
            iProg = iProg + 5
            If iProg > 95 Then iProg = 95
            prg1.Value = iProg
            Do Until .EOF
               Select Case !LOITYPE
                  Case 1
                     sActivity = "Begining Balance"
                  Case 3
                     sActivity = "Shipped Item"
                  Case 4
                     sActivity = "Returned Item"
                  Case 5
                     sActivity = "Canceled SO Item"
                  Case 6
                     sActivity = "Completed MO"
                  Case 7
                     sActivity = "Closed MO"
                  Case 9
                     sActivity = "Pick Request"
                  Case 10
                     sActivity = "Actual Pick"
                  Case 11
                     sActivity = "Pick On Dock"
                  Case 12
                     sActivity = "Canceled Pick Req"
                  Case 13
                     sActivity = "Pick Surplus"
                  Case 14
                     sActivity = "Open PO Item"
                  Case 15
                     sActivity = "PO Receipt"
                  Case 16
                     sActivity = "Canceled PO Item"
                  Case 17
                     sActivity = "Invoiced PO Item"
                  Case 18
                     sActivity = "On Dock"
                  Case 19
                     sActivity = "Manual Adjustment"
                  Case 21
                     sActivity = "Restocked Pick"
                  Case 22
                     sActivity = "Scrapped Pick Item"
                  Case 23
                     sActivity = "Pick Substitute"
                  Case 21
                     sActivity = "Pick From Freight"
                  Case 32
                     sActivity = "Inventory Transfer"
                  Case 38
                     sActivity = "Canceled MO Completion"
                  Case Else
                     sActivity = "Undocumented Activity"
               End Select
               
               If Trim(!LOIPSNUMBER) <> "" Then
                  sPackSlip = "" & Trim(!LOIPSNUMBER) _
                              & "-" & Format(!LOIPSITEM, "##0")
               Else
                  sPackSlip = ""
               End If
               RdoRpt.AddNew
               RdoRpt!LoiSysNumber = "" & Trim(!LOINUMBER)
               RdoRpt!LoiSysRecord = !LOIRECORD
               RdoRpt!LoiSysQuantity = !LOIQUANTITY
               RdoRpt!LoiPlannedDate = Format(!LOIPDATE, "mm/dd/yy")
               RdoRpt!LoiSysDate = Format(!LOIADATE, "mm/dd/yy")
               RdoRpt!LoiComments = "" & Trim(!LOICOMMENT) & " " & sPackSlip
               RdoRpt!LoiActivity = sActivity
               RdoRpt.Update
               .MoveNext
               sProcName = "FillReportd"
            Loop
            ClearResultSet RdoLot
         End With
      End If
   Next
   On Error Resume Next
   prg1.Value = 100
   'RdoRpt.Close
   Set RdoLot = Nothing
   Set RdoRpt = Nothing
   
End Sub

Private Sub PrepareReport()
   On Error GoTo DiaErr1
   MouseCursor 13
   prg1.Visible = True
   iProg = 10
   prg1.Value = iProg
   sProcName = "Preparereport"
   FillReportTableh
   MouseCursor 0
   PrintReport
   prg1.Visible = False
   Exit Sub
   
DiaErr1:
   prg1.Visible = False
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   
End Sub
