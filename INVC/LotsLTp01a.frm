VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form LotsLTp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lots By Part Number"
   ClientHeight    =   4185
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   8325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4185
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdShowDocLokImage 
      Height          =   375
      Left            =   7080
      Picture         =   "LotsLTp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Retrieve Document Imaging System Image"
      Top             =   1800
      Width           =   375
   End
   Begin VB.CheckBox chkIncludeLotComments 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   9
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTp01a.frx":F172
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Tag             =   "3"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Get Matching User Lots In The Date Range"
      Top             =   1800
      Width           =   1035
   End
   Begin VB.ComboBox cmbLot 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Lot From List Or Blank For All"
      Top             =   1800
      Width           =   3840
   End
   Begin VB.OptionButton optUsr 
      Caption         =   "User Lot Number"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   2640
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton optSys 
      Caption         =   "System Lot Number"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
   End
   Begin VB.ComboBox txtEnd 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Leading Char Search  (*  In Front Is A Legal Wild Card)"
      Top             =   960
      Width           =   3075
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   315
      Left            =   4920
      Picture         =   "LotsLTp01a.frx":F920
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   960
      Width           =   350
   End
   Begin VB.CheckBox chkDetails 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   3000
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   7080
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7080
      TabIndex        =   10
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "LotsLTp01a.frx":FC62
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "LotsLTp01a.frx":FDE0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7260
      Top             =   3480
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4185
      FormDesignWidth =   8325
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1800
      TabIndex        =   27
      Top             =   3720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label d 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Lot Comments"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   3360
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "( Blank For ALL)"
      Height          =   285
      Index           =   8
      Left            =   5760
      TabIndex        =   25
      Top             =   2280
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Number (User)"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   22
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Uses Actual Transaction Date)"
      Height          =   288
      Index           =   1
      Left            =   5760
      TabIndex        =   21
      Top             =   1320
      Width           =   2412
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   19
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Lots From"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "( Blank For ALL)"
      Height          =   285
      Index           =   6
      Left            =   5760
      TabIndex        =   15
      Top             =   960
      Width           =   2145
   End
   Begin VB.Label d 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Detail"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   1785
   End
End
Attribute VB_Name = "LotsLTp01a"
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
'3/10/05 Added cmbLot and code
'3/11/05 Added sLots to PrintReport to not show "ALL" in cmblot
'7/7/05 Corrected date problems in fill and prepare and fading optPrn/optDis
'8/5/05 Added Location
'9/15/05 Added Inventory Transfer to report table (32)
'1/24/06 Correct Query (PrintReport)
'3/7/07 Corrected Lot Number Query (PrintReport) 7.2.2
'3/23/07 Expanded Activity notes per JLH (FillReportTabled) 7.2.6
Option Explicit
Dim bOnLoad As Byte

Dim iProg As Integer
Dim iTotalLots As Integer
Dim sLots(2000) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd




Private Sub cmbLot_LostFocus()
   cmbLot = CheckLen(cmbLot, 50)
   If Trim(cmbLot) = "" Then cmbLot = "ALL"
   If cmbLot = "ALL" Then cmdShowDocLokImage.Enabled = False
   If cmbLot <> "ALL" And Len(cmbLot) > 0 Then cmdShowDocLokImage.Enabled = True
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   cmbLot.Clear
   cmdShowDocLokImage.Enabled = False
   z1(5).Enabled = True
   cmbLot.Enabled = True
   cmdSel.Enabled = True
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



Private Sub cmdSel_Click()
   GetLots
   
End Sub

Private Sub cmdShowDocLokImage_Click()
    Dim AdoSysLot As ADODB.Recordset
    Dim sSysLot As String
    
    On Error Resume Next
    
    sSql = "SELECT LOTNUMBER FROM LohdTable WHERE LOTUSERLOTID='" & cmbLot & "' "
    bSqlRows = clsADOCon.GetDataSet(sSql, AdoSysLot, ES_FORWARD)
    If bSqlRows Then
        sSysLot = "" & AdoSysLot!lotNumber
        If Len(sSysLot) > 0 Then
             SendFindToDockLok (sSysLot)
        End If
    End If
    Set AdoSysLot = Nothing
End Sub

Private Sub SendFindToDockLok(sSystemLot As String)
    Dim doclok As DocLokIntegrator
    Dim bSentOk As Boolean
  
    Set doclok = New DocLokIntegrator
    
    If doclok.Installed Then
        doclok.OpenXMLFile "Retrieve", "Material Purchase Orders"
        ' MM TODO
        'doclok.AddXMLIndex "Document Type", "Material Certification"
        doclok.AddXMLIndex "System Lot Number", sSystemLot
        
        doclok.CloseXMLFile
        bSentOk = doclok.SendXMLFileToDocLok
        
    Else
        MsgBox "Document Imaging System is not installed"
    End If
    Set doclok = Nothing
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   MouseCursor 0
   If DocumentLokEnabled Then cmdShowDocLokImage.Visible = True Else cmdShowDocLokImage.Visible = False
   cmdShowDocLokImage.Enabled = False
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

Private Sub PrintReport()
   Dim sLots As String
   MouseCursor 13
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init

   On Error GoTo DiaErr1
   If Trim(txtPrt) = "ALL" Then cmbLot = "ALL"
   If cmbLot = "" Then cmbLot = "ALL"
   If Trim(cmbLot) <> "ALL" Then sLots = cmbLot
   
   'SetMdiReportsize MdiSect
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "Includes='" & txtPrt _
'                        & "... From " & txtBeg & " Through  " & txtEnd & "'"
'   MdiSect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
'   MdiSect.Crw.Formulas(3) = "ShowLotComments=" & Me.chkIncludeLotComments

   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowLotComments"
   aFormulaName.Add "ShowDetails"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(txtPrt _
                        & "... From " & txtBeg & " Through  " & txtEnd) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add Me.chkIncludeLotComments
   aFormulaValue.Add chkDetails.Value
   
   If optSys Then
      sCustomReport = GetCustomReport("invlt01a")
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
'      MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   Else
      sCustomReport = GetCustomReport("invlt01b")
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
'      MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   End If
   
'   If chkDetails.value = vbUnchecked Then
'      MdiSect.Crw.SectionFormat(0) = "GROUPHDR.1.0;F;;;"
'      MdiSect.Crw.SectionFormat(1) = "DETAIL.0.0;F;;;"
'   Else
'      MdiSect.Crw.SectionFormat(0) = "GROUPHDR.1.0;T;;;"
'      MdiSect.Crw.SectionFormat(1) = "DETAIL.1.0;T;;;"
'   End If
   
   sSql = "{EsReportLots01h.LotLocation} LIKE '" & Trim(txtLoc) & "*' "
   If sLots <> "" Then sLots = "AND {EsReportLots01h.LotUsrNumber} = '" & sLots & "'"
   cCRViewer.SetReportSelectionFormula sSql & sLots
   cCRViewer.SetDbTableConnection
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   'cCRViewer.CRViewerSize Me
   'cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

   'MdiSect.Crw.SelectionFormula = sSql & sLots
   'SetCrystalAction Me
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
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 5)
   txtPrt = "ALL"
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(Abs(optSys.Value)) & chkDetails.Value & Me.chkIncludeLotComments.Value
   SaveSetting "Esi2000", "EsiInvc", "lt01", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiInvc", "lt01", sOptions) & "0000"
   
   optSys.Value = Val(Mid(sOptions, 1, 1))      'optusr will be opposite
   chkDetails.Value = Val(Mid(sOptions, 2, 1))
   chkIncludeLotComments.Value = Val(Mid(sOptions, 3, 1))
End Sub

Private Sub chkDetails_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrepareReport
   
End Sub


Private Sub optPrn_Click()
   PrepareReport
   
End Sub




Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
   
End Sub


Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
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
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then
      txtPrt = "ALL"
      cmbLot.Clear
      cmdShowDocLokImage.Enabled = False
      z1(5).Enabled = False
      cmbLot.Enabled = False
      cmdSel.Enabled = False
   Else
      z1(5).Enabled = True
      cmbLot.Enabled = True
      cmdSel.Enabled = True
   End If
   
   
End Sub



Private Sub FillReportTableh()
   
   On Error GoTo DiaErr1
   
   Dim AdoLot As ADODB.Recordset
   Dim AdoRpt As ADODB.Recordset
      
   Dim b As Byte
   Dim bRows As Byte
   Dim sLotType As String
   Dim sPartNo As String
   Dim sPartDs As String
   Dim sVnick As String
   Dim sVname As String
   
   Dim sBegDate As String
   Dim sEndDate As String
   Dim unitCost As Currency
   
   Erase sLots
   iTotalLots = 0
   sProcName = "FillReporth"
   sSql = "truncate table EsReportLots01h"
   clsADOCon.ExecuteSql sSql
   sSql = "truncate table EsReportLots01d"
   clsADOCon.ExecuteSql sSql
   
   If txtBeg <> "ALL" Then
      sBegDate = Format(txtBeg, "mm/dd/yyyy 00:00")
   Else
      'sBegDate = "01/01/95 00:00"
      sBegDate = "01/01/1995"
   End If
   If txtEnd <> "ALL" Then
      sEndDate = Format(txtEnd, "mm/dd/yyyy 23:59")
   Else
      sEndDate = "12/31/2025 23:59"
   End If
   If txtPrt = "ALL" Then sPartNo = "" Else sPartNo = Compress(txtPrt)
   sSql = "SELECT * FROM LohdTable JOIN LoitTable ON LOTNUMBER=LOINUMBER" & vbCrLf _
      & "WHERE LOIRECORD=(SELECT MAX(a.LOIRECORD) FROM LoitTable a WHERE LOTNUMBER = a.LOINUMBER)" & vbCrLf _
      & "AND LOIPARTREF LIKE '" & sPartNo & "%'" & vbCrLf _
      & "AND LOIADATE BETWEEN '" & sBegDate & "' AND '" & sEndDate & "'"

'   sSql = "SELECT * FROM LohdTable JOIN LoitTable ON LOTNUMBER=LOINUMBER" & vbCrLf _
'      & "WHERE LOIRECORD=1" & vbCrLf _
'
'      '& "AND LOIPARTREF LIKE '" & sPartNo & "%'" & vbCrLf _
'
'      '& "AND LOIADATE BETWEEN '" & sBegDate & "' AND '" & sEndDate & "')"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoLot, ES_FORWARD)
   If bSqlRows Then
      With AdoLot
         bRows = 0
         b = 1
         sSql = "SELECT * FROM EsReportLots01h WHERE " _
                & "LotSysNumber=''"
         bSqlRows = clsADOCon.GetDataSet(sSql, AdoRpt, ES_KEYSET)
         Do Until .EOF
            iProg = iProg + 5
            If iProg > 40 Then iProg = 40
            prg1.Value = iProg
            iTotalLots = iTotalLots + 1
            If iTotalLots > 1999 Then
               bRows = 1
               Exit Do
            End If
            sLots(iTotalLots) = "" & Trim(!lotNumber)
            AdoRpt.AddNew
            Select Case !LOITYPE
               Case 15
                  sLotType = "Purchase Order Receipt"
               Case 6
                  sLotType = "MO Receipt"
               Case 19
                  sLotType = "Manual Inventory Adjustment"
               Case 40
                  sLotType = "Return Part to Vendor"
               Case Else
                  sLotType = "Misc Incoming"
            End Select
            sPartNo = "" & Trim(!LotPartRef)
            sPartNo = GetThisPartNumber(sPartNo, sPartDs)
            AdoRpt!LotSysNumber = "" & Trim(!lotNumber)
            AdoRpt!LotUsrNumber = "" & Trim(!LOTUSERLOTID)
            AdoRpt!LotType = sLotType
            AdoRpt!LotPartNumber = sPartNo
            AdoRpt!LotPartDesc = sPartDs
            AdoRpt!LotComment = "" & Trim(!LOTCOMMENTS)
            AdoRpt!LotPlannedDate = Format(!LOTPDATE, "mm/dd/yyyy")
            AdoRpt!LotSysDate = Format(!LotADate, "mm/dd/yyyy")
            AdoRpt!LotBeginQty = Format(!LOTORIGINALQTY, ES_QuantityDataFormat)
            AdoRpt!LOTRemainQty = Format(!LOTREMAININGQTY, ES_QuantityDataFormat)
            unitCost = IIf(IsNull(!LotUnitCost), 0, (!LotUnitCost))
            AdoRpt!LotUnitCost = Format(unitCost, ES_QuantityDataFormat)
            AdoRpt!LOTLOCATION = "" & Trim(!LOTLOCATION)
            
            If Not IsNull(!LOTEXPIRESON) Then
                AdoRpt!LotExpDate = Format(!LOTEXPIRESON, "mm/dd/yy")
            End If
            
            'PO?
            If !LOTPO > 0 Then
               sVnick = GetPOInformation(!LOTPO)
               sVnick = GetThisVendor(sVnick, sVname)
               AdoRpt!LotVendorNick = sVnick
               AdoRpt!LotVendorName = sVname
               AdoRpt!LotPONum = "PO " & Format(!LOTPO, "00000") _
                                 & "-" & Format$(!LOTPOITEM, "###0") & Trim(!LOTPOITEM)
            Else
               'MO?
               If !LOTMORUNNO > 0 Then
                  sPartNo = "" & Trim(!LOTMOPARTREF)
                  sPartNo = GetThisPartNumber(sPartNo, sPartDs)
                  AdoRpt!LotMONumber = "MO " & Trim(sPartNo) & " Run " _
                                       & Format$(!LOTMORUNNO, "00000")
                  AdoRpt!LotMOPartDesc = "" & sPartDs
               End If
            End If
            AdoRpt.Update
            sProcName = "FillReporth"
            .MoveNext
         Loop
         ClearResultSet AdoLot
      End With
   End If
   If bRows = 0 Then
      iProg = 40
      prg1.Value = iProg
      On Error Resume Next
      AdoRpt.Close
      Set AdoLot = Nothing
      Set AdoRpt = Nothing
      If b = 1 Then FillReportTabled
   Else
      MsgBox "Too Many Rows Of Data.  Please Tighten Your" & vbCr _
         & "Report Date Range And Try Again.", _
         vbInformation, Caption
      Erase sLots
      prg1.Value = 0
      prg1.Visible = False
      On Error Resume Next
      AdoRpt.Close
      Set AdoLot = Nothing
      Set AdoRpt = Nothing
   End If
   
   Exit Sub


DiaErr1:
   prg1.Visible = False
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetThisPartNumber(sPartNum, sDesc) As String
   Dim AdoPrt As ADODB.Recordset
   sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable " _
          & "WHERE PARTREF='" & sPartNum & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoPrt, ES_FORWARD)
   If bSqlRows Then
      With AdoPrt
         GetThisPartNumber = "" & Trim(!PartNum)
         sDesc = "" & Trim(!PADESC)
         ClearResultSet AdoPrt
      End With
   End If
   Set AdoPrt = Nothing
End Function


Private Function GetPOInformation(lPO As Long) As String
   Dim AdoPoi As ADODB.Recordset
   sSql = "SELECT PONUMBER,POVENDOR FROM PohdTable " _
          & "WHERE PONUMBER=" & lPO & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoPoi, ES_FORWARD)
   If bSqlRows Then
      With AdoPoi
         GetPOInformation = "" & Trim(!POVENDOR)
         ClearResultSet AdoPoi
      End With
   End If
   Set AdoPoi = Nothing
   
End Function

Private Sub FillReportTabled()
   Dim AdoLot As ADODB.Recordset
   Dim AdoRpt As ADODB.Recordset
   Dim iList As Integer
   Dim sPackSlip As String
   Dim sActivity As String
   Dim sString As String
   Dim sSONumber As String
   
   
   sProcName = "FillReportd"
   sSql = "SELECT * FROM EsReportLots01d WHERE " _
          & "LoiSysNumber=''"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoRpt, ES_KEYSET)
   For iList = 1 To iTotalLots
      sSql = "SELECT * FROM LoitTable WHERE LOINUMBER='" & sLots(iList) _
             & "' ORDER BY LOIRECORD"
      bSqlRows = clsADOCon.GetDataSet(sSql, AdoLot, ES_FORWARD)
      If bSqlRows Then
         With AdoLot
            iProg = iProg + 5
            If iProg > 95 Then iProg = 95
            prg1.Value = iProg
            Do Until .EOF
               Select Case !LOITYPE
                  Case 1
                     sActivity = "Begining Balance"
                  Case 3
                     sActivity = "Shipped Item "
                  Case 4
                     sActivity = "Returned Item"
                  Case 5
                     sActivity = "Canceled SO Item"
                  Case 6
                     sActivity = "Completed MO " & Trim(!LOIMOPARTREF) & " Run" & str$(!LOIMORUNNO)
                  Case 7
                     sActivity = "Closed MO " & Trim(!LOIMOPARTREF) & " Run" & str$(!LOIMORUNNO)
                  Case 9
                     sActivity = "Pick Request " & Trim(!LOIMOPARTREF) & " Run" & str$(!LOIMORUNNO)
                  Case 10
                     sActivity = "Actual Pick " & Trim(!LOIMOPARTREF) & " Run" & str$(!LOIMORUNNO)
                  Case 11
                     sActivity = "Canceled All MO Picks " & Trim(!LOIMOPARTREF) & " Run" & str$(!LOIMORUNNO)
                  Case 12
                     sActivity = "Canceled Pick Item " & Trim(!LOIMOPARTREF) & " Run" & str$(!LOIMORUNNO)
                  Case 13
                     sActivity = "Pick Surplus " & Trim(!LOIMOPARTREF) & " Run" & str$(!LOIMORUNNO)
                  Case 14
                     sActivity = "Open PO Item PO" _
                                 & Format$(!LOIPONUMBER, "000000") & "-" & !LOIPOITEM & !LOIPOREV
                  Case 15
                     sActivity = "PO Receipt PO" _
                                 & Format$(!LOIPONUMBER, "000000") & "-" & !LOIPOITEM & !LOIPOREV
                  Case 16
                     sActivity = "Canceled PO Item PO" _
                                 & Format$(!LOIPONUMBER, "000000") & "-" & !LOIPOITEM & !LOIPOREV
                  Case 17
                     sActivity = "Invoiced PO Item PO" _
                                 & Format$(!LOIPONUMBER, "000000") & "-" & !LOIPOITEM & !LOIPOREV
                  Case 18
                     sActivity = "On Dock "
                  Case 19
                     sActivity = "Manual Adjustment "
                  Case 21
                     sActivity = "Restocked Pick " & Trim(!LOIMOPARTREF) & " Run" & str$(!LOIMORUNNO)
                  Case 22
                     sActivity = "Scrapped Pick Item " & Trim(!LOIMOPARTREF) & " Run" & str$(!LOIMORUNNO)
                  Case 23
                     sActivity = "Pick Substitute " & Trim(!LOIMOPARTREF) & " Run" & str$(!LOIMORUNNO)
                  Case 21
                     sActivity = "Pick From Freight "
                  Case 25
                     sSONumber = GetSONumber(!LOIPSNUMBER)
                     sActivity = "Shipment " & "SO " & Trim(sSONumber) & "-" & Trim(!LOIPSNUMBER) & "-" & str$(!LOIPSITEM)
                     'sActivity = "Shipment " & Trim(!LOIPSNUMBER) & "-" & str$(!LOIPSITEM)
                  Case 30
                     sActivity = "Cycle Count Reconciliation"
                  Case 32
                     sActivity = "Inventory Transfer "
                  Case 33
                     sActivity = "Cancel A Packing Slip " & Trim(!LOIPSNUMBER) & "-" & Trim$(str$(!LOIPSITEM))
                  Case 38
                     sActivity = "Canceled MO Completion"
                  Case 40
                     sActivity = "Return Part to Vendor"
                  Case Else
                     sActivity = "Undocumented Activity"
               End Select
               
               sPackSlip = ""
               If Trim(!LOIPSNUMBER) <> "" Then sPackSlip = " " & Trim(!LOIPSNUMBER) & "-" & Format(!LOIPSITEM, "##0")
               
               AdoRpt.AddNew
               AdoRpt!LoiSysNumber = "" & Trim(!LOINUMBER)
               AdoRpt!LoiSysRecord = !LOIRECORD
               AdoRpt!LoiSysQuantity = !LOIQUANTITY
               AdoRpt!LoiPlannedDate = Format(!LOIPDATE, "mm/dd/yyyy")
               AdoRpt!LoiSysDate = Format(!LOIADATE, "mm/dd/yyyy")
             
               Select Case !LOITYPE
               Case 19
'                    sString = GetInventoryComment(Trim(str("" & !LoiActivity)))
                    sString = GetInventoryComment(Trim("" & !LoiActivity))
                    If Len(sString) = 0 Then sString = "" & Trim(!LOICOMMENT)
                    'If Len(Trim("" & !INREF2)) = 0 Then sString = "" & Trim(!LOICOMMENT) Else sString = "" & !INREF2
               Case Else
                    sString = "" & Trim(!LOICOMMENT)
               End Select
               
               If Len(sString) > 40 Then
                  sString = Left(sString, 40)
               End If
               AdoRpt!LoiComments = sString
               AdoRpt!LoiActivity = sActivity
               AdoRpt.Update
               .MoveNext
               sProcName = "FillReportd"
            Loop
            ClearResultSet AdoLot
         End With
      End If
   Next
   On Error Resume Next
   prg1.Value = 100
   AdoRpt.Close
   Set AdoLot = Nothing
   Set AdoRpt = Nothing
   
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


Private Sub GetLots()
   Dim sBeg As String
   Dim sEnd As String
   cmbLot.Clear
   cmdShowDocLokImage.Enabled = False
   
   If txtBeg = "ALL" Then sBeg = "01/01/95" Else sBeg = txtBeg
   If txtEnd = "ALL" Then sEnd = "12/31/25" Else sEnd = txtEnd
   
   sSql = "SELECT LOTNUMBER,LOTUSERLOTID FROM LohdTable WHERE " _
          & "(LOTPARTREF='" & Compress(txtPrt) & "' AND " _
          & "LOTPDATE BETWEEN '" & sBeg & " 00:00' AND '" _
          & sEnd & " 23:59') ORDER BY LOTNUMBER"
   cmbLot.AddItem "ALL"
   LoadComboBox cmbLot
   cmbLot = cmbLot.List(0)
   
   On Error Resume Next
   If cmbLot.ListCount > 0 Then
    cmbLot.SetFocus
    cmdShowDocLokImage.Enabled = True
   End If
   
End Sub


Private Function DocumentLokEnabled() As Boolean
    Dim AdoDocLok As ADODB.Recordset
    Dim iDL As Integer
    
    DocumentLokEnabled = False
    sSql = "SELECT CODOCLOK FROM ComnTable WHERE COREF=1"
    bSqlRows = clsADOCon.GetDataSet(sSql, AdoDocLok, ES_FORWARD)
    If bSqlRows Then
        iDL = 0 & AdoDocLok!CODOCLOK
        If iDL = 1 Then DocumentLokEnabled = True
    End If
    Set AdoDocLok = Nothing
End Function

Private Function GetSONumber(ByVal strPSNum As String) As String
   Dim RdoPrt As ADODB.Recordset
    If Len(strPSNum) Then
         sSql = "select Distinct(PISONUMBER) from psitTable where pipackslip = '" & strPSNum & "'"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
     If bSqlRows Then
          With RdoPrt
             GetSONumber = "" & Trim(!PISONUMBER)
             ClearResultSet RdoPrt
          End With
     End If
    Else
        GetSONumber = ""
   End If
End Function
