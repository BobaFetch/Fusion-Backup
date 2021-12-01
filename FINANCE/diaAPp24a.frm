VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPp24a 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View A Cash Disbursement"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   Begin VB.ComboBox cboChkNum 
      Height          =   315
      Left            =   1860
      TabIndex        =   20
      Tag             =   "4"
      Top             =   2520
      Width           =   1395
   End
   Begin VB.ComboBox cboEnd 
      Height          =   315
      Left            =   1860
      TabIndex        =   3
      Tag             =   "4"
      Top             =   2040
      Width           =   1395
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   600
      Width           =   1335
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "diaAPp24a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "diaAPp24a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Tag             =   "1"
      Top             =   1200
      Width           =   1395
   End
   Begin VB.ComboBox cboVendor 
      Height          =   315
      Left            =   1860
      TabIndex        =   0
      Tag             =   "3"
      Top             =   345
      Width           =   1555
   End
   Begin VB.ComboBox cboStart 
      Height          =   315
      Left            =   1860
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1620
      Width           =   1395
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   14
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
      PictureUp       =   "diaAPp24a.frx":0308
      PictureDn       =   "diaAPp24a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   15
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
      PictureUp       =   "diaAPp24a.frx":0594
      PictureDn       =   "diaAPp24a.frx":06DA
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5940
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3090
      FormDesignWidth =   6855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
      Height          =   165
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   2550
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   2100
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank If Unknown)"
      Height          =   255
      Index           =   4
      Left            =   3420
      TabIndex        =   17
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1860
      TabIndex        =   13
      Top             =   780
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank If Unknown)"
      Height          =   255
      Index           =   0
      Left            =   3420
      TabIndex        =   12
      Top             =   1620
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disbursement Amount"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1260
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Nickname"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank If Unknown)"
      Height          =   255
      Index           =   6
      Left            =   3420
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "diaAPp24a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
'
' diaAPp24a - View a Cash Disbursement
'
' Created: 1/26/06 (TEL)
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim sMsg As String
Public bRemote As Byte

' Key Handeling
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboEnd_Change()
   ' Get all Check number for that Vendor
   GetCheckNumber

End Sub

Private Sub cboStart_Change()
   ' Get all Check number for that Vendor
   GetCheckNumber
End Sub

Private Sub cboVendor_Change()
   ' Get all Check number for that Vendor
   GetCheckNumber
End Sub

'*************************************************************************************

Private Sub cboVendor_Click()
   'FindVendor Me
   lblName = GetVendorName(cboVendor)
End Sub

Private Sub cboVendor_GotFocus()
   ComboGotFocus cboVendor
End Sub

Private Sub cboVendor_KeyUp(KeyCode As Integer, Shift As Integer)
   ComboKeyUp cboVendor, KeyCode
End Sub

Private Sub cboVendor_LostFocus()
   '    FindVendor Me
   lblName = GetVendorName(cboVendor)
   ' Get all Check number for that Vendor
   GetCheckNumber
   
End Sub


Private Sub Form_Load()
   '    GetOptions
   If bRemote Then
      Me.WindowState = vbMinimized
   Else
      FormLoad Me
      FormatControls
      bOnLoad = True
   End If
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      GetOptions
      bOnLoad = False
   End If
End Sub

Public Sub FillCombo()
   On Error GoTo DiaErr1
   
   'FillVendors Me
   LoadComboWithVendors cboVendor
   'FindVendor Me
   lblName = GetVendorName(cboVendor)
   ' Get all Check number for that Vendor
   GetCheckNumber
   
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   SaveOptions
   If Not bRemote Then
      FormUnload
   End If
   bRemote = False
   Set diaARp09a = Nothing
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub optDis_Click()
   If Not bRemote Then
      PrintReport
   End If
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub txtAmt_Change()
   ' Get all Check number for that Vendor
   GetCheckNumber

End Sub

Private Sub txtAmt_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtAmt_LostFocus()
   txtAmt = Format(txtAmt, CURRENCYMASK)
End Sub

Private Sub cboStart_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboStart_GotFocus()
   SelectFormat Me
End Sub

Private Sub cboStart_LostFocus()
   If Trim(cboStart) <> "" Then
      cboStart = CheckDate(cboStart)
   End If
End Sub

Private Sub cboEnd_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboEnd_GotFocus()
   SelectFormat Me
End Sub

Private Sub cboEnd_LostFocus()
   If Trim(cboEnd) <> "" Then
      cboEnd = CheckDate(cboEnd)
   End If
End Sub

Public Sub PrintReport()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   On Error GoTo DiaErr1
   '    If sVendor <> "" Then
   '        cboVendor = sVendor
   '    End If
   If Trim(cboVendor) = "" Then
      sMsg = "Please Select A Vendor."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   MouseCursor 13
   
'   SetMdiReportsize MdiSect
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'   MdiSect.Crw.Formulas(2) = "Title1 = 'From: " & Trim(cboStart) & " To: " & Trim(cboEnd) & "'"
'   MdiSect.Crw.Formulas(3) = "Title2 = '" & Trim(txtAmt) & "'"
'   MdiSect.Crw.Formulas(4) = "Title3 = 'Vendor " & Trim(cboVendor) & ": " & lblName & "'"
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    aFormulaName.Add "Title3"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'From: " & CStr(Trim(cboStart) & " To: " & Trim(cboEnd)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Trim(txtAmt)) & "'")
    aFormulaValue.Add CStr("'Vendor " & CStr(Trim(cboVendor) & ": " & lblName) & "'")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finap13.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   
   ' removed DCDEBIT <> 0 2/28/06
   '    sSql = "{JritTable.DCDEBIT} <> 0 and {ChksTable.CHKVENDOR} = '" _
   '        & Compress(cboVendor.Text) & "'"
   
   sSql = "{ChksTable.CHKVENDOR} = '" _
          & Compress(cboVendor.Text) & "'"
   '    If sNum <> "" Then
   '        bRemote = True
   '        optDis = True
   '        sSql = sSql & " AND {ChksTable.CACHECKNO} = '" & sNum & "'"
   '    Else
   If Trim(cboStart) <> "" Then
      sSql = sSql & " and {ChksTable.CHKACTUALDATE} >= DateTime ('" _
             & Trim(cboStart) & "')"
   End If
   If Trim(cboEnd) <> "" Then
      sSql = sSql & " and {ChksTable.CHKACTUALDATE} <= DateTime ('" _
             & Trim(cboEnd) & "')"
   End If
   If Trim(txtAmt) <> "" Then
      sSql = sSql & " and ccur({ChksTable.CHKAMOUNT}) = ccur('" _
             & Trim(txtAmt.Text) & "')"
   End If
   
   If Trim(cboChkNum) <> "<ALL>" Then
      sSql = sSql & " and ccur({ChksTable.CHKNUMBER}) = ccur('" _
             & Trim(cboChkNum) & "')"
   End If
   sSql = sSql & " and {JrhdTable.MJTYPE} in ['XC', 'CC']"
   '    End If
'   MdiSect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
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

Public Sub PrintReport1()
   On Error GoTo DiaErr1
   '    If sVendor <> "" Then
   '        cboVendor = sVendor
   '    End If
   If Trim(cboVendor) = "" Then
      sMsg = "Please Select A Vendor."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   MouseCursor 13
   
   'SetMdiReportsize MdiSect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1 = 'From: " & Trim(cboStart) & " To: " & Trim(cboEnd) & "'"
   MdiSect.crw.Formulas(3) = "Title2 = '" & Trim(txtAmt) & "'"
   MdiSect.crw.Formulas(4) = "Title3 = 'Vendor " & Trim(cboVendor) & ": " & lblName & "'"
   
   sCustomReport = GetCustomReport("finap13.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   ' removed DCDEBIT <> 0 2/28/06
   '    sSql = "{JritTable.DCDEBIT} <> 0 and {ChksTable.CHKVENDOR} = '" _
   '        & Compress(cboVendor.Text) & "'"
   
   sSql = "{ChksTable.CHKVENDOR} = '" _
          & Compress(cboVendor.Text) & "'"
   '    If sNum <> "" Then
   '        bRemote = True
   '        optDis = True
   '        sSql = sSql & " AND {ChksTable.CACHECKNO} = '" & sNum & "'"
   '    Else
   If Trim(cboStart) <> "" Then
      sSql = sSql & " and {ChksTable.CHKACTUALDATE} >= DateTime ('" _
             & Trim(cboStart) & "')"
   End If
   If Trim(cboEnd) <> "" Then
      sSql = sSql & " and {ChksTable.CHKACTUALDATE} <= DateTime ('" _
             & Trim(cboEnd) & "')"
   End If
   If Trim(txtAmt) <> "" Then
      sSql = sSql & " and ccur({ChksTable.CHKAMOUNT}) = ccur('" _
             & Trim(txtAmt.Text) & "')"
   End If
   
   If Trim(cboChkNum) <> "<ALL>" Then
      sSql = sSql & " and ccur({ChksTable.CHKNUMBER}) = ccur('" _
             & Trim(cboChkNum) & "')"
   End If
   
   '    End If
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & "StartDate", cboStart
   SaveSetting "Esi2000", "EsiFina", Me.Name & "EndDate", cboEnd
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Vendor", cboVendor
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Amount", txtAmt
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim defaultDate As String
   defaultDate = Format(Date, "mm/dd/yyyy")
   cboStart = GetSetting("Esi2000", "EsiFina", Me.Name & "StartDate", defaultDate)
   cboEnd = GetSetting("Esi2000", "EsiFina", Me.Name & "EndDate", defaultDate)
   cboVendor = GetSetting("Esi2000", "EsiFina", Me.Name & "Vendor", cboVendor.List(0))
   lblName = GetVendorName(cboVendor)
   txtAmt = GetSetting("Esi2000", "EsiFina", Me.Name & "Amount", 0)
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub

Private Sub GetCheckNumber()

    Dim sVendor As String
    Dim sAmt As String
    Dim sStartDate As String
    Dim sEnddate As String
    Dim sChkSQL As String
    
    If Trim(cboVendor.Text) <> "" Then
        sVendor = cboVendor.Text
    Else
        ' Atleast Vendor name should be present.
        Exit Sub
    End If
    
    sStartDate = Trim(cboStart.Text)
    sEnddate = Trim(cboEnd.Text)
    sAmt = Trim(txtAmt.Text)
    
    sChkSQL = "Select CHKNUMBER from ChksTable " & vbCrLf _
          & " WHERE CHKVENDOR = '" & sVendor & "'"
    
    
    If sStartDate <> "" And _
            sEnddate <> "" Then
        sChkSQL = sChkSQL & " AND CHKACTUALDATE BETWEEN " & vbCrLf _
                    & "'" & sStartDate & "' AND '" & sEnddate & "'"
    End If
    
    If sAmt <> "" Then
        sChkSQL = sChkSQL & " AND CHKAMOUNT = '" & sAmt & "'"
    End If
    
    ' Load the Check combo box
    LoadComboWithSQL cboChkNum, sChkSQL, True
    
    MdiSect.ActiveForm.Refresh
End Sub
