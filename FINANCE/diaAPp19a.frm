VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPp19a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Analysis (Report)"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optVnd 
      Caption         =   "    "
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   6000
      Width           =   735
   End
   Begin VB.OptionButton optDte 
      Caption         =   "    "
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   5760
      Width           =   735
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.ComboBox cmbvnd 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CheckBox chkVoidOnly 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   5160
      Width           =   735
   End
   Begin VB.CheckBox chkReCap 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   4920
      Width           =   735
   End
   Begin VB.CheckBox chkUnCleared 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   4680
      Width           =   735
   End
   Begin VB.CheckBox chkCleared 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   4440
      Width           =   735
   End
   Begin VB.CheckBox chkComp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   3960
      Width           =   735
   End
   Begin VB.CheckBox ChkExt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtBegNum 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Tag             =   "1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtEndNum 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox txtBegDte 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox txtEndDte 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4920
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Save And Exit"
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4920
      TabIndex        =   20
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaAPp19a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaAPp19a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   17
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
      PictureUp       =   "diaAPp19a.frx":0308
      PictureDn       =   "diaAPp19a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   18
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
      PictureUp       =   "diaAPp19a.frx":0594
      PictureDn       =   "diaAPp19a.frx":06DA
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   5520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6450
      FormDesignWidth =   6090
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort:"
      Height          =   285
      Index           =   19
      Left            =   120
      TabIndex        =   42
      Top             =   5520
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   18
      Left            =   120
      TabIndex        =   41
      Top             =   3720
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   17
      Left            =   3120
      TabIndex        =   40
      Top             =   3000
      Width           =   1785
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   39
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Account"
      Height          =   285
      Index           =   16
      Left            =   120
      TabIndex        =   38
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   15
      Left            =   3120
      TabIndex        =   37
      Top             =   2160
      Width           =   1785
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   36
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Vendor"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   35
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "By Vendor"
      Height          =   285
      Index           =   14
      Left            =   240
      TabIndex        =   34
      Top             =   6000
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "By Date"
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   33
      Top             =   5760
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Void Checks Only"
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   32
      Top             =   5160
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Missing Check Recap"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   31
      Top             =   4920
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Un-Cleared Checks"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   30
      Top             =   4680
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cleared Checks"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   29
      Top             =   4440
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "A/P Computer Checks"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   28
      Top             =   3960
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "A/P External Checks"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   27
      Top             =   4200
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   3120
      TabIndex        =   26
      Top             =   480
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   25
      Top             =   1320
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   1680
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Check #"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Check #"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   480
      Width           =   2145
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaAPp19a"
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

'***************************************************************************************
' diaAPp19a - Check Analysis
'
' Notes:
'
' Created: 11/06/01 (nth)
' Revisions:
'   12/20/02 (nth) Revised reports and added SaveOptions / GetOptions
'   12/20/02 (nth) Added CreateChkCodeTable and AssignChkCodes
'   10/22/03 (nth) Added customreport\
'   04/13/04 (nth) fixed selection formula runtime error
'
'***************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sMsg As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'***************************************************************************************

Private Sub cmbAct_Click()
   FindAccount Me
End Sub

Private Sub cmbAct_LostFocus()
   cmbAct = CheckLen(cmbAct, 12)
   If Trim(cmbAct) = "" Then
      lblDsc = "Multiple Accounts Selected."
   Else
      FindAccount Me
   End If
End Sub

Private Sub cmbVnd_Click()
   FindVendor Me
End Sub

Private Sub cmbVnd_LostFocus()
   cmbvnd = CheckLen(cmbvnd, 10)
   If Trim(cmbvnd) = "" Then
      lblNme = "Multiple Vendors Selected."
   Else
      FindVendor Me
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      FillVendors Me
      FillCombo ' Cash Accounts
      ReopenJet
      CreateChkCodeTable
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = True
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
   Set diaAPp19a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub txtBegDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEndDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub PrintReport()
   Dim sWindows As String
   Dim sTemp As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   On Error GoTo DiaErr1
   If ChkExt = vbUnchecked And chkComp = vbUnchecked Then
      sMsg = "Please Select A/P Computer Checks" & vbCrLf _
             & "A/P External Checks Or Both."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   MouseCursor 13
   
   optPrn.enabled = False
   optDis.enabled = False
   
   AssignChkCodes ' fill temp jet db with ! or *
   
   ReopenJet
   
   sWindows = GetWindowsDir()
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "StartDate"
    aFormulaName.Add "EndDate"
    aFormulaName.Add "StartCheck"
    aFormulaName.Add "EndCheck"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtBegDte) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtEndDte) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtBegNum) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtEndNum) & "'")
   
   sSql = ""
   
   If Trim(txtBegNum) <> "" Then _
           sSql = sSql & "{ChksTable.CHKNUMBER} >= '" & Trim(txtBegNum) & "' AND "
   If Trim(txtEndNum) <> "" Then _
           sSql = sSql & "{ChksTable.CHKNUMBER} <= '" & Trim(txtEndNum) & "' AND "
   If Trim(txtBegDte) <> "" Then _
           sSql = sSql & "{ChksTable.CHKACTUALDATE} >= #" & txtBegDte & "# AND "
   If Trim(txtEndDte) <> "" Then _
           sSql = sSql & "{ChksTable.CHKACTUALDATE} <= #" & txtEndDte & "# AND "
   
   If chkVoidOnly = vbChecked Then
        aFormulaName.Add "IncludeComputer"
        aFormulaValue.Add "'N'"
        aFormulaName.Add "IncludeExternal"
        aFormulaValue.Add "'N'"
        aFormulaName.Add "IncludeUnClearedChecks"
        aFormulaValue.Add "'N'"
        aFormulaName.Add "IncludeClearedChecks"
        aFormulaValue.Add "'N'"
        aFormulaName.Add "VoidOnly"
        aFormulaValue.Add "'Y'"
        
      sSql = sSql & "isdate(totext({ChksTable.CHKVOIDDATE})) AND "
   Else
        aFormulaName.Add "VoidOnly"
        aFormulaValue.Add "'N'"
   End If
   
   If chkComp = vbUnchecked Or ChkExt = vbUnchecked Then
      If chkComp = vbChecked Then
            aFormulaName.Add "IncludeComputer"
            aFormulaValue.Add "'Y'"
         'sSql = sSql & "left({JritTable.DCHEAD},2) = 'CC' AND "
         sSql = sSql & "{chkstable.CHKTYPE} = 2 AND "
         
      Else
            aFormulaName.Add "IncludeComputer"
            aFormulaValue.Add "'N'"
      End If
      
      If ChkExt = vbChecked Then
            aFormulaName.Add "IncludeExternal"
            aFormulaValue.Add "'Y'"
         'sSql = sSql & "left({JritTable.DCHEAD},2) = 'XC' AND "
         sSql = sSql & "{chkstable.CHKTYPE} IN [1,3] AND "
      Else
            aFormulaName.Add "IncludeExternal"
            aFormulaValue.Add "'N'"
      End If
   Else
       aFormulaName.Add "IncludeExternal"
       aFormulaValue.Add "'Y'"
       aFormulaName.Add "IncludeComputer"
       aFormulaValue.Add "'Y'"
      'sSql = sSql & "left({JritTable.DCHEAD},2) IN ['CC','XC'] AND "
      sSql = sSql & "{chkstable.CHKTYPE} IN [1,2,3] AND "
      
   End If
   
   If chkUnCleared = vbChecked And chkCleared = vbChecked Then
       aFormulaName.Add "IncludeUnClearedChecks"
       aFormulaValue.Add "'Y'"
       aFormulaName.Add "IncludeClearedChecks"
       aFormulaValue.Add "'Y'"
   ElseIf chkUnCleared = vbChecked And chkCleared = vbUnchecked Then
       aFormulaName.Add "IncludeUnClearedChecks"
       aFormulaValue.Add "'Y'"
       aFormulaName.Add "IncludeClearedChecks"
       aFormulaValue.Add "'N'"
      sSql = sSql & "not isdate(totext({ChksTable.CHKCLEARDATE})) AND "
   ElseIf chkUnCleared = vbUnchecked And chkCleared = vbChecked Then
       aFormulaName.Add "IncludeUnClearedChecks"
       aFormulaValue.Add "'N'"
       aFormulaName.Add "IncludeClearedChecks"
       aFormulaValue.Add "'Y'"
      sSql = sSql & "isdate(totext({ChksTable.CHKCLEARDATE})) AND "
   Else
       aFormulaName.Add "IncludeUnClearedChecks"
       aFormulaValue.Add "'Y'"
       aFormulaName.Add "IncludeClearedChecks"
       aFormulaValue.Add "'Y'"
   End If
   
   If chkReCap = vbChecked Then
       aFormulaName.Add "IncludeReCap"
       aFormulaValue.Add "'Y'"
   Else
       aFormulaName.Add "IncludeReCap"
       aFormulaValue.Add "'N'"
   End If
   
   If Trim(cmbvnd) <> "" Then
      aFormulaName.Add "CheckVendor"
      aFormulaValue.Add CStr("'" & CStr(Trim(cmbvnd)) & "'")
      sSql = sSql & "{ChksTable.CHKVENDOR} = '" & Compress(cmbvnd) & "' AND "
   Else
      aFormulaName.Add "CheckVendor"
      aFormulaValue.Add CStr("'ALL'")
   End If
   
   If Trim(cmbAct) <> "" Then
      aFormulaName.Add "CheckAccount"
      aFormulaValue.Add CStr("'" & CStr(Trim(cmbAct)) & "'")
      'sSql = sSql & "{JritTable.DCACCTNO} = '" & Compress(cmbAct) & "' AND "
      sSql = sSql & "{ChksTable.CHKACCT} = '" & Compress(cmbAct) & "' AND "
   Else
      aFormulaName.Add "CheckAccount"
      aFormulaValue.Add CStr("'ALL'")
   End If
   
   'sSql = sSql & " {JritTable.DCCREDIT} = 0 AND left({JritTable.dcdesc},4) <> 'VOID'"
   'sSql = sSql & " left({JritTable.dcdesc},4) <> 'VOID'"
   sSql = sSql & " {ChksTable.CHKVOID} <> 1"
   
   sTemp = sSql
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
   If optDte.Value = True Then
      aFormulaName.Add "SortByDate"
      aFormulaValue.Add CStr("'Y'")
      aFormulaName.Add "SortByVendor"
      aFormulaValue.Add CStr("'N'")
      sCustomReport = GetCustomReport("finch07a.rpt")
   Else
      aFormulaName.Add "SortByDate"
      aFormulaValue.Add CStr("'N'")
      aFormulaName.Add "SortByVendor"
      aFormulaValue.Add CStr("'Y'")
      sCustomReport = GetCustomReport("finch07b.rpt")
   End If
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sTemp
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   optPrn.enabled = True
   optDis.enabled = True
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   optPrn.enabled = True
   optDis.enabled = True
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sWindows As String
   Dim sCustomReport As String
   Dim sTemp As String
   
   On Error GoTo DiaErr1
   If ChkExt = vbUnchecked And chkComp = vbUnchecked Then
      sMsg = "Please Select A/P Computer Checks" & vbCrLf _
             & "A/P External Checks Or Both."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   MouseCursor 13
   
   optPrn.enabled = False
   optDis.enabled = False
   
   AssignChkCodes ' fill temp jet db with ! or *
   
   ReopenJet
   
   'SetMdiReportsize MdiSect
   sWindows = GetWindowsDir()
   
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "StartDate='" & txtBegDte & "'"
   MdiSect.crw.Formulas(3) = "EndDate='" & txtEndDte & "'"
   MdiSect.crw.Formulas(4) = "StartCheck='" & txtBegNum & "'"
   MdiSect.crw.Formulas(5) = "EndCheck='" & txtEndNum & "'"
   
   sSql = ""
   
   If Trim(txtBegNum) <> "" Then _
           sSql = sSql & "{ChksTable.CHKNUMBER} >= '" & Trim(txtBegNum) & "' AND "
   If Trim(txtEndNum) <> "" Then _
           sSql = sSql & "{ChksTable.CHKNUMBER} <= '" & Trim(txtEndNum) & "' AND "
   If Trim(txtBegDte) <> "" Then _
           sSql = sSql & "{ChksTable.CHKACTUALDATE} >= #" & txtBegDte & "# AND "
   If Trim(txtEndDte) <> "" Then _
           sSql = sSql & "{ChksTable.CHKACTUALDATE} <= #" & txtEndDte & "# AND "
   
   If chkVoidOnly = vbChecked Then
      MdiSect.crw.Formulas(6) = "IncludeComputer='N'"
      MdiSect.crw.Formulas(7) = "IncludeExternal='N'"
      MdiSect.crw.Formulas(8) = "IncludeUnClearedChecks='N'"
      MdiSect.crw.Formulas(9) = "IncludeClearedChecks='N'"
      MdiSect.crw.Formulas(11) = "VoidOnly='Y'"
      sSql = sSql & "isdate(totext({ChksTable.CHKVOIDDATE})) AND "
   Else
      MdiSect.crw.Formulas(11) = "VoidOnly='N'"
   End If
   
   If chkComp = vbUnchecked Or ChkExt = vbUnchecked Then
      If chkComp = vbChecked Then
         MdiSect.crw.Formulas(6) = "IncludeComputer='Y'"
         sSql = sSql & "left({JritTable.DCHEAD},2) = 'CC' AND "
      Else
         MdiSect.crw.Formulas(6) = "IncludeComputer='N'"
      End If
      
      If ChkExt = vbChecked Then
         MdiSect.crw.Formulas(7) = "IncludeExternal='Y'"
         sSql = sSql & "left({JritTable.DCHEAD},2) = 'XC' AND "
      Else
         MdiSect.crw.Formulas(7) = "IncludeExternal='N'"
      End If
   Else
      MdiSect.crw.Formulas(7) = "IncludeExternal='Y'"
      MdiSect.crw.Formulas(6) = "IncludeComputer='Y'"
      sSql = sSql & "left({JritTable.DCHEAD},2) IN ['CC','XC'] AND "
   End If
   
   If chkUnCleared = vbChecked And chkCleared = vbChecked Then
      MdiSect.crw.Formulas(8) = "IncludeUnClearedChecks='Y'"
      MdiSect.crw.Formulas(9) = "IncludeClearedChecks='Y'"
   ElseIf chkUnCleared = vbChecked And chkCleared = vbUnchecked Then
      MdiSect.crw.Formulas(8) = "IncludeUnClearedChecks='Y'"
      MdiSect.crw.Formulas(9) = "IncludeClearedChecks='N'"
      sSql = sSql & "not isdate(totext({ChksTable.CHKCLEARDATE})) AND "
   ElseIf chkUnCleared = vbUnchecked And chkCleared = vbChecked Then
      MdiSect.crw.Formulas(8) = "IncludeUnClearedChecks='N'"
      MdiSect.crw.Formulas(9) = "IncludeClearedChecks='Y'"
      sSql = sSql & "isdate(totext({ChksTable.CHKCLEARDATE})) AND "
   Else
      MdiSect.crw.Formulas(8) = "IncludeUnClearedChecks='Y'"
      MdiSect.crw.Formulas(9) = "IncludeClearedChecks='Y'"
   End If
   
   If chkReCap = vbChecked Then
      MdiSect.crw.Formulas(10) = "IncludeReCap='Y'"
   Else
      MdiSect.crw.Formulas(10) = "IncludeReCap='N'"
   End If
   
   If Trim(cmbvnd) <> "" Then
      MdiSect.crw.Formulas(12) = "CheckVendor='" & Trim(cmbvnd) & "'"
      sSql = sSql & "{ChksTable.CHKVENDOR} = '" & Compress(cmbvnd) & "' AND "
   Else
      MdiSect.crw.Formulas(12) = "CheckVendor='ALL'"
   End If
   
   If Trim(cmbAct) <> "" Then
      MdiSect.crw.Formulas(13) = "CheckAccount='" & Trim(cmbAct) & "'"
      sSql = sSql & "{JritTable.DCACCTNO} = '" & Compress(cmbAct) & "' AND "
   Else
      MdiSect.crw.Formulas(13) = "CheckAccount='ALL'"
   End If
   
   sSql = sSql & " {JritTable.DCCREDIT} = 0 AND left({JritTable.dcdesc},4) <> 'VOID'"
   
   sTemp = sSql
   
   If optDte.Value = True Then
      MdiSect.crw.Formulas(14) = "SortByDate='Y'"
      MdiSect.crw.Formulas(15) = "SortByVendor='N'"
      sCustomReport = GetCustomReport("finch07a.rpt")
   Else
      MdiSect.crw.Formulas(14) = "SortByDate='N'"
      MdiSect.crw.Formulas(15) = "SortByVendor='Y'"
      sCustomReport = GetCustomReport("finch07b.rpt")
   End If
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   MdiSect.crw.SelectionFormula = sTemp
   'SetCrystalAction Me
   
   optPrn.enabled = True
   optDis.enabled = True
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   optPrn.enabled = True
   optDis.enabled = True
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillCombo()
   Dim rdoAct As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
   
   If bSqlRows Then
      With rdoAct
         Do While Not .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
         .Cancel
      End With
      cmbAct.ListIndex = 0
      FindAccount Me
   End If
   Set rdoAct = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub SaveOptions()
   Dim sOptions As String * 9
   Dim sDte As String * 1
   Dim sVnd As String * 1
   
   If optDte.Value = False Then sDte = "0" Else sDte = "1"
   If optVnd.Value = False Then sVnd = "0" Else sVnd = "1"
   
   sOptions = RTrim(chkComp.Value) _
              & RTrim(ChkExt.Value) _
              & RTrim(chkCleared.Value) _
              & RTrim(chkUnCleared.Value) _
              & RTrim(chkReCap.Value) _
              & RTrim(chkVoidOnly.Value) _
              & sDte _
              & sVnd
   
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   
   If Len(Trim(sOptions)) > 0 Then
      chkComp.Value = Val(Left(sOptions, 1))
      ChkExt.Value = Val(Mid(sOptions, 2, 1))
      chkCleared.Value = Val(Mid(sOptions, 3, 1))
      chkUnCleared.Value = Val(Mid(sOptions, 4, 1))
      chkReCap.Value = Val(Mid(sOptions, 5, 1))
      chkVoidOnly.Value = Val(Mid(sOptions, 6, 1))
      
      If Mid(sOptions, 7, 1) = 0 Then
         optDte.Value = False
      Else
         optDte.Value = True
      End If
      
      If Mid(sOptions, 8, 1) = 0 Then
         optVnd.Value = False
      Else
         optVnd.Value = True
      End If
   End If
   
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub CreateChkCodeTable()
   Dim NewTb1 As TableDef
   Dim NewTb2 As TableDef
   Dim NewIdx1 As Index
   
   On Error Resume Next
   
   JetDb.Execute "DROP TABLE ChkCodeTable"
   
   'Fields. Note that we allow empties
   Set NewTb1 = JetDb.CreateTableDef("ChkCodeTable")
   With NewTb1
      'Check Number
      .Fields.Append .CreateField("Check", dbText, 12)
      'Code
      .Fields.Append .CreateField("Code", dbText, 1)
   End With
   
   'add the table and indexes to Jet.
   JetDb.TableDefs.Append NewTb1
   Set NewTb2 = JetDb!ChkCodeTable
   With NewTb2
      Set NewIdx1 = .CreateIndex
      With NewIdx1
         .Name = "CheckIdx"
         .Fields.Append .CreateField("Check")
      End With
      .Indexes.Append NewIdx1
   End With
   
End Sub

Private Sub AssignChkCodes()
   Dim sTemp As String
   Dim RdoChk As ADODB.Recordset
   Dim DbChk As Recordset
   Dim lLastChk As Long
   Dim lDif As Long
   
   On Error GoTo DiaErr1
   sSql = "SELECT CHKNUMBER as ChkNum FROM ChksTable "
   sTemp = ""
   
   If Trim(txtBegNum) <> "" Then _
           sTemp = sTemp & "CHKNUMBER >= '" & Trim(txtBegNum) & "' AND "
   If Trim(txtEndNum) <> "" Then _
           sTemp = sTemp & "CHKNUMBER <= '" & Trim(txtEndNum) & "' AND "
   If Trim(txtBegDte) <> "" Then _
           sTemp = sTemp & "CHKACTUALDATE >= " & txtBegDte & " AND "
   If Trim(txtEndDte) <> "" Then _
           sTemp = sTemp & "CHKACTUALDATE <= " & txtEndDte
   
   If Len(sTemp) Then sTemp = "WHERE " & sTemp
   
   If Right(Trim(sTemp), 4) = " AND" Then
      sTemp = Left(sTemp, Len(sTemp) - 4)
   End If
   
   sSql = sSql & sTemp & " ORDER BY ChkNum"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   
   If bSqlRows Then
      JetDb.Execute "DELETE * FROM ChkCodeTable"
      Set DbChk = JetDb.OpenRecordset("ChkCodeTable", dbOpenDynaset)
      
      With RdoChk
         lLastChk = CLng(!Chknum) - 1
         While Not .EOF
            If IsNumeric(!Chknum) Then
               DbChk.AddNew
               DbChk!Check = CStr(!Chknum)
               lDif = (CLng(!Chknum) - lLastChk)
               If lDif = 0 Then
                  DbChk!code = "!"
               ElseIf lDif = 1 Then
                  DbChk!code = " "
               Else
                  DbChk!code = "*"
               End If
               DbChk.Update
               lLastChk = CLng(!Chknum)
            End If
            .MoveNext
         Wend
         .Cancel
         JetDb.Close
      End With
   Else
      
   End If
   
   Set RdoChk = Nothing
   Set JetDb = Nothing
   
   Exit Sub
DiaErr1:
   sProcName = "AssignChkCodes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
