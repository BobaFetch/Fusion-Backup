VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPp07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Invoice Register (Report)"
   ClientHeight    =   3855
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Contains Vendors With Invoices"
      Top             =   1560
      Width           =   1555
   End
   Begin VB.TextBox txtInv 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Tag             =   "3"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CheckBox chkItm 
      Caption         =   "___"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkSrt 
      Caption         =   "___"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "4"
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5400
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   6
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
      PictureUp       =   "diaAPp07a.frx":0000
      PictureDn       =   "diaAPp07a.frx":0146
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
      FormDesignHeight=   3855
      FormDesignWidth =   6585
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   13
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
      PictureUp       =   "diaAPp07a.frx":028C
      PictureDn       =   "diaAPp07a.frx":03D2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL)"
      Height          =   285
      Index           =   7
      Left            =   3720
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices Beginning With"
      Height          =   405
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Default Sort By Vendor)"
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Invoice Items"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By Invoice Date"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Beginning Date"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   915
   End
End
Attribute VB_Name = "diaAPp07a"
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
' diaAPp07a - Vendor Invoice Register (Report)
'
' Notes:
'
' Created: (nth)
' Revisions:
'   12/20/02 (nth) Revised and brought up to spec's.
'   08/16/04 (nth) Added printer to saveoptions and getoptions.
'   08/31/04 (nth) Added leading characters search per AUBCOR.
'
'*************************************************************************************

Option Explicit
Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodVendor As Byte
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


'*************************************************************************************

Private Sub cmbVnd_Click()
   If cmbVnd <> "ALL" Then
      bGoodVendor = FindVendor(Me)
   Else
      lblNme = "All Vendors.."
   End If
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If Len(cmbVnd) = 0 Then cmbVnd = "ALL"
   If cmbVnd <> "ALL" Then
      bGoodVendor = FindVendor(Me)
   Else
      lblNme = "All Vendors.."
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub FillCombo()
   Dim RdoVed As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT VIVENDOR,VEREF,VENICKNAME " _
          & "FROM VihdTable,VndrTable WHERE VIVENDOR=VEREF ORDER BY VIVENDOR"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      With RdoVed
         cmbVnd = "ALL"
         AddComboStr cmbVnd.hwnd, "ALL"
         Do Until .EOF
            AddComboStr cmbVnd.hwnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
      End With
   End If
   lblNme = "All Vendors."
   Set RdoVed = Nothing
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   txtEnd = Format(Now, "mm/dd/yy")
   txtBeg = Format(Now, "mm/01/yy")
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   GetOptions
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set diaAPp07a = Nothing
End Sub

Private Sub PrintReport()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   On Error GoTo DiaErr1
   MouseCursor 13
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finap07.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   sSql = "{VihdTable.VIDATE} >= #" & txtBeg & "# AND {VihdTable.VIDATE} <= #" _
          & txtEnd & "#"
   If cmbVnd <> "ALL" Then
      sSql = sSql & "AND {VihdTable.VIVENDOR} = '" & Compress(cmbVnd) & "'"
   End If
   If Trim(txtInv) <> "" Then
      sSql = sSql & " AND {VihdTable.VINO} LIKE '" & txtInv & "*'"
   End If
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'From" & CStr(txtBeg & " To " & txtEnd) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")

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
   sProcName = "printrep"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCustomReport As String
   On Error GoTo DiaErr1
   MouseCursor 13
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("finap07.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{VihdTable.VIDATE} >= #" & txtBeg & "# AND {VihdTable.VIDATE} <= #" _
          & txtEnd & "#"
   If cmbVnd <> "ALL" Then
      sSql = sSql & "AND {VihdTable.VIVENDOR} = '" & Compress(cmbVnd) & "'"
   End If
   If Trim(txtInv) <> "" Then
      sSql = sSql & " AND {VihdTable.VINO} LIKE '" & txtInv & "*'"
   End If
   'sSql = sSql & "LEFT({JritTable.DCHEAD},2) = 'PJ'"
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Includes='From " & txtBeg & " To " & txtEnd & "'"
   MdiSect.crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "printrep"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = chkItm.Value & chkSrt.Value
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   On Error Resume Next
   If Len(Trim(sOptions)) > 0 Then
      chkItm.Value = Mid(sOptions, 1, 1)
      chkSrt.Value = Mid(sOptions, 2, 1)
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
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

Private Sub txtInv_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtInv_LostFocus()
   txtInv = CheckLen(txtInv, 20)
End Sub
