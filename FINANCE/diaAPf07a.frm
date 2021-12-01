VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPf07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "1099's"
   ClientHeight    =   2820
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2820
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMinimum 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Tag             =   "1"
      Text            =   "0.00"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox chkDetail 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1980
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboStart 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "4"
      Top             =   840
      Width           =   1350
   End
   Begin VB.ComboBox cboEnd 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1200
      Width           =   1350
   End
   Begin VB.CheckBox chkTestPattern 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   2340
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboVendor 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Contains Vendors With PO's Not Invoiced"
      Top             =   480
      Width           =   1560
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4740
      Top             =   1860
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2820
      FormDesignWidth =   5835
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4680
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4680
      TabIndex        =   9
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaAPf07a.frx":0000
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
         Picture         =   "diaAPf07a.frx":017E
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
      TabIndex        =   10
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
      PictureUp       =   "diaAPf07a.frx":0308
      PictureDn       =   "diaAPf07a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   11
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
      PictureUp       =   "diaAPf07a.frx":0594
      PictureDn       =   "diaAPf07a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail"
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   18
      Top             =   1980
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   17
      Top             =   900
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   16
      Top             =   1260
      Width           =   825
   End
   Begin VB.Label chkSkip600 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Pattern"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   2340
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   2
      Left            =   180
      TabIndex        =   14
      Top             =   540
      Width           =   885
   End
   Begin VB.Label chkSkip600 
      BackStyle       =   0  'Transparent
      Caption         =   "Skip less than"
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   13
      Top             =   1620
      Width           =   1125
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaAPf07a"
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

'*********************************************************************************
' diaAPf07a - 1099's
'
' Notes:
'
' Created: 01/29/03 (nth)
' Revisions:
'   03/04/04 (nth) Added detail for Linda
'   03/05/04 (nth) Added cash basis version of report for Linda
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   LoadComboWithVendors cboVendor, True
   sCurrForm = Caption
   cboStart = Format(Now, "mm/01/yy")
   cboEnd = Format(Now, "mm/31/yy")
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
   FormUnload
   Set diaARp15a = Nothing
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   MouseCursor 13
   On Error GoTo DiaErr1
   
   optPrn.enabled = False
   optDis.enabled = False
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   'SetMdiReportsize MdiSect
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finap16.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Phone"
    aFormulaName.Add "Fax"
    aFormulaName.Add "CoAddress1"
    aFormulaName.Add "CoAddress2"
    aFormulaName.Add "CoAddress3"
    aFormulaName.Add "CoAddress4"
    aFormulaName.Add "StartDate"
    aFormulaName.Add "EndDate"
    aFormulaName.Add "Minimum"
    aFormulaName.Add "ShowDetail"
    aFormulaName.Add "TestPattern"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'1099'")
    aFormulaValue.Add CStr("''")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Phone) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Fax) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(1)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(2)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(3)) & "'")
    aFormulaValue.Add CStr("'" & CStr(Co.Addr(4)) & "'")

    aFormulaValue.Add CStr("'" & CStr(cboStart) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboEnd) & "'")
    aFormulaValue.Add CStr(txtMinimum.Text)
    aFormulaValue.Add CStr(chkDetail.Value)
    aFormulaValue.Add CStr(chkTestPattern.Value)
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
'    sSql = "{VndrTable.VE1099} = 1.00 and {ChksTable.CHKVOID} <> 1.00 AND " _
'            & "{ChksTable.CHKACTUALDATE} >= #" & Trim(cboStart) & "# and " _
'            & "{ChksTable.CHKACTUALDATE} <= #" & Trim(cboEnd) & "#"

    sSql = "{VndrTable.VE1099} = 1.00 and {ChksTable.CHKVOID} <> 1.00 AND " _
            & "{ChksTable.CHKPOSTDATE} >= #" & Trim(cboStart) & "# and " _
            & "{ChksTable.CHKPOSTDATE} <= #" & Trim(cboEnd) & "#"
    
    cCRViewer.SetReportSelectionFormula (sSql)
    cCRViewer.CRViewerSize Me
    
    ' Set report parameter
    cCRViewer.SetDbTableConnection


    cCRViewer.OpenCrystalReportObject Me, aFormulaName

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aRptParaType
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   optPrn.enabled = True
   optDis.enabled = True
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   MouseCursor 13
   On Error GoTo DiaErr1
   optPrn.enabled = False
   optDis.enabled = False
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("finap16.rpt")
   
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Title1='1099'"
   MdiSect.crw.Formulas(2) = "Title2=''"
   MdiSect.crw.Formulas(3) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(4) = "Phone='" & Co.Phone & "'"
   'MdiSect.crw.Formulas(5) = "Fax='" & Co.Fax & "'"
   MdiSect.crw.Formulas(5) = "CoAddress1='" & Co.Addr(1) & "'"
   MdiSect.crw.Formulas(6) = "CoAddress2='" & Co.Addr(2) & "'"
   MdiSect.crw.Formulas(7) = "CoAddress3='" & Co.Addr(3) & "'"
   MdiSect.crw.Formulas(8) = "CoAddress4='" & Co.Addr(4) & "'"
   MdiSect.crw.Formulas(9) = "StartDate='" & cboStart & "'"
   MdiSect.crw.Formulas(10) = "EndDate='" & cboEnd & "'"
   MdiSect.crw.Formulas(11) = "Minimum=" & txtMinimum.Text
   MdiSect.crw.Formulas(12) = "ShowDetail=" & chkDetail.Value
   MdiSect.crw.Formulas(13) = "TestPattern=" & chkTestPattern.Value
   
   'MdiSect.crw.SelectionFormula = sSql
  ' SetCrystalAction Me
   optPrn.enabled = True
   optDis.enabled = True
   MouseCursor 0
   Exit Sub
DiaErr1:
   MsgBox ("Here5")
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   MsgBox ("Here6")
   DoModuleErrors Me
   MsgBox ("Here7")
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = "0" & CStr(chkDetail) & CStr(chkTestPattern)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & "StartDate", cboStart
   SaveSetting "Esi2000", "EsiFina", Me.Name & "EndDate", cboEnd
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Minimum", txtMinimum
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Vendor", cboVendor
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, "000")
   'chkSkip.Value = CInt(Mid(sOptions, 1, 1))
   chkDetail.Value = CInt(Mid(sOptions, 2, 1))
   chkTestPattern.Value = CInt(Mid(sOptions, 3, 1))
   
   Dim defaultDate As String
   defaultDate = Format(Date, "mm/dd/yyyy")
   cboStart = GetSetting("Esi2000", "EsiFina", Me.Name & "StartDate", defaultDate)
   cboEnd = GetSetting("Esi2000", "EsiFina", Me.Name & "EndDate", defaultDate)
   txtMinimum = GetSetting("Esi2000", "EsiFina", Me.Name & "Minimum", "600.00")
   cboVendor = GetSetting("Esi2000", "EsiFina", Me.Name & "Vendor", cboVendor.List(0))
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub cboEnd_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboEnd_LostFocus()
   cboEnd = CheckDate(cboEnd)
End Sub

Private Sub cboStart_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboStart_LostFocus()
   cboStart = CheckDate(cboStart)
End Sub
