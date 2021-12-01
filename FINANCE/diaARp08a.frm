VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARp08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unprinted Invoices (Report)"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtEDte 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ComboBox txtSDte 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5640
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2475
      FormDesignWidth =   6315
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5160
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Save And Exit"
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5160
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaARp08a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaARp08a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   7
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
      PictureUp       =   "diaARp08a.frx":0308
      PictureDn       =   "diaARp08a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   8
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
      PictureUp       =   "diaARp08a.frx":0594
      PictureDn       =   "diaARp08a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaARp08a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*********************************************************************************
'
' diaPar08a - Unprinted Invoices
'
' Created: 11/29/01 (nth)
' Revisions:
'   10/28/03 (JCW!) (Fixed Report and Selection Query\Bad Joins)
'   08/15/04 (nth) Added getoptions and saveoptions
'
'
'*********************************************************************************
Option Explicit

Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = True
   sCurrForm = Caption
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   SaveOptions
   FormUnload
   Set diaARp08a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txteDte = Format(ES_SYSDATE, "mm/dd/yy")
   txtsDte = Left(txteDte, 3) & "01" & Right(txteDte, 3)
   
End Sub

Private Sub PrintReport()
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
    sCustomReport = GetCustomReport("finar08.rpt")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
   
   ' Set report titles and headers
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'List Of Unprinted Invoices'")
    aFormulaValue.Add CStr("'From" & CStr(Trim(txtsDte) & "  Through " & Trim(txteDte)) & "'")
   
   sSql = ""
   If Trim(txtsDte) <> "" Then
      sSql = sSql & "{CihdTable.INVDATE} >= #" & Trim(txtsDte) & "# AND "
   End If
   
   If Trim(txteDte) <> "" Then
      sSql = sSql & "{CihdTable.INVDATE} <= #" & Trim(txteDte) & "# AND "
   End If
   
   sSql = sSql & "cstr({CihdTable.INVPRINTED}) = ''  AND {JrhdTable.MJTYPE} = 'SJ'"
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
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
   
   ' Error handeling
DiaErr1:
   
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCustomReport As String
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   
   sCustomReport = GetCustomReport("finar08.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   ' Set report titles and headers
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1='List Of Unprinted Invoices'"
   MdiSect.crw.Formulas(3) = "Title2='From " & Trim(txtsDte) & "  Through " & Trim(txteDte) & "'"
   
   sSql = ""
   If Trim(txtsDte) <> "" Then
      sSql = sSql & "{CihdTable.INVDATE} >= #" & Trim(txtsDte) & "# AND "
   End If
   
   If Trim(txteDte) <> "" Then
      sSql = sSql & "{CihdTable.INVDATE} <= #" & Trim(txteDte) & "# AND "
   End If
   
   sSql = sSql & "cstr({CihdTable.INVPRINTED}) = ''  AND {JrhdTable.MJTYPE} = 'SJ'"
   MdiSect.crw.SelectionFormula = sSql
   
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
   ' Error handeling
DiaErr1:
   
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub txtEDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEdte_LostFocus()
   txteDte = CheckDate(txteDte)
End Sub

Private Sub txtSDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtSDte_LostFocus()
   txtsDte = CheckDate(txtsDte)
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub
