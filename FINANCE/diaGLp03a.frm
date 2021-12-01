VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Journals (Report)"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTyp 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Tag             =   "0"
      ToolTipText     =   "Journal Prefix"
      Top             =   1080
      Width           =   1605
   End
   Begin VB.ComboBox cmbFyr 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Fiscal Year"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaGLp03a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "diaGLp03a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5520
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2865
      FormDesignWidth =   6255
   End
   Begin VB.ComboBox cmbJrn 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5040
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
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
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLp03a.frx":0308
      PictureDn       =   "diaGLp03a.frx":044E
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
      PictureUp       =   "diaGLp03a.frx":0594
      PictureDn       =   "diaGLp03a.frx":06DA
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "(Optional)"
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   14
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal Year"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Journal"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "diaGLp03a"
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

'************************************************************************************
' diaGLp03a - General Journal (Report)
'
' Notes:
'
' Created: (nth)
' Revsions:
' 10/22/03 (nth) Removed message box "no journals found"
' 06/24/04 (nth) Added optional parameters to allow printreport to be called remotely
' 09/28/04 (nth) Added fiscal year and prefix filters.
'
'************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sMsg As String
Public bRemote As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Public Sub PrintReport(Optional sJournal As String)
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
'   SetMdiReportsize MdiSect
   
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "RequestBy='" & sInitials & "'"
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   
   sCustomReport = GetCustomReport("fingl03.rpt")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = "{GjhdTable.GJNAME} = '" & Trim(cmbJrn) & "'"
    
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

'   MdiSect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub PrintReport1(Optional sJournal As String)
   Dim sCustomReport As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   'SetMdiReportsize MdiSect
   If Len(sJournal) > 0 Then cmbJrn = sJournal
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='" & sInitials & "'"
   
   sCustomReport = GetCustomReport("fingl03.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = "{GjhdTable.GJNAME} = '" & Trim(cmbJrn) & "'"
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
Public Sub GetJrnl()
   Dim rdoJrn As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT GJDESC FROM GjhdTable WHERE GJNAME = '" & cmbJrn & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn)
   If bSqlRows Then
      With rdoJrn
         lblDsc = "" & Trim(!GJDESC)
         .Cancel
      End With
   Else
      lblDsc = ""
   End If
   Set rdoJrn = Nothing
   Exit Sub
DiaErr1:
   sProcName = "GetJrnl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbFyr_LostFocus()
   FillJrnl
End Sub

Private Sub cmbjrn_Click()
   GetJrnl
   GetJrnl
End Sub

Private Sub cmbJrn_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbjrn_LostFocus()
   cmbJrn = CheckLen(cmbJrn, 12)
   GetJrnl
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
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   If bRemote Then
      Me.WindowState = vbMinimized
   Else
      FormLoad Me
      bOnLoad = True
   End If
   FormatControls
   GetOptions
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   If Not bRemote Then
      FormUnload
   End If
   bRemote = False
   Set diaGLp03a = Nothing
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

Private Sub ShowPrinters_MouseUp(Button As Integer, Shift As Integer, _
                                 x As Single, y As Single)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub FillCombo()
   'Dim rdoYr As ADODB.Recordset
   'Dim rdoJrn As ADODB.Recordset
   On Error GoTo DiaErr1
   sProcName = "FillFiscalYe"
   FillFiscalYears Me
   sProcName = "filljrnl"
   FillJrnl
   Exit Sub
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillJrnl()
   Dim rdoJrn As ADODB.Recordset
   sProcName = "FillJrnl"
   cmbJrn.Clear
'   sSql = "SELECT GJNAME FROM GjhdTable WHERE YEAR(GJPOST) = " & cmbFyr
'   If Len(txtTyp) Then
'      sSql = sSql & " AND GJNAME LIKE '" & Trim(txtTyp) & "%'"
'   End If
   
   sSql = "SELECT GJNAME" & vbCrLf _
      & "FROM GjhdTable" & vbCrLf _
      & "JOIN GlfyTable ON GJPOST BETWEEN FYSTART AND FYEND" & vbCrLf _
      & "AND FYYEAR = " & cmbFyr & vbCrLf
   If Trim(txtTyp) <> "" Then
      sSql = sSql & "WHERE GJNAME LIKE '" & txtTyp & "%'"
   End If
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         Do Until .EOF
            AddComboStr cmbJrn.hwnd, "" & Trim(!GJNAME)
            .MoveNext
         Loop
         .Cancel
      End With
      cmbJrn.ListIndex = 0
   End If
   Set rdoJrn = Nothing
   GetJrnl
   Exit Sub
End Sub

Private Sub txtTyp_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtTyp_LostFocus()
   txtTyp = CheckLen(txtTyp, 12)
   FillJrnl
End Sub
