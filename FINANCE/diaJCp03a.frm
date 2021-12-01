VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaJCp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MO Overhead Summary"
   ClientHeight    =   2430
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2430
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6120
      TabIndex        =   11
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaJCp03a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaJCp03a.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Contains Qualified Runs"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Qualified Part Numbers (CO)"
      Top             =   1080
      Width           =   3060
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   2
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
      PictureUp       =   "diaJCp03a.frx":0308
      PictureDn       =   "diaJCp03a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   9
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
      PictureUp       =   "diaJCp03a.frx":0594
      PictureDn       =   "diaJCp03a.frx":06DA
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   15
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Qty"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   14
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   720
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblStu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   1125
      Width           =   1095
   End
End
Attribute VB_Name = "diaJCp03a"
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
'
' diaJCp03a - MO Cost Analysis Report
'
' Created: (cjs)
' Revisions:
'   06/11/03 (nth) Added VITADDRES to PO cost on report per incident 17887
'   05/07/04 (nth) Removed jet DB logic use subreport instead
'   05/18/04 (nth) Added options from MCS see dbm23
' 4/11/05 TEL - formatted date passed to MO Cost Analysis (finjc01.rpt) as mm/dd/yy
' 6/8/05 TEL - allow selection of closed runs
'
'************************************************************************************

'Dim rdoQry As rdoQuery
Dim AdoCmdObj As ADODB.Command

Dim bOnLoad As Byte
Dim bGoodMo As Byte

Dim lRunno As Long
Dim SPartRef As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Public Sub GetStatusRun()
   Dim RdoStu As ADODB.Recordset
   On Error GoTo DiaErr1
   SPartRef = Compress(cmbPrt)
   sSql = "SELECT RUNSTATUS,RUNQTY from RunsTable WHERE RUNREF = '" _
          & SPartRef & "' AND RUNNO=" & cmbRun & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStu, ES_FORWARD)
   If bSqlRows Then
      With RdoStu
         lblStu = "" & Trim(!RUNSTATUS)
         lblQty = "" & Trim(!RUNQTY)
         .Close
      End With
   Else
      lblStu = ""
      lblQty = ""
   End If
   Set RdoStu = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getstatusRun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillFormRuns()
   Dim RdoRns As ADODB.Recordset
   Dim SPartRef As String
   cmbRun.Clear
   SPartRef = Compress(cmbPrt)
   'rdoQry(0) = SPartRef
   AdoCmdObj.parameters(0) = SPartRef
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoCmdObj)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
      End With
   Else
   End If
   If cmbRun.ListCount > 0 Then
      cmbRun = Format(cmbRun.List(0), "####0")
      GetStatusRun
   End If
   On Error Resume Next
   Set RdoRns = Nothing
   
   Exit Sub
   
DiaErr1:
   sProcName = "fillformru"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbPrt_Click()
   LocalFindPart Me
   FillFormRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   LocalFindPart Me
   FillFormRuns
   
End Sub


Private Sub cmbRun_Click()
   If Val(cmbRun) > 0 Then GetStatusRun Else _
          lblStu = ""
   
End Sub


Private Sub cmbRun_LostFocus()
   If Val(cmbRun) > 0 Then GetStatusRun Else _
          lblStu = ""
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub FillCombo()
   Dim RdoPcl As ADODB.Recordset
   Dim sTempPart As String
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PALEVEL,RUNREF," _
          & "RUNSTATUS FROM PartTable,RunsTable WHERE " _
          & "RUNREF=PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPcl)
   If bSqlRows Then
      With RdoPcl
         cmbPrt = "" & Trim(!PARTNUM)
         lblDsc = "" & Trim(!PADESC)
         Do Until .EOF
            If sTempPart <> Trim(!PARTNUM) Then
               'cmbPrt.AddItem "" & Trim(!PARTNUM)
               AddComboStr cmbPrt.hWnd, "" & Trim(!PARTNUM)
               sTempPart = Trim(!PARTNUM)
            End If
            .MoveNext
         Loop
      End With
      If cmbPrt.ListCount > 0 Then FillFormRuns
   Else
      MsgBox "No Matching Runs Recorded.", _
         vbInformation, Caption
   End If
   On Error Resume Next
   Set RdoPcl = Nothing
   cmbPrt.SetFocus
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
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim i As Integer
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   GetOptions
   '    sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
   '        & "RunsTable WHERE RUNREF = ? " _
   '        & "AND (RUNSTATUS<>'CA' AND RUNSTATUS<>'CL')  "
   
   'allow selection of closed runs
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND (RUNSTATUS<>'CA')  "
   
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim prmRunRef As ADODB.Parameter
   Set prmRunRef = New ADODB.Parameter
   prmRunRef.Type = adChar
   prmRunRef.SIZE = 30
   AdoCmdObj.parameters.Append prmRunRef
   
   'Set rdoQry = RdoCon.CreateQuery("", sSql)
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
   Set diaJCp03a = Nothing
End Sub
Private Sub PrintReport()
   'Dim sCustomReport As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   

   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "MoPart"
   aFormulaName.Add "MoRun"
   aFormulaName.Add "MoQty"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & cmbPrt _
                        & " , Run No " & cmbRun & "'")
   aFormulaValue.Add CStr("'" & cmbPrt & "'")
   aFormulaValue.Add CStr(cmbRun)
   aFormulaValue.Add CStr("'" & lblQty & "'")
   
   'MdiSect.Crw.SelectionFormula = sSql
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
  
   sCustomReport = GetCustomReport("finjc03")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   cCRViewer.SetDbTableConnection
   
   aRptPara.Add CStr(Compress(cmbPrt))
   aRptPara.Add CStr(cmbRun)
   aRptPara.Add CStr(lblQty)
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   
   ' Set report parameter
   
   cCRViewer.CRViewerSize Me
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType    'must happen AFTER SetDbTableConnection call!
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
  
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
End Sub

Public Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
   Else
   End If
End Sub



Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
      lblStu = ""
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optAct_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
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


Private Sub LocalFindPart(frm As Form, Optional sGetPart As String)
   Dim RdoPrt As ADODB.Recordset
   If sGetPart = "" Then
      sGetPart = Compress(frm.cmbPrt)
   Else
      sGetPart = Compress(sGetPart)
   End If
   On Error Resume Next
   If Len(sGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable " _
             & "WHERE PARTREF='" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
      If bSqlRows Then
         With RdoPrt
            frm.cmbPrt = "" & Trim(!PARTNUM)
            frm.lblDsc.ForeColor = frm.ForeColor
            frm.lblDsc = "" & Trim(!PADESC)
         End With
      Else
         frm.lblDsc.ForeColor = ES_RED
         frm.cmbPrt = "NONE"
         frm.lblDsc = "*** Part Number Wasn't Found ***"
         
      End If
   Else
      frm.cmbPrt = "NONE"
   End If
   Set RdoPrt = Nothing
End Sub


