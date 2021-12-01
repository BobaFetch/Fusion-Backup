VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLp15a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cash Account Reconciliation (Report)"
   ClientHeight    =   2595
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2595
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtFrm 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton CmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5040
      TabIndex        =   4
      Top             =   240
      Width           =   1335
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Display The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Print The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   5
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
      PictureUp       =   "diaGLp15a.frx":0000
      PictureDn       =   "diaGLp15a.frx":0146
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   6
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
      PictureUp       =   "diaGLp15a.frx":028C
      PictureDn       =   "diaGLp15a.frx":03D2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Last Reconciled Date)"
      Height          =   405
      Index           =   3
      Left            =   4440
      TabIndex        =   14
      Top             =   1560
      Width           =   1905
   End
   Begin VB.Label lblThr 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   12
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   825
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "diaGLp15a"
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
' diaGLp15a - Account Reconciliation (Report)
'
' Notes:
'
' Created: (nth) 09/22/04
' Revisions:
' 11/15/04 (nth) Add a from date per JEVINT
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbAct_Click()
   GetAccount
End Sub

Private Sub cmbAct_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbAct_LostFocus()
   GetAccount
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
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
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
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
   Set diaGLp15a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   Dim rdoAct As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
   If bSqlRows Then
      With rdoAct
         While Not .EOF
            AddComboStr cmbAct.hwnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Wend
      End With
      cmbAct.ListIndex = 0
   End If
   Set rdoAct = Nothing
   Exit Sub
DiaErr1:
   sProcName = "fillcomb"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
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
   sCustomReport = GetCustomReport("fingl15")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "From"
    aFormulaName.Add "Through"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtFrm) & "'")
    aFormulaValue.Add CStr("'" & CStr(lblThr) & "'")
    
   sSql = "{GlacTable.GLACCTREF} = '" & Compress(cmbAct) & "'"
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
DiaErr1:
   sProcName = "printrep"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCustomReport As String
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("fingl15")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   
   MdiSect.crw.Formulas(2) = "From='" & txtFrm & "'"
   MdiSect.crw.Formulas(3) = "Through='" & lblThr & "'"
   
   sSql = "{GlacTable.GLACCTREF} = '" & Compress(cmbAct) & "'"
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "printrep"
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

Private Sub txtFrm_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtFrm_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtFrm_LostFocus()
   txtFrm = CheckDate(txtFrm)
End Sub


Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub

Private Sub GetAccount()
   Dim rdoAct As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT GLDESCR,GLRECDATE FROM GlacTable WHERE GLACCTREF = '" _
          & Compress(cmbAct) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
   If bSqlRows Then
      With rdoAct
         lblDsc.ForeColor = Me.ForeColor
         lblDsc = "" & Trim(.Fields(0))
         lblThr = Format(.Fields(1), "mm/dd/yy")
         txtFrm = Format(.Fields(1), "mm/01/yy")
         .Cancel
      End With
   Else
      lblDsc.ForeColor = ES_RED
      lblDsc = "*** Invalid Account Number ***"
      lblThr = ""
      txtFrm = Format(ES_SYSDATE, "mm/dd/yy")
   End If
   Set rdoAct = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getaccount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
