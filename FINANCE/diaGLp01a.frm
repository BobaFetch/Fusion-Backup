VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaGLp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chart Of Accounts (Report)"
   ClientHeight    =   2910
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2910
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optAct 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtLvl 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   285
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaGLp01a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaGLp01a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Text            =   "ALL"
      Top             =   1080
      Width           =   1935
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
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
      PictureUp       =   "diaGLp01a.frx":0308
      PictureDn       =   "diaGLp01a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2910
      FormDesignWidth =   7260
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
      PictureUp       =   "diaGLp01a.frx":0594
      PictureDn       =   "diaGLp01a.frx":06DA
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Inactive"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4200
      TabIndex        =   13
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Index           =   8
      Left            =   3720
      TabIndex        =   12
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Zero Or 9 For All)"
      Height          =   285
      Index           =   3
      Left            =   4800
      TabIndex        =   11
      Top             =   1560
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through Level"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   9
      Top             =   1080
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Accounts"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "diaGLp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'************************************************************************************
' diaGLp01a - Chart Of Accounts (Report)
'
' Notes:
'
' Created: (cjs)
' Revsions:
'   08/15/03 (nth) Added the bChart flag
'   03/29/04 (nth) fixed chart of level filter.
'
'************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bClose As Byte
Dim vAccounts(10, 4) As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Private Sub cmbAct_Click()
   lblTyp = vAccounts(cmbAct.ListIndex, 3)
End Sub

Private Sub cmbAct_LostFocus()
   Dim b As Boolean
   Dim i As Integer
   On Error Resume Next
   cmbAct = CheckLen(cmbAct, 12)
   For i = 0 To cmbAct.ListCount - 1
      If cmbAct = cmbAct.List(i) Then b = True
   Next
   If Not b Then
      cmbAct = cmbAct.List(0)
      cmbAct.ListIndex = 0
   End If
   If cmbAct.ListIndex > -1 Then
      lblTyp = vAccounts(cmbAct.ListIndex, 3)
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
End Sub

Private Sub FillCombo()
   Dim i As Integer
   Dim RdoGlm As ADODB.Recordset
   Dim sAccount As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM GlmsTable WHERE COACCTREC=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm)
   If bSqlRows Then
      vAccounts(0, 0) = "ALL"
      vAccounts(0, 1) = "ALL"
      vAccounts(0, 2) = "All Accounts"
      vAccounts(0, 3) = "0"
      With RdoGlm
         i = 1
         vAccounts(i, 0) = "" & Trim(!COASSTREF)
         vAccounts(i, 1) = "" & Trim(!COASSTACCT)
         vAccounts(i, 2) = "" & Trim(!COASSTDESC)
         vAccounts(i, 3) = Format(!COASSTTYPE, "0")
         
         i = 2
         vAccounts(i, 0) = "" & Trim(!COLIABREF)
         vAccounts(i, 1) = "" & Trim(!COLIABACCT)
         vAccounts(i, 2) = "" & Trim(!COLIABDESC)
         vAccounts(i, 3) = Format(!COLIABTYPE, "0")
         
         i = 3
         vAccounts(i, 0) = "" & Trim(!COEQTYREF)
         vAccounts(i, 1) = "" & Trim(!COEQTYACCT)
         vAccounts(i, 2) = "" & Trim(!COEQTYDESC)
         vAccounts(i, 3) = Format(!COEQTYTYPE, "0")
         
         i = 4
         vAccounts(i, 0) = "" & Trim(!COINCMREF)
         vAccounts(i, 1) = "" & Trim(!COINCMACCT)
         vAccounts(i, 2) = "" & Trim(!COINCMDESC)
         vAccounts(i, 3) = Format(!COINCMTYPE, "0")
         
         i = 6
         vAccounts(i, 0) = "" & Trim(!COEXPNREF)
         vAccounts(i, 1) = "" & Trim(!COEXPNACCT)
         vAccounts(i, 2) = "" & Trim(!COEXPNDESC)
         vAccounts(i, 3) = Format(!COEXPNTYPE, "0")
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = 5
            vAccounts(i, 0) = "" & Trim(!COCOGSREF)
            vAccounts(i, 1) = "" & Trim(!COCOGSACCT)
            vAccounts(i, 2) = "" & Trim(!COCOGSDESC)
            vAccounts(i, 3) = Format(!COCOGSTYPE, "0")
         End If
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = 7
            vAccounts(i, 0) = "" & Trim(!COOINCREF)
            vAccounts(i, 1) = "" & Trim(!COOINCACCT)
            vAccounts(i, 2) = "" & Trim(!COOINCDESC)
            vAccounts(i, 3) = Format(!COOINCTYPE, "0")
         End If
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = 8
            vAccounts(i, 0) = "" & Trim(!COOEXPREF)
            vAccounts(i, 1) = "" & Trim(!COOEXPACCT)
            vAccounts(i, 2) = "" & Trim(!COOEXPDESC)
            vAccounts(i, 3) = Format(!COOEXPTYPE, "0")
         End If
         
         sAccount = "" & Trim(!COCOGSREF)
         If Len(sAccount) Then
            i = 9
            vAccounts(i, 0) = "" & Trim(!COFDTXREF)
            vAccounts(i, 1) = "" & Trim(!COFDTXACCT)
            vAccounts(i, 2) = "" & Trim(!COFDTXDESC)
            vAccounts(i, 3) = Format(!COFDTXTYPE, "0")
         End If
         .Cancel
      End With
   End If
   iTotal = i
   For i = 0 To iTotal
      AddComboStr cmbAct.hWnd, Format$(vAccounts(i, 1))
   Next
   If cmbAct.ListCount > 0 Then
      cmbAct = cmbAct.List(0)
      lblTyp = vAccounts(0, 3)
   End If
   Set RdoGlm = Nothing
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
   ReopenJet
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
   On Error Resume Next
   Set DbAct = Nothing
   'JetDb.Execute "DROP TABLE ActrpTable"
   Set diaGLp01a = Nothing
End Sub

Private Sub PrintReport()
   Dim sCustomReport As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim iShowAcct As Integer
   
   On Error GoTo DiaErr1
   
   If (cmbAct = "ALL") Then
      iShowAcct = "0"
   Else
      iShowAcct = CInt(cmbAct)
   End If
   
   sCustomReport = GetCustomReport("fingl01")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   
   cCRViewer.SetReportTitle = "fingl01.rpt"
   cCRViewer.ShowGroupTree False
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "Level"
   aFormulaName.Add "ShowTopLevelAcct"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & cmbAct & " " _
                        & "And Levels Through " & txtLvl & "...'")
   aFormulaValue.Add CInt(Val(txtLvl))
   aFormulaValue.Add CInt(iShowAcct)
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   cCRViewer.CRViewerSize Me
   ' Set report parameter
'MsgBox CStr(cCRViewer.crxReport.ParameterFields.count)
 
   cCRViewer.SetDbTableConnection True
   
'MsgBox CStr(cCRViewer.crxReport.ParameterFields.count)

   ' report parameter
   aRptPara.Add CStr(optAct)
   aRptParaType.Add CStr("String")
   ' Set report parameter
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType      'must happen AFTER SetDbTableConnection call!
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   
   Exit Sub
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


'Private Sub PrintReport1()
'   Dim sWindows As String
'   Dim sCustomReport As String
'
'   On Error GoTo DiaErr1
'
'   MouseCursor 13
'
'   SetMdiReportsize MdiSect
'
'   ReopenJet
'
'   sWindows = GetWindowsDir()
'   MdiSect.crw.DataFiles(0) = sWindows & "\temp\esifina.mdb"
'   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.crw.Formulas(1) = "Includes='" & cmbAct & " " _
'                       & "And Levels Through " & txtLvl & "...'"
'   MdiSect.crw.Formulas(2) = "Level=" & txtLvl
'  sCustomReport = GetCustomReport("fingl01.rpt")
'   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
'
'   SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtLvl = Format(Val(txtLvl), "0")
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(txtLvl) & Trim(str(optAct.Value))
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      txtLvl = Left(sOptions, 1)
      optAct.Value = Mid(sOptions, 2, 1)
   Else
      txtLvl = "0"
      optAct.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub

Private Sub optAct_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub

Private Sub optDis_Click()
   'BuildAccounts
   PrintReport
End Sub

Private Sub optPrn_Click()
   'BuildAccounts
   PrintReport
End Sub

Private Sub txtLvl_LostFocus()
   txtLvl = CheckLen(txtLvl, 1)
   txtLvl = Format(Abs(Val(txtLvl)), "0")
End Sub

