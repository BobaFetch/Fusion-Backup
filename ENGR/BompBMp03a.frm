VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form BompBMp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parts List Used On"
   ClientHeight    =   3660
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3660
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkShowBomComments 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   3300
      Width           =   735
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6140
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6140
      TabIndex        =   16
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "BompBMp03a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "BompBMp03a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   2700
      Width           =   735
   End
   Begin VB.CheckBox optAsn 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Revision (Blank For Default)"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   2100
      Width           =   735
   End
   Begin VB.CheckBox OptExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.ComboBox cmbPls 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "3"
      Top             =   1110
      Width           =   3345
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3660
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Comments"
      Height          =   165
      Index           =   9
      Left            =   240
      TabIndex        =   22
      Top             =   3300
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   288
      Index           =   8
      Left            =   5400
      TabIndex        =   20
      Top             =   1464
      Width           =   852
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6240
      TabIndex        =   19
      Top             =   1464
      Width           =   300
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Top             =   1460
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Comments"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   2700
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned PL Rev's Only"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Part Desc"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   252
      Index           =   3
      Left            =   5400
      TabIndex        =   11
      Top             =   1128
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   2100
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1125
      Width           =   1815
   End
End
Attribute VB_Name = "BompBMp03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/4/05 Corrected the selection criteria in FillCombo
Option Explicit
Dim bGoodList As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiEngr", "BompBMp03a", sOptions) & "00000000"
   optDsc = Mid(sOptions, 1, 1)
   OptExt = Mid(sOptions, 2, 1)
   optCmt = Mid(sOptions, 3, 1)
   optAsn = Mid(sOptions, 4, 1)
   chkShowBomComments = Mid(sOptions, 5, 1)
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = optDsc & OptExt & optCmt & optAsn & chkShowBomComments
   SaveSetting "Esi2000", "EsiEngr", "BompBMp03a", Trim(sOptions)
End Sub

Private Sub cmbPls_Click()
   FillBomhRev cmbPls
   GetPartRevision
   
End Sub

Private Sub cmbPls_LostFocus()
   cmbPls = CheckLen(cmbPls, 30)
   FillBomhRev cmbPls
   GetPartRevision
   
End Sub


Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   
End Sub




Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbPls.Clear
   sSql = "SELECT DISTINCT PARTREF,PARTNUM FROM " _
          & "BmplTable,PartTable WHERE BMPARTREF=PARTREF " _
          & "ORDER BY PARTREF"
   LoadComboBox cmbPls
   If cmbPls.ListCount > 0 Then cmbPls = cmbPls.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   bOnLoad = 1
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
   SaveCurrentSelections
   FormUnload
   Set BompBMp03a = Nothing
   
End Sub




Private Sub PrintReport()
   Dim sPartNumber As String
   MouseCursor 13
   On Error GoTo Eng01
   sPartNumber = Compress(cmbPls)
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowBomComments"
   aFormulaName.Add "ShowListComments"
   aFormulaName.Add "ShowDescription"
   aFormulaName.Add "ShowExtendedDescription"
   
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkShowBomComments) & "'")
   aFormulaValue.Add optCmt
   aFormulaValue.Add optDsc
   aFormulaValue.Add OptExt

   sCustomReport = GetCustomReport("engbm03")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{BmplTable.BMPARTREF}='" & sPartNumber & "' " _
          & "AND {BmplTable.BMREV}='" & cmbRev & "' "
   If optAsn Then sSql = sSql & "AND {PartTable1.PABOMREV}<>'' "
   
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
   
Eng01:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Eng02
Eng02:
   DoModuleErrors Me
   
End Sub



Private Sub optAsn_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   Dim bGoodRout As Byte
   MouseCursor 13
   bGoodList = GetList()
   If bGoodList = 0 Then
      MouseCursor 0
      MsgBox "Parts List Wasn't Found.", vbExclamation, Caption
      Exit Sub
   Else
      PrintReport
   End If
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Private Function GetList() As Byte
   Dim RdoLst As ADODB.Recordset
   Dim sPartNumber As String
   GetList = 1
   Exit Function
   On Error GoTo DiaErr1
   sPartNumber = Compress(cmbPls)
   sSql = "SELECT BMHREF,BMHREV FROM BmhdTable WHERE BMHREF='" _
          & sPartNumber & "' AND BMHREV='" & Trim(cmbRev) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst)
   If bSqlRows Then
      With RdoLst
         cmbRev = "" & Trim(!BMHREV)
         ClearResultSet RdoLst
      End With
      GetList = 1
   Else
      GetList = 0
   End If
   Set RdoLst = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getlist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub GetPartRevision()
   Dim RdoRev As ADODB.Recordset
   Dim sPartNumber As String
   
   On Error GoTo DiaErr1
   sPartNumber = Compress(cmbPls)
   sSql = "SELECT PARTREF,PADESC,PALEVEL,PABOMREV FROM " _
          & "PartTable WHERE PARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRev)
   If bSqlRows Then
      With RdoRev
         cmbRev = "" & Trim(!PABOMREV)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = "" & Trim(str(!PALEVEL))
         ClearResultSet RdoRev
      End With
   Else
      cmbRev = ""
   End If
   Set RdoRev = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpartrev"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optPrn_Click()
   Dim bGoodRout As Byte
   MouseCursor 13
   bGoodList = GetList()
   If Not bGoodList Then
      MouseCursor 0
      MsgBox "Parts List Wasn't Found.", vbExclamation, Caption
      Exit Sub
   Else
      PrintReport
   End If
   
End Sub
