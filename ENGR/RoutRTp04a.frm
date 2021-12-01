VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTp04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routings Used On Report"
   ClientHeight    =   2835
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2835
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTp04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "RoutRTp04a.frx":07AE
      Height          =   340
      Left            =   5520
      Picture         =   "RoutRTp04a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Parts Assigned To This Routing"
      Top             =   960
      Width           =   350
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "RoutRTp04a.frx":15FA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "RoutRTp04a.frx":1784
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "3"
      Top             =   960
      Width           =   3375
   End
   Begin VB.CheckBox OptCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   645
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2430
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2835
      FormDesignWidth =   7095
   End
   Begin VB.Label txtDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Description"
      Height          =   285
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Top             =   1800
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing Number"
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   990
      Width           =   1695
   End
End
Attribute VB_Name = "RoutRTp04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbRte_Click()
   GetRout
   
End Sub

Private Sub cmbRte_LostFocus()
   cmbRte = CheckLen(cmbRte, 30)
   
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

Private Sub cmdVew_Click()
   If cmdVew Then
      RteTree.Show
      cmdVew = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillRoutings
      If cmbRte.ListCount > 0 Then GetRout
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   
   FormatControls
   GetOptions
   bOnLoad = 1
   
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
   Set RoutRTp04a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim bGoodRout As Byte
   Dim sRout As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   bGoodRout = GetRout
   If Not bGoodRout Then
      MouseCursor 0
      If Len(Trim(cmbRte)) = 0 Then Exit Sub
      MsgBox "That Routing Wasn't Found.", vbInformation, Caption
      Exit Sub
   End If
   sRout = Compress(cmbRte)
   MouseCursor 13
   On Error GoTo DiaErr1
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("engrt04")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowComments"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Includes " & CStr(cmbRte) & "...'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add OptCmt.value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RthdTable.RTREF}='" & sRout & "' "
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
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Function GetRout()
   Dim RdoRte As ADODB.Recordset
   Dim sRout As String
   sRout = Compress(cmbRte)
   GetRout = False
   If Len(sRout) = 0 Then Exit Function
   On Error Resume Next
   sSql = "Qry_GetToolRout '" & sRout & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte)
   If bSqlRows Then
      With RdoRte
         GetRout = True
         cmbRte = "" & Trim(RdoRte!RTNUM)
         txtDsc = "" & Trim(RdoRte!RTDESC)
         ClearResultSet RdoRte
      End With
   Else
      GetRout = False
      txtDsc = ""
      cmbRte.SetFocus
   End If
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   MouseCursor 13
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   MouseCursor 13
   PrintReport
   
End Sub



Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "rt04", sOptions)
   If Len(sOptions) > 0 Then OptCmt.value = Val(Left(sOptions, 1))
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(OptCmt.value)
   SaveSetting "Esi2000", "EsiEngr", "rt04", Trim(sOptions)
   
End Sub
