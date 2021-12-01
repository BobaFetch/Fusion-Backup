VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routing Report"
   ClientHeight    =   2775
   ClientLeft      =   2115
   ClientTop       =   1155
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optShowToollist 
      Height          =   195
      Left            =   2040
      TabIndex        =   14
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "RoutRTp01a.frx":07AE
      Height          =   340
      Left            =   5520
      Picture         =   "RoutRTp01a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Parts Assigned To This Routing"
      Top             =   980
      Width           =   350
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "RoutRTp01a.frx":15FA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "RoutRTp01a.frx":1778
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox optDsc 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   645
   End
   Begin VB.CheckBox OptCmt 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   645
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   2070
      TabIndex        =   0
      Tag             =   "3"
      Top             =   990
      Width           =   3345
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2775
      FormDesignWidth =   7095
   End
   Begin VB.Label Label1 
      Caption         =   "Tooling Lists?"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1785
   End
   Begin VB.Label txtDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2070
      TabIndex        =   10
      Top             =   1340
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Descriptions?"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing Comments?"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   990
      Width           =   1815
   End
End
Attribute VB_Name = "RoutRTp01a"
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
Dim bGoodRout As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   
End Sub

Private Sub cmbRte_Click()
   bGoodRout = GetRout()
   
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
      If Len(sCurrRout) Then cmbRte = sCurrRout
      bGoodRout = GetRout()
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
   If bGoodRout Then
      sCurrRout = cmbRte
      SaveSetting "Esi2000", "EsiEngr", "CurrentRouting", Trim(sCurrRout)
   Else
      sCurrRout = ""
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set RoutRTp01a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sRout As String
   sRout = Compress(cmbRte)
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
    Dim sCustomReport As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    Dim strIncludes As String
    Dim strRequestBy As String
   
    sCustomReport = GetCustomReport("engrt01")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowPartDesc"
    aFormulaName.Add "ShowCmt"
    aFormulaName.Add "ShowToolList"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")

    strRequestBy = "'Requested By: " & sInitials & "'"
    aFormulaValue.Add CStr(strRequestBy)
   
    aFormulaValue.Add CStr(optDsc)
    aFormulaValue.Add CStr(OptCmt)
    aFormulaValue.Add CStr(optShowToollist)
   
   sSql = "{RthdTable.RTREF}='" & sRout & "' "
   
    ' Set Formula values
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.SetReportDistinctRecords True
    ' set the report Selection
    cCRViewer.SetReportSelectionFormula (sSql)
    cCRViewer.CRViewerSize Me
    
    ' Set report parameter
    cCRViewer.SetDbTableConnection


    cCRViewer.OpenCrystalReportObject Me, aFormulaName

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aRptParaType
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

Private Function GetRout() As Byte
   Dim RdoRte As ADODB.Recordset
   Dim sReportRout As String
   
   On Error GoTo DiaErr1
   sReportRout = Compress(cmbRte)
   
   sSql = "Qry_GetToolRout '" & sReportRout & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         GetRout = True
         cmbRte = "" & Trim(!RTNUM)
         txtDsc = "" & Trim(!RTDESC)
         ClearResultSet RdoRte
      End With
   Else
      GetRout = False
      txtDsc = ""
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
   Dim bGoodRout As Byte
   MouseCursor 13
   bGoodRout = GetRout()
   If Not bGoodRout Then
      MouseCursor 0
      MsgBox "Routing Wasn't Found.", vbExclamation, Caption
      Exit Sub
   Else
      PrintReport
   End If
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   MouseCursor 13
   bGoodRout = GetRout()
   If Not bGoodRout Then
      MouseCursor 0
      MsgBox "Routing Wasn't Found.", vbExclamation, Caption
      Exit Sub
   Else
      PrintReport
   End If
   
End Sub



Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "rt01", sOptions)
   If Len(sOptions) > 0 Then
      optDsc.value = Val(Left(sOptions, 1))
      OptCmt.value = Val(Mid(sOptions, 2, 1))
      If Len(sOptions) < 3 Then optShowToollist.value = 0 Else optShowToollist.value = Val(Mid(sOptions, 3, 1))
   End If
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(optDsc.value) _
              & RTrim(OptCmt.value) & RTrim(optShowToollist.value)
   SaveSetting "Esi2000", "EsiEngr", "rt01", Trim(sOptions)
   
End Sub
