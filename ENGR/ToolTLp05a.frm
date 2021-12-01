VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ToolTLp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tool Lists Used On Manufacturing Orders"
   ClientHeight    =   3120
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3120
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLp05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Tag             =   "4"
      Top             =   2040
      Width           =   1250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4320
      TabIndex        =   3
      Tag             =   "4"
      Top             =   2040
      Width           =   1250
   End
   Begin VB.ComboBox cmbMos 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains MO Operations With Tool Lists (Not Closed Or Canceled)"
      Top             =   1080
      WhatsThisHelpID =   100
      Width           =   3345
   End
   Begin VB.ComboBox cmbLst 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Tool Lists"
      Top             =   1560
      WhatsThisHelpID =   100
      Width           =   3345
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   2640
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ToolTLp05a.frx":07AE
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
         Picture         =   "ToolTLp05a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   3120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3120
      FormDesignWidth =   7335
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   7
      Left            =   3120
      TabIndex        =   16
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scheduled Start From"
      Height          =   315
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Orders"
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool List(s)"
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   5
      Left            =   5880
      TabIndex        =   12
      Top             =   1560
      Width           =   1404
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5880
      TabIndex        =   10
      Top             =   1080
      Width           =   1404
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Tag             =   " "
      Top             =   2400
      Width           =   1425
   End
End
Attribute VB_Name = "ToolTLp05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT OPREF,OPTOOLLIST,RUNREF,RUNSTATUS," _
          & "PARTREF,PARTNUM FROM RnopTable,RunsTable,PartTable " _
          & "WHERE (OPREF=RUNREF AND RUNREF=PARTREF AND OPTOOLLIST<>'' " _
          & "AND RUNSTATUS NOT LIKE 'C%') ORDER BY RUNREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         AddComboStr cmbMos.hwnd, "ALL"
         Do Until .EOF
            AddComboStr cmbMos.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   If cmbMos.ListCount > 0 Then cmbMos = cmbMos.List(0) _
                                         Else cmbMos = "ALL"
   
   sSql = "Qry_FillToolListCombo"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      AddComboStr cmbLst.hwnd, "ALL"
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbLst.hwnd, "" & Trim(!TOOLLIST_NUM)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   If cmbLst.ListCount > 0 Then cmbLst = cmbLst.List(0) _
                                         Else cmbLst = "ALL"
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
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

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
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
   FormUnload
   Set ToolTLp05a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sPart As String
   Dim sList As String
   Dim sBDate As String
   Dim sEDate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If cmbMos = "" Then cmbMos = "ALL"
   If cmbLst = "" Then cmbLst = "ALL"
   If cmbMos <> "ALL" Then sPart = Compress(cmbMos)
   If cmbLst <> "ALL" Then sList = Compress(cmbLst)
   If Not IsDate(txtBeg) Then
      sBDate = "1995,01,01"
   Else
      sBDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEDate = "2024,12,31"
   Else
      sEDate = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDetails"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Manufacturing Orders " & CStr(cmbMos & ", Tool Lists " _
                        & cmbLst) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add optDet.value
   sCustomReport = GetCustomReport("engtl05")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = ""
   sSql = "{RnopTable.OPREF} LIKE '" & sPart & "*' AND " _
          & "{TlhdTable.TOOLLIST_REF} LIKE '" & sList & "*' " _
          & "AND ({RunsTable.RUNSCHED} In Date(" & sBDate & ") " _
          & "To Date(" & sEDate & ")) "
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


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
 '  txtBeg = Format(ES_SYSDATE, "mm/dd/yy")
 '  txtEnd = Format(ES_SYSDATE + 30, "mm/dd/yy")
   txtBeg = ""
   txtEnd = ""
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiEngr", "tl05", Trim(optDet.value)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   sOptions = Trim(GetSetting("Esi2000", "EsiEngr", "tl05", sOptions))
   On Error Resume Next
   If Trim(sOptions) <> "" Then optDet.value = Left(sOptions, 1)
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub
