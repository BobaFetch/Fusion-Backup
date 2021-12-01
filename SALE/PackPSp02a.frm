VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Packing Slip Log"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7140
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtEdt 
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox txtBdt 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1700
      Width           =   1215
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   14
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PackPSp02a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PackPSp02a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   490
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.CheckBox optAdr 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtEps 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1340
      Width           =   915
   End
   Begin VB.TextBox txtBps 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   1020
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   7140
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   195
      Index           =   6
      Left            =   5880
      TabIndex        =   15
      Top             =   1680
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(If Different)"
      Height          =   195
      Index           =   5
      Left            =   3600
      TabIndex        =   12
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Addresses"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   10
      Top             =   1700
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Beginning Date"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1700
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending PS Number"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting PS Number"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1020
      Width           =   1905
   End
End
Attribute VB_Name = "PackPSp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'New 10/2/03
Option Explicit

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   GetPsDates
End Sub

'10/2/03

Private Sub GetPsDates()
   Dim RdoGdt As ADODB.Recordset
   sSql = "SELECT MIN(PSPRINTED) FROM PshdTable WHERE PSPRINTED " _
          & "IS NOT NULL"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGdt, ES_FORWARD)
   If bSqlRows Then
      txtBdt = Format(RdoGdt.Fields(0), "mm/dd/yyyy")
   Else
      txtEdt = Format(ES_SYSDATE, "mm/dd/yyyy")
   End If
   txtEdt = Format(ES_SYSDATE, "mm/dd/yyyy")
   
   Set RdoGdt = Nothing
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "Sh02", sOptions)
   If Len(sOptions) > 0 Then
      'txtBdt = Left(sOptions, 8)
      'txtEdt = Mid(sOptions, 9, 8)
      optAdr.Value = Val(Mid(sOptions, 17, 1))
   End If
   txtBps = "ALL"
   txtEps = "ALL"
   
End Sub


Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = txtBdt _
              & txtEdt _
              & RTrim(optAdr.Value)
   SaveSetting "Esi2000", "EsiSale", "sh02", Trim(sOptions)
   
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
   MdiSect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
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
   FormUnload
   Set PackPSp02a = Nothing
   
End Sub


Private Sub PrintReport()
   Dim sBegDt As String
   Dim sBegPs As String
   Dim sEndDt As String
   Dim sEndPs As String
   Dim sBDate As String
   Dim sEDate As String
   
   If Trim(txtBps) = "" Then txtBps = "ALL"
   If Trim(txtEps) = "" Then txtEps = "ALL"
   
   If Len(txtBps) = 0 Or txtBps = "ALL" Then
      sBegPs = ""
      txtBps = "ALL"
   Else
      sBegPs = txtBps
   End If
   If Len(txtEps) = 0 Or txtEps = "ALL" Then
      sEndPs = "zzz"
      txtEps = "ALL"
   Else
      sEndPs = txtEps
   End If
   If IsDate(txtBdt) Then
      sBegDt = Format(txtBdt, "yyyy,mm,dd")
      sBDate = txtBdt
   Else
      sBegDt = "1995,01,01"
      sBDate = "ALL"
   End If
   
   If IsDate(txtEdt) Then
      sEndDt = Format(txtEdt, "yyyy,mm,dd")
      sEDate = txtEdt
   Else
      sEndDt = "2024,12,31"
      sEDate = "ALL"
   End If
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
 
   
    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("sleps02")
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = "sleps02"
    cCRViewer.ShowGroupTree False

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtBps) & CStr(txtEps) & "... And " _
                            & "Dates From " & CStr(sBDate) & " To " & CStr(sEDate) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    
 
   sSql = "{PshdTable.PSNUMBER} in '" & sBegPs & "' to '" _
          & "" & sEndPs & "'" _
          & "AND {PshdTable.PSPRINTED} in Date(" & sBegDt _
           & ") to Date(" & sEndDt & ")"
    
    cCRViewer.SetReportSelectionFormula sSql
   
   If optAdr.Value = vbUnchecked Then
      cCRViewer.SetReportSection "GroupFooterSection1", True
      cCRViewer.SetReportSection "GroupFooterSection2", True
   Else
      cCRViewer.SetReportSection "GroupFooterSection1", False
      cCRViewer.SetReportSection "GroupFooterSection2", False
   End If
   
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
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


Private Sub optAdr_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optAdr_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   MouseCursor 11
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   MouseCursor 11
   PrintReport
   
End Sub


Private Sub txtBdt_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBdt_LostFocus()
   If Len(Trim(txtBdt)) = 0 Then txtBdt = "ALL"
   If txtBdt <> "ALL" Then txtBdt = CheckDateEx(txtBdt)
   
End Sub

Private Sub txtBps_LostFocus()
   txtBps = CheckLen(txtBps, 8)
   If Len(txtBps) = 0 Then txtBps = "ALL"
   
End Sub

Private Sub txtEdt_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEdt_LostFocus()
   If Len(Trim(txtEdt)) = 0 Then txtEdt = "ALL"
   If Trim(txtEdt) <> "ALL" Then txtEdt = CheckDate(txtEdt)
   
End Sub

Private Sub txtEps_LostFocus()
   txtEps = CheckLen(txtEps, 8)
   If Len(txtEps) = 0 Then txtEps = "ALL"
   
End Sub
