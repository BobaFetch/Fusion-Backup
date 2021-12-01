VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order Log By Date"
   ClientHeight    =   2430
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRp05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PurcPRp05a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PurcPRp05a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox optCls 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5760
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   1800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2430
      FormDesignWidth =   6885
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   9
      Left            =   5400
      TabIndex        =   12
      Top             =   1080
      Width           =   1400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By Vendor Type"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Skip Closed PO's"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "PurcPRp05a"
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
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "pr05", sOptions)
   If Len(sOptions) > 0 Then
      optCls.Value = Val(Left(sOptions, 1))
      optTyp.Value = Val(Mid(sOptions, 2, 1))
   End If
   If Trim(txtEnd) = "" Then txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   If Trim(txtBeg) = "" Then txtBeg = "01/01/" & Right(txtEnd, 4)
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sBeg As String * 8
   Dim sEnd As String * 8
   sBeg = txtBeg
   sEnd = txtEnd
   'Save by Menu Option
   sOptions = RTrim(optCls.Value) _
              & RTrim(optTyp.Value) _
              & sBeg & sEnd
   SaveSetting "Esi2000", "EsiProd", "pr05", Trim(sOptions)
   
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
   MDISect.lblBotPanel = Caption
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
   Set PurcPRp05a = Nothing
   
End Sub
Private Sub PrintReport()
    Dim sLastPo As String
    Dim sBegDte As String
    Dim sEndDte As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   If Not IsDate(txtBeg) Then
      sBegDte = "1995,01,01"
   Else
      sBegDte = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEndDte = "2024,12,31"
   Else
      sEndDte = Format(txtEnd, "yyyy,mm,dd")
   End If
   On Error GoTo Ppr04

    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaValue.Add CStr("'Purchase Orders From" & txtBeg & " Ending " & txtEnd & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
   
   If optTyp.Value = vbChecked Then
      sCustomReport = GetCustomReport("prdpr05a")
   '   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
   Else
      sCustomReport = GetCustomReport("prdpr05")
'      MDISect.Crw.ReportFileName = sReportPath & sCustomReport
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
   End If
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{PohdTable.PODATE} in Date(" & sBegDte _
          & ") to Date(" & sEndDte & ")"
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
Ppr04:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub optCls_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optCls_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub optTyp_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optTyp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub
