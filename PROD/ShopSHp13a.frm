VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp13a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Status (Report)"
   ClientHeight    =   2985
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2985
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "3"
      Top             =   960
      Width           =   3255
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   3075
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp13a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optCom 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Tag             =   "4"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   315
      Left            =   5400
      Picture         =   "ShopSHp13a.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part"
      Top             =   960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox cmbPrt1 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Tag             =   "3"
      Top             =   960
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5640
      TabIndex        =   7
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp13a.frx":0AF0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ShopSHp13a.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2985
      FormDesignWidth =   6780
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   6
      Left            =   5280
      TabIndex        =   17
      Top             =   1800
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   1668
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Comments"
      Height          =   288
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1668
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   288
      Index           =   2
      Left            =   3120
      TabIndex        =   13
      Top             =   1800
      Width           =   948
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "With Dates From "
      Height          =   288
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1788
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Of Part Number"
      Height          =   288
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1788
   End
End
Attribute VB_Name = "ShopSHp13a"
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

'*********************************************************************************
' diaGARMAN
'
' Notes: Created for GARMAN
'
' Created: 06/29/04 (nth)
' Revisions:
'5/25/06 (CJS) Rebuilt the entire report, selection criteria and groups
'        Added sum of open MO's
Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "CMBPRT"
      ViewParts.txtPrt = cmbPrt
      ViewParts.Show
   End If
   
End Sub

Private Sub cmbPrt_Click()
   If bCancel = 0 Then GetCurrentPart cmbPrt, lblDsc
   bCancel = 0
   
End Sub

Private Sub cmbPrt_Change()
   If bCancel = 0 Then GetCurrentPart cmbPrt, lblDsc
   bCancel = 0
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = 1
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "CMBPRT"
   ViewParts.txtPrt = cmbPrt
   ViewParts.Show
   
End Sub

Private Sub cmdFnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
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
   If bOnLoad Then
      bOnLoad = 0
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombo cmbPrt
      
      cmbPrt = ""
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   sCurrForm = Caption
   'BackColor =1B1DC2D4D0DC
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
   Set ShopSHp13a = Nothing
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 5) = "*** P" Then lblDsc.ForeColor = ES_RED _
           Else lblDsc.ForeColor = vbBlack
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 5)
   
End Sub

Private Sub PrintReport()
   Dim cRunQtyRemain As Currency
   Dim sCustomReport As String
   Dim sBegDate As String
   Dim sEndDate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   
   If lblDsc.ForeColor = ES_RED Then
      MsgBox "This Report Requires A Valid Part Number.", _
         vbInformation, Caption
      Exit Sub
   End If
   MouseCursor 13
   On Error GoTo DiaErr1
   sProcName = ""
   cRunQtyRemain = GetSumMOQuantity()
   If Not IsDate(txtBeg) Then
      sBegDate = "1995,01,01"
   Else
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEndDate = "2024,12,31"
   Else
      sEndDate = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("prdsh19.rpt")
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "Status='" & Val(optCom) & "'"
'   MDISect.Crw.Formulas(2) = "RequestBy='Requested By: " _
'                        & sInitials & "'"
'   MDISect.Crw.Formulas(3) = "OpenMo='" & Format$(cRunQtyRemain, "######0.000") & "'"
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Status"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "OpenMo"
    aFormulaName.Add "ShowComments"
   
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add optCom.Value
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & CStr(Format$(cRunQtyRemain, "######0.000")) & "'")
    aFormulaValue.Add optCom.Value
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
    sSql = "{SoitTable.ITSCHED} in Date(" & sBegDate & ") to Date(" & sEndDate & ") " _
          & "AND {PartTable.PARTREF}='" & Compress(cmbPrt) & "'" _
          & " AND {SoitTable.ITINVOICE} = 0 and {SoitTable.ITPSNUMBER} = '' "
          
'   If optCom.value = vbChecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.0.0;T;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.0.0;F;;;"
'   End If
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   If sProcName = "" Then sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiProd", "sh19", optCom.Value
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = Trim(GetSetting("Esi2000", "EsiProd", "sh19", sOptions))
   If sOptions = "" Then optCom.Value = vbChecked _
                 Else optCom.Value = sOptions
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Trim(txtBeg) = "" Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then _
      txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Trim(txtEnd) = "" Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then _
      txtEnd = CheckDateEx(txtEnd)
   
End Sub



Private Function GetSumMOQuantity() As Currency
   Dim RdoQty As ADODB.Recordset
   sSql = "select sum(RUNREMAININGQTY) AS RunsRemaining FROM " _
          & "RunsTable WHERE (RUNREF='" & Compress(cmbPrt) & "' " _
          & "AND RUNSTATUS NOT LIKE 'C%')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoQty, ES_FORWARD)
   If bSqlRows Then
      With RdoQty
         If Not IsNull(!RunsRemaining) Then
            GetSumMOQuantity = !RunsRemaining
         Else
            GetSumMOQuantity = 0
         End If
      End With
      ClearResultSet RdoQty
   Else
      GetSumMOQuantity = 0
   End If
   Set RdoQty = Nothing
   If Err > 0 Then sProcName = "getsummoqu"
   
End Function

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

Private Sub txtPrt_LostFocus()
   If bCancel = 0 Then GetCurrentPart txtPrt, lblDsc
   bCancel = 0
End Sub

