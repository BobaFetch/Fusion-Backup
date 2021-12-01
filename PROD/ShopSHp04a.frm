VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Orders By Part"
   ClientHeight    =   3240
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3240
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optStatCode 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2400
      TabIndex        =   30
      Top             =   2775
      Width           =   735
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optSta 
      Caption         =   "RL"
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   2
      ToolTipText     =   "Released"
      Top             =   1600
      Width           =   615
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4560
      TabIndex        =   10
      Tag             =   "4"
      Top             =   1920
      Width           =   1250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Tag             =   "4"
      Top             =   1920
      Width           =   1250
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Manufacturing Orders"
      Top             =   960
      Width           =   3545
   End
   Begin VB.CheckBox optOps 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optQty 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   3105
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2400
      TabIndex        =   11
      Top             =   2535
      Width           =   735
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CA"
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   8
      ToolTipText     =   "Canceled"
      Top             =   1600
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CL"
      Height          =   255
      Index           =   5
      Left            =   6000
      TabIndex        =   7
      ToolTipText     =   "Closed "
      Top             =   1600
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CO"
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Complete"
      Top             =   1600
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PC"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   5
      ToolTipText     =   "Picked Complete"
      Top             =   1600
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PP"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   4
      ToolTipText     =   "Picked Partial"
      Top             =   1600
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PL"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      ToolTipText     =   "Pick List"
      Top             =   1600
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "SC"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Scheduled"
      Top             =   1600
      Width           =   615
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   14
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp04a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   490
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ShopSHp04a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   490
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3240
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show MO Status Code"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   31
      Top             =   2760
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   6000
      TabIndex        =   29
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   9
      Left            =   6000
      TabIndex        =   27
      Top             =   1920
      Width           =   1545
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   26
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   25
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   960
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Operation Information"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Partial Completion Quantities"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Comments"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   21
      Top             =   2520
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Orders From"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   1920
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Status:"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   1600
      Width           =   2025
   End
End
Attribute VB_Name = "ShopSHp04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/28/05 Changed date handling
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress(3) As New EsiKeyBd
Private txtGotFocus(3) As New EsiKeyBd




Private Sub cmbPrt_Click()
   If cmbPrt.ListCount > 0 Then cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt.ListCount > 0 Then cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   If Trim(cmbPrt) = "" Then cmbPrt = "ALL"
   
   
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
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "PartTable,RunsTable WHERE PARTREF=RUNREF " _
          & "ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   Else
      lblDsc = "*** No Part Numbers Recorded ***"
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
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
   If Len(Trim(cmbPrt)) Then cUR.CurrentPart = cmbPrt
   SaveCurrentSelections
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ShopSHp03a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sBegDate As String
   Dim sEndDate As String
   Dim sPart As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sSubSql As String
   
   MouseCursor 13
   CheckOptions
   On Error Resume Next
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
   On Error GoTo DiaErr1
   
   
   If cmbPrt = "ALL" Then sPart = "" Else sPart = Compress(cmbPrt)
   
    If optStatCode.Value Then
      sCustomReport = GetCustomReport("prdsh04b")
    Else
      sCustomReport = GetCustomReport("prdsh04a")
    End If

   'MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowComments"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & txtBeg & " Through " & txtEnd & "...'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add optCmt.Value
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
   sSql = "{RunsTable.RUNREF} LIKE '" & sPart & "*' "

   sSql = sSql & " AND {RunsTable.RUNSCHED} in Date(" & sBegDate & ") to Date(" & sEndDate & ") "
   If optSta(0).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'SC' "
   If optSta(1).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PL' "
   If optSta(2).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PP' "
   If optSta(3).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PC' "
   If optSta(4).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CO' "
   If optSta(5).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CL' "
   If optSta(6).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CA' "
   If optSta(7).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'RL' "
'   If optCmt.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.0.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.0.0;T;;;"
'   End If
'   MDISect.Crw.SelectionFormula = sSql
'
'   cCRViewer.SetReportSelectionFormula sSql
'
'   sSubSql = "{?@MOPart}= {?Pm-RunsTable.RUNREF} and {?@RunNo} = {?Pm-RunsTable.RUNNO}"
'
'    'set the sub sql variable pass the sub report name
'   cCRViewer.SetSubRptSelFormula "GetNextOpWC", sSubSql
   
   aRptPara.Add CStr("255T45033")
   aRptPara.Add CStr("3")
'
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("Int")
'
'   ' set the sub sql variable pass the sub report name
   cCRViewer.SetSubRptDBParameters "GetNextOpWC", aRptPara, aRptParaType


   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   
'   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FormatControls()
   Set txtGotFocus(0).esCmbGotfocus = txtBeg
   Set txtGotFocus(1).esCmbGotfocus = txtEnd
   Set txtGotFocus(2).esCmbGotfocus = cmbPrt
   
   Set txtKeyPress(0).esCmbKeyDate = txtBeg
   Set txtKeyPress(1).esCmbKeyDate = txtEnd
   Set txtKeyPress(2).esCmbKeyCase = cmbPrt
'   txtBeg = Format(ES_SYSDATE, "mm/01/yy")
'   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtBeg = ""
   txtEnd = ""
   
End Sub

Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   Dim sPart As String * 30
   sPart = cmbPrt
   'Save by Menu Option
   For iList = 0 To 5
      sOptions = sOptions & Trim(str(optSta(iList).Value))
   Next
   sOptions = sOptions & Trim(str(optSta(iList).Value))
   sOptions = sOptions & Trim(str(optCmt.Value))
   sOptions = sOptions & sPart
   SaveSetting "Esi2000", "EsiProd", "sh04", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh04", sOptions)
   If Len(sOptions) > 0 Then
      For iList = 1 To 6
         optSta(iList - 1) = Mid$(sOptions, iList, 1)
      Next
      optSta(iList - 1) = Mid$(sOptions, iList, 1)
      optCmt.Value = Val(Mid(sOptions, iList + 1, 1))
   End If
   
End Sub



Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      'lblDsc.ForeColor = ES_RED
      lblDsc = "*** ALL PARTS ***"
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optOps_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub optQty_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optSta_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) > 3 Then
      txtBeg = CheckDateEx(txtBeg)
   Else
      txtBeg = "ALL"
   End If
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) > 3 Then
      txtEnd = CheckDateEx(txtEnd)
   Else
      txtEnd = "ALL"
   End If
   
End Sub



Private Sub CheckOptions()
   Dim bByte As Byte
   Dim iList As Integer
   
   For iList = 0 To 5
      If optSta(iList).Value = vbChecked Then
         bByte = True
         Exit For
      End If
   Next
   If optSta(iList).Value = vbChecked Then bByte = True
   
   If Not bByte Then
      For iList = 0 To 6
         optSta(iList).Value = vbChecked
      Next
   End If
   
End Sub