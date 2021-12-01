VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Orders By Date"
   ClientHeight    =   4635
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7320
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4635
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optStatCode 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2400
      TabIndex        =   32
      Top             =   3255
      Width           =   735
   End
   Begin VB.ComboBox cmbSortBy 
      Height          =   315
      ItemData        =   "ShopSHp03a.frx":0000
      Left            =   2400
      List            =   "ShopSHp03a.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp03a.frx":003D
      Style           =   1  'Graphical
      TabIndex        =   29
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
      TabIndex        =   1
      ToolTipText     =   "Released"
      Top             =   1080
      Width           =   615
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2400
      TabIndex        =   8
      Tag             =   "4"
      Top             =   1560
      Width           =   1250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4440
      TabIndex        =   9
      Tag             =   "4"
      Top             =   1560
      Width           =   1250
   End
   Begin VB.CheckBox optOps 
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optQty 
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   3465
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2400
      TabIndex        =   12
      Top             =   3015
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CA"
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   7
      ToolTipText     =   "Canceled"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CL"
      Height          =   255
      Index           =   5
      Left            =   6000
      TabIndex        =   6
      ToolTipText     =   "Closed"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CO"
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   5
      ToolTipText     =   "Complete"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PC"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   4
      ToolTipText     =   "Picked Complete"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PP"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "Picked Partial"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PL"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      ToolTipText     =   "Pick List"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "SC"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      ToolTipText     =   "Scheduled"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   15
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp03a.frx":07EB
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ShopSHp03a.frx":0969
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
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
      FormDesignHeight=   4635
      FormDesignWidth =   7320
   End
   Begin VB.Label lblDateType 
      Caption         =   "(Scheduled Date)"
      Height          =   375
      Left            =   2400
      TabIndex        =   34
      Top             =   2000
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show MO Status Code"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   33
      Top             =   3240
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Sort Report by"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   9
      Left            =   5880
      TabIndex        =   28
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Operation Information"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Partial Completion Quantities"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Comments"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   25
      Top             =   3000
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   24
      Top             =   2760
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   21
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Orders From"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   1560
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Status:"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   1080
      Width           =   2025
   End
End
Attribute VB_Name = "ShopSHp03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/28/05 Changed date selections
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress(2) As New EsiKeyBd
Private txtGotFocus(2) As New EsiKeyBd


Private Sub cmbSortBy_Click()
   lblDateType = "(Scheduled Date)"
   If (cmbSortBy.ListIndex = 0) Then
      lblDateType = "(Complete Date)"
   End If
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
   Show
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
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
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

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
   
    If optStatCode.Value Then
      sCustomReport = GetCustomReport("prdsh03b")
    Else
      sCustomReport = GetCustomReport("prdsh03a")
    End If
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowComments"
    aFormulaName.Add "ShowExDescription"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "SortBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtBeg) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    'aFormulaValue.Add CStr("'" & CStr(optCmt.value) & "'")
    aFormulaValue.Add optCmt.Value
    aFormulaValue.Add optExt.Value
    aFormulaValue.Add optDsc.Value
    
    aFormulaValue.Add CStr(cmbSortBy.ListIndex)
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
   sSql = ""
   
   If (cmbSortBy.ListIndex = 0) Then
      sSql = "{RunsTable.RUNSCHED} in Date(" & sBegDate & ") to Date(" & sEndDate & ") "
   Else
      sSql = "{RunsTable.RUNSTART} in Date(" & sBegDate & ") to Date(" & sEndDate & ") "
   End If
   
   If optSta(0).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'SC' "
   If optSta(1).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PL' "
   If optSta(2).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PP' "
   If optSta(3).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PC' "
   If optSta(4).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CO' "
   If optSta(5).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CL' "
   If optSta(6).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CA' "
   If optSta(7).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'RL' "
   
'   MDISect.Crw.SelectionFormula = sSql
 '  SetCrystalAction Me
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
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Set txtGotFocus(0).esCmbGotfocus = txtBeg
   Set txtGotFocus(1).esCmbGotfocus = txtEnd
   
   Set txtKeyPress(0).esCmbKeyDate = txtBeg
   Set txtKeyPress(1).esCmbKeyDate = txtEnd
'   txtBeg = Format(ES_SYSDATE, "mm/01/yy")
'   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtBeg = ""
   txtEnd = ""
   
End Sub

Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   
   'Save by Menu Option
   For iList = 0 To 6
      sOptions = sOptions & Trim(str(optSta(iList).Value))
   Next
   sOptions = sOptions & Trim(str(optSta(iList).Value))
   sOptions = sOptions & Trim(str(optDsc.Value)) _
              & Trim(str(optExt.Value)) & Trim(str(optCmt.Value))
   sOptions = sOptions & Trim(str(cmbSortBy.ListIndex))
   
   SaveSetting "Esi2000", "EsiProd", "sh03", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   
   On Error Resume Next
   'Get By Menu Option
   sOptions = GetSetting("Esi2000", "EsiProd", "sh03", sOptions)
   If Len(sOptions) > 0 Then
      For iList = 1 To 7
         optSta(iList - 1) = Mid$(sOptions, iList, 1)
      Next
      optSta(iList - 1) = Mid$(sOptions, iList, 1)
      optDsc.Value = Val(Mid(sOptions, iList + 1, 1))
      optExt.Value = Val(Mid(sOptions, iList + 2, 1))
      optCmt.Value = Val(Mid(sOptions, iList + 3, 1))
      cmbSortBy.ListIndex = Val(Mid(sOptions, iList + 4, 1))
   Else
      For iList = 1 To 7
         optSta(iList).Value = vbChecked
      Next
      cmbSortBy.ListIndex = 0
   End If
   
End Sub



Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
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



Private Sub CheckOptions()
   Dim bByte As Byte
   Dim iList As Integer
   
   For iList = 0 To 6
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
