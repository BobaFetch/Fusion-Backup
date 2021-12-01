VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PurcPRp08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order History By Vendor"
   ClientHeight    =   5490
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7230
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5490
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbFilterBy 
      Height          =   315
      ItemData        =   "PurcPRp08a.frx":0000
      Left            =   2160
      List            =   "PurcPRp08a.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Tag             =   "9"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CheckBox optShowMonthYearTotals 
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      ToolTipText     =   "Display a Subtotal by Due Date per Vendor"
      Top             =   2880
      Width           =   495
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Enter Product Class to Print (4 Characters)"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Enter Product Code to Print (6 Characters)"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRp08a.frx":0026
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtPrt 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Leading Character Search Or Select Contains Part Numbers On Purchase Orders"
      Top             =   1080
      Width           =   3540
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   1560
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "PurcPRp08a.frx":07D4
      Height          =   315
      Left            =   6720
      Picture         =   "PurcPRp08a.frx":0B16
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   3960
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optT16 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   10
      Top             =   4440
      Width           =   735
   End
   Begin VB.CheckBox optT17 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   9
      Top             =   4200
      Width           =   735
   End
   Begin VB.CheckBox optT15 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   3960
      Width           =   735
   End
   Begin VB.CheckBox optT14 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   3720
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Tag             =   "4"
      ToolTipText     =   "Ending Date the Items are Due"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CheckBox optItm 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   5160
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   4920
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   4680
      Width           =   735
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Tag             =   "4"
      ToolTipText     =   "Starting Date the Items are Due"
      Top             =   2520
      Width           =   1215
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
         Picture         =   "PurcPRp08a.frx":0E58
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
         Picture         =   "PurcPRp08a.frx":0FD6
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Tag             =   "3"
      Top             =   360
      Width           =   1555
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   3120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5490
      FormDesignWidth =   7230
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter/Sort by"
      Height          =   285
      Index           =   20
      Left            =   240
      TabIndex        =   42
      Top             =   2160
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal by Month/Year"
      Height          =   285
      Index           =   17
      Left            =   240
      TabIndex        =   41
      Top             =   2880
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   19
      Left            =   4200
      TabIndex        =   40
      Top             =   1800
      Width           =   1350
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   18
      Left            =   4200
      TabIndex        =   39
      Top             =   1440
      Width           =   1350
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Class"
      Height          =   285
      Index           =   16
      Left            =   240
      TabIndex        =   38
      Top             =   1800
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   285
      Index           =   15
      Left            =   240
      TabIndex        =   37
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label lblVEName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   36
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   14
      Left            =   5880
      TabIndex        =   34
      Top             =   2520
      Width           =   1350
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   13
      Left            =   5880
      TabIndex        =   33
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   30
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Canceled items"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   29
      Top             =   4440
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoiced Items"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   28
      Top             =   4200
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Received Items"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   27
      Top             =   3960
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open Items"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   26
      Top             =   3720
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   25
      Top             =   5160
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   4920
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   23
      Top             =   4680
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   22
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   21
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   2
      Left            =   3600
      TabIndex        =   20
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From "
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   1305
   End
End
Attribute VB_Name = "PurcPRp08a"
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
Dim bGoodType As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbCde_LostFocus()
    cmbCde = CheckLen(cmbCde, 6)
    If Len(Trim(cmbCde)) = 0 Then cmbCde = "ALL"
End Sub

Private Sub cmbCls_LostFocus()
    cmbCls = CheckLen(cmbCls, 4)
    If Len(Trim(cmbCls)) = 0 Then cmbCls = "ALL"
End Sub

Private Sub cmbVnd_Click()
   GetThisVendor
   
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If cmbVnd = "" Then cmbVnd = "ALL"
   GetThisVendor
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdFnd_Click()
   optVew.Value = vbChecked
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   ViewParts.Show
   
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
   If bOnLoad Then
      AddComboStr cmbVnd.hwnd, "ALL"
      FillCombo
      
      'cmbCde.AddItem "ALL"
      FillProductCodes
      'cmbCls.AddItem "ALL"
      FillProductClasses
      'cmbCde = "ALL"
      'cmbCls = "ALL"
     
      
      FillVendors
      GetThisVendor
      cmbVnd = "ALL"
      GetOptions
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   'GetOptions
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
   Set PurcPRp08a = Nothing
   
End Sub

Private Sub PrintReport()
    Dim sBeg As String
    Dim sEnd As String
    Dim sPart As String
    Dim sType As String
    Dim sVendor As String
    
    Dim sCode As String
    Dim sClass As String

    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Not IsDate(txtBeg) Then
      sBeg = "1995,01,01"
   Else
      sBeg = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEnd = "2024,12,31"
   Else
      sEnd = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   sPart = Compress(txtPrt)
   If txtPrt = "ALL" Then sPart = ""
   
   'If Trim(cmbVnd) = "" Then cmbVnd = "ALL"
   If cmbVnd <> "ALL" Then sVendor = Compress(cmbVnd) Else sVendor = ""
   
   If cmbCde = "ALL" Then sCode = "" Else sCode = Compress(cmbCde)
   If cmbCls = "ALL" Then sClass = "" Else sClass = Compress(cmbCls)
   
   MouseCursor 13
   On Error GoTo DiaErr1
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowExtendedDescription"
    aFormulaName.Add "ShowItem"
    aFormulaName.Add "ShowMonthYearTotals"
    aFormulaName.Add "SortByPODate"
    
    aFormulaValue.Add CStr("'" & sFacility & "'")
    aFormulaValue.Add CStr("'From " & txtBeg & " To " & txtEnd & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add optDsc.Value
    aFormulaValue.Add optExt.Value
    aFormulaValue.Add optItm.Value
    aFormulaValue.Add optShowMonthYearTotals.Value
    If cmbFilterBy.ListIndex = 1 Then aFormulaValue.Add CStr("1") Else aFormulaValue.Add CStr("0")
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("prdpr10")
 '  MDISect.Crw.ReportFileName = sReportPath & sCustomReport
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    If cmbFilterBy.ListIndex = 0 Then sSql = "{PoitTable.PIPDATE} " Else sSql = "{PohdTable.PODATE} "
    sSql = sSql & " in Date(" & sBeg & ") to Date(" & sEnd & ") "
    
    sSql = sSql & " AND {PartTable.PAPRODCODE} LIKE '" & sCode & "*' "
    sSql = sSql & " AND {PartTable.PACLASS} LIKE '" & sClass & "*' "
    
    sType = "AND ("
   If optT14.Value = vbChecked Then
      sSql = sSql & sType & " {PoitTable.PITYPE}=14 "
      sType = "OR"
   End If
   If optT15.Value = vbChecked Then
      sSql = sSql & sType & " {PoitTable.PITYPE}=15 "
      sType = "OR"
   End If
   If optT17.Value = vbChecked Then
      sSql = sSql & sType & " {PoitTable.PITYPE}=17 "
      sType = "OR"
   End If
   If optT16.Value = vbChecked Then
      sSql = sSql & sType & " {PoitTable.PITYPE}=16) "
   Else
      sSql = sSql & ") "
   End If
   If sVendor <> "" Then
      sSql = sSql & "AND {VndrTable.VEREF}='" & sVendor & "' "
   End If
   If sPart <> "ALL" Then
      sSql = sSql & "AND {PoitTable.PIPART} LIKE '" & sPart & "*' "
   End If
   
'   If optDsc.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(2) = "GROUPFTR.3.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(2) = "GROUPFTR.3.0;T;;;"
'   End If
'   If optExt.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(3) = "GROUPFTR.2.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(3) = "GROUPFTR.2.0;T;;;"
'   End If
'   If optItm.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(4) = "GROUPFTR.1.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(4) = "GROUPFTR.1.0;T;;;"
'   End If
   
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
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtBeg = "01/01/" & Format(ES_SYSDATE, "yyyy")
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = RTrim(optDsc.Value) _
              & RTrim(optExt.Value) _
              & RTrim(optItm.Value) _
              & RTrim(optT14.Value) _
              & RTrim(optT15.Value) _
              & RTrim(optT17.Value) _
              & RTrim(optT16.Value) _
              & Left(RTrim(txtPrt) & Space(30), 30) _
              & Left(RTrim(cmbCde) & Space(6), 6) _
              & Left(RTrim(cmbCls) & Space(4), 4) _
              & Left(RTrim(txtBeg) & Space(8), 10) _
              & Left(RTrim(txtEnd) & Space(8), 10) _
              & Right("0" & LTrim(str(optShowMonthYearTotals.Value)), 1) _
              & cmbFilterBy.ListIndex
              
   SaveSetting "Esi2000", "EsiProd", "pr10", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim iPartLen As Integer
   
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "pr10", sOptions)
   If Len(sOptions) Then
      If Len(sOptions) < 38 Then sOptions = Left(sOptions & Space(38), 38) & "ALL   " & "ALL " & "ALL     " & "ALL     " & "0"
      optDsc.Value = Val(Left(sOptions, 1))
      optExt.Value = Val(Mid(sOptions, 2, 1))
      optItm.Value = Val(Mid(sOptions, 3, 1))
      optT14.Value = Val(Mid(sOptions, 4, 1))
      optT15.Value = Val(Mid(sOptions, 5, 1))
      optT17.Value = Val(Mid(sOptions, 6, 1))
      optT16.Value = Val(Mid(sOptions, 7, 1))
      txtPrt = Trim(Mid(sOptions, 8, 30))
      cmbCde = Trim(Mid(sOptions, 38, 6))
      cmbCls = Trim(Mid(sOptions, 44, 4))
      txtBeg = Trim(Mid(sOptions, 48, 10))
      txtEnd = Trim(Mid(sOptions, 58, 10))
      optShowMonthYearTotals.Value = Val(Mid(sOptions, 68, 1))
      If Len(sOptions) > 68 Then cmbFilterBy.ListIndex = Val(Mid(sOptions, 65, 1)) Else cmbFilterBy.ListIndex = 0
   Else
      cmbFilterBy.ListIndex = 0
   End If
   
   
End Sub

Private Sub optDis_Click()
   bGoodType = CheckTypes()
   If bGoodType Then PrintReport
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optItm_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   bGoodType = CheckTypes()
   If bGoodType Then PrintReport
   
End Sub

Private Sub optT14_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optT15_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optT16_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optT17_KeyPress(KeyAscii As Integer)
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

Private Function CheckTypes() As Byte
   Dim bByte As Byte
   If Len(txtPrt) = 0 Then txtPrt = "ALL"
   If optT14.Value = vbChecked Then bByte = True
   If optT15.Value = vbChecked Then bByte = True
   If optT17.Value = vbChecked Then bByte = True
   If optT16.Value = vbChecked Then bByte = True
   If bByte Then
      CheckTypes = True
   Else
      CheckTypes = False
      MsgBox "Requires At Least One Item Type.", vbInformation, Caption
   End If
   Exit Function
   
DiaErr1:
   sProcName = "checktypes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If txtPrt = "" Then txtPrt = "ALL"
   
End Sub

Private Sub FillCombo()
   'Added 10/16/02
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PIPART,PARTREF,PARTNUM " _
          & "FROM PoitTable,PartTable WHERE PIPART=PARTREF " _
          & "ORDER BY PIPART"
   LoadComboBox txtPrt, 1
   If txtPrt.ListCount > 0 Then txtPrt = txtPrt.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

