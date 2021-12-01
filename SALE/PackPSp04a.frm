VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSp04a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipped Items By Part Number"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6405
   Visible         =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "PackPSp04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSp04a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "PackPSp04a.frx":0938
      Height          =   315
      Left            =   4680
      Picture         =   "PackPSp04a.frx":0C7A
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   960
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "3"
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox cmbcde 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Tag             =   "8"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CheckBox optDol 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox optIco 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox cmbStartDate 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cmbEndDate 
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   11
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PackPSp04a.frx":0FBC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   490
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PackPSp04a.frx":1146
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5280
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   3720
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4380
      FormDesignWidth =   6405
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   27
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   252
      Index           =   12
      Left            =   240
      TabIndex        =   25
      Top             =   2520
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5280
      TabIndex        =   22
      Top             =   1680
      Width           =   1188
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   11
      Left            =   5280
      TabIndex        =   21
      Top             =   2160
      Width           =   1188
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   20
      Top             =   2160
      Width           =   1545
   End
   Begin VB.Label z1 
      Caption         =   "Sort Parts By Highest Dollar"
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   19
      Top             =   3240
      Width           =   2052
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipments From"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   1725
      Width           =   1305
   End
   Begin VB.Label z1 
      Caption         =   "Ext. Descriptions"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Width           =   1452
   End
   Begin VB.Label z1 
      Caption         =   "Item Comments"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   3000
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   15
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   252
      Index           =   1
      Left            =   5280
      TabIndex        =   14
      Top             =   960
      Width           =   1092
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "PackPSp04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of                     ***
'*** ESI Software Engineering Inc, Seattle, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***
'See the UpdateTables prodecure for database revisions
' PackPSp04a - Shipped Items by part number
'
' Created: 12/31/03 (JCW)
' Revisions:
' 01/22/04 (JCW) Fixed Spelling/Layout/naming struct./report names
' 4/27/05 Removed Combo (CJS)
' 3/23/06 Corrected Formulae and added groupings
' 5/10/06 See Print Report
Option Explicit
Dim bOnLoad As Byte
Private txtKeyPress(3) As New EsiKeyBd
Private txtGotFocus(3) As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbcde_LostFocus()
   If Trim(cmbcde) = "" Then cmbcde = "ALL"
   
End Sub


Private Sub cmbPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "CMBPRT"
      ViewParts.txtPrt = cmbPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   cmbPrt = txtPrt
   cmbPrt_LostFocus
End Sub

Private Sub cmbPrt_LostFocus()
   Dim sOldPart As String
   cmbPrt = CheckLen(cmbPrt, 30)
   sOldPart = cmbPrt
   If cmbPrt = "" Then cmbPrt = "ALL"
   If Trim(cmbPrt) <> "ALL" Then
      cmbPrt = CheckLen(cmbPrt, 30)
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
      If lblDsc.ForeColor = ES_RED Then lblDsc = ""
      If cmbPrt = "" Then cmbPrt = sOldPart
   Else
      lblDsc = "Range Of Parts Selected."
      cmbPrt = "ALL"
   End If
   
End Sub

Private Sub cmbPrt_Click()
   Dim sOldPart As String
   cmbPrt = CheckLen(cmbPrt, 30)
   sOldPart = cmbPrt
   If cmbPrt = "" Then cmbPrt = "ALL"
   If Trim(cmbPrt) <> "ALL" Then
      cmbPrt = CheckLen(cmbPrt, 30)
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
      If lblDsc.ForeColor = ES_RED Then lblDsc = ""
      If cmbPrt = "" Then cmbPrt = sOldPart
   Else
      lblDsc = "Range Of Parts Selected."
      cmbPrt = "ALL"
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "CMBPRT"
   ViewParts.txtPrt = cmbPrt
   optVew.Value = vbChecked
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
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      'Fillcombo
      FillProductCodes
      GetOptions
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombo cmbPrt
      
      FormatControls
     bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub txtPrt_Change()
   cmbPrt = txtPrt
   
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   'FormatControls
   cmbStartDate = Format(Now, "mm/01/yyyy")
   cmbEndDate = Format(Now, "mm/dd/yyyy")
   bOnLoad = 1
   'GetOptions
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PackPSp04a = Nothing
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
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   If cmbcde = "" Then cmbcde = "ALL"
   If cmbPrt = "" Then cmbPrt = "ALL"
   If cmbPrt <> "ALL" Then sPart = Compress(cmbPrt)
   If IsDate(cmbStartDate) Then
      sBegDate = Format(cmbStartDate, "yyyy,mm,dd")
   Else
      sBegDate = "1995,01,01"
   End If
   
   If IsDate(cmbEndDate) Then
      sEndDate = Format(cmbEndDate, "yyyy,mm,dd")
   Else
      sEndDate = "2024,12,31"
   End If
   On Error GoTo DiaErr1
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    aFormulaName.Add "Part"
    aFormulaName.Add "ProductClass"
    aFormulaName.Add "ShowDollars"
    aFormulaName.Add "ExtDesc"
    aFormulaName.Add "Ico"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'Shipped Items By Part'")
    aFormulaValue.Add CStr("'Shipped From " & CStr(cmbStartDate & "  Through " & cmbEndDate) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbPrt) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbcde) & "'")
    aFormulaValue.Add CStr("'" & CStr(optDol.Value) & "'")
    aFormulaValue.Add CStr("'" & CStr(optExt.Value) & "'")
    aFormulaValue.Add CStr("'" & CStr(optIco.Value) & "'")
   
   If optDol.Value = vbChecked Then
      sCustomReport = GetCustomReport("sleSh01a.rpt")
   Else
      sCustomReport = GetCustomReport("sleSh01b.rpt")
   End If
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{SoitTable.ITCANCELED} = 0 AND " _
          & "({CihdTable.INVNO} <> 0 OR trim(cstr({PsitTable.PIPACKSLIP})) <> '') " _
          & "and trim(cstr({@Shipped})) <> '' AND {@shipped} In Date('" _
          & sBegDate & "') To Date('" & sEndDate & "') " _
          & " AND (LEN(LTRIM({SoitTable.ITPSNUMBER})) <> 0)"
   sSql = sSql & " AND {PartTable.PARTREF} Like '" & sPart & "*' "
   
   If Trim(cmbcde) <> "ALL" Then _
           sSql = sSql & " AND {PartTable.PAPRODCODE} Like '" & cmbcde & "*'"
   
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
   
   Set txtGotFocus(0).esCmbGotfocus = cmbStartDate
   Set txtGotFocus(1).esCmbGotfocus = cmbEndDate
   Set txtGotFocus(2).esCmbGotfocus = cmbPrt
   
   Set txtKeyPress(0).esCmbKeyDate = cmbStartDate
   Set txtKeyPress(1).esCmbKeyDate = cmbEndDate
   Set txtKeyPress(2).esCmbKeyCase = cmbPrt
   cmbPrt = "ALL"
   cmbcde = "ALL"
   lblDsc = "Range Of Parts Selected."
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub




Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub cmbStartDate_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub cmbStartDate_LostFocus()
   If Len(Trim(cmbStartDate)) = 0 Then cmbStartDate = "ALL"
   If cmbStartDate <> "ALL" Then cmbStartDate = CheckDateEx(cmbStartDate)
   
End Sub


Private Sub cmbEndDate_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub cmbEndDate_LostFocus()
   If Len(Trim(cmbEndDate)) = 0 Then cmbEndDate = "ALL"
   If Trim(cmbEndDate) <> "ALL" Then cmbEndDate = CheckDateEx(cmbEndDate)
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(optExt.Value) _
              & RTrim(optIco.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & "_Printer", lblPrinter
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optExt.Value = Val(Mid(sOptions, 1, 1))
      optIco.Value = Val(Mid(sOptions, 2, 1))
   Else
      optExt.Value = vbUnchecked
      optIco.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & "_Printer", lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
   
End Sub

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

