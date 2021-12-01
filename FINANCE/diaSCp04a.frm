VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaSCp04a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proposed Vs. Current Standard Cost (Report)"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7335
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   320
      Left            =   4200
      Picture         =   "diaSCp04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "Show BOM Structure"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.ComboBox cmbCls 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Tag             =   "3"
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox cmbCde 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise Product Code (6 Char)"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtVar 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CheckBox optCst 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   3825
      Width           =   855
   End
   Begin VB.CheckBox optZro 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   4580
      Width           =   855
   End
   Begin VB.CheckBox optDesc 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   4080
      Width           =   855
   End
   Begin VB.CheckBox optExt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   4335
      Width           =   855
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   7
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   8
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   9
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   10
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   11
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox ChkTyp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   12
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   20
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaSCp04a.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Print The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaSCp04a.frx":04CC
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Display The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.TextBox cmbPrt1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox optVew 
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   4440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5055
      FormDesignWidth =   7335
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   46
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaSCp04a.frx":064A
      PictureDn       =   "diaSCp04a.frx":0790
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   47
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaSCp04a.frx":08D6
      PictureDn       =   "diaSCp04a.frx":0A1C
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   48
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Class"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   42
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   13
      Left            =   4800
      TabIndex        =   41
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   12
      Left            =   4800
      TabIndex        =   40
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   39
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%     (Blank For All)"
      Height          =   285
      Index           =   10
      Left            =   4440
      TabIndex        =   38
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Items With Greater Variance Than:"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   37
      Top             =   1440
      Width           =   3225
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   4680
      TabIndex        =   36
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Detail"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   35
      Top             =   3825
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parts With Zero QOH"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   34
      Top             =   4575
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Description"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   33
      Top             =   4080
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Ext Description"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   32
      Top             =   4335
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   31
      Top             =   3480
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Part Types:"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   30
      Top             =   3120
      Width           =   1425
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   29
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   28
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   27
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   26
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   25
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   24
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   23
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label zTyp 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   22
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "diaSCp04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaSCp04a -
'
' Notes:
'
' Created: 02/24/04 (JCW)
'
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
End Sub



Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductCodes Me
      FillProductClasses Me
      FillPartCombo cmbPrt
      cmbPrt = ""
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   GetOptions
   bOnLoad = True
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub cmbPrt_Click()
   If Trim(cmbPrt) <> "" Then
      cmbPrt = CheckLen(cmbPrt, 30)
      FindPart Me
   Else
      lblDsc = "*** Multiple Parts Selected ***"
      lblDsc.ForeColor = &H80000012
   End If
   
End Sub

Private Sub cmbPrt_Change()
   If Trim(cmbPrt) <> "" Then
      cmbPrt = CheckLen(cmbPrt, 30)
      FindPart Me
   Else
      lblDsc = "*** Multiple Parts Selected ***"
      lblDsc.ForeColor = &H80000012
   End If
   
End Sub

Private Sub txtVar_LostFocus()
   If Trim(txtVar) <> "" Then
      txtVar = Format(NumberFix(Trim(txtVar)), CURRENCYMASK)
   End If
   
End Sub

Private Sub ChkTyp_GotFocus(Index As Integer)
   zTyp(Index).BorderStyle = 1
   
End Sub

Private Sub ChkTyp_LostFocus(Index As Integer)
   zTyp(Index).BorderStyle = 0
   
End Sub

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 12)
   If Not bValidElement(cmbCde) Then
      cmbCde = ""
   End If
   
End Sub

Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 12)
   If Not bValidElement(cmbCls) Then
      cmbCls = ""
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaSCp04a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim i As Integer
   Dim sInclude As String
   Dim sTitle As String
   Dim sTitle2 As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   On Error GoTo DiaErr1
   
'   SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("finsc04")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
'   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = "1=1 "
   
   If Trim(cmbPrt) <> "" Then
      sSql = sSql & " AND {PartTable.PARTREF} = '" & Compress(cmbPrt) & "' "
      sTitle = sTitle & "Part: " & cmbPrt & " , "
   Else
      sTitle = sTitle & "Part: ALL , "
   End If
   
   If Trim(cmbCde) <> "" Then
      sSql = sSql & " AND {PartTable.PAPRODCODE} = '" & Compress(cmbCde) & "' "
      sTitle = sTitle & "Product Code: " & cmbCde & " , "
   Else
      sTitle = sTitle & "Product Code: ALL , "
   End If
   
   If Trim(cmbCls) <> "" Then
      sSql = sSql & " AND {PartTable.PACLASS} = '" & Compress(cmbCls) & "' "
      sTitle = sTitle & "Product Class: " & cmbCls & " , "
   Else
      sTitle = sTitle & "Product Class: ALL , "
   End If
   
   If Trim(txtVar) <> "" Then
      sSql = sSql & " AND {@VARIANCE} > " & CCur(txtVar) & " "
      sTitle2 = sTitle2 & "Greater Variance Than: " & txtVar & " , "
   End If
   
   sTitle2 = sTitle2 & "Showing Zero QOH Items? "
   If optZro.Value = vbUnchecked Then
      sSql = sSql & " AND {PartTable.PAQOH} > 0 "
      sTitle2 = sTitle2 & "N"
   Else
      sTitle2 = sTitle2 & "Y"
   End If
   
   For i = 0 To 7
      If ChkTyp(i).Value = vbChecked Then
         sInclude = sInclude & Trim(zTyp(i)) & ","
      End If
   Next
   If Trim(sInclude) <> "" Then
      sInclude = Left(sInclude, Len(sInclude) - 1)
      sSql = sSql & " AND {PartTable.PALEVEL} IN[" & sInclude & "] "
   Else
      sInclude = "1,2,3,4,5,6,7,8"
   End If
   
   
'   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.crw.Formulas(1) = "Title1='Proposed Standard Vs. Current Standard Cost'"
'   MdiSect.crw.Formulas(2) = "Title2='" & sTitle & "'"
'   MdiSect.crw.Formulas(3) = "RequestedBy='Requested By: " & sInitials & "'"
'   MdiSect.crw.Formulas(4) = "CostDet='" & CStr(optCst.Value) & "'"
'   MdiSect.crw.Formulas(5) = "Desc='" & CStr(optDesc.Value) & "'"
'   MdiSect.crw.Formulas(6) = "ExtDesc='" & CStr(optExt.Value) & "'"
'   MdiSect.crw.Formulas(7) = "Title3='Part Types: " & sInclude & "'"
'   MdiSect.crw.Formulas(8) = "Title4='" & sTitle2 & "'"
'   MdiSect.crw.SelectionFormula = sSql
'   SetCrystalAction Me
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    aFormulaName.Add "RequestedBy"
    aFormulaName.Add "CostDet"
    aFormulaName.Add "Desc"
    aFormulaName.Add "ExtDesc"
    aFormulaName.Add "Title3"
    aFormulaName.Add "Title4"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Proposed Standard Vs. Current Standard Cost'")
    aFormulaValue.Add CStr("'" & CStr(sTitle) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & CStr(optCst.Value) & "'")
    aFormulaValue.Add CStr("'" & CStr(optDesc.Value) & "'")
    aFormulaValue.Add CStr("'" & CStr(optExt.Value) & "'")
    aFormulaValue.Add CStr("'Part Types: " & CStr(sInclude) & "'")
    aFormulaValue.Add CStr("'" & CStr(sTitle2) & "'")
   
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   
   Exit Sub
   
DiaErr1:
   sProcName = "Printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport1()
   Dim sCustomReport As String
   Dim i As Integer
   Dim sInclude As String
   Dim sTitle As String
   Dim sTitle2 As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("finsc04")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = "1=1 "
   
   If Trim(cmbPrt) <> "" Then
      sSql = sSql & " AND {PartTable.PARTREF} = '" & Compress(cmbPrt) & "' "
      sTitle = sTitle & "Part: " & cmbPrt & " , "
   Else
      sTitle = sTitle & "Part: ALL , "
   End If
   
   If Trim(cmbCde) <> "" Then
      sSql = sSql & " AND {PartTable.PAPRODCODE} = '" & Compress(cmbCde) & "' "
      sTitle = sTitle & "Product Code: " & cmbCde & " , "
   Else
      sTitle = sTitle & "Product Code: ALL , "
   End If
   
   If Trim(cmbCls) <> "" Then
      sSql = sSql & " AND {PartTable.PACLASS} = '" & Compress(cmbCls) & "' "
      sTitle = sTitle & "Product Class: " & cmbCls & " , "
   Else
      sTitle = sTitle & "Product Class: ALL , "
   End If
   
   If Trim(txtVar) <> "" Then
      sSql = sSql & " AND {@VARIANCE} > " & CCur(txtVar) & " "
      sTitle2 = sTitle2 & "Greater Variance Than: " & txtVar & " , "
   End If
   
   sTitle2 = sTitle2 & "Showing Zero QOH Items? "
   If optZro.Value = vbUnchecked Then
      sSql = sSql & " AND {PartTable.PAQOH} > 0 "
      sTitle2 = sTitle2 & "N"
   Else
      sTitle2 = sTitle2 & "Y"
   End If
   
   For i = 0 To 7
      If ChkTyp(i).Value = vbChecked Then
         sInclude = sInclude & Trim(zTyp(i)) & ","
      End If
   Next
   If Trim(sInclude) <> "" Then
      sInclude = Left(sInclude, Len(sInclude) - 1)
      sSql = sSql & " AND {PartTable.PALEVEL} IN[" & sInclude & "] "
   Else
      sInclude = "1,2,3,4,5,6,7,8"
   End If
   
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Title1='Proposed Standard Vs. Current Standard Cost'"
   MdiSect.crw.Formulas(2) = "Title2='" & sTitle & "'"
   MdiSect.crw.Formulas(3) = "RequestedBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(4) = "CostDet='" & CStr(optCst.Value) & "'"
   MdiSect.crw.Formulas(5) = "Desc='" & CStr(optDesc.Value) & "'"
   MdiSect.crw.Formulas(6) = "ExtDesc='" & CStr(optExt.Value) & "'"
   MdiSect.crw.Formulas(7) = "Title3='Part Types: " & sInclude & "'"
   MdiSect.crw.Formulas(8) = "Title4='" & sTitle2 & "'"
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "Printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Public Sub SaveOptions()
   Dim sOptions As String
   Dim i As Integer
   On Error Resume Next
   
   For i = 0 To 7
      sOptions = sOptions & CStr(ChkTyp(i).Value)
   Next
   sOptions = sOptions & CStr(optCst.Value)
   sOptions = sOptions & CStr(optDesc.Value)
   sOptions = sOptions & CStr(optExt.Value)
   sOptions = sOptions & CStr(optZro.Value)
   
   SaveSetting "Esi2000", "Esifina", Me.Name, Trim(sOptions)
   
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   Dim i As Integer
   On Error Resume Next
   
   sOptions = GetSetting("Esi2000", "Esifina", Me.Name, sOptions)
   If Len(Trim(sOptions)) = 0 Then sOptions = "111111111101"
   For i = 0 To 7
      ChkTyp(i).Value = Mid(sOptions, i + 1, 1)
   Next
   optCst.Value = Mid(sOptions, 9, 1)
   optDesc.Value = Mid(sOptions, 10, 1)
   optExt.Value = Mid(sOptions, 11, 1)
   optZro.Value = Mid(sOptions, 12, 1)
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub cmdFnd_Click()
   optVew.Value = vbChecked
   ViewParts.Show
   
End Sub

'************************************* FUNCTIONS *******************************

Private Function NumberFix(sNumber As String) As String
   Dim i As Integer
   Dim sRight As String
   On Error GoTo DiaErr1
   'Fixes Commas on the far left, removes multiple decimals
   
   If Left(sNumber, 1) = "," Then
      sNumber = Right(sNumber, Len(sNumber) - 1)
   End If
   
   For i = 1 To Len(sNumber)
      If InStr(Right(sNumber, i), ".") Then
         sRight = Right(sNumber, i)
         Exit For
      End If
   Next
   i = 0
   
   sNumber = Left(sNumber, Len(sNumber) - Len(sRight))
   
   RemoveSymbols sNumber
   RemoveSymbols sRight
   If Trim(sNumber) <> "" Or Trim(sRight) <> "" Then
      NumberFix = sNumber & "." & sRight
   Else
      NumberFix = "0"
   End If
   Exit Function
   
DiaErr1:
   sProcName = "Numberfix"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub RemoveSymbols(sNum As String)
   Dim i As Integer
   On Error GoTo DiaErr1
   
   For i = 1 To Len(sNum)
      If InStr(Left(sNum, i), ".") Or InStr(Left(sNum, i), "-") Or InStr(Left(sNum, i), "+") Then
         'delete decimal and return the string without
         sNum = Left(sNum, i - 1) & Right(sNum, Len(sNum) - i)
         i = i - 1
      End If
   Next
   sNum = Left(sNum, 13)
   Exit Sub
   
DiaErr1:
   sProcName = "RemoveSymbols"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Function bValidElement(cmbCombo As ComboBox) As Boolean 'ADJUSTED TO COMPARE STRINGS
   Dim i As Integer
   On Error GoTo DiaErr1
   If cmbCombo.ListCount > 0 Then
      For i = 0 To cmbCombo.ListCount - 1
         If Trim(UCase(cmbCombo.List(i))) = Trim(UCase(cmbCombo.Text)) Then
            bValidElement = True
            cmbCombo.ListIndex = i
         End If
      Next
   End If
   Exit Function
   
DiaErr1:
   sProcName = "bValidElement"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
