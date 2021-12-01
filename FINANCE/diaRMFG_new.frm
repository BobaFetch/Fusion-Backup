VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaRMFG_new 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Raw Material/Finished Goods Inventory New"
   ClientHeight    =   4620
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4620
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optIna 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   3960
      Width           =   855
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Tag             =   "9"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox optLot 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   3600
      Width           =   855
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Tag             =   "8"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Tag             =   "4"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox typ 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5640
      TabIndex        =   12
      Top             =   480
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaRMFG_new.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaRMFG_new.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   15
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaRMFG_new.frx":0308
      PictureDn       =   "diaRMFG_new.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   16
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
      PictureUp       =   "diaRMFG_new.frx":0594
      PictureDn       =   "diaRMFG_new.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Inactive and Obsolete Part"
      Height          =   405
      Index           =   7
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Code"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   26
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL) "
      Height          =   285
      Index           =   9
      Left            =   3600
      TabIndex        =   25
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Head Only"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Class"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL) "
      Height          =   285
      Index           =   10
      Left            =   3600
      TabIndex        =   22
      Top             =   2280
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "As Of Date"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaRMFG_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*********************************************************************************
' diaRMFG - Raw Material Finished Goods
'
' Notes:
'
' Created: 02/09/04 (nth)
' Revisions:
'   02/12/04 (nth) Added part class and lot detail options.
'   05/04/04 (nth) Added include parts with no QOH option.
'
'*********************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************
Private Sub cmbCde_LostFocus()
   If Trim(cmbCde) = "" Then cmbCde = "ALL"
End Sub


Private Sub cmbCls_LostFocus()
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
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

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaRMFG = Nothing
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

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim sCustomReport As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   On Error GoTo whoops
    
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
   If Trim(cmbCde) = "" Then cmbCde = "ALL"
   'get custom report name if one has been defined
   sCustomReport = GetCustomReport("finRMFG_new.rpt")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   
   cCRViewer.SetReportTitle = "finRMFG_new.rpt"
   cCRViewer.ShowGroupTree False
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "ShowPartDesc"
   aFormulaName.Add "ShowExtDesc"
   'aFormulaName.Add "ShowSummary"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
   aFormulaValue.Add CStr("'As Of " & txtDte & "'")
   aFormulaValue.Add CInt(optDsc)
   aFormulaValue.Add CInt(optExt)
   'aFormulaValue.Add CInt(chkSummary)
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{RptRMFGoods;1.RPTDATEQTY} > 0"
   cCRViewer.SetReportSelectionFormula (sSql)
   
   cCRViewer.CRViewerSize Me
   ' Set report parameter
   cCRViewer.SetDbTableConnection True
   ' report parameter
   aRptPara.Add CStr(txtDte)
   aRptPara.Add CStr(cmbCls.Text)
   aRptPara.Add CStr(cmbCde.Text)
   aRptPara.Add CStr("0")
   aRptPara.Add CStr(typ(1))
   aRptPara.Add CStr(typ(2))
   aRptPara.Add CStr(typ(3))
   aRptPara.Add CStr(typ(4))
   aRptPara.Add CStr(optIna)
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   
   ' Set report parameter
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType    'must happen AFTER SetDbTableConnection call!
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   Exit Sub
   
whoops:
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

'Private Sub PrintReport()
'   Dim sCustomerReport As String
'   Dim sType As String
'   Dim b As Byte
'
'   MouseCursor 13
'   On Error GoTo DiaErr1
'
'   If Trim(cmbCls) = "" Then cmbCls = "ALL"
'
'   'setmdireportsizemdisect
'
'   For b = 1 To 4
'      If typ(b) = vbChecked Then
'         sType = sType & CStr(b) & ","
'      End If
'   Next
'   If Len(sType) Then
'      sType = Left(sType, Len(sType) - 1)
'   End If
'
'   If optLot = vbChecked Then
'      sCustomReport = GetCustomReport("finRMFGa.rpt")
'   Else
'      sCustomReport = GetCustomReport("finRMFGb.rpt")
'   End If
'   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
'
'   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " _
'                        & sInitials & "'"
'   MdiSect.crw.Formulas(2) = "AsOf='" & txtDte & "'"
'   MdiSect.crw.Formulas(3) = "Title1='As Of " & txtDte & "'"
'   MdiSect.crw.Formulas(4) = "Title2='Includes Part Types " & sType & "'"
'   MdiSect.crw.Formulas(5) = "Title3='For Part Class " & cmbCls & "'"
'
'   MdiSect.crw.Formulas(6) = "Dsc=" & optDsc
'   MdiSect.crw.Formulas(7) = "Ext=" & optExt
'   MdiSect.crw.Formulas(8) = "QOH=" & optQOH
'
'   sSql = "{InvaTable.INADATE}<=cdate('" & txtDte & _
'          "') AND {PartTable.PALEVEL} IN [" & sType & "]"
'   If UCase(cmbCls) <> "ALL" Then
'      sSql = sSql & " AND {PartTable.PACLASS}='" & cmbCls & "'"
'   End If
'   'If optLot = vbUnchecked Then
'   '    sSql = sSql & " AND {PartTable.PAQOH} > 0"
'   'End If
'   MdiSect.crw.SelectionFormula = sSql
'
'   'setcrystalaction me
'
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printrep"
'   CurrError.Number = Err
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'End Sub
'
Private Sub FillCombo()
   FillProductClasses Me
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   
   ' Save by Menu Option
   sOptions = Trim(txtDte.Text) _
              & RTrim(typ(1).Value) _
              & RTrim(typ(2).Value) _
              & RTrim(typ(3).Value) _
              & RTrim(typ(4).Value) _
              & RTrim(optDsc.Value) _
              & RTrim(optExt.Value) _
              & RTrim(optLot.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As String
   On Error Resume Next
   dToday = CInt(Mid(Format(Now, "mm/dd/yy"), 4, 2))

   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
     
     If dToday < 21 Then
      txtDte = Mid(sOptions, 1, 8)
     Else
      txtDte = GetMonthEnd(Format(Now, "mm/dd/yy"))
     End If

      typ(1).Value = Val(Mid(sOptions, 9, 1))
      typ(2).Value = Val(Mid(sOptions, 10, 1))
      typ(3).Value = Val(Mid(sOptions, 11, 1))
      typ(4).Value = Val(Mid(sOptions, 12, 1))
      optDsc.Value = Val(Mid(sOptions, 13, 1))
      optExt.Value = Val(Mid(sOptions, 14, 1))
      optLot.Value = Val(Mid(sOptions, 15, 1))
   Else
      typ(1).Value = vbChecked
      typ(2).Value = vbChecked
      typ(3).Value = vbChecked
      typ(4).Value = vbChecked
      optDsc.Value = vbUnchecked
      optExt.Value = vbUnchecked
      optLot.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub
