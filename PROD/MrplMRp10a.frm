VERSION 5.00
Begin VB.Form MrplMRp10a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MO Early/Late Report For Purchased Parts"
   ClientHeight    =   2070
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2070
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "MrplMRp10a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
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
      Picture         =   "MrplMRp10a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4440
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "MrplMRp10a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "MrplMRp10a.frx":0AB6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   11
      Left            =   6120
      TabIndex        =   14
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   13
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   12
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last MRP"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   10
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblMrp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblUsr 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   6
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "MrplMRp10a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/19/06 Revised report and selections. Removed extra report.
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd




'Private Sub cmbCde_LostFocus()
'   cmbCde = CheckLen(cmbCde, 6)
'   If cmbCde = "" Then cmbCde = "ALL"
'
'End Sub
'
'
'Private Sub cmbCls_LostFocus()
'   cmbCls = CheckLen(cmbCls, 6)
'   If cmbCls = "" Then cmbCls = "ALL"
'
'End Sub
'


'Private Sub cmbPart_LostFocus()
'    cmbPart = CheckLen(cmbPart, 30)
'    If Trim(cmbPart) = "" Then cmbPart = "ALL"
'End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

'Private Sub cmdFnd_Click()
'   ViewParts.lblControl = "TXTPRT"
'   ViewParts.txtPrt = txtPrt
'   optVew.Value = vbChecked
'   ViewParts.Show
'
'End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

'Private Sub FillCombos()
'    On Error Resume Next
'    sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
'        & "FROM PartTable  " _
'        & "INNER JOIN MrplTable ON MrplTable.MRP_PARTREF=PartTable.PARTREF " _
'        & " WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0 " _
'        & "ORDER BY PARTREF"
'    LoadComboBox cmbPart, 0
'    cmbPart = "ALL"
'    If Trim(cmbPart) = "" Then cmbPart = "ALL"
'End Sub
'
Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetMRPDates
      GetLastMrp
'      cmbCde.AddItem "ALL"
'      FillProductCodes
'      If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
'      cmbCls.AddItem "ALL"
'      FillProductClasses
'      If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
      
'      FillCombos
      
      bOnLoad = 0
   End If
   If optVew.Value = vbChecked Then
      optVew.Value = vbUnchecked
      Unload ViewParts
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
   'SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set MrplMRp10a = Nothing
   
End Sub




Private Sub PrintReport()
'    Dim sParts As String
'    Dim sCode As String
'    Dim sClass As String
    Dim sBDate As String
    Dim sEDate As String
    Dim sBegDate As String
    Dim sEndDate As String
    Dim sMbe As String
    
    Dim sCustomReport As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    
    Dim aRptSubRptPara As New Collection
    Dim aRptSubRptParaType As New Collection
    
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    'Dim strIncludes As String
    Dim strDateDev As String
    Dim sSubSql As String
    Dim sPAMake As String
    
    MouseCursor 13
    On Error GoTo DiaErr1
    GetMRPCreateDates sBegDate, sEndDate

    If Trim(txtBeg) = "" Then txtBeg = "ALL"
    If Trim(txtEnd) = "" Then txtEnd = "ALL"
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

'    If Trim(txtPrt) = "" Then txtPrt = "ALL"
'    If Trim(cmbPart) = "" Then cmbPart = "ALL"
'
'    If Trim(cmbCde) = "" Then cmbCde = "ALL"
'    If Trim(cmbCls) = "" Then cmbCls = "ALL"
'    If Trim(cmbPart) = "ALL" Then sParts = "" Else sParts = Compress(cmbPart)
'    If Trim(cmbCde) = "ALL" Then sCode = "" Else sCode = Compress(cmbCde)
'    If Trim(cmbCls) = "ALL" Then sClass = "" Else sClass = Compress(cmbCls)



    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("prdEarlyLatePart4")

    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "DateDeveloped"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")

'    strIncludes = Trim(cmbPart) & ", Prod Code(s) " & cmbCde & ", Class(es) " _
'                            & cmbCls
'    aFormulaValue.Add CStr("'" & CStr(strIncludes) & "...'")
    aFormulaValue.Add CStr("")
    aFormulaValue.Add CStr("'" & CStr(sInitials) & "'")

    strDateDev = "'MRP Created  " & sBegDate & " For Requirements Through " & sEndDate & "'"
    aFormulaValue.Add CStr(strDateDev)

   ' Set Formula values
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
   ' report parameter
   aRptPara.Add CStr(txtBeg)
   aRptPara.Add CStr(txtEnd)
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection

   cCRViewer.SetReportDBParameters aRptPara, aRptParaType   'must happen AFTER SetDbTableConnection call!
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aRptParaType
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
'   txtPrt = "ALL"
End Sub

'Private Sub SaveOptions()
'   Dim sOptions As String
'   Dim sCode As String * 6
'   Dim sClass As String * 4
'   sCode = cmbCde
'   sClass = cmbCls
'   SaveSetting "Esi2000", "EsiProd", "Prdmr10", sOptions
'   SaveSetting "Esi2000", "EsiProd", "Prdmr10", lblPrinter
'
'End Sub
'
'Private Sub GetOptions()
'   Dim sOptions As String
'   On Error Resume Next
'   sOptions = GetSetting("Esi2000", "EsiProd", "Prdmr10", sOptions)
'   If Len(Trim(sOptions)) > 0 Then
'      cmbCde = Mid$(sOptions, 1, 6)
'      cmbCls = Mid$(sOptions, 7, 4)
'   End If
'   lblPrinter = GetSetting("Esi2000", "EsiProd", "Prdmr10", lblPrinter)
'   If lblPrinter = "" Then lblPrinter = "Default Printer"
'
'End Sub
'
'
Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub chkExtDesc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub chkExceptions_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
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

'Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF4 Then
'      ViewParts.lblControl = "TXTPRT"
'      ViewParts.txtPrt = txtPrt
'      optVew.Value = vbChecked
'      ViewParts.Show
'   End If
'
'End Sub

'Private Sub txtPrt_LostFocus()
'   txtPrt = CheckLen(txtPrt, 30)
'   If Trim(txtPrt) = "" Then txtPrt = "ALL"
'
'End Sub



'Least to greatest dates 10/12/01

Private Sub GetMRPDates()

   Dim RdoDte As ADODB.Recordset
    sSql = "SELECT MIN(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtBeg = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtBeg.ToolTipText = "Earliest Date By Default"
   
   sSql = "SELECT MAX(MRP_PARTDATERQD) FROM MrplTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtEnd = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtEnd.ToolTipText = "Latest Date By Default"
   Set RdoDte = Nothing
End Sub
