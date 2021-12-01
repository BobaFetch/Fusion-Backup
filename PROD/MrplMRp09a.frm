VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form MrplMRp09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MO Early/Late Report"
   ClientHeight    =   4305
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   8805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4305
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   495
      Left            =   1680
      TabIndex        =   34
      Top             =   2400
      Width           =   2775
      Begin VB.OptionButton optMbe 
         Caption         =   "M"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "B"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   37
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "E"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   36
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "ALL"
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   35
         Top             =   200
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CheckBox chkType 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   33
      Top             =   3120
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   32
      Top             =   3120
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   31
      Top             =   3120
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   30
      Top             =   3120
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.ComboBox cmbPart 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "MrplMRp09a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
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
      Picture         =   "MrplMRp09a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CheckBox chkExtDesc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   3840
      Width           =   735
   End
   Begin VB.CheckBox chkDesc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   3480
      Width           =   735
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "MrplMRp09a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "MrplMRp09a.frx":0AB6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4305
      FormDesignWidth =   8805
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make, Buy, Either"
      Height          =   285
      Index           =   14
      Left            =   240
      TabIndex        =   40
      Top             =   2640
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types"
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   39
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   27
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   9
      Left            =   5640
      TabIndex        =   26
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   25
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   24
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   23
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   22
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Codes"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Description"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   3480
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   2385
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last MRP"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblMrp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblUsr 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   360
      Width           =   615
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   1005
      Width           =   1425
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   11
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "MrplMRp09a"
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




Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If cmbCde = "" Then cmbCde = "ALL"
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 6)
   If cmbCls = "" Then cmbCls = "ALL"
   
End Sub



Private Sub cmbPart_LostFocus()
    cmbPart = CheckLen(cmbPart, 30)
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub

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

Private Sub FillCombos()
    On Error Resume Next
    sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
        & "FROM PartTable  " _
        & "INNER JOIN MrplTable ON MrplTable.MRP_PARTREF=PartTable.PARTREF " _
        & " WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0 " _
        & "ORDER BY PARTREF"
    LoadComboBox cmbPart, 0
    cmbPart = "ALL"
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetMRPDates
      GetLastMrp
      cmbCde.AddItem "ALL"
      FillProductCodes
      If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
      cmbCls.AddItem "ALL"
      FillProductClasses
      If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
      
      FillCombos
      
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
   GetOptions
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
   Set MrplMRp09a = Nothing
   
End Sub




Private Sub PrintReport()
    Dim sParts As String
    Dim sCode As String
    Dim sClass As String
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
    Dim strIncludes As String
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
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
    
    If Trim(cmbCde) = "" Then cmbCde = "ALL"
    If Trim(cmbCls) = "" Then cmbCls = "ALL"
    If Trim(cmbPart) = "ALL" Then sParts = "" Else sParts = Compress(cmbPart)
    If Trim(cmbCde) = "ALL" Then sCode = "" Else sCode = Compress(cmbCde)
    If Trim(cmbCls) = "ALL" Then sClass = "" Else sClass = Compress(cmbCls)



    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("prdEarlyLate")

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

    strIncludes = Trim(cmbPart) & ", Prod Code(s) " & cmbCde & ", Class(es) " _
                            & cmbCls
    aFormulaValue.Add CStr("'" & CStr(strIncludes) & "...'")
    aFormulaValue.Add CStr("'" & CStr(sInitials) & "'")

    strDateDev = "'MRP Created  " & sBegDate & " For Requirements Through " & sEndDate & "'"
    aFormulaValue.Add CStr(strDateDev)

   If optMbe(0).Value = True Then
      sPAMake = "M"
      sMbe = "Make"
      sSql = sSql & "AND {PartTable.PAMAKEBUY}='M'"
      sSubSql = sSubSql & "AND {PartTable.PAMAKEBUY}='M'"
   ElseIf optMbe(1).Value = True Then
      sPAMake = "B"
      sMbe = "Buy"
      sSql = sSql & "AND {PartTable.PAMAKEBUY}='B'"
      sSubSql = sSubSql & "AND {PartTable.PAMAKEBUY}='B'"
   ElseIf optMbe(2).Value = True Then
      sPAMake = "E"
      sMbe = "Either"
      sSql = sSql & "AND {PartTable.PAMAKEBUY}='E'"
      sSubSql = sSubSql & "AND {PartTable.PAMAKEBUY}='E'"
   Else
      sPAMake = "A"
      sMbe = "Make, Buy And Either"
   End If
   
   aFormulaName.Add "Mbe"
   aFormulaValue.Add CStr("'" & sMbe & "'")
    
    
   'select part types
   Dim types As String
   Dim includes As String
   includes = ""
   Dim I As Integer
   For I = 1 To 3
     If Me.chkType(I).Value = vbChecked Then
        includes = includes & " " & I
     End If
   Next
   
   If includes = "" Then
     MsgBox "No part types selected"
     Exit Sub
   End If

   aFormulaName.Add "PartInc"
   aFormulaValue.Add CStr("'" & includes & "'")
    
   aFormulaName.Add "ShowExDesc"
   aFormulaValue.Add chkExtDesc.Value
        
   aFormulaName.Add "ShowPartDesc"
   aFormulaValue.Add chkDesc.Value

   ' Set Formula values
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
   aRptPara.Add CStr(txtBeg)
   aRptPara.Add CStr(txtEnd)
   aRptPara.Add CStr(chkType(1))
   aRptPara.Add CStr(chkType(2))
   aRptPara.Add CStr(chkType(3))
   aRptPara.Add CStr("0")
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   aRptParaType.Add CStr("Int")
   
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

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sCode As String * 6
   Dim sClass As String * 4
   sCode = cmbCde
   sClass = cmbCls
   sOptions = sCode & sClass & Trim(str(Val(chkExtDesc.Value))) _
              & Trim(str(Val(chkDesc.Value)))
   SaveSetting "Esi2000", "EsiProd", "Prdmr09", sOptions
   SaveSetting "Esi2000", "EsiProd", "Pmr09", lblPrinter
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "Prdmr01", sOptions)
   If Len(Trim(sOptions)) > 0 Then
      cmbCde = Mid$(sOptions, 1, 6)
      cmbCls = Mid$(sOptions, 7, 4)
      chkExtDesc.Value = Val(Mid$(sOptions, 11, 1))
      chkDesc.Value = Val(Mid$(sOptions, 12, 1))
   End If
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Pmr09", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub


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

