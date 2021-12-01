VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form BookBLp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backlog By Part With Prices"
   ClientHeight    =   3960
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3960
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "BookBLp01a.frx":0000
      Height          =   315
      Left            =   4800
      Picture         =   "BookBLp01a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1080
      Width           =   350
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BookBLp01a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cmbCls 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   6
      Top             =   3360
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "BookBLp01a.frx":0E32
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
         Picture         =   "BookBLp01a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3960
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   10
      Left            =   5400
      TabIndex        =   22
      Top             =   1800
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   9
      Left            =   5400
      TabIndex        =   21
      Top             =   2520
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   7
      Left            =   5400
      TabIndex        =   20
      Top             =   2160
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Class"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cutoff Date (Sched)"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Tag             =   " "
      Top             =   3360
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Desc"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Tag             =   " "
      Top             =   3120
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5400
      TabIndex        =   12
      Top             =   1080
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1665
   End
End
Attribute VB_Name = "BookBLp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/25/05 Changed dates and Options
'4/25/05 Turned off Part Combo
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbcde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If Len(cmbCde) = 0 Then cmbCde = "ALL"
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 4)
   If Len(cmbCls) = 0 Then cmbCls = "ALL"
   
End Sub

Private Sub cmbPrt_Click()
   GetPart
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt = "" Then cmbPrt = "ALL"
   GetPart
   
End Sub

Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPrt = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPrt
   ViewParts.Show
End Sub

Private Sub txtPrt_Change()
   cmbPrt = txtPrt
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub


Private Sub txtPrt_LostFocus()
   cmbPrt = txtPrt
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   GetPart
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

Private Sub cmdVew_Click(Index As Integer)
   ViewParts.lblControl = "CMBPRT"
   ViewParts.txtPrt = cmbPrt
   ViewParts.Show
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductCodes
      FillProductClasses
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombo cmbPrt
      
      FormatControls
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
 '  FormatControls
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
   Set BookBLp01a = Nothing
   
End Sub
Private Sub PrintReport()
   Dim sEnd As String
   Dim sPart As String
   Dim sCode As String
   Dim sClass As String
   
   If cmbPrt <> "ALL" Then sPart = Compress(cmbPrt)
   If Not IsDate(txtEnd) Then sEnd = "2024,12,31" Else _
                 sEnd = Format(txtEnd, "yyyy,mm,dd")
   If cmbCde <> "ALL" Then sCode = Compress(cmbCde)
   If cmbCls <> "ALL" Then sClass = Compress(cmbCls)
   MouseCursor 13
   On Error GoTo DiaErr1
   
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
 
    sCustomReport = GetCustomReport("slebl01")
   
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = "slebl01"
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowComments"
    
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Part Number(s)" & CStr(cmbPrt) & " Through " & _
                                CStr(txtEnd) & "... '")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add optDsc
    aFormulaValue.Add optCmt
  
   sSql = "{PartTable.PARTREF} LIKE '" & sPart & "*' " _
          & "AND {SoitTable.ITSCHED}<= Date(" & sEnd & ") " _
          & "AND {PartTable.PAPRODCODE} LIKE '" & sCode & "*' " _
          & "AND {PartTable.PACLASS} LIKE '" & sClass & "*' " _
          & " AND {SoitTable.ITPSNUMBER} = '' " _
          & " AND {SoitTable.ITINVOICE} = .000 " _
          & " AND {SoitTable.ITPSSHIPPED} = 0 " _
          & " AND {SoitTable.ITCANCELED} = .000"
    
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
   'SetCrystalAction Me
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
   cmbCls = "ALL"
   cmbCde = "ALL"
   cmbPrt = "ALL"
   GetPart
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(str(optDsc.Value)) & Trim(str(optCmt.Value))
   SaveSetting "Esi2000", "EsiSale", "bL01", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = Trim(GetSetting("Esi2000", "EsiSale", "bL01", Trim(sOptions)))
   If Len(sOptions) = 2 Then
      optDsc = Left$(sOptions, 1)
      optCmt = Right$(sOptions, 1)
   Else
      optDsc.Value = vbChecked
      optCmt.Value = vbChecked
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 12) = "*** No Parts" Then
      lblDsc.ForeColor = ES_RED
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


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub



Private Sub GetPart()
   Dim RdoPrt As ADODB.Recordset
   sSql = "Qry_GetPartNumberBasics '" & Compress(cmbPrt) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(.Fields(1))
         If Len(cmbPrt) > 3 Then
            lblDsc = "" & Trim(.Fields(2))
         Else
            lblDsc = "*** Range Of Parts Selected ***"
         End If
         ClearResultSet RdoPrt
      End With
   Else
      lblDsc = "*** Range Of Parts Selected ***"
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Trim(txtEnd) <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
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

