VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Center Queue"
   ClientHeight    =   3390
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3390
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbSortDetailsBy 
      Height          =   315
      ItemData        =   "ShopSHp08a.frx":0000
      Left            =   1920
      List            =   "ShopSHp08a.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp08a.frx":002B
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtWcn 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter A New Work Center"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CheckBox optGrp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   2520
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   2
      Top             =   2040
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp08a.frx":07D9
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ShopSHp08a.frx":0957
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3390
      FormDesignWidth =   7335
   End
   Begin VB.Label Label1 
      Caption         =   "Sort Details by"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   7
      Left            =   4200
      TabIndex        =   16
      Top             =   1440
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cutoff Date"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Tag             =   " "
      Top             =   1440
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Tag             =   " "
      Top             =   2040
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Comments"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Tag             =   " "
      Top             =   2280
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   4200
      TabIndex        =   11
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   360
      Picture         =   "ShopSHp08a.frx":0AE1
      ToolTipText     =   "Chart Results"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chart"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Tag             =   " "
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "ShopSHp08a"
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

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

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
   sSql = "Qry_FillWorkCentersAll"
   LoadComboBox txtWcn
   txtWcn = "ALL"
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
   Set ShopSHp08a = Nothing
   
End Sub

Private Sub PrintReport()
    Dim sCenter As String
    Dim sDate As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If Trim(txtWcn) = "" Then txtWcn = "ALL"
   If txtWcn = "ALL" Then
      sCenter = ""
   Else
      sCenter = Compress(txtWcn)
   End If
   If Not IsDate(txtDte) Then
      sDate = "2024,12,31"
   Else
      sDate = Format(txtDte, "yyyy,mm,dd")
   End If
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowComments"
    aFormulaName.Add "ShowGroup"
    aFormulaName.Add "SortByDate"
  
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtWcn) & "...'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    
    If optDsc.Value = vbUnchecked Then
       aFormulaValue.Add ""
    Else
       aFormulaValue.Add ""
    End If
     aFormulaValue.Add OptCmt.Value
     aFormulaValue.Add optGrp.Value

    If (cmbSortDetailsBy.ListIndex = 1) Then aFormulaValue.Add CStr("1") Else aFormulaValue.Add CStr("0")
    
    sCustomReport = GetCustomReport("prdsh09")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{WcntTable.WCNREF} Like '" & sCenter & "*' AND " _
          & "{RnopTable.OPSCHEDDATE}<=Date(" & sDate & ")"
          
   sSql = sSql & " AND {RnopTable.OPSCHEDDATE} <> Date (0, 0, 0) and " _
                 & " {RnopTable.OPSHOP} <> '' and " _
                   & "{RnopTable.OPCOMPLETE} = 0 and " _
                   & "{RunsTable.RUNSTATUS} in ['SC', 'PC', 'PP', 'PL', 'RL']"
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
   txtWcn = "ALL"
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = Trim(str(optDsc.Value)) _
              & Trim(str(OptCmt.Value)) _
              & Trim(str(optGrp.Value)) _
              & Trim(str(cmbSortDetailsBy.ListIndex))
              
   SaveSetting "Esi2000", "EsiProd", "sh09", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh09", sOptions)
   If Len(sOptions) Then
      optDsc.Value = Val(Mid(sOptions, 1, 1))
      OptCmt.Value = Val(Mid(sOptions, 2, 1))
      optGrp.Value = Val(Mid(sOptions, 3, 1))
      If Len(sOptions) < 4 Then cmbSortDetailsBy.ListIndex = 0 Else cmbSortDetailsBy.ListIndex = Val(Mid(sOptions, 4, 1))
   End If
   
End Sub

Private Sub Image1_Click()
   If optGrp.Value = vbChecked Then
      optGrp.Value = vbUnchecked
   Else
      optGrp.Value = vbChecked
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

Private Sub optGrp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtDte_LostFocus()
   If txtDte = "" Then txtDte = "ALL"
   If txtDte <> "ALL" Then txtDte = CheckDateEx(txtDte)
   
End Sub


Private Sub txtWcn_LostFocus()
   txtWcn = CheckLen(txtWcn, 12)
   If Len(txtWcn) = 0 Then txtWcn = "ALL"
   
End Sub
