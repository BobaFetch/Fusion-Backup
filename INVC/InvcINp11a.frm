VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InvcINp11a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "No Inventory Movement Report"
   ClientHeight    =   4140
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optSrtByPrt 
      Caption         =   "PartNumber"
      Height          =   255
      Left            =   1920
      TabIndex        =   32
      Top             =   2160
      Width           =   1215
   End
   Begin VB.OptionButton optSrtByLoc 
      Caption         =   "Location"
      Height          =   255
      Left            =   3360
      TabIndex        =   31
      Top             =   2160
      Width           =   975
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   3315
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Select Product Class From List"
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   4920
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      Tag             =   "4"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINp11a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Leading Char Search  (*  In Front Is A Legal Wild Card)"
      Top             =   870
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "InvcINp11a.frx":07AE
      Height          =   315
      Left            =   5400
      Picture         =   "InvcINp11a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   840
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optZro 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   12
      Top             =   3765
      Width           =   735
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6600
      TabIndex        =   22
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "InvcINp11a.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "InvcINp11a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox typ 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   9
      Top             =   2640
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   8
      Top             =   2640
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   7
      Top             =   2640
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   2640
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   10
      Top             =   3240
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   3465
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   360
      Left            =   6600
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
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
      FormDesignHeight=   4140
      FormDesignWidth =   7695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By"
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   33
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Codes"
      Height          =   285
      Index           =   14
      Left            =   3720
      TabIndex        =   30
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      Height          =   285
      Index           =   13
      Left            =   600
      TabIndex        =   29
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "No Movement Through"
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   28
      Top             =   1725
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "No Movement From"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   27
      Top             =   1725
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   10
      Left            =   6240
      TabIndex        =   26
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   8
      Left            =   5880
      TabIndex        =   24
      Top             =   900
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Zero Rem Quantities"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   23
      Top             =   3765
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   21
      Top             =   3000
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types?"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   20
      Top             =   2640
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   3495
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   3240
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   870
      Width           =   1215
   End
End
Attribute VB_Name = "InvcINp11a"
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

Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If cmbCde = "" Then cmbCde = "ALL"
   
End Sub

Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 4)
   If cmbCls = "" Then cmbCls = "ALL"
   
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "in08", sOptions)
   If Len(sOptions) > 0 Then
      For iList = 1 To 4
         typ(iList) = Mid$(sOptions, iList, 1)
      Next
      optDsc.Value = Val(Mid(sOptions, iList, 1))
      optExt.Value = Val(Mid(sOptions, iList + 1, 1))
      optZro.Value = Val(Mid(sOptions, iList + 2, 1))
      cmbCls = Mid(sOptions, iList + 3, 4)
      If cmbCls = "" Then cmbCls = "ALL"
      cmbCde = Mid$(iList + 7, 6)
      If cmbCde = "" Then cmbCde = "ALL"
   Dim sCode As String * 6
      
   End If
   
End Sub


Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   Dim sClass As String * 4
   Dim sCode As String * 6
   
   sCode = cmbCde
   sClass = cmbCls
   
   'Save by Menu Option
   For iList = 1 To 3
      sOptions = sOptions & Trim(str(typ(iList).Value))
   Next
   sOptions = sOptions & Trim(str(typ(iList).Value))
   sOptions = sOptions & Trim(str(optDsc.Value))
   sOptions = sOptions & Trim(str(optExt.Value)) _
              & Trim(str(optZro.Value)) & sClass & sCode
   
   SaveSetting "Esi2000", "EsiProd", "in08", Trim(sOptions)
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
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
   MouseCursor 0
   MdiSect.lblBotPanel = Caption
   
   If bOnLoad Then
    ' Load the product code
    cmbCde.AddItem "ALL"
    FillProductCodes
    FillPartCombo cmbPrt
    cmbPrt = "ALL"
    If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
    
    cmbCls.AddItem "ALL"
    FillProductClasses
    If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
        bOnLoad = 0
   End If
   
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
  Set InvcINp11a = Nothing
   
End Sub




Private Sub PrintReport()
    Dim iList As Integer
    Dim sPart As String
    Dim sClass As String
    Dim sPartCode As String
    Dim sQual As String
    Dim sBegDate As String
    Dim sEndDate  As String
    Dim sDateDev As String
     
    Dim sIncludes As String
    Dim sIncludes1 As String
    Dim sPartType1 As String
    Dim sPartType2 As String
    Dim sPartType3 As String
    Dim sPartType4 As String
    
    MouseCursor 13
    'SetMdiReportsize MdiSect
    
    If (txtBeg = "" Or txtEnd = "") Then
          MsgBox "You Need Select - No Activity From Date and projected till Dates.", vbInformation, Caption
          Exit Sub
    End If
    sPart = Compress(cmbPrt)
    If Len(sPart) = 0 Then
       cmbPrt = "ALL"
       sPart = ""
    Else
       If sPart = "ALL" Then sPart = ""
    End If
    
    sClass = Compress(cmbCls)
    If Len(sClass) = 0 Then
       sClass = ""
    Else
       If sClass = "ALL" Then sClass = ""
    End If
     
    sPartCode = Compress(cmbCde)
    If Len(sPartCode) = 0 Then
       sPartCode = ""
    Else
       If sPartCode = "ALL" Then sPartCode = ""
    End If
    
    If optZro.Value = vbUnchecked Then
       sQual = "<"
    Else
       sQual = "<="
    End If
    
     sIncludes = "'Includes Part(s) " & cmbPrt & "... Part Type(s) '"
     sIncludes1 = "'Part Class " & sClass & " And Part Code " & sPartCode
     
     sIncludes1 = sIncludes1 & " - Part Type "
    
    If (typ(1).Value = vbChecked) Then
        sPartType1 = typ(1).Value
        sIncludes1 = sIncludes1 & "1 "
    Else
        sPartType1 = 0
    End If
    If (typ(2).Value = vbChecked) Then
        sPartType2 = typ(2).Value
        sIncludes1 = sIncludes1 & ",2 "
    Else
        sPartType2 = 0
    End If
    If (typ(3).Value = vbChecked) Then
        sPartType3 = typ(3).Value
        sIncludes1 = sIncludes1 & ",3 "
    Else
        sPartType3 = 0
    End If
    If (typ(4).Value = vbChecked) Then
        sPartType4 = typ(4).Value
        sIncludes1 = sIncludes1 & ",4 "
    Else
        sPartType4 = 0
    End If
    
    sIncludes1 = sIncludes1 & "'"
   
    Dim sCustomReport As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    
    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("invInactNotMrp.rpt")
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    
    cCRViewer.SetReportTitle = "invInactive.rpt"
    cCRViewer.ShowGroupTree False

    Dim iSort As Integer
    
    If (optSrtByLoc.Value = True) Then
        iSort = 1
    Else
        iSort = 0
    End If

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowPartDesc"
    aFormulaName.Add "ShowExtDesc"
    aFormulaName.Add "IncludeZeroQty"
    aFormulaName.Add "Title2"
    aFormulaName.Add "Title3"
    aFormulaName.Add "sortBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
    aFormulaValue.Add CInt(optDsc)
    aFormulaValue.Add CInt(optExt)
    aFormulaValue.Add CInt(optZro)
    aFormulaValue.Add CStr(sIncludes)
    aFormulaValue.Add CStr(sIncludes1)
    aFormulaValue.Add iSort

    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection True
    ' report parameter
    aRptPara.Add CStr(txtBeg)
    aRptPara.Add CStr(txtEnd)
    aRptPara.Add CStr(sClass)
    aRptPara.Add CStr(sPartCode)
    aRptPara.Add CStr(optZro)
    aRptPara.Add CStr(sPartType1)
    aRptPara.Add CStr(sPartType2)
    aRptPara.Add CStr(sPartType3)
    aRptPara.Add CStr(sPartType4)
    
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("Int")
    aRptParaType.Add CStr("Int")
    aRptParaType.Add CStr("Int")
    aRptParaType.Add CStr("Int")
    aRptParaType.Add CStr("Int")
    ' Set report parameter
    cCRViewer.SetReportDBParameters aRptPara, aRptParaType
    'cCRViewer.SetTableConnection aRptPara
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   Dim iList As Integer
   MouseCursor 13
   For iList = 1 To 4
      If typ(iList).Value = vbChecked Then Exit For
   Next
   If (iList = 5) Then
      MouseCursor 0
      MsgBox "You Need At Least One Part Type.", vbInformation, Caption
      On Error Resume Next
      typ(1).SetFocus
   Else
      PrintReport
   End If
   
End Sub


Private Sub optPrn_Click()
   Dim iList As Integer
   MouseCursor 13
   For iList = 1 To 4
      If typ(iList).Value = vbChecked Then Exit For
   Next
   If iList = 5 Then
      MouseCursor 0
      MsgBox "You Need At Least One Part Type.", vbInformation, Caption
      On Error Resume Next
      typ(1).SetFocus
   Else
      PrintReport
   End If
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDateEx(txtBeg)
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDateEx(txtEnd)
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Len(txtPrt) = 0 Then txtPrt = "ALL"
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) = 0 Then cmbPrt = "ALL"

End Sub

Private Sub typ_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
