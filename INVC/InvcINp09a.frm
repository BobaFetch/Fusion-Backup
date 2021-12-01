VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form InvcINp09a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Excess Inventory Report"
   ClientHeight    =   3225
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3225
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   1020
      Width           =   1335
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINp09a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1260
      TabIndex        =   1
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Leading Char Search  (*  In Front Is A Legal Wild Card)"
      Top             =   360
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "InvcINp09a.frx":07AE
      Enabled         =   0   'False
      Height          =   315
      Left            =   4620
      Picture         =   "InvcINp09a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   360
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6360
      TabIndex        =   19
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "InvcINp09a.frx":0E32
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
         Picture         =   "InvcINp09a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   11
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
      Left            =   2760
      TabIndex        =   6
      Top             =   1560
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   5
      Top             =   1560
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   13
      Top             =   1560
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   2340
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   2565
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   360
      Left            =   6360
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4680
      Top             =   2340
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3225
      FormDesignWidth =   7590
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Codes"
      Height          =   285
      Index           =   14
      Left            =   2880
      TabIndex        =   24
      Top             =   1020
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   23
      Top             =   1020
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   10
      Left            =   5280
      TabIndex        =   22
      Top             =   1020
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   5100
      TabIndex        =   20
      Top             =   420
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2100
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types?"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   2595
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   2340
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   14
      Top             =   390
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "InvcINp09a"
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
      For iList = 1 To 8
         typ(iList) = Mid$(sOptions, iList, 1)
      Next
      optDsc.Value = Val(Mid(sOptions, iList, 1))
      optExt.Value = Val(Mid(sOptions, iList + 1, 1))
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
              & sClass & sCode
   
   SaveSetting "Esi2000", "EsiProd", "in09", Trim(sOptions)
   
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
    If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
    
    cmbCls.AddItem "ALL"
    FillProductClasses
    FillPartCombo cmbPrt
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
  Set InvcINp09a = Nothing
   
End Sub




Private Sub PrintReport()
    Dim iList As Integer
    Dim sPart As String
    Dim sClass As String
    Dim sPartCode As String
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
     
     GetMRPCreateDates sBegDate, sEndDate
     If (Trim(sBegDate) <> "" And Trim(sEndDate) <> "") Then
         sDateDev = "'MRP Created  " & sBegDate & " For Requirements Through " & sEndDate & "'"
     Else
         sDateDev = "'Report based on last MRP run'"
     End If
    
    Dim sCustomReport As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    
    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("invExcess.rpt")
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    
    cCRViewer.SetReportTitle = "invExcess.rpt"
    cCRViewer.ShowGroupTree False

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowPartDesc"
    aFormulaName.Add "ShowExtDesc"
    aFormulaName.Add "DateDeveloped"
    aFormulaName.Add "Title2"
    aFormulaName.Add "Title3"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
    aFormulaValue.Add CInt(optDsc)
    aFormulaValue.Add CInt(optExt)
    aFormulaValue.Add CStr(sDateDev)
    aFormulaValue.Add CStr(sIncludes)
    aFormulaValue.Add CStr(sIncludes1)

    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection True
    ' report parameter
    aRptPara.Add CStr(cmbCls.Text)
    aRptPara.Add CStr(cmbCde.Text)
    aRptPara.Add CStr(sPartType1)
    aRptPara.Add CStr(sPartType2)
    aRptPara.Add CStr(sPartType3)
    aRptPara.Add CStr(sPartType4)
    
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("String")
    aRptParaType.Add CStr("Int")
    aRptParaType.Add CStr("Int")
    aRptParaType.Add CStr("Int")
    aRptParaType.Add CStr("Int")
    ' Set report parameter
    cCRViewer.SetReportDBParameters aRptPara, aRptParaType   'must happen AFTER SetDbTableConnection call!
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

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   Dim iList As Integer
   MouseCursor 13
   For iList = 1 To 4
      If typ(iList).Value = vbChecked Then Exit For
   Next
   If (iList = 5) And (typ(iList).Value <> vbChecked) Then
      MouseCursor 0
      MsgBox "You Need At Least One Part Type.", vbInformation, Caption
      On Error Resume Next
      typ(1).SetFocus
   Else
      PrintReport
   End If
   
End Sub

Private Sub optZro_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
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

