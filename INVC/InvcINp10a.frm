VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InvcINp10a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ITAR and EAR Parts"
   ClientHeight    =   4155
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4155
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optQoh 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2400
      TabIndex        =   32
      ToolTipText     =   "Shows Only Part Numbers With No Product Code"
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox optITAR 
      Height          =   195
      Left            =   2400
      TabIndex        =   29
      Top             =   2520
      Width           =   240
   End
   Begin VB.CheckBox optEAR 
      Height          =   195
      Left            =   2400
      TabIndex        =   28
      Top             =   2880
      Width           =   240
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINp10a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtCde 
      Height          =   285
      Left            =   4725
      TabIndex        =   2
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Leading Char(s) Or Blank For All"
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "InvcINp10a.frx":07AE
      Height          =   315
      Left            =   5520
      Picture         =   "InvcINp10a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1080
      Width           =   350
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6720
      TabIndex        =   21
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "InvcINp10a.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "InvcINp10a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox typ 
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   5760
      TabIndex        =   10
      Top             =   1800
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   5280
      TabIndex        =   9
      Top             =   1800
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   8
      Top             =   1800
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   7
      Top             =   1800
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   6
      Top             =   1800
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   5
      Top             =   1800
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   4
      Top             =   1800
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.TextBox txtCls 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Leading Char(s) Or Blank For All"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Tag             =   "3"
      Text            =   "ALL"
      Top             =   1110
      Width           =   3075
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2400
      TabIndex        =   11
      Top             =   3240
      Width           =   735
   End
   Begin VB.CheckBox OptCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   360
      Left            =   6720
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7560
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4155
      FormDesignWidth =   7860
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Only Items With Qoh"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   33
      ToolTipText     =   "Shows Only Part Numbers With No Product Code"
      Top             =   2160
      Width           =   2145
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "Include ITAR Part"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   31
      Top             =   2520
      Width           =   1845
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "Include EAR Part"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   30
      Top             =   2880
      Width           =   2085
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code(s)"
      Height          =   285
      Index           =   8
      Left            =   3240
      TabIndex        =   26
      Top             =   1455
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL)"
      Height          =   285
      Index           =   7
      Left            =   6120
      TabIndex        =   25
      Top             =   1440
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL)"
      Height          =   285
      Index           =   6
      Left            =   6120
      TabIndex        =   24
      Top             =   1080
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types?"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   20
      Top             =   1800
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Class(es)"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   1455
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   3240
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   1110
      Width           =   1815
   End
End
Attribute VB_Name = "InvcINp10a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/24/02 added Qoh
'2/28/05 Changed date handling
'1/25/07 Fixed Only Items With QOH CheckBox
Option Explicit

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtCde = "ALL"
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   Dim sOptions2 As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "in01", sOptions)
   If Len(sOptions) > 0 Then
      For iList = 1 To 8
         typ(iList) = Mid$(sOptions, iList, 1)
      Next
      optDsc.Value = Val(Mid(sOptions, iList, 1))
      OptCmt.Value = Val(Mid(sOptions, iList + 1, 1))
      txtCls = Mid(sOptions, iList + 2, 4)
      If txtCls = "" Then txtCls = "ALL"
   End If
   
   sOptions2 = GetSetting("Esi2000", "EsiProd", "in01a", Trim(sOptions))
   If Len(Trim(sOptions2)) = 0 Then sOptions2 = "1"
   optQoh.Value = Val(sOptions2)
   
End Sub


Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   Dim sClass As String * 4
   sClass = txtCls
   
   'Save by Menu Option
   For iList = 1 To 7
      sOptions = sOptions & Trim(str(typ(iList).Value))
   Next
   sOptions = sOptions & Trim(str(typ(iList).Value))
   sOptions = sOptions & Trim(str(optDsc.Value)) _
              & Trim(str(OptCmt.Value)) & sClass
   SaveSetting "Esi2000", "EsiProd", "in01", Trim(sOptions)
   SaveSetting "Esi2000", "EsiProd", "in01a", Trim(optQoh.Value)
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
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
   MouseCursor 0
   MdiSect.lblBotPanel = Caption
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set InvcINp10a = Nothing
   
End Sub
Private Sub PrintReport()
    Dim iList As Integer
    Dim sPart As String
    Dim sClass As String
    Dim sCode As String
   Dim sqlITAR As String
   Dim sqlEAR As String
    
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection

   If Len(Trim(txtPrt)) = 0 Then txtPrt = "ALL"
   If txtPrt <> "ALL" Then sPart = Compress(txtPrt)
   If txtCls <> "ALL" Then sClass = Compress(txtCls)
   If txtCde <> "ALL" Then sCode = Compress(txtCde)

    On Error GoTo whoops

    'get custom report name if one has ben defined
    sCustomReport = GetCustomReport("prdin10.rpt")
  
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = "prdin10.rpt"
    cCRViewer.ShowGroupTree False

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "IncOnlyITAR"
    aFormulaName.Add "IncOnlyEAR"
    
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowComments"

    'aFormulaName.Add "ShowPartDesc"
    'aFormulaName.Add "ShowExtDesc"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")

    
    sSql = "'Includes Part(s) " & txtPrt & "... Part Type(s) "
    For iList = 1 To 7
       If typ(iList).Value = vbChecked Then sSql = sSql & str(iList) & ","
    Next
    If typ(iList).Value = vbChecked Then sSql = sSql & str(iList) & ","
    iList = Len(sSql)
    sSql = Left(sSql, iList - 1) & " And Classe(s) " & txtCls & "...'"

    
    aFormulaValue.Add CStr(sSql)
    
    aFormulaValue.Add optITAR.Value
    aFormulaValue.Add optEAR.Value
    
    aFormulaValue.Add optDsc.Value
    aFormulaValue.Add OptCmt.Value
    
    'aFormulaValue.Add CInt(optDsc.value)
    'aFormulaValue.Add CInt(OptCmt.value)

    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

    sSql = "{PartTable.PARTREF} Like '" & sPart & "*' " _
           & "AND {PartTable.PACLASS} like '" & sClass & "*' "
    
    sSql = sSql & "AND {PartTable.PAPRODCODE} like '" & sCode & "*' "
    
    If optQoh.Value = vbChecked Then _
                      sSql = sSql & "AND {PartTable.PAQOH}>0 "
    
    If typ(1).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>1 "
    If typ(2).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>2 "
    If typ(3).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>3 "
    If typ(4).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>4 "
    If typ(5).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>5 "
    If typ(6).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>6 "
    If typ(7).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>7 "
    If typ(8).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>8 "
    
'sSql = cCRViewer.GetReportSelectionFormula
   
   If (optITAR.Value = 1) Then
      sqlITAR = " {PartTable.PAITARRPT} = 1"
   End If
   
   If (optEAR.Value = 1) Then
      sqlEAR = " {PartTable.PAEARRPT} = 1"
   End If

   If (sqlITAR <> "") Then
      sSql = sSql & " AND (" & sqlITAR
   End If
   
   If (sqlEAR <> "") Then
      If (sqlITAR <> "") Then
         sSql = sSql & " OR " & sqlEAR & ")"
      Else
         sSql = sSql & " AND (" & sqlEAR & ")"
      End If
   Else
      If (sqlITAR <> "") Then
         sSql = sSql & ")"
      End If
   End If
   
    ' set the report section
    cCRViewer.SetReportSelectionFormula (sSql)

    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    
    ' report parameter
'    If optDsc.value = vbUnchecked Then
'        cCRViewer.SetReportSection "DetailSection1", True
'    Else
'        cCRViewer.SetReportSection "DetailSection1", False
'    End If
'    If OptCmt.value = vbUnchecked Then
'        cCRViewer.SetReportSection "GroupFooterSection1", True
'        cCRViewer.SetReportSection "GroupFooterSection1", True
'    Else
'        cCRViewer.SetReportSection "GroupFooterSection2", False
'        cCRViewer.SetReportSection "GroupFooterSection2", False
'    End If

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

Private Sub PrintReport1()
    Dim iList As Integer
    Dim sPart As String
    Dim sClass As String
    Dim sCode As String
    
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection

   If Len(Trim(txtPrt)) = 0 Then txtPrt = "ALL"
   If txtPrt <> "ALL" Then sPart = Compress(txtPrt)
   If txtCls <> "ALL" Then sClass = Compress(txtCls)
   If txtCde <> "ALL" Then sCode = Compress(txtCde)

    On Error GoTo whoops

    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("prdin01.rpt")


  
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = "prdin01.rpt"
    cCRViewer.ShowGroupTree False

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    'aFormulaName.Add "ShowPartDesc"
    'aFormulaName.Add "ShowExtDesc"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    
    sSql = "'Includes Part(s) " & txtPrt & "... Part Type(s) "
    For iList = 1 To 7
       If typ(iList).Value = vbChecked Then sSql = sSql & str(iList) & ","
    Next
    If typ(iList).Value = vbChecked Then sSql = sSql & str(iList) & ","
    iList = Len(sSql)
    sSql = Left(sSql, iList - 1) & " And Classe(s) " & txtCls & "...'"

    
    aFormulaValue.Add CStr(sSql)
    'aFormulaValue.Add CInt(optDsc)
    'aFormulaValue.Add CInt(optExt)

    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

    sSql = "{PartTable.PARTREF} Like '" & sPart & "*' " _
           & "AND {PartTable.PACLASS} like '" & sClass & "*' "
    sSql = sSql & "AND {PartTable.PAPRODCODE} like '" & sCode & "*' "
    
    If optQoh.Value = vbChecked Then _
                      sSql = sSql & "AND {PartTable.PAQOH}>0 "
    If typ(1).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>1 "
    If typ(2).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>2 "
    If typ(3).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>3 "
    If typ(4).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>4 "
    If typ(5).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>5 "
    If typ(6).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>6 "
    If typ(7).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>7 "
    If typ(8).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>8 "
    ' set the report section
    cCRViewer.SetReportSelectionFormula (sSql)

    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    
    ' report parameter
    If optDsc.Value = vbUnchecked Then
        cCRViewer.SetReportSection "DetailSection1", True
    Else
        cCRViewer.SetReportSection "DetailSection1", False
    End If
    If OptCmt.Value = vbUnchecked Then
        cCRViewer.SetReportSection "GroupFooterSection1", True
        cCRViewer.SetReportSection "GroupFooterSection1", True
    Else
        cCRViewer.SetReportSection "GroupFooterSection2", False
        cCRViewer.SetReportSection "GroupFooterSection2", False
    End If

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
'   Dim iList As Integer
'   Dim sPart As String
'   Dim sClass As String
'   Dim sCode As String
'
'   MouseCursor 13
'   SetMdiReportsize MdiSect
'   If Len(Trim(txtPrt)) = 0 Then txtPrt = "ALL"
'   If txtPrt <> "ALL" Then sPart = Compress(txtPrt)
'   If txtCls <> "ALL" Then sClass = Compress(txtCls)
'   If txtCde <> "ALL" Then sCode = Compress(txtCde)
'
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
'   sCustomReport = GetCustomReport("prdin01.rpt")
'   On Error GoTo DiaErr1
'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
'   sSql = "{PartTable.PARTREF} Like '" & sPart & "*' " _
'          & "AND {PartTable.PACLASS} like '" & sClass & "*' "
'   If optCode.value = vbUnchecked Then
'      sSql = sSql & "AND {PartTable.PAPRODCODE} like '" & sCode & "*' "
'   Else
'      sSql = sSql & "AND {PartTable.PAPRODCODE}='' "
'   End If
'   If optQoh.value = vbChecked Then _
'                     sSql = sSql & "AND {PartTable.PAQOH}>0 "
'   If typ(1).value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>1 "
'   If typ(2).value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>2 "
'   If typ(3).value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>3 "
'   If typ(4).value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>4 "
'   If typ(5).value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>5 "
'   If typ(6).value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>6 "
'   If typ(7).value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>7 "
'   If typ(8).value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>8 "
'   MdiSect.Crw.SelectionFormula = sSql
'   sSql = "Includes='Includes Part(s) " & txtPrt & "... Part Type(s) "
'   For iList = 1 To 7
'      If typ(iList).value = vbChecked Then sSql = sSql & str(iList) & ","
'   Next
'   If typ(iList).value = vbChecked Then sSql = sSql & str(iList) & ","
'   iList = Len(sSql)
'   sSql = Left(sSql, iList - 1) & " And Classe(s) " & txtCls & "...'"
'   MdiSect.Crw.Formulas(1) = sSql
'
'   If optDsc.value = vbUnchecked Then
'      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
'   Else
'      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
'   End If
'   If OptCmt.value = vbUnchecked Then
'      MdiSect.Crw.SectionFormat(1) = "GROUPFTR.0.0;F;;;"
'      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.0.1;F;;;"
'   Else
'      MdiSect.Crw.SectionFormat(1) = "GROUPFTR.0.0;T;;;"
'      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.0.1;T;;;"
'   End If
'   SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub




























Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   Dim iList As Integer
   MouseCursor 13
   For iList = 1 To 8
      If typ(iList).Value = vbChecked Then Exit For
   Next
   If iList = 9 Then
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
   For iList = 1 To 8
      If typ(iList).Value = vbChecked Then Exit For
   Next
   If iList = 9 Then
      MouseCursor 0
      MsgBox "You Need At Least One Part Type.", vbInformation, Caption
      On Error Resume Next
      typ(1).SetFocus
   Else
      PrintReport
   End If
   
End Sub

Private Sub txtCde_LostFocus()
   txtCde = CheckLen(txtCde, 6)
   If Len(txtCde) = 0 Then txtCde = "ALL"
   
End Sub


Private Sub txtCls_LostFocus()
   txtCls = CheckLen(txtCls, 4)
   If Len(txtCls) = 0 Then txtCls = "ALL"
   
End Sub


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
   If Len(txtPrt) = 0 Then
        txtPrt = "ALL"
   Else
        txtPrt = UCase(txtPrt)
   End If
   
End Sub


Private Sub typ_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
