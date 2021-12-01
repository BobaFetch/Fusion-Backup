VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InvcINp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parts By Part Description"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   3075
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "InvcINp02a.frx":07AE
      Height          =   315
      Left            =   5280
      Picture         =   "InvcINp02a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   960
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6720
      TabIndex        =   22
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "InvcINp02a.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "InvcINp02a.frx":0FBC
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox typ 
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   10
      Top             =   1680
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   5040
      TabIndex        =   9
      Top             =   1680
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   8
      Top             =   1680
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   7
      Top             =   1680
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   6
      Top             =   1680
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   5
      Top             =   1680
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   4
      Top             =   1680
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox typ 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.TextBox txtCls 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Tag             =   "3"
      Text            =   "ALL"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Tag             =   "2"
      Text            =   "ALL"
      ToolTipText     =   "The Search Is By Uppercase (* Leading Wild Card)"
      Top             =   990
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   11
      Top             =   2280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox OptCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   2520
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
      Left            =   6360
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   7860
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL)"
      Height          =   285
      Index           =   7
      Left            =   5760
      TabIndex        =   26
      Top             =   1320
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL)"
      Height          =   285
      Index           =   6
      Left            =   5760
      TabIndex        =   25
      Top             =   1080
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   21
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types?"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   20
      Top             =   1680
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Class(es)"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions?"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   2520
      Width           =   2115
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Description(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   990
      Width           =   1695
   End
End
Attribute VB_Name = "InvcINp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   Dim sClass As String * 4
   sClass = txtCls
   
   'Save by Menu Option
   For iList = 1 To 8
      sOptions = sOptions & Trim(str(typ(iList).Value))
   Next
   sOptions = sOptions & Trim(str(optDsc.Value)) _
              & Trim(str(OptCmt.Value)) & sClass
   SaveSetting "Esi2000", "EsiProd", "in02", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "in02", sOptions)
   If Len(sOptions) > 0 Then
      For iList = 1 To 8
         typ(iList) = Mid$(sOptions, iList, 1)
      Next
      optDsc.Value = Val(Mid(sOptions, iList, 1))
      OptCmt.Value = Val(Mid(sOptions, iList + 1, 1))
      txtCls = Mid(sOptions, iList + 2, 4)
      If txtCls = "" Then txtCls = "ALL"
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.optDsc.Value = vbChecked
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
   FillCombo cmbPrt
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
   Set InvcINp02a = Nothing
   
End Sub

Public Sub FillCombo(cmbPrt As ComboBox)
   'Dim rdoPart As rdoResultset
   Dim rdoPart As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAQOH FROM PartTable ORDER BY PADESC,PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPart, ES_FORWARD)
   
   If bSqlRows Then
      With rdoPart
         While Not .EOF
            AddComboStr cmbPrt.hwnd, "" & Trim(!PADESC)
            .MoveNext
         Wend
         .Cancel
      End With
   End If
   Set rdoPart = Nothing
   cmbPrt.ListIndex = 0
   Exit Sub
   
DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Sub


Private Sub PrintReport()
   Dim iList As Integer
   Dim sPart As String
   Dim sClass As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   
   MouseCursor 13
   sPart = Trim(cmbPrt)
   If Len(sPart) = 0 Then
      cmbPrt = "ALL"
      sPart = ""
   Else
      If sPart = "ALL" Then sPart = ""
   End If
   
   sClass = Compress(txtCls)
   If Len(sClass) = 0 Then
      txtCls = "ALL"
      sClass = ""
   Else
      If sClass = "ALL" Then sClass = ""
   End If
   
   sCustomReport = GetCustomReport("prdin02.rpt")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.ShowGroupTree False
   On Error GoTo DiaErr1
   sSql = "UpperCase({PartTable.PADESC}) Like '" & UCase$(sPart) & "*' " _
          & "AND {PartTable.PACLASS} like '" & sClass & "*' "
   If typ(1).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>1 "
   If typ(2).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>2 "
   If typ(3).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>3 "
   If typ(4).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>4 "
   If typ(5).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>5 "
   If typ(6).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>6 "
   If typ(7).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>7 "
   If typ(8).Value = vbUnchecked Then sSql = sSql & "AND {PartTable.PALEVEL}<>8 "
   sSql = sSql & " AND {PartTable.PATOOL} = 0"
   cCRViewer.SetReportSelectionFormula (sSql)
   sSql = "'Includes Description(s) " & cmbPrt & "... Part Type(s) "
   For iList = 1 To 7
      If typ(iList).Value = vbChecked Then sSql = sSql & str(iList) & ","
   Next
   If typ(iList).Value = vbChecked Then sSql = sSql & str(iList) & ","
   iList = Len(sSql)
   sSql = Left(sSql, iList - 1) & " And Classe(s) " & txtCls & "...'"
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowDescription"
    aFormulaName.Add "ShowComments"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr(sSql)
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add optDsc.Value
    aFormulaValue.Add OptCmt.Value
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    cCRViewer.CRViewerSize Me
    cCRViewer.SetDbTableConnection
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
   If Len(txtPrt) = 0 Then txtPrt = "ALL"
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) = 0 Then cmbPrt = "ALL"
   
End Sub

Private Sub typ_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
