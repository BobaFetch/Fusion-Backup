VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InvcINp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revised Parts"
   ClientHeight    =   3180
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7215
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   3075
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "InvcINp03a.frx":07AE
      Height          =   315
      Left            =   5280
      Picture         =   "InvcINp03a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1440
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtCls 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Leading Char Search  (*  In Front Is A Legal Wild Card)"
      Top             =   1440
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   2400
      Width           =   735
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
         Picture         =   "InvcINp03a.frx":0E32
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
         Picture         =   "InvcINp03a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   2760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3180
      FormDesignWidth =   7215
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   5
      Left            =   5760
      TabIndex        =   18
      Top             =   2160
      Width           =   1308
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   4
      Left            =   5760
      TabIndex        =   17
      Top             =   1800
      Width           =   1308
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   3
      Left            =   5760
      TabIndex        =   16
      Top             =   1440
      Width           =   1308
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions?"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions?"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   1785
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   1785
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revisions From Date"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1785
   End
End
Attribute VB_Name = "InvcINp03a"
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
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

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
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then bOnLoad = 0
   MouseCursor 0
   FillPartCombo cmbPrt
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtPrt = "ALL"
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
   Set InvcINp03a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sDate As String
   Dim sClass As String
   Dim sPart As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   If Len(Trim(cmbPrt)) = 0 Then cmbPrt = "ALL"
   If Len(Trim(txtCls)) = 0 Then txtCls = "ALL"
   If cmbPrt = "ALL" Then
      sPart = ""
   Else
      sPart = cmbPrt
   End If
   If txtCls = "ALL" Then
      sClass = ""
   Else
      sClass = txtCls
   End If
   If Not IsDate(txtDte) Then
      sDate = "1995,01,01"
   Else
      sDate = Format(txtDte, "yyyy,mm,dd")
   End If
   MouseCursor 13
   
   On Error GoTo DiaErr1
   sCustomReport = GetCustomReport("prdin03")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.ShowGroupTree False
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowDesc"
    aFormulaName.Add "ShowExDesc"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Revised From " & CStr(txtDte _
                        & " And Classe(s) " & txtCls) & "...'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add optDsc.Value
    aFormulaValue.Add optExt.Value
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{PartTable.PARTREF} LIKE '" & sPart & "*' " _
          & "AND {PartTable.PACLASS} LIKE '" & sClass & "*' " _
          & "AND {PartTable.PAREVDATE}>= Date(" & sDate & ")"
   sSql = sSql & " and {PartTable.PATOOL} = 0"
   cCRViewer.SetReportSelectionFormula (sSql)
   'cCRViewer.CRViewerSize Me
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



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(str(optDsc.Value)) _
              & Trim(str(optDsc.Value)) _
              & Trim(txtCls)
   SaveSetting "Esi2000", "EsiProd", "in03", sOptions
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "in03", sOptions)
   iList = Len(Trim(sOptions))
   If iList > 0 Then
      optDsc.Value = Val(Mid(sOptions, 1, 1))
      optExt.Value = Val(Mid(sOptions, 2, 1))
      txtCls = Mid(sOptions, 3, iList - 2)
   Else
      txtPrt = "ALL"
      txtCls = "ALL"
      optDsc.Value = vbChecked
      optExt.Value = vbChecked
   End If
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub txtCls_LostFocus()
   txtCls = CheckLen(txtCls, 4)
   If Len(txtCls) = 0 Then txtCls = "ALL"
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDte_LostFocus()
   If Len(txtDte) Then txtDte = CheckDateEx(txtDte)
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Len(txtPrt) = 0 Then txtPrt = "ALL"
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) = 0 Then cmbPrt = "ALL"
End Sub

