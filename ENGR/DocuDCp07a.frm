VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCp07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documents Used On MO's"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCp07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.ComboBox cmbDoc 
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List - Contains Documents Assigned To MO's"
      Top             =   1080
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "DocuDCp07a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "DocuDCp07a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6000
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   7095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Tag             =   " "
      Top             =   2280
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Tag             =   " "
      Top             =   2040
      Width           =   1425
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   1440
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Tag             =   " "
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "Document"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "DocuDCp07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'7/22/04 New
Option Explicit
Dim bOnLoad As Byte
Dim bGoodDoc As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT RUNDLSDOCREF,DOREF,DONUM FROM " _
          & "RndlTable,DdocTable WHERE (RUNDLSDOCREF<>'' " _
          & "AND RUNDLSDOCREF=DOREF) ORDER BY RUNDLSDOCREF"
   LoadComboBox cmbDoc, 1
   If cmbDoc.ListCount > 0 Then cmbDoc = cmbDoc.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbDoc_Click()
   bGoodDoc = GetThisDocument()
   
End Sub

Private Sub cmbDoc_LostFocus()
   cmbDoc = CheckLen(cmbDoc, 30)
   bGoodDoc = GetThisDocument()
   
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
   Set DocuDCp07a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sBook As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   
   On Error GoTo DiaErr1
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDesc"
   aFormulaName.Add "ShowExDesc"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbDoc) & "...'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optDsc.value
   aFormulaValue.Add OptExt.value
   
   sCustomReport = GetCustomReport("engdc07")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{RndlTable.RUNDLSDOCREF} = '" & Compress(cmbDoc) & "' "
   
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
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim$(str$(optDsc.value)) & Trim$(str$(OptExt.value))
   SaveSetting "Esi2000", "EsiEngr", "dc07", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "dc07", sOptions)
   If Len(sOptions) Then
      optDsc.value = Val(Left$(sOptions, 1))
      OptExt.value = Val(Right$(sOptions, 1))
   Else
      optDsc.value = vbChecked
      OptExt.value = vbChecked
   End If
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 5) = "*** D" Then _
           lblDsc.ForeColor = ES_RED Else _
           lblDsc.ForeColor = vbBlack
   
End Sub

Private Sub optDis_Click()
   If lblDsc.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Document.", _
         vbInformation, Caption
   Else
      PrintReport
   End If
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   If lblDsc.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Document.", _
         vbInformation, Caption
   Else
      PrintReport
   End If
   
End Sub

Private Function GetThisDocument() As Byte
   Dim RdoDoc As ADODB.Recordset
   sSql = "SELECT DOREF,DONUM,DODESCR FROM DdocTable " _
          & "WHERE DOREF='" & Compress(cmbDoc) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_FORWARD)
   If bSqlRows Then
      With RdoDoc
         lblDsc = "" & Trim(!DODESCR)
         GetThisDocument = 1
         ClearResultSet RdoDoc
      End With
   Else
      lblDsc = "*** Document Wasn't Found ***"
      GetThisDocument = 0
   End If
   Set RdoDoc = Nothing
   Exit Function
DiaErr1:
   sProcName = "getthisdocu"
   GetThisDocument = 0
   lblDsc = "*** Document Wasn't Found ***"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
