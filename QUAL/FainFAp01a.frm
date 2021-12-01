VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form FainFAp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "First Article Inspection Report"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "FainFAp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   6120
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "8"
      Text            =   " "
      ToolTipText     =   "Select Revision From List"
      Top             =   1080
      Width           =   945
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "FainFAp01a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "FainFAp01a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   11
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Number"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Drawing/Documents"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Tag             =   " "
      Top             =   1800
      Width           =   1425
   End
End
Attribute VB_Name = "FainFAp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'4/19/06 Stored procedures Rejection Tags
'4/19/06 Stored procedures First Article Inspection
Option Explicit
Dim bGoodReport As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetReport() As Byte
   Dim RdoRep As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FA_REF,FA_NUMBER,FA_DESCRIPTION FROM " _
          & "FahdTable WHERE FA_REF='" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRep, ES_FORWARD)
   If bSqlRows Then
      With RdoRep
         cmbPrt = "" & Trim(!FA_NUMBER)
         lblDsc = "" & Trim(!FA_DESCRIPTION)
         GetReport = 1
         ClearResultSet RdoRep
      End With
      FillRevisions
   Else
      GetReport = 0
      lblDsc = "*** Report Wasn't Found ***"
   End If
   Set RdoRep = Nothing
   Exit Function
   
DiaErr1:
   bGoodReport = 0
   sProcName = "getreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillRevisions()
   cmbRev.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT FA_REF,FA_REVISION FROM " _
          & "FahdTable WHERE FA_REF='" & Compress(cmbPrt) & "'"
   LoadComboBox cmbRev
   If cmbRev.ListCount > 0 Then cmbRev = cmbRev.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillrevs "
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub cmbPrt_Click()
   bGoodReport = GetReport()
   
End Sub


Private Sub cmbPrt_LostFocus()
   bGoodReport = GetReport()
   
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
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MdiSect.lblBotPanel = Caption
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
   Set FainFAp01a = Nothing
   
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
   
   FillDocuments
   On Error GoTo DiaErr1
   sCustomReport = GetCustomReport("quafa01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   If optDet.value = vbChecked Then
      aFormulaName.Add "ShowDocs"
      aFormulaValue.Add CStr("'1'")
   Else
      aFormulaName.Add "ShowDocs"
      aFormulaValue.Add CStr("'0'")
   End If
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   sSql = "{FahdTable.FA_REF}='" & Compress(cmbPrt) & "' " _
          & "AND {FahdTable.FA_REVISION}='" & cmbRev & "'"
   cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport1()
   Dim sBook As String
   MouseCursor 13
   
   FillDocuments
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("quafa01")
   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   If optDet.value = vbChecked Then
      MdiSect.Crw.Formulas(1) = "ShowDocs='1'"
   Else
      MdiSect.Crw.Formulas(1) = "ShowDocs='0'"
   End If
   sSql = "{FahdTable.FA_REF}='" & Compress(cmbPrt) & "' " _
          & "AND {FahdTable.FA_REVISION}='" & cmbRev & "'"
   MdiSect.Crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   
End Sub

Private Sub GetOptions()
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillFirstArticles"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillDocuments()
   Dim RdoDoc As ADODB.Recordset
   Dim iList As Integer
   Dim sDoc(11, 4) As String
   
   If optDet.value = vbChecked Then
      sSql = "SELECT * FROM FadcTable WHERE FA_DOCNUMBER='" _
             & Compress(cmbPrt) & "' AND FA_DOCREVISION='" _
             & cmbRev & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_FORWARD)
      If bSqlRows Then
         With RdoDoc
            Do Until .EOF
               iList = iList + 1
               sDoc(iList, 1) = "" & Trim(!FA_DOCDESCRIPTION)
               sDoc(iList, 2) = "" & Trim(!FA_DOCSHEET)
               sDoc(iList, 3) = "" & Trim(!FA_DOCCHANGE)
               .MoveNext
            Loop
            ClearResultSet RdoDoc
         End With
      End If
   End If
   sSql = "SELECT * FROM EsReportFarp01 WHERE FA_RECORD=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_KEYSET)
   If bSqlRows Then
      With RdoDoc
         !FA_DOCNUMBER = Compress(cmbPrt)
         !FA_DOCDESCRIPTION1 = sDoc(1, 1)
         !FA_DOCSHEET1 = sDoc(1, 2)
         !FA_DOCCHANGE1 = sDoc(1, 3)
         
         !FA_DOCDESCRIPTION2 = sDoc(2, 1)
         !FA_DOCSHEET2 = sDoc(2, 2)
         !FA_DOCCHANGE2 = sDoc(2, 3)
         
         !FA_DOCDESCRIPTION3 = sDoc(3, 1)
         !FA_DOCSHEET3 = sDoc(3, 2)
         !FA_DOCCHANGE3 = sDoc(3, 3)
         
         !FA_DOCDESCRIPTION4 = sDoc(4, 1)
         !FA_DOCSHEET4 = sDoc(4, 2)
         !FA_DOCCHANGE4 = sDoc(4, 3)
         
         !FA_DOCDESCRIPTION5 = sDoc(5, 1)
         !FA_DOCSHEET5 = sDoc(5, 2)
         !FA_DOCCHANGE5 = sDoc(5, 3)
         
         !FA_DOCDESCRIPTION6 = sDoc(6, 1)
         !FA_DOCSHEET6 = sDoc(6, 2)
         !FA_DOCCHANGE6 = sDoc(6, 3)
         
         !FA_DOCDESCRIPTION7 = sDoc(7, 1)
         !FA_DOCSHEET7 = sDoc(7, 2)
         !FA_DOCCHANGE7 = sDoc(7, 3)
         
         !FA_DOCDESCRIPTION8 = sDoc(8, 1)
         !FA_DOCSHEET8 = sDoc(8, 2)
         !FA_DOCCHANGE8 = sDoc(8, 3)
         
         !FA_DOCDESCRIPTION9 = sDoc(9, 1)
         !FA_DOCSHEET9 = sDoc(9, 2)
         !FA_DOCCHANGE9 = sDoc(9, 3)
         
         !FA_DOCDESCRIPTION10 = sDoc(10, 1)
         !FA_DOCSHEET10 = sDoc(10, 2)
         !FA_DOCCHANGE10 = sDoc(10, 3)
         .Update
         ClearResultSet RdoDoc
      End With
   End If
   Set RdoDoc = Nothing
   
End Sub
