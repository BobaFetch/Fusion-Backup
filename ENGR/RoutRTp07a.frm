VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTp07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pictures By Routing Operation"
   ClientHeight    =   2925
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2925
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbRte 
      Height          =   288
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Routings And Operations With Pictures"
      Top             =   960
      WhatsThisHelpID =   100
      Width           =   3345
   End
   Begin VB.ComboBox cmbOpno 
      Height          =   288
      Left            =   6360
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Contains Routings And Operations With Pictures"
      Top             =   960
      WhatsThisHelpID =   100
      Width           =   948
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTp07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   1920
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "RoutRTp07a.frx":07AE
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
         Picture         =   "RoutRTp07a.frx":092C
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
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2925
      FormDesignWidth =   7455
   End
   Begin VB.Label txtDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1800
      TabIndex        =   13
      Top             =   1320
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation"
      Height          =   372
      Index           =   4
      Left            =   5400
      TabIndex        =   12
      Top             =   960
      Width           =   912
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   288
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   912
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing "
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   1272
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "OP Comments"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1788
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Tag             =   " "
      Top             =   1680
      Width           =   1428
   End
End
Attribute VB_Name = "RoutRTp07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'7/24/06 New
'8/14/06 Added To Prod/Shop Floor
Option Explicit
Dim bOnLoad As Byte
Dim bGoodRout As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   sSql = "SELECT DISTINCT OPREF,OPNO,RTNUM FROM RtpcTable,RthdTable " _
          & "WHERE (OPREF=RTREF AND OPPICTURE IS NOT NULL)"
   LoadComboBox cmbRte, 1
   If cmbRte.ListCount = 0 Then
      MsgBox "There Are No Operations With Pictures To Report.", _
         vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetRouting() As Byte
   Dim RdoRte As ADODB.Recordset
   cmbOpno.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT RTREF,RTNUM,RTDESC FROM RthdTable WHERE " _
          & "RTREF='" & Compress(cmbRte) & " '"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         cmbRte = "" & Trim(!RTNUM)
         txtDsc = "" & Trim(!RTDESC)
         GetRouting = 1
         .Cancel
      End With
      ClearResultSet RdoRte
   Else
      txtDsc = "*** Routing Isn't Listed ***"
      GetRouting = 0
   End If
   Set RdoRte = Nothing
   If GetRouting Then FillOperations
   Exit Function
   
DiaErr1:
   sProcName = "getrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub cmbOpno_LostFocus()
   cmbOpno = Format(Abs(Val(cmbOpno)), "000")
   Dim bByte As Byte
   Dim iList As Integer
   cmbOpno = Format(Abs(Val(cmbOpno)), "000")
   If cmbOpno.ListCount > 0 Then
      For iList = 0 To cmbOpno.ListCount - 1
         If cmbOpno.List(iList) = cmbOpno Then bByte = 1
      Next
      If bByte = 0 Then
         cmbOpno = cmbOpno.List(0)
         MsgBox "Select An Operation With A Picture From The List.", _
            vbInformation, Caption
      End If
   Else
      MsgBox "No Operations With A Picture Found.", _
         vbInformation, Caption
   End If
   
End Sub


Private Sub cmbRte_Click()
   bGoodRout = GetRouting()
   
End Sub


Private Sub cmbRte_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   If cmbRte.ListCount > 0 Then
      For iList = 0 To cmbRte.ListCount - 1
         If cmbRte = cmbRte.List(iList) Then bByte = 1
      Next
      If bByte = 0 Then
         MsgBox "Select The Routing From The List.", _
            vbInformation, Caption
         cmbRte = cmbRte.List(0)
         bGoodRout = GetRouting()
      End If
   End If
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
   Set RoutRTp07a = Nothing
   
End Sub




Private Sub PrintReport()
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim cCRViewer As EsCrystalRptViewer
   
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("engrt07")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "OPNO"
   aFormulaName.Add "RequestBy"
   'aFormulaName.Add "ShowDetails"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & cmbOpno & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   'aFormulaValue.Add optDet.value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RtpcTable.OPREF} = '" & Compress(cmbRte) & "' AND {RtpcTable.OPNO} =" _
          & Val(cmbOpno) & " "
   
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


'Private Sub PrintReport()
'   MouseCursor 13
'   On Error GoTo DiaErr1
'   SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "OPNO='" & cmbOpno & "'"
'   MDISect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
'   sCustomReport = GetCustomReport("engrt07")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   sSql = "{RtpcTable.OPREF} = '" & Compress(cmbRte) & "' AND {RtpcTable.OPNO} =" _
'          & Val(cmbOpno) & " "
'   If optDet.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.0;T;;;"
'   End If
'   MDISect.Crw.SelectionFormula = sSql
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













Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   On Error Resume Next
   SaveSetting "Esi2000", "EsiEngr", "rt07", Trim(str(optDet.value))
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "rt07", sOptions)
   If Len(sOptions) > 0 Then _
          optDet.value = Val(sOptions)
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   If bGoodRout = 1 Then _
                  PrintReport Else MsgBox "Requires A Valid Selection.", _
                  vbInformation, Caption
   
End Sub


Private Sub optPrn_Click()
   If bGoodRout = 1 Then _
                  PrintReport Else MsgBox "Requires A Valid Selection.", _
                  vbInformation, Caption
   
End Sub





Private Sub FillOperations()
   On Error GoTo DiaErr1
   cmbOpno.Clear
   sSql = "SELECT OPREF,OPNO FROM RtpcTable WHERE (OPREF='" _
          & Compress(cmbRte) & " ' AND OPPICTURE IS NOT NULL)"
   LoadNumComboBox cmbOpno, "000", 1
   Exit Sub
   
DiaErr1:
   sProcName = "filloperations"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtDsc_Change()
   If Left(txtDsc, 6) = "*** Ro" Then _
           txtDsc.ForeColor = ES_RED Else txtDsc.ForeColor = vbBlack
   
End Sub
