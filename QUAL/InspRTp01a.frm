VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inspection Reports (Report)"
   ClientHeight    =   3870
   ClientLeft      =   1770
   ClientTop       =   1140
   ClientWidth     =   7140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3870
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbRes 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   13
      Tag             =   "8"
      ToolTipText     =   "Select Resposibility Code From List"
      Top             =   2640
      Width           =   1675
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   12
      Tag             =   "8"
      ToolTipText     =   "Select Characteristic Code From List"
      Top             =   1800
      Width           =   1675
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optFrm 
      Caption         =   "From"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   1
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optDis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         Picture         =   "InspRTp01a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   560
         Picture         =   "InspRTp01a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6120
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3870
      FormDesignWidth =   7140
   End
   Begin VB.ComboBox cmbTag 
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Tag Number"
      Top             =   900
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Responsibility"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discrepancy"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   18
      Top             =   1800
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   4200
      TabIndex        =   17
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   4200
      TabIndex        =   16
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   3000
      Width           =   2940
   End
   Begin VB.Label lblCde 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   2160
      Width           =   2940
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Detail"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   2025
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tag Type"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspection Report"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   915
      Width           =   1725
   End
End
Attribute VB_Name = "InspRTp01a"
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

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(str(optDsc.Value))
   SaveSetting "Esi2000", "EsiQual", "rj01", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiQual", "rj01", sOptions)
   If Len(sOptions) > 0 Then
      optDsc.Value = Val(Left(sOptions, 1))
   Else
      optDsc.Value = vbChecked
   End If
   
End Sub



Private Sub cmbTag_Click()
   GetTag
   
End Sub

Private Sub cmbTag_LostFocus()
   cmbTag = CheckLen(cmbTag, 12)
   GetTag
   
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


Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillRejectionTags"
   LoadComboBox cmbTag
   If cmbTag.ListCount > 0 Then
      If optFrm.Value = vbChecked Then
         cmbTag = InspRTe01b.lblTag
         Unload InspRTe01b
      Else
         cmbTag = cmbTag.List(0)
      End If
      GetTag
   End If
   
   sSql = "Qry_FillDescripancyCodes"
   LoadComboBox cmbCde
   cmbCde = ""
   
   sSql = "Qry_FillReasonCodes"
   LoadComboBox cmbRes
   cmbRes = ""
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetRespCode()
   Dim RdoRsp As ADODB.Recordset
   If Trim(cmbRes) = "" Then
      lblRes = "*** ALL ***"
      Exit Sub
   End If
   On Error GoTo DiaErr1
   sSql = "SELECT RESREF,RESNUM,RESDESC FROM RjrsTable " _
          & "WHERE RESREF='" & Compress(cmbRes) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRsp, ES_FORWARD)
   If bSqlRows Then
      With RdoRsp
         lblRes = "" & Trim(!RESDESC)
         ClearResultSet RdoRsp
      End With
   Else
      lblRes = "*** ALL ***"
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getrespco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetCharaCode()
   Dim RdoCha As ADODB.Recordset
   
   If Trim(cmbCde) = "" Then
      lblCde = "*** ALL ***"
      Exit Sub
   End If
   
   On Error GoTo DiaErr1
   sSql = "SELECT CDEREF,CDENUM,CDEDESC FROM RjcdTable " _
          & "WHERE CDEREF='" & Compress(cmbCde) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCha, ES_FORWARD)
   If bSqlRows Then
      With RdoCha
         lblCde = "" & Trim(!CDEDESC)
         ClearResultSet RdoCha
      End With
   Else
      lblCde = "*** ALL ***"
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getcharco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillCombo
      GetCharaCode
      GetRespCode
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
   If optFrm.Value = vbChecked Then
      optFrm = vbUnchecked
      InspRTe01b.optNew.Value = vbUnchecked
      InspRTe01b.Caption = "Revise " & InspRTe01b.Caption
      InspRTe01b.lblTag = cmbTag
      InspRTe01b.lblType = lblTyp & " Tag "
      InspRTe01b.Show
   End If
   
End Sub

Private Sub cmbCde_Click()
   If Not bOnLoad Then GetCharaCode
End Sub


Private Sub cmbCde_LostFocus()
   GetCharaCode
End Sub

Private Sub cmbRes_Click()
   If Not bOnLoad Then GetRespCode
   
End Sub


Private Sub cmbRes_LostFocus()
   GetRespCode
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If optFrm.Value = vbUnchecked Then FormUnload
   Set InspRTp01a = Nothing
   
End Sub




Private Sub PrintReport()
   Dim sError$
   Dim sTagNum As String
   Dim sDesp As String
   Dim sRes As String
   Dim strRptFi
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
      
   MouseCursor 13
   'SetMdiReportsize MdiSect
   sTagNum = Compress(cmbTag)
   sDesp = Compress(cmbCde)
   sRes = Compress(cmbRes)
   On Error GoTo DiaErr1
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   
   If ES_CUSTOM = "JEVCO" Then
      If Left(lblTyp, 1) = "V" Then
         cCRViewer.SetReportFileName sReportPath, "jevrj01v.rpt"
      Else
         cCRViewer.SetReportFileName sReportPath, "jevrj01c.rpt"
      End If
   Else
      If Left(lblTyp, 1) = "V" Then
         sCustomReport = GetCustomReport("quarj01v")
         cCRViewer.SetReportFileName sCustomReport, sReportPath
      Else
         sCustomReport = GetCustomReport("quarj01c")
         cCRViewer.SetReportFileName sCustomReport, sReportPath
      End If
      cCRViewer.SetReportTitle = sCustomReport
   End If
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "ShowDetail"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add optDsc.Value
   
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   'sSql = "{RjhdTable.REJREF}= '" & sTagNum & "' "
   
   sSql = "{RjhdTable.REJREF} LIKE '" & sTagNum & "*' AND " _
         & "{RjitTable.RITCHARCODE} LIKE '" & sDesp & "*' AND " _
         & "{RjitTable.RITRESPCODE} LIKE '" & sRes & "*'"
         
   cCRViewer.SetReportSelectionFormula sSql
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
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


Private Sub lblTyp_Change()
   If Left(lblTyp, 6) = "*** Ta" Then
      lblTyp.ForeColor = ES_RED
   Else
      lblTyp.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optDis_Click()
   MouseCursor 13
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFrm_Click()
   'never visible - called from InspRTe01b
   
End Sub

Private Sub optPrn_Click()
   MouseCursor 13
   PrintReport
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetTag()
   Dim RdoTag As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetRejectionTag '" & Compress(cmbTag) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTag, ES_FORWARD)
   If bSqlRows Then
      With RdoTag
         lblTyp.Width = 1575
         cmbTag = "" & Trim(!REJNUM)
         Select Case !REJTYPE
            Case "C" 'Customer
               lblTyp = "Customer"
            Case "I" 'Internal
               lblTyp = "Internal"
            Case "V" 'Vendor
               lblTyp = "Vendor"
            Case "M" 'MRB
               lblTyp = "MRB"
         End Select
         ClearResultSet RdoTag
      End With
   Else
      lblTyp.Width = 2075
      lblTyp = "*** ALL ***"
            
   End If
   Set RdoTag = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "gettag"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
