VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ppiESp04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Of Formulae"
   ClientHeight    =   2895
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2895
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ppiESp04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Tag             =   "3"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ppiESp04a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ppiESp04a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   4
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
      FormDesignHeight=   2895
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Only Blank Work Centers)"
      Height          =   288
      Index           =   4
      Left            =   4440
      TabIndex        =   11
      Top             =   1920
      Width           =   2148
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1920
      TabIndex        =   10
      Top             =   1440
      Width           =   3012
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center(s)"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Formulae Without WorkCenters"
      Height          =   492
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1788
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   4440
      TabIndex        =   7
      Top             =   1080
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Tag             =   " "
      Top             =   1680
      Width           =   1428
   End
End
Attribute VB_Name = "ppiESp04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'1/18/06 New (PROPLA Custom)
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   sSql = "SELECT DISTINCT WCNREF,WCNNUM FROM WcntTable " _
          & "ORDER BY WCNREF"
   LoadComboBox cmbWcn
   If cmbWcn.ListCount > 0 Then
      cmbWcn = cmbWcn.List(0)
      GetWorkCenter
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub cmbWcn_Click()
   GetWorkCenter
   
End Sub

Private Sub cmbWcn_LostFocus()
   cmbWcn = Trim(cmbWcn)
   If cmbWcn = "" Then
      cmbWcn = "ALL"
      lblDsc = "All Work Centers With Formulae"
   Else
      If cmbWcn <> "NONE" Then GetWorkCenter
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
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
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
   Set ppiESp04a = Nothing
   
End Sub




Private Sub PrintReport()
   MouseCursor 13
   Dim sFormula As String
   On Error GoTo DiaErr1
   
   If cmbWcn = "ALL" Or cmbWcn = "NONE" Then sFormula = "" _
               Else sFormula = Compress(cmbWcn)
   
   SetMdiReportsize MDISect
   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   If optDet.value = vbUnchecked Then
      MDISect.Crw.Formulas(1) = "Includes='Work Center(s) " & cmbWcn & "...'"
   Else
      MDISect.Crw.Formulas(1) = "Includes='Formulae With No Work Center...'"
   End If
   MDISect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   sCustomReport = GetCustomReport("enges07a")
   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   If optDet.value = vbUnchecked Then
      sSql = "{EsfrTable.FORMULA_CENTER} LIKE '" & sFormula & "*' " _
             & "AND {EsfrTable.FORMULA_CENTER}<>''"
   Else
      sSql = "{EsfrTable.FORMULA_CENTER}='' "
   End If
   MDISect.Crw.SelectionFormula = sSql
   SetCrystalAction Me
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
   On Error Resume Next
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   
End Sub

Private Sub optDet_Click()
   If optDet.value = vbChecked Then
      cmbWcn.Enabled = False
      cmbWcn = "NONE"
      lblDsc = "*** Formulae With No Work Center ***"
   Else
      cmbWcn.Enabled = True
      cmbWcn = "ALL"
      lblDsc = "All Work Centers With Formulae"
   End If
   
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





Private Sub GetWorkCenter()
   Dim RdoWcn As ADODB.Recordset
   If cmbWcn <> "ALL" And cmbWcn <> "NONE" Then
      sSql = "SELECT WCNREF,WCNDESC FROM WcntTable WHERE WCNREF='" & Compress(cmbWcn) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoWcn, ES_FORWARD)
      If bSqlRows Then
         lblDsc = "" & Trim(RdoWcn!WCNDESC)
         RdoWcn.Cancel
         ClearResultSet RdoWcn
      Else
         lblDsc = "Range Of Work Centers"
      End If
   End If
   Set RdoWcn = Nothing
   
End Sub
