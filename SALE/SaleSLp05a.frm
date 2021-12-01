VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Order Register"
   ClientHeight    =   3195
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7125
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3195
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "SaleSLp05a.frx":0000
      DownPicture     =   "SaleSLp05a.frx":04F2
      Enabled         =   0   'False
      Height          =   372
      Left            =   0
      MaskColor       =   &H00000000&
      Picture         =   "SaleSLp05a.frx":09E4
      Style           =   1  'Graphical
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   0
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "SaleSLp05a.frx":0ED6
      DownPicture     =   "SaleSLp05a.frx":13C8
      Enabled         =   0   'False
      Height          =   372
      Left            =   0
      MaskColor       =   &H00000000&
      Picture         =   "SaleSLp05a.frx":18BA
      Style           =   1  'Graphical
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   384
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLp05a.frx":1DAC
      Style           =   1  'Graphical
      TabIndex        =   64
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   61
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "SaleSLp05a.frx":255A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "SaleSLp05a.frx":26D8
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox optitm 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   20
      Top             =   2400
      Width           =   615
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   25
      Left            =   6840
      TabIndex        =   58
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   24
      Left            =   6600
      TabIndex        =   56
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   23
      Left            =   6360
      TabIndex        =   54
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   22
      Left            =   6120
      TabIndex        =   52
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   21
      Left            =   5880
      TabIndex        =   50
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   20
      Left            =   5640
      TabIndex        =   48
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   19
      Left            =   5400
      TabIndex        =   46
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   18
      Left            =   5160
      TabIndex        =   44
      Top             =   1920
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   17
      Left            =   4920
      TabIndex        =   42
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   16
      Left            =   4680
      TabIndex        =   40
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   15
      Left            =   4440
      TabIndex        =   38
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   14
      Left            =   4200
      TabIndex        =   34
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   13
      Left            =   3960
      TabIndex        =   33
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   12
      Left            =   3720
      TabIndex        =   32
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   11
      Left            =   3480
      TabIndex        =   17
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   10
      Left            =   3240
      TabIndex        =   16
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   9
      Left            =   3000
      TabIndex        =   15
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   8
      Left            =   2760
      TabIndex        =   14
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   7
      Left            =   2520
      TabIndex        =   13
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   6
      Left            =   2280
      TabIndex        =   12
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1920
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   1800
      TabIndex        =   10
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   9
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   1920
      Width           =   375
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   375
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3195
      FormDesignWidth =   7125
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   5280
      TabIndex        =   63
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Sales Orders Items"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   60
      Top             =   2400
      Width           =   2028
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      Height          =   255
      Index           =   25
      Left            =   6840
      TabIndex        =   59
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   255
      Index           =   24
      Left            =   6600
      TabIndex        =   57
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   23
      Left            =   6360
      TabIndex        =   55
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      Height          =   255
      Index           =   22
      Left            =   6120
      TabIndex        =   53
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      Height          =   255
      Index           =   21
      Left            =   5880
      TabIndex        =   51
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      Height          =   255
      Index           =   20
      Left            =   5640
      TabIndex        =   49
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      Height          =   255
      Index           =   19
      Left            =   5400
      TabIndex        =   47
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   255
      Index           =   18
      Left            =   5160
      TabIndex        =   45
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   255
      Index           =   17
      Left            =   4920
      TabIndex        =   43
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      Height          =   255
      Index           =   16
      Left            =   4680
      TabIndex        =   41
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   255
      Index           =   15
      Left            =   4440
      TabIndex        =   39
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      Height          =   255
      Index           =   14
      Left            =   4200
      TabIndex        =   37
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      Height          =   255
      Index           =   13
      Left            =   3960
      TabIndex        =   36
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   255
      Index           =   12
      Left            =   3720
      TabIndex        =   35
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      Height          =   255
      Index           =   11
      Left            =   3480
      TabIndex        =   31
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      Height          =   255
      Index           =   10
      Left            =   3240
      TabIndex        =   30
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   29
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   28
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   27
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   26
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   25
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   24
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   23
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   19
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   18
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   6
      Top             =   1680
      Width           =   160
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Sales Orders:"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Start Date"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   915
   End
End
Attribute VB_Name = "SaleSLp05a"
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
Dim sIncludes As String
Dim sSelections As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
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
   MdiSect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim b As Integer
   Dim iList As Integer
   FormLoad Me
   FormatControls
   
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 5)
   b = 1
   
   For iList = 0 To 25
      b = b + 1
      With optTyp(iList)
         .TabIndex = b
         .ToolTipText = "Check Or Space Bar To Select"
      End With
      With lblAlp(iList)
         .Width = 160
         .Left = optTyp(iList).Left + 20
      End With
   Next
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
   Set SaleSLp05a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sBegDte As String
   Dim sEndDte As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   If Not IsDate(txtBeg) Then
      sBegDte = "1995,01,01"
   Else
      sBegDte = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEndDte = "2024,12,31"
   Else
      sEndDte = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   GetSelections
   If sSelections = "" Then
      MouseCursor 0
      MsgBox "Requires One Sales Order Type.", vbInformation, Caption
      Exit Sub
   End If
   On Error GoTo DiaErr1
   
  
  aFormulaName.Add "CompanyName"
  aFormulaName.Add "Includes"
  aFormulaName.Add "RequestBy"
  aFormulaName.Add "ShowItem"
  
  aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
  aFormulaValue.Add CStr("'Booked " & CStr(txtBeg _
                        & " To " & txtEnd & " Types: " & sIncludes) & "'")
  aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
  aFormulaValue.Add optItm.Value
  Set cCRViewer = New EsCrystalRptViewer
  cCRViewer.Init
  sCustomReport = GetCustomReport("sleco06")
  cCRViewer.SetReportFileName sCustomReport, sReportPath
  cCRViewer.SetReportTitle = sCustomReport
  cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
  
   sSql = "{SohdTable.SODATE} in Date(" & sBegDte _
          & ") to Date(" & sEndDte & ")" _
          & " AND {SoitTable.ITCANCELED} = 0"
   
   
   cCRViewer.SetReportSelectionFormula sSql & sSelections
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

Private Sub optDis_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub optitm_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optItm_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub optTyp_GotFocus(Index As Integer)
   lblAlp(Index).BorderStyle = 1
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim a As Integer
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "Sl06", sOptions)
   If Len(sOptions) > 0 Then
      For iList = 0 To 25
         a = a + 1
         optTyp(iList).Value = Val(Mid(sOptions, a, 1))
      Next
   End If
   
End Sub

Private Sub optTyp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optTyp_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optTyp_LostFocus(Index As Integer)
   lblAlp(Index).BorderStyle = 0
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Trim(txtEnd) <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub

Private Sub GetSelections()
   Dim bByte As Byte
   Dim iList As Integer
   'gets the checked selections
   sIncludes = ""
   On Error GoTo DiaErr1
   sSelections = "AND ("
   For iList = 0 To 24
      If optTyp(iList) = vbChecked Then
         sIncludes = sIncludes & lblAlp(iList) & ","
         If bByte Then sSelections = sSelections & " OR "
         bByte = True
         sSelections = sSelections & "{SohdTable.SOTYPE}='" _
                       & lblAlp(iList) & "' "
      End If
   Next
   If optTyp(iList) = vbChecked Then
      sIncludes = sIncludes & lblAlp(iList)
      If bByte Then sSelections = sSelections & " OR "
      sSelections = sSelections & "{SohdTable.SOTYPE}='" _
                    & lblAlp(iList) & "'"
   End If
   If bByte = True Then sSelections = sSelections & ")" Else sSelections = ""
   If Right(sIncludes, 1) = "," Then
      iList = Len(sIncludes) - 1
      sIncludes = Left(sIncludes, iList)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getselect"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   For iList = 0 To 25
      sOptions = sOptions & RTrim(optTyp(iList).Value)
   Next
   'Save by Menu Option
   SaveSetting "Esi2000", "EsiSale", "sl06", Trim(sOptions)
   
End Sub
