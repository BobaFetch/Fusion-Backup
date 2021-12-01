VERSION 5.00
Begin VB.Form ShopSHp16a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Backlog with MO Status"
   ClientHeight    =   3765
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3765
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   13
      Left            =   2040
      TabIndex        =   43
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   14
      Left            =   2280
      TabIndex        =   42
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   15
      Left            =   2520
      TabIndex        =   41
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   16
      Left            =   2760
      TabIndex        =   40
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   17
      Left            =   3000
      TabIndex        =   39
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   18
      Left            =   3240
      TabIndex        =   38
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   19
      Left            =   3480
      TabIndex        =   37
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   20
      Left            =   3720
      TabIndex        =   36
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   21
      Left            =   3960
      TabIndex        =   35
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   22
      Left            =   4200
      TabIndex        =   34
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   23
      Left            =   4440
      TabIndex        =   33
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   24
      Left            =   4680
      TabIndex        =   32
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   25
      Left            =   4920
      TabIndex        =   31
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      BackColor       =   &H00000000&
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   30
      Top             =   2280
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   29
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   2520
      TabIndex        =   28
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   2760
      TabIndex        =   27
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   26
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   3240
      TabIndex        =   25
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   6
      Left            =   3480
      TabIndex        =   24
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   7
      Left            =   3720
      TabIndex        =   23
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   8
      Left            =   3960
      TabIndex        =   22
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   9
      Left            =   4200
      TabIndex        =   21
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   10
      Left            =   4440
      TabIndex        =   20
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   11
      Left            =   4680
      TabIndex        =   19
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   12
      Left            =   4920
      TabIndex        =   18
      Top             =   2280
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.ComboBox cmbDivision 
      Height          =   315
      Left            =   2100
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "8"
      Top             =   1140
      Width           =   1095
   End
   Begin VB.ComboBox cmbThroughDate 
      Height          =   315
      Left            =   4020
      TabIndex        =   2
      Tag             =   "4"
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cmbFromDate 
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Tag             =   "4"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp16a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2100
      TabIndex        =   6
      Top             =   3360
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox cmbAsOfDate 
      Height          =   315
      Left            =   2100
      TabIndex        =   0
      Tag             =   "4"
      Top             =   300
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5580
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5580
      TabIndex        =   10
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp16a.frx":07AE
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
         Picture         =   "ShopSHp16a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2100
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "8"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      Height          =   255
      Index           =   13
      Left            =   2040
      TabIndex        =   70
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      Height          =   255
      Index           =   14
      Left            =   2280
      TabIndex        =   69
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   255
      Index           =   15
      Left            =   2520
      TabIndex        =   68
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      Height          =   255
      Index           =   16
      Left            =   2760
      TabIndex        =   67
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   255
      Index           =   17
      Left            =   3000
      TabIndex        =   66
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   255
      Index           =   18
      Left            =   3240
      TabIndex        =   65
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      Height          =   255
      Index           =   19
      Left            =   3480
      TabIndex        =   64
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      Height          =   255
      Index           =   20
      Left            =   3720
      TabIndex        =   63
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      Height          =   255
      Index           =   21
      Left            =   3960
      TabIndex        =   62
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      Height          =   255
      Index           =   22
      Left            =   4200
      TabIndex        =   61
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   23
      Left            =   4440
      TabIndex        =   60
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   255
      Index           =   24
      Left            =   4680
      TabIndex        =   59
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      Height          =   255
      Index           =   25
      Left            =   4920
      TabIndex        =   58
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   57
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   56
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   55
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   54
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   53
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   52
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   51
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   50
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      Height          =   255
      Index           =   8
      Left            =   3960
      TabIndex        =   49
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   48
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      Height          =   255
      Index           =   10
      Left            =   4440
      TabIndex        =   47
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   46
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   255
      Index           =   12
      Left            =   4920
      TabIndex        =   45
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Types"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   44
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   17
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   4
      Left            =   3300
      TabIndex        =   16
      Tag             =   " "
      Top             =   780
      Width           =   705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "For SOs scheduled from"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Tag             =   " "
      Top             =   780
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Tag             =   " "
      Top             =   3360
      Width           =   1725
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "As of Date"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Tag             =   " "
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1425
   End
End
Attribute VB_Name = "ShopSHp16a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/28/05 Changed date handling
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCst_Change()
   If cmbCst = "" Then cmbCst = "ALL"
End Sub

Private Sub cmbDivision_LostFocus()
   If cmbDivision = "" Then
      cmbDivision = "ALL"
      Exit Sub
   End If
   
   Dim b As Byte
   Dim iList As Integer
   cmbDivision = CheckLen(cmbDivision, 10)
   For iList = 0 To cmbDivision.ListCount - 1
      If cmbDivision = cmbDivision.List(iList) Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbDivision = cmbDivision.List(0)
   End If
   
End Sub

Private Sub cmbFromDate_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub cmbFromDate_LostFocus()
   If cmbFromDate = "" Then cmbFromDate = "ALL"
End Sub

Private Sub cmbThroughDate_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub cmbThroughDate_LostFocus()
   If cmbThroughDate = "" Then cmbThroughDate = "ALL"
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

Private Sub FillCombos()
   On Error GoTo DiaErr1
   
   'fill division combobox
   sSql = "select distinct rtrim(SODIVISION) from SohdTable order by rtrim(SODIVISION)"
   LoadComboBox cmbDivision, -1
   If cmbDivision.ListCount > 0 Then
      cmbDivision = cmbDivision.List(0)
   Else
      cmbDivision = "No Sales Orders"
      cmbDivision.ForeColor = ES_RED
   End If
   
   'fill customer combobox
   FillCustomers Me
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   
   If bOnLoad Then
      FillCombos
      bOnLoad = 0
   End If
   cmbFromDate = ""
   cmbThroughDate = ""
   cmbDivision = ""
   cmbCst = ""
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
   Set ShopSHp16a = Nothing
   
End Sub

Private Sub PrintReport()

    Dim A As Byte
    Dim b As Byte
    Dim C As Byte

    Dim sCust As String
    Dim sAsOfDate As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error Resume Next
   sCust = Compress(cmbDivision)
   If Not IsDate(cmbAsOfDate) Then
      sAsOfDate = "1995,01,01"
   Else
      sAsOfDate = Format(cmbAsOfDate, "yyyy,mm,dd")
   End If
   
   For b = 0 To 25
      If optTyp(b).Value = vbChecked Then C = 1
   Next
   If C = 0 Then
      MsgBox "No SO Types Selected.", vbInformation, Caption
      Exit Sub
   End If
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "AsOf"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ExtendedDesc"
    aFormulaName.Add "Division"
    aFormulaName.Add "FromDate"
    aFormulaName.Add "ToDate"
    aFormulaName.Add "Customer"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbAsOfDate) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add optDsc.Value
    aFormulaValue.Add CStr("'" & CStr(cmbDivision.Text) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbFromDate) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbThroughDate) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbCst) & "'")
   
    sCustomReport = GetCustomReport("slebl08")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = "({Vw_Slebl08.CUNICKNAME} = {@Customer} OR {@Customer} = 'ALL' OR {@Customer} = '')and " _
           & " {Vw_Slebl08.ITINVOICE} = 0.00 and " _
            & " {Vw_Slebl08.ITPSSHIPPED} <> 1.00 and " _
             & " {Vw_Slebl08.ITCUSTREQ} >= {@dFromDate} and " _
              & " {Vw_Slebl08.ITCUSTREQ} <= {@dToDate} and " _
               & " {SoitTable.ITPSNUMBER} = '' and " _
                & " {Vw_Slebl08.ITCANCELED} <> 1.00 and " _
                 & "({Vw_Slebl08.SODIVISION} = {@sDivision} OR {@sDivision} = 'ALL' OR {@sDivision} = '')"

   sSql = sSql & " AND ("
   
   A = 65
   C = 0
   For b = 0 To 25
      If optTyp(b).Value = vbChecked Then
         If C = 1 Then
            sSql = sSql & "OR {Vw_Slebl08.SOTYPE}='" & Chr$(A) & "' "
         Else
            sSql = sSql & "{Vw_Slebl08.SOTYPE}='" & Chr$(A) & "' "
         End If
         C = 1
      End If
      A = A + 1
   Next
   sSql = sSql & ")"

'   MDISect.Crw.Formulas(3) = "ExtendedDesc='" & optDsc.value & "'"
'   MDISect.Crw.Formulas(4) = "Division='" & cmbDivision.Text & "'"
'   MDISect.Crw.Formulas(5) = "FromDate='" & cmbFromDate & "'"
'   MDISect.Crw.Formulas(6) = "ToDate='" & cmbThroughDate & "'"
'   MDISect.Crw.Formulas(7) = "Customer='" & cmbCst & "'"
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   
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
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cmbAsOfDate = Format(Now, "MM/DD/yyyy")
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiProd", "sh16", Trim(str(optDsc.Value))
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh06", sOptions)
   If Len(sOptions) Then
      optDsc.Value = Val(sOptions)
   Else
      optDsc.Value = vbChecked
   End If
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub cmbasofdate_DropDown()
   ShowCalendarEx Me
End Sub


Private Sub cmbasofdate_LostFocus()
   If Len(Trim(cmbAsOfDate)) > 3 Then
      cmbAsOfDate = CheckDateEx(cmbAsOfDate)
   Else
      cmbAsOfDate = "ALL"
   End If
   
End Sub
