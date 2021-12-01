VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form BookBKp18a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Availability Report"
   ClientHeight    =   4290
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4290
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "BookBKp18a.frx":0000
      Height          =   315
      Left            =   4920
      Picture         =   "BookBKp18a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   71
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   600
      Width           =   350
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reporting Date"
      Height          =   615
      Left            =   2040
      TabIndex        =   68
      Top             =   1200
      Width           =   4695
      Begin VB.OptionButton optReportDate 
         Caption         =   "Customer Request Date"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   70
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optReportDate 
         Caption         =   "Scheduled Ship Date"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BookBKp18a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optExc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   2040
      TabIndex        =   65
      Top             =   3480
      Width           =   735
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "5"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   13
      Left            =   2040
      TabIndex        =   49
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   14
      Left            =   2280
      TabIndex        =   48
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   15
      Left            =   2520
      TabIndex        =   47
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   16
      Left            =   2760
      TabIndex        =   46
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   17
      Left            =   3000
      TabIndex        =   45
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   18
      Left            =   3240
      TabIndex        =   44
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   19
      Left            =   3480
      TabIndex        =   43
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   20
      Left            =   3720
      TabIndex        =   42
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   21
      Left            =   3960
      TabIndex        =   41
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   22
      Left            =   4200
      TabIndex        =   40
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   23
      Left            =   4440
      TabIndex        =   39
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   24
      Left            =   4680
      TabIndex        =   38
      Top             =   3120
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   25
      Left            =   4920
      TabIndex        =   37
      Top             =   3120
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
      TabIndex        =   3
      Top             =   2640
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   2760
      TabIndex        =   6
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   22
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   3240
      TabIndex        =   21
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   6
      Left            =   3480
      TabIndex        =   20
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   7
      Left            =   3720
      TabIndex        =   19
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   8
      Left            =   3960
      TabIndex        =   18
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   9
      Left            =   4200
      TabIndex        =   17
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   10
      Left            =   4440
      TabIndex        =   16
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   11
      Left            =   4680
      TabIndex        =   15
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   12
      Left            =   4920
      TabIndex        =   14
      Top             =   2640
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Tag             =   "4"
      ToolTipText     =   "Contains The Last Scheduled Delivery Date On Record"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Requires A Valid Part Number"
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
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
         Picture         =   "BookBKp18a.frx":0E32
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
         Picture         =   "BookBKp18a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4320
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4290
      FormDesignWidth =   7260
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   2040
      TabIndex        =   67
      Top             =   3840
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Beginning Balance From Totals"
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   64
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   63
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      Height          =   255
      Index           =   13
      Left            =   2040
      TabIndex        =   62
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      Height          =   255
      Index           =   14
      Left            =   2280
      TabIndex        =   61
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   255
      Index           =   15
      Left            =   2520
      TabIndex        =   60
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      Height          =   255
      Index           =   16
      Left            =   2760
      TabIndex        =   59
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   255
      Index           =   17
      Left            =   3000
      TabIndex        =   58
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   255
      Index           =   18
      Left            =   3240
      TabIndex        =   57
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      Height          =   255
      Index           =   19
      Left            =   3480
      TabIndex        =   56
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      Height          =   255
      Index           =   20
      Left            =   3720
      TabIndex        =   55
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      Height          =   255
      Index           =   21
      Left            =   3960
      TabIndex        =   54
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      Height          =   255
      Index           =   22
      Left            =   4200
      TabIndex        =   53
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   23
      Left            =   4440
      TabIndex        =   52
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   255
      Index           =   24
      Left            =   4680
      TabIndex        =   51
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      Height          =   255
      Index           =   25
      Left            =   4920
      TabIndex        =   50
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   36
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   35
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   34
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   33
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   32
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   31
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   30
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   29
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      Height          =   255
      Index           =   8
      Left            =   3960
      TabIndex        =   28
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   27
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      Height          =   255
      Index           =   10
      Left            =   4440
      TabIndex        =   26
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   25
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   255
      Index           =   12
      Left            =   4920
      TabIndex        =   24
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Types"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   23
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cutoff Date"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   1665
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   1665
   End
End
Attribute VB_Name = "BookBKp18a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'12/20/05 New
'12/21/05 Added SOTYPES and built in queries
'1/5/06 Added Custom Features for INTCOA (including GetLastDate)
'03/10/2010 BBS Fixed Ticket #24749
Option Explicit
'Dim RdoQry1     As rdoQuery
'Dim RdoQry2 As rdoQuery
Dim cmdObj As ADODB.Command
Dim RdoEsi As ADODB.Recordset
Dim RdoRpt As ADODB.Recordset 'KeySet
Dim bOnLoad As Byte

Dim iRow As Integer

Dim cBalance As Currency
Dim cQtyIn As Currency
Dim cQtyOut As Currency

Dim vDate As Variant
Dim sPartNumber As String
Dim sIncludes As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetPart()
   Dim RdoPrt As ADODB.Recordset
   sSql = "Qry_GetPartNumberBasics '" & Compress(cmbPrt) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(.Fields(1))
         If Len(cmbPrt) > 0 Then
            lblDsc = "" & Trim(.Fields(2))
         Else
            lblDsc = "*** Part Number Wasn't Found ***"
         End If
         ClearResultSet RdoPrt
      End With
   Else
      lblDsc = "*** Part Number Wasn't Found ***"
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbPrt_Click()
   GetPart
End Sub

Private Sub txtPrt_LostFocus()
   cmbPrt = txtPrt
   GetPart
End Sub

Private Sub cmbPrt_LostFocus()
   GetPart
End Sub


Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPrt = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPrt
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
   If bOnLoad Then CreateReportTable
   bOnLoad = 0
   MouseCursor 0
   
   Dim bPartSearch As Boolean
   
   bPartSearch = GetPartSearchOption
   SetPartSearchOption (bPartSearch)
   
   If (Not bPartSearch) Then FillPartCombo cmbPrt
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   '4/24/06
   '    sSql = "SELECT CUREF,CUNAME FROM CustTable where CUREF = ?"
   '    Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   '    RdoQry1.MaxRows = 1
   
   sSql = "SELECT VEREF,VEBNAME FROM VndrTable WHERE VEREF= ? "
   
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   
   Dim prmObj As ADODB.Parameter
   
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adChar
   prmObj.Size = 10
   cmdObj.parameters.Append prmObj

'   Set RdoQry2 = RdoCon.CreateQuery("", sSql)
'   RdoQry2.MaxRows = 1
   
     
   bOnLoad = 1
    
    optReportDate(0).Value = True   'BBS Added on 03/11/2010 for Ticket # 24749
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   sSql = "truncate table EsReportBook18"
   clsADOCon.ExecuteSql sSql
   'Set RdoQry2 = Nothing
   'Set RdoRpt = Nothing
   'Set RdoEsi = Nothing
   Set cmdObj = Nothing
   Set RdoEsi = Nothing
   Set RdoRpt = Nothing
   
   FormUnload
   Set BookBKp18a = Nothing
   
End Sub
Private Sub PrintReport()
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    
    sCustomReport = GetCustomReport("slebk18")
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = "slebk18"
    cCRViewer.ShowGroupTree False

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "Includes2"
    aFormulaName.Add "RequestBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Includes Sales Order Types:" & CStr(sIncludes) & "'")
    aFormulaValue.Add CStr("'And Items On Or Before" & CStr(txtEnd) & "...'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    
   If optExc.Value = vbChecked Then
      aFormulaName.Add "BegBal"
      aFormulaValue.Add CStr("'Excludes Beginning Balance From Totals '")
   Else
      aFormulaName.Add "BegBal"
      aFormulaValue.Add CStr("'Includes Beginning Balance In Totals'")
   End If
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
    ' print the copies
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   prg1.Visible = False
   
   Exit Sub
   
DiaErr1:
   prg1.Visible = False
   sSql = "truncate table EsReportBook18"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtEnd = GetLastSODate()
   txtEnd.ToolTipText = "Contains The Last Scheduled Delivery Date On Record"
   For b = 0 To 24
      optTyp(b).TabIndex = b + 3
   Next
   optTyp(b).TabIndex = b + 3
   optExc.TabIndex = b + 4
   
End Sub

Private Sub SaveOptions()
   Dim b As Byte
   Dim sOptions As String
   On Error Resume Next
   For b = 0 To 25
      sOptions = sOptions & Trim$(optTyp(b).Value)
   Next
   SaveSetting "Esi2000", "EsiSale", "bk18", Trim(sOptions)
   SaveSetting "Esi2000", "EsiSale", "bk18a", Trim(Val(optExc.Value))
   
End Sub

Private Sub GetOptions()
   Dim b As Byte
   Dim sOptions As String
   Dim sExclude As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "bk18", Trim(sOptions))
   If Len(Trim(sOptions)) > 0 Then
      For b = 0 To 24
         optTyp(b).Value = Val(Mid$(sOptions, b + 1, 1))
      Next
      optTyp(b).Value = Val(Mid$(sOptions, b + 1, 1))
   End If
   sExclude = GetSetting("Esi2000", "EsiSale", "bk18a", Trim(sExclude))
   If sExclude = "" Then
      If ES_CUSTOM = "INTCOA" Then optExc.Value = vbChecked
   Else
      optExc.Value = Val(sExclude)
   End If
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 12) = "*** Part Num" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub optDis_Click()
   GetPart
   If lblDsc.ForeColor = ES_RED Then MsgBox "Requires A Valid Part Number.", _
                         vbInformation, Caption Else BuildReport
   
End Sub


Private Sub optExc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   GetPart
   If lblDsc.ForeColor = ES_RED Then MsgBox "Requires A Valid Part Number.", _
                         vbInformation, Caption Else BuildReport
   
   
End Sub

Private Sub optTyp_GotFocus(Index As Integer)
   lblAlp(Index).BorderStyle = 1
   
End Sub

Private Sub optTyp_LostFocus(Index As Integer)
   lblAlp(Index).BorderStyle = 0
   
End Sub

Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Trim(txtEnd) = "" Then
      txtEnd = "ALL"
   Else
      txtEnd = CheckDateEx(txtEnd)
   End If
   
End Sub

'EsReportBook18

Private Sub CreateReportTable()
   
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT Row from EsReportBook18 where Row=12"
   Dim success As Boolean
   success = clsADOCon.ExecuteSql(sSql)  'rdExecDirect
'   If clsADOCon.ADOErrNum <> 0 Then
   If Not success Then
      clsADOCon.ADOErrNum = 0
      Err.Clear
      'BBS Added on 03/11/2010 for Ticket #24749
      sSql = "create table EsReportBook18 (" _
             & "Row smallint null default(1)," _
             & "Col1 char(60) null default('')," _
             & "Col2 char(24) null default('')," _
             & "Col3 char(30) null default('')," _
             & "Col4 char(20) null default('')," _
             & "Col5 char(20) null default('')," _
             & "Col6 char(20) null default('')," _
             & "Col7 char(4) null default('')," _
             & "Col8 char(20) null default(''))"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "create unique clustered index RowRef on EsReportBook18 " _
                & "(Row) WITH FILLFACTOR=80"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "create index PartRef on EsReportBook18 " _
                & "(Col3) WITH FILLFACTOR=80"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
   End If
   Err.Clear
   clsADOCon.ADOErrNum = 0
   
End Sub

Private Sub BuildReport()
   Dim bByte As Byte
   Dim sTemp As String
   
    ' truncate the table
   sSql = "truncate table EsReportBook18"
   clsADOCon.ExecuteSql sSql
   
   bByte = TestReportTable
   If bByte = 1 Then
      MsgBox "The Table Is In Use. Try Again In A Few Minutes.", _
         vbInformation, Caption
      Exit Sub
   End If
   MouseCursor 13
   If IsDate(txtEnd) Then
      If Len(txtEnd) = 8 And Mid(txtEnd, 6, 1) = "/" Then sTemp = Left(txtEnd, 6) & "20" & Right(txtEnd, 2) Else sTemp = txtEnd
      vDate = Format(sTemp, "yyyy-mm-dd 23:59")
   Else
      vDate = "2025-12-31"
   End If
   
   For bByte = 0 To 24
      lblAlp(bByte).BorderStyle = 0
   Next
   lblAlp(bByte).BorderStyle = 0
   
   sPartNumber = Compress(cmbPrt)
   optPrn.Enabled = False
   optDis.Enabled = False
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM EsReportBook18"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_KEYSET)
   iRow = 0
   cBalance = 0
   cQtyOut = 0
   
   prg1.Value = 10
   prg1.Visible = True
   sProcName = "getbalance"
   GetBegBalance
   Sleep 100
   
   prg1.Value = 20
   sProcName = "getsalesorders"
   GetSalesOrders
   Sleep 100
   
   prg1.Value = 30
   sProcName = "getpicks"
   GetPicks
   Sleep 100
   
   prg1.Value = 40
   sProcName = "getmanufactu"
   GetManufacturingOrders
   Sleep 100
   
   prg1.Value = 50
   sProcName = "getpurchaseo"
   GetPurchaseOrders
   Sleep 100
   
   prg1.Value = 80
   sProcName = "getpartdata"
   GetPartData
   Sleep 100
   
   RdoRpt.Close
   sSql = "UPDATE EsReportBook18 SET Col3='" & sPartNumber & "'"
    clsADOCon.ExecuteSql sSql 'rdExecDirect
   prg1.Value = 100
   Sleep 500
   PrintReport
   Exit Sub
   
DiaErr1:
   prg1.Visible = False
   optPrn.Enabled = True
   optDis.Enabled = True
   sSql = "truncate table EsReportBook18"
    clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub GetBegBalance()
   sSql = "SELECT ISNULL(PAQOH, 0) as PAQOH FROM PartTable WHERE PARTREF='" & sPartNumber & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEsi, ES_FORWARD)
   If bSqlRows Then
      With RdoEsi
         iRow = iRow + 1
         If optExc.Value = vbChecked Then
            cBalance = 0
         Else
            cBalance = !PAQOH
         End If
         With RdoRpt
            .AddNew
            !Row = iRow
            !Col1 = String$(8, " ") & "Current Quantity On Hand:"
            If RdoEsi!PAQOH > 0 Then
               !Col6 = Format$(RdoEsi!PAQOH, "0.000")
            Else
               !Col6 = Format$(0.0001, "0.000")
            End If
            .Update
         End With
         ClearResultSet RdoEsi
      End With
   End If
   
End Sub

Public Sub GetSalesOrders()
   Dim b As Byte
   Dim bByte As Byte
   Dim bList As Byte
   Dim cSoBalance As Currency
   Dim sCustomer As String
   Dim sDateField As String 'BBS Added on 03/10/2010 for Ticket #24749
   
   If (optReportDate(0).Value = True) Then sDateField = "ITSCHED" Else sDateField = "ITCUSTREQ"  'BBS Added ITCUSTREQ on 03/10/2010 for Ticket #24749

   
   sIncludes = ""
   'BBS Added on 03/11/2010 for Ticket #24749
   sSql = "SELECT SONUMBER,SOTYPE,SOCUST,ITSO,ITNUMBER,ITREV," _
          & "ITPART,ITQTY, ITSCHED, ITCUSTREQ FROM SohdTable,SoitTable WHERE (SONUMBER=ITSO " _
          & "AND ITPART='" & sPartNumber & "' AND ITPSITEM>0 AND ITPSSHIPPED=0 " _
          & "AND ITINVOICE=0 AND ITCANCELED=0) AND ("
   For bList = 0 To 25
      If optTyp(bList).Value = vbChecked Then
         bByte = bByte + 1
         sIncludes = sIncludes & Chr$(bList + 65)
         If bByte = 1 Then
            sSql = sSql & "SOTYPE='" & Chr$(bList + 65) & "'"
         Else
            sSql = sSql & " OR SOTYPE='" & Chr$(bList + 65) & "'"
         End If
      End If
   Next
   sSql = sSql & ") AND " & sDateField & " <='" & vDate & "' ORDER BY " & sDateField  'BBS Changed on 03/10/2010 for Ticket #24749
   'Sect 1 on Packslips not shipped
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEsi, ES_FORWARD)
   If bSqlRows Then
      b = 1
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = ""
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      
      RdoRpt!Col4 = "Scheduled" & String(11, " ")
      RdoRpt!Col8 = "Requested" & String(11, " ")
      RdoRpt!Col6 = "Projected" & String(11, " ")
      
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = "Customer"
      RdoRpt!Col2 = "Sales Order"
      RdoRpt!Col4 = "Date"
      RdoRpt!Col5 = "Quantity"
      RdoRpt!Col6 = "Balance"
      RdoRpt!Col8 = "Date"  'BBS Added on 03/11/2010 for Ticket #24749
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = String$(60, "-")
      RdoRpt!Col2 = String$(24, "-")
      RdoRpt!Col4 = String$(20, "-")
      RdoRpt!Col8 = String$(20, "-")
      RdoRpt!Col5 = String$(20, "-")
      RdoRpt!Col6 = String$(20, "-")
      
      RdoRpt.Update
      With RdoEsi
         Do Until .EOF
            With RdoRpt
               iRow = iRow + 1
               sCustomer = xCustomer(RdoEsi!SOCUST)
               .AddNew
               If ES_CUSTOM <> "INTCOA" Then
                  ' MM Not need to reduce the Balance as this is been packsliped.
                  'cBalance = cBalance - RdoEsi!ITQTY
                  cSoBalance = cSoBalance + RdoEsi!ITQty
               End If
               !Row = iRow
               !Col1 = sCustomer
               !Col2 = RdoEsi!SOTYPE & Format$(RdoEsi!SoNumber, SO_NUM_FORMAT) & "-" & Trim$(RdoEsi!ITNUMBER) & RdoEsi!itrev
               !Col4 = Format$(RdoEsi!itsched, "mm/dd/yy")
               !Col8 = Format$(RdoEsi!itcustreq, "mm/dd/yy") 'BBS Changed on 03/11/2010 for Ticket #24749
               !Col5 = Format$(RdoEsi!ITQty, "0.000")
               '!Col6 = cBalance
               !Col6 = Format$(cBalance, "0.000")
               !Col7 = "*"
               .Update
            End With
            .MoveNext
         Loop
         .Cancel
      End With
      ClearResultSet RdoEsi
   Else
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = ""
      RdoRpt.Update
   End If
   'Sect 2 other Sales Order Items
   bByte = 0
'BBS Changed on 03/10/2010 for Ticket #24749
   sSql = "SELECT SONUMBER,SOTYPE,SOCUST,ITSO,ITNUMBER,ITREV," _
          & "ITPART,ITQTY,ITSCHED, ITCUSTREQ FROM SohdTable,SoitTable WHERE (SONUMBER=ITSO " _
          & "AND ITPART='" & sPartNumber & "' AND ITPSITEM=0 AND ITPSSHIPPED=0 " _
          & "AND ITCANCELED=0) AND ("
   For bList = 0 To 25
      If optTyp(bList).Value = vbChecked Then
         bByte = bByte + 1
         If bByte = 1 Then
            sSql = sSql & "SOTYPE='" & Chr$(bList + 65) & "'"
         Else
            sSql = sSql & " OR SOTYPE='" & Chr$(bList + 65) & "'"
         End If
      End If
   Next
   sSql = sSql & ") ORDER BY " & sDateField  'BBS Changed on 03/10/2010 for Ticket #24749
   
   Debug.Print sSql
   'Sect 2
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEsi, ES_FORWARD)
   If bSqlRows Then
      If b = 0 Then
         iRow = iRow + 1
         RdoRpt.AddNew
         RdoRpt!Row = iRow
         RdoRpt!Col1 = ""
         RdoRpt.Update
         
         iRow = iRow + 1
         RdoRpt.AddNew
         RdoRpt!Row = iRow
         RdoRpt!Col4 = "Scheduled" & String(11, " ")
         RdoRpt!Col8 = "Requested" & String(11, " ")
         RdoRpt!Col6 = "Projected" & String(11, " ")
         RdoRpt.Update
         
         iRow = iRow + 1
         RdoRpt.AddNew
         RdoRpt!Row = iRow
         RdoRpt!Col1 = "Customer"
         RdoRpt!Col2 = "Sales Order"
         RdoRpt!Col4 = "Date"
         RdoRpt!Col8 = "Date"
         RdoRpt!Col5 = "Quantity"
         RdoRpt!Col6 = "Balance"
         RdoRpt.Update
         
         iRow = iRow + 1
         RdoRpt.AddNew
         RdoRpt!Row = iRow
         RdoRpt!Col1 = String$(60, "-")
         RdoRpt!Col2 = String$(24, "-")
         RdoRpt!Col4 = String$(20, "-")
         RdoRpt!Col8 = String$(20, "-")
         RdoRpt!Col5 = String$(20, "-")
         RdoRpt!Col6 = String$(20, "-")
         
         RdoRpt.Update
      End If
      With RdoEsi
         Do Until .EOF
            iRow = iRow + 1
            With RdoRpt
               If Format(RdoEsi!itcustreq, "YYYY-MM-DD") <= vDate Then   'BBS Changed from ITSCHED to ITCUSTREQ on 03/10/2010 for Ticket #24749
                  sCustomer = xCustomer(RdoEsi!SOCUST)
                  .AddNew
                  cBalance = cBalance - RdoEsi!ITQty
                  cSoBalance = cSoBalance + RdoEsi!ITQty
                  !Row = iRow
                  !Col1 = sCustomer
                  !Col2 = RdoEsi!SOTYPE & Format$(RdoEsi!SoNumber, SO_NUM_FORMAT) & "-" & Trim$(RdoEsi!ITNUMBER) & RdoEsi!itrev
                  !Col4 = Format$(RdoEsi!itsched, "mm/dd/yy") 'BBS Changed on 03/11/2010 for Ticket #24749
                  !Col8 = Format$(RdoEsi!itcustreq, "mm/dd/yy") 'BBS Changed on 03/11/2010 for Ticket #24749
                  !Col5 = Format$(RdoEsi!ITQty, "0.000")
                  !Col6 = Format$(cBalance, "0.000") 'cBalance
                  !Col7 = ""
                  .Update
               End If
            End With
            .MoveNext
         Loop
         .Cancel
      End With
      ClearResultSet RdoEsi
   Else
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = ""
      RdoRpt.Update
      
   End If
   With RdoRpt
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col5 = String(20, "-")
      .Update
      
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col1 = String$(8, " ") & "Total Required For Sales Orders:"
      !Col5 = Format$(cSoBalance, "0.000")
      !Col6 = Format$(cBalance, "0.000") 'cBalance
      .Update
      
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col5 = String(20, "-")
      .Update
   End With
   cQtyOut = cSoBalance
   
   
End Sub

Private Function xCustomer(CustomerName) As String
   Dim RdoCst As ADODB.Recordset
   
   sSql = "Qry_GetCustomerBasics '" & CustomerName & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then xCustomer = "" & Trim(RdoCst.Fields(1))
   ClearResultSet RdoCst
   
End Function

Public Sub GetPicks()
   Dim cMoPicks As Currency
  ' sSql = "SELECT PKPARTREF,PKMOPART,PKMORUN,PKTYPE,PKPDATE,PKPQTY,PKAQTY " _
  '        & "PARTREF,PARTNUM FROM MopkTable,PartTable WHERE (PKMOPART=PARTREF " _
  '        & "AND PKPARTREF='" & sPartNumber & "' AND PKTYPE<>12 AND PKAQTY=0) " _
  '        & "AND PKPDATE<='" & vDate & "' ORDER BY PKPDATE,PKMOPART,PKMORUN"
  '
   sSql = "SELECT PKPARTREF,PKMOPART,PKMORUN,PKTYPE,PKPDATE,PKPQTY,PKAQTY, " _
          & " PARTREF,PARTNUM, RUNSTATUS FROM MopkTable " _
          & " INNER JOIN PartTable ON PKMOPART=PARTREF " _
          & " INNER JOIN RunsTable ON PKMOPART=RUNREF AND PKMORUN=RUNNO " _
          & " WHERE PKPARTREF='" & sPartNumber & "' AND PKTYPE<>12 AND PKAQTY=0 " _
          & " AND PKPDATE<='" & vDate & "' AND RUNSTATUS NOT LIKE 'C%' " _
          & " ORDER BY PKPDATE,PKMOPART,PKMORUN"
          
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEsi, ES_FORWARD)
   If bSqlRows Then
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = ""
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col4 = "Scheduled" & String(11, " ")
      RdoRpt!Col6 = "Projected" & String(11, " ")
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = "Manufacturing Order Picks"
      RdoRpt!Col2 = "Run"
      RdoRpt!Col4 = "Date"
      RdoRpt!Col5 = "Quantity"
      RdoRpt!Col6 = "Balance"
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = String$(60, "-")
      RdoRpt!Col2 = String$(24, "-")
      RdoRpt!Col4 = String$(20, "-")
      RdoRpt!Col5 = String$(20, "-")
      RdoRpt!Col6 = String$(20, "-")
      RdoRpt.Update
      With RdoEsi
         Do Until .EOF
            With RdoRpt
               iRow = iRow + 1
               cBalance = cBalance - RdoEsi!PKPQTY
               cMoPicks = cMoPicks + RdoEsi!PKPQTY
               .AddNew
               !Row = iRow
               !Col1 = "" & Trim(RdoEsi!PartNum)
               !Col2 = "Run " & Format(RdoEsi!PKMORUN, "####0")
               !Col4 = Format$(RdoEsi!PKPDATE, "mm/dd/yy")
               !Col5 = Format$(RdoEsi!PKPQTY, "0.000")
               !Col6 = Format$(cBalance, "0.000")
               .Update
            End With
            .MoveNext
         Loop
         .Cancel
      End With
      ClearResultSet RdoEsi
   End If
   cQtyOut = cQtyOut + cMoPicks
   With RdoRpt
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col5 = String(20, "-")
      .Update
      
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col1 = String$(8, " ") & "Total Required For MO Picks:"
      !Col5 = Format$(cMoPicks, "0.000")
      !Col6 = Format$(cBalance, "0.000") 'cBalance
      .Update
      
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col5 = String(20, "-")
      .Update
      iRow = iRow + 1
      
      .AddNew
      !Row = iRow
      !Col1 = String(20, " ")
      .Update
      
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col1 = String$(4, " ") & "Total Quantity Required:"
      !Col5 = Format$(cQtyOut, "0.000")
      !Col6 = Format$(cBalance, "0.000") 'cBalance
      .Update
   End With
   
   
End Sub

Private Sub GetManufacturingOrders()
   Dim cMoRuns As Currency
   sSql = "SELECT RUNREF,RUNNO,RUNSCHED,RUNSTATUS,RUNREMAININGQTY " _
          & "FROM RunsTable WHERE (RUNREF='" & sPartNumber & "' " _
          & "AND RUNSTATUS NOT LIKE 'C%') AND RUNSCHED<='" _
          & vDate & "' ORDER BY RUNSCHED,RUNNO"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEsi, ES_FORWARD)
   If bSqlRows Then
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = ""
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col4 = "Scheduled" & String(11, " ")
      RdoRpt!Col6 = "Projected" & String(11, " ")
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = "Manufacturing Order"
      RdoRpt!Col2 = "Run"
      RdoRpt!Col4 = "Date"
      RdoRpt!Col5 = "Quantity"
      RdoRpt!Col6 = "Balance"
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = String$(60, "-")
      RdoRpt!Col2 = String$(24, "-")
      RdoRpt!Col4 = String$(20, "-")
      RdoRpt!Col5 = String$(20, "-")
      RdoRpt!Col6 = String$(20, "-")
      RdoRpt.Update
      With RdoEsi
         Do Until .EOF
            With RdoRpt
               iRow = iRow + 1
               cBalance = cBalance + RdoEsi!RUNREMAININGQTY
               cMoRuns = cMoRuns + RdoEsi!RUNREMAININGQTY
               .AddNew
               !Row = iRow
               !Col1 = sPartNumber
               !Col2 = "Run " & Format(RdoEsi!Runno, "####0") & " Status " & Trim(RdoEsi!RUNSTATUS)
               !Col4 = Format$(RdoEsi!RUNSCHED, "mm/dd/yy")
               !Col5 = Format$(RdoEsi!RUNREMAININGQTY, "0.000")
               !Col6 = Format$(cBalance, "0.000") 'cBalance
               .Update
            End With
            .MoveNext
         Loop
         .Cancel
      End With
      ClearResultSet RdoEsi
   End If
   cQtyIn = cMoRuns
   With RdoRpt
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col5 = String(20, "-")
      .Update
      
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col1 = String$(8, " ") & "Total Available From Manufacturing Orders:"
      !Col5 = Format$(cMoRuns, "0.000")
      !Col6 = Format$(cBalance, "0.000") 'cBalance
      .Update
   End With
   
End Sub

Private Sub GetPurchaseOrders()
   Dim cPoItems As Currency
   Dim sVendor As String
   'BBS Added PITYPE<>39 to the query below on 8/28/2010 for Ticket #37239 to keep Cancelled items from showing
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPART,PIPDATE," _
          & "PIPQTY,PONUMBER,POVENDOR FROM PoitTable,PohdTable WHERE " _
          & "(PIPART='" & sPartNumber & "' AND PINUMBER=PONUMBER AND PIAQTY=0 " _
          & "AND PITYPE<>39 AND PITYPE<>16 AND PITYPE<>17) AND PIPDATE<='" _
          & vDate & "' ORDER BY PIPDATE"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEsi, ES_FORWARD)
   If bSqlRows Then
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = ""
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col4 = "Scheduled" & String(11, " ")
      RdoRpt!Col6 = "Projected" & String(11, " ")
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = "Vendor "
      RdoRpt!Col2 = "PO Item"
      RdoRpt!Col4 = "Date"
      RdoRpt!Col5 = "Quantity"
      RdoRpt!Col6 = "Balance"
      RdoRpt.Update
      
      iRow = iRow + 1
      RdoRpt.AddNew
      RdoRpt!Row = iRow
      RdoRpt!Col1 = String$(60, "-")
      RdoRpt!Col2 = String$(24, "-")
      RdoRpt!Col4 = String$(20, "-")
      RdoRpt!Col5 = String$(20, "-")
      RdoRpt!Col6 = String$(20, "-")
      RdoRpt.Update
      With RdoEsi
         Do Until .EOF
            With RdoRpt
               iRow = iRow + 1
               sVendor = xVendor(RdoEsi!POVENDOR)
               .AddNew
               cBalance = cBalance + RdoEsi!PIPQTY
               cPoItems = cPoItems + RdoEsi!PIPQTY
               !Row = iRow
               !Col1 = sVendor
               !Col2 = "PO " & Format$(RdoEsi!PONumber, "00000") & "-" & Trim$(RdoEsi!PIITEM) & RdoEsi!PIREV
               !Col4 = Format$(RdoEsi!PIPDATE, "mm/dd/yy")
               !Col5 = Format$(RdoEsi!PIPQTY, "0.000")
               !Col6 = Format$(cBalance, "0.000") 'cBalance
               .Update
            End With
            .MoveNext
         Loop
         .Cancel
      End With
      ClearResultSet RdoEsi
   End If
   cQtyIn = cQtyIn + cPoItems
   With RdoRpt
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col5 = String(20, "-")
      .Update
      
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col1 = String$(8, " ") & "Total Available From Purchase Orders:"
      !Col5 = Format$(cPoItems, "0.000")
      !Col6 = Format$(cBalance, "0.000") 'cBalance
      .Update
      
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col5 = String(20, "-")
      .Update
      
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col1 = String(20, " ")
      .Update
      
      iRow = iRow + 1
      .AddNew
      !Row = iRow
      !Col1 = String$(4, " ") & "Total Quantity In/Available:"
      !Col5 = Format$(cQtyIn, "0.000")
      !Col6 = Format$(cBalance, "0.000") 'cBalance
      .Update
   End With
   
End Sub

Private Function xVendor(VendorName) As String
   Dim RdoVnd As ADODB.Recordset
   'RdoQry2(0) = VendorName
   'bSqlRows = clsADOCon.GetQuerySet(RdoVnd, RdoQry2, ES_FORWARD)
   cmdObj.parameters(0).Value = VendorName
   bSqlRows = clsADOCon.GetQuerySet(RdoVnd, cmdObj, ES_FORWARD, True)
   If bSqlRows Then xVendor = "" & Trim(RdoVnd.Fields(1))
   ClearResultSet RdoVnd
   
End Function

Public Sub GetPartData()
   sSql = "SELECT PARTREF,PASTDCOST,PAPRICE,PAFLOWTIME,PALEADTIME " _
          & "FROM PartTable WHERE PARTREF='" & sPartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEsi, ES_FORWARD)
   If bSqlRows Then
      With RdoRpt
         iRow = iRow + 1
         .AddNew
         !Row = iRow
         !Col1 = " "
         .Update
         
         iRow = iRow + 1
         .AddNew
         !Row = iRow
         !Col1 = String(8, " ") & "Standard Cost:"
         !Col5 = Format$(RdoEsi!PASTDCOST, "0.0000")
         .Update
         
         iRow = iRow + 1
         .AddNew
         !Row = iRow
         !Col1 = String(8, " ") & "Unit Price:"
         !Col5 = Format$(RdoEsi!PAPRICE, "0.0000")
         .Update
         
         iRow = iRow + 1
         .AddNew
         !Row = iRow
         !Col1 = String(8, " ") & "Lead Time:"
         !Col5 = Format$(RdoEsi!PALEADTIME, "##0") & " Days"
         .Update
         
         iRow = iRow + 1
         .AddNew
         !Row = iRow
         !Col1 = String(8, " ") & "Flow Time:"
         !Col5 = Format$(RdoEsi!PAFLOWTIME, "##0") & " Days"
         .Update
      End With
      ClearResultSet RdoEsi
   End If
End Sub

'Test it and drop it if it isn't being used

Private Function TestReportTable() As Byte
   On Error Resume Next
   sSql = "INSERT INTO EsReportBook18 (Col1) VALUES(' ')"
      clsADOCon.ExecuteSql sSql
   If Err > 0 Then
      TestReportTable = 1
   Else
      sSql = "truncate table EsReportBook18"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   
End Function


Private Function GetLastSODate() As String
   Dim RdoDate As ADODB.Recordset
   Dim sDateField As String
   
   If optReportDate(0).Value = True Then sDateField = "ITSCHED" Else sDateField = "ITCUSTREQ"
   
   On Error Resume Next
   'BBS Changed from ITSCHED to ITCUSTREQ on 03/10/2010 for Ticket #24749
   sSql = "SELECT MAX(" & sDateField & ") AS LastDate FROM SoitTable " _
          & "WHERE (ITCANCELED=0 AND ITPSSHIPPED=0)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDate, ES_FORWARD)
   If bSqlRows Then
      With RdoDate
         If Not IsNull(!LastDate) Then
            GetLastSODate = Format$(!LastDate, "mm/dd/yyyy")
         Else
            GetLastSODate = Format(Now + 365, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDate
      End With
   Else
      GetLastSODate = Format(Now + 365, "mm/dd/yyyy")
   End If
   
   Set RdoDate = Nothing
   
End Function

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function


