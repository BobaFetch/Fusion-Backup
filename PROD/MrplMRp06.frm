VERSION 5.00
Begin VB.Form MrplMRp06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Availablity Report with MRP"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optMODet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   84
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "MrplMRp06.frx":0000
      Height          =   315
      Left            =   4320
      Picture         =   "MrplMRp06.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   83
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   480
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      Top             =   480
      Width           =   2895
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.ComboBox cmbLvl 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "MrplMRp06.frx":0684
      Left            =   1920
      List            =   "MrplMRp06.frx":069A
      TabIndex        =   81
      Text            =   "7"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   78
      Top             =   5160
      Width           =   375
   End
   Begin VB.CheckBox OptCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   77
      Top             =   5520
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   72
      Top             =   4560
      Width           =   2775
      Begin VB.OptionButton optMbe 
         Caption         =   "ALL"
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   76
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "E"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   75
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "B"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   74
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "M"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5400
      TabIndex        =   67
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5400
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.CheckBox optExc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   1920
      TabIndex        =   64
      Top             =   3720
      Width           =   735
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   12
      Left            =   4800
      TabIndex        =   36
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   11
      Left            =   4560
      TabIndex        =   35
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   10
      Left            =   4320
      TabIndex        =   34
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   9
      Left            =   4080
      TabIndex        =   33
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   8
      Left            =   3840
      TabIndex        =   32
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   7
      Left            =   3600
      TabIndex        =   31
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   6
      Left            =   3360
      TabIndex        =   30
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   3120
      TabIndex        =   29
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   2880
      TabIndex        =   28
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   2640
      TabIndex        =   27
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   26
      Top             =   2760
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   25
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
      Left            =   1920
      TabIndex        =   24
      Top             =   2760
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   25
      Left            =   4800
      TabIndex        =   23
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   24
      Left            =   4560
      TabIndex        =   22
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   23
      Left            =   4320
      TabIndex        =   21
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   22
      Left            =   4080
      TabIndex        =   20
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   21
      Left            =   3840
      TabIndex        =   19
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   20
      Left            =   3600
      TabIndex        =   18
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   19
      Left            =   3360
      TabIndex        =   17
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   18
      Left            =   3120
      TabIndex        =   16
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   17
      Left            =   2880
      TabIndex        =   15
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   16
      Left            =   2640
      TabIndex        =   14
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   15
      Left            =   2400
      TabIndex        =   13
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   14
      Left            =   2160
      TabIndex        =   12
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "5"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   13
      Left            =   1920
      TabIndex        =   11
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Tag             =   "4"
      ToolTipText     =   "Contains The Last Scheduled Delivery Date On Record"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reporting Date"
      Height          =   615
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
      Begin VB.OptionButton optReportDate 
         Caption         =   "Scheduled Ship Date"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optReportDate 
         Caption         =   "Customer Request Date"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox cmbPrt1 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Requires A Valid Part Number"
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MRP MO and PO Detail"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   85
      Top             =   5880
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Levels Through"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   82
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   80
      Top             =   5160
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include BOM comments"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   79
      Top             =   5520
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make, Buy, Either"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   71
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   70
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Beginning Balance From Totals"
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   65
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Types"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   63
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   255
      Index           =   12
      Left            =   4800
      TabIndex        =   62
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      Height          =   255
      Index           =   11
      Left            =   4560
      TabIndex        =   61
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      Height          =   255
      Index           =   10
      Left            =   4320
      TabIndex        =   60
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      Height          =   255
      Index           =   9
      Left            =   4080
      TabIndex        =   59
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      Height          =   255
      Index           =   8
      Left            =   3840
      TabIndex        =   58
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   57
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   56
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   55
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   54
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   53
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   52
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   51
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   50
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      Height          =   255
      Index           =   25
      Left            =   4800
      TabIndex        =   49
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   255
      Index           =   24
      Left            =   4560
      TabIndex        =   48
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   23
      Left            =   4320
      TabIndex        =   47
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      Height          =   255
      Index           =   22
      Left            =   4080
      TabIndex        =   46
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      Height          =   255
      Index           =   21
      Left            =   3840
      TabIndex        =   45
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      Height          =   255
      Index           =   20
      Left            =   3600
      TabIndex        =   44
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      Height          =   255
      Index           =   19
      Left            =   3360
      TabIndex        =   43
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   255
      Index           =   18
      Left            =   3120
      TabIndex        =   42
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   255
      Index           =   17
      Left            =   2880
      TabIndex        =   41
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      Height          =   255
      Index           =   16
      Left            =   2640
      TabIndex        =   40
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   255
      Index           =   15
      Left            =   2400
      TabIndex        =   39
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      Height          =   255
      Index           =   14
      Left            =   2160
      TabIndex        =   38
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      Height          =   255
      Index           =   13
      Left            =   1920
      TabIndex        =   37
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cutoff Date"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   9
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Part Number"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "MrplMRp06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/21/2010 New

Option Explicit

Dim bOnLoad As Byte

Dim iRow As Integer

Dim vDate As Variant
Dim sPartNumber As String

Dim iOrder As Integer
Dim sIns As String
Dim sBomRev As String

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

Private Sub cmbPrt_Change()
   GetPart
End Sub

Private Sub cmbPrt_LostFocus()
   GetPart
End Sub


Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdFnd_Click()
    ViewParts.lblControl = "CMBPRT"
    ViewParts.txtPrt = cmbPrt
    ViewParts.Show
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   bOnLoad = 0
   MouseCursor 0
   
   Dim bPartSearch As Boolean
   
   bPartSearch = GetPartSearchOption
   SetPartSearchOption (bPartSearch)
   
   If (Not bPartSearch) Then FillPartCombo cmbPrt
   cmbPrt = ""
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   optMbe(3).Value = True
   optReportDate(0).Value = True   'BBS Added on 03/11/2010 for Ticket # 24749
   
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
   On Error Resume Next
   FormUnload
   Set MrplMRp06 = Nothing
   
End Sub


Private Sub PrintReport()
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   Dim i As Integer
   Dim sSOTypes As String
   Dim BOMSql As String
   Dim bLvl As Byte
      
   
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    Dim sCutoffDate As String
    Dim sTemp As String
   
    If optMODet.Value = 1 Then
        sCustomReport = GetCustomReport("prdmr06det")
    Else
        sCustomReport = GetCustomReport("prdmr06")
    End If
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False
    
    If Compress(txtEnd) = "ALL" Then
        sCutoffDate = GetLastSODate()
    Else
        sCutoffDate = Format(txtEnd, "mm/dd/yyyy")
    End If

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes1"
    aFormulaName.Add "Includes2"
    aFormulaName.Add "Includes3"
    
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "CutoffDate"
    aFormulaName.Add "ExcludeBegBalance"
    aFormulaName.Add "SalesOrderDateField"
    aFormulaName.Add "ShowPartDesc"
    aFormulaName.Add "ShowBOMCmt"
        
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    If optReportDate(0).Value = True Then sTemp = "Scheduled Ship Date" Else sTemp = "Customer Required Date"
    
    'Includes1 formula
    aFormulaValue.Add CStr("'Includes Items by " & sTemp & " On Or Before " & CStr(txtEnd) & "'")
    
    'Build Includes2 formula
    sTemp = ""
    If optExc.Value = vbChecked Then sTemp = sTemp & "Exclude Beginning Balance Totals,"
    sTemp = sTemp & " BOM Levels Through " & cmbLvl & ","
    
    bLvl = Val(cmbLvl) + 1
    BOMSql = "{MrpbTable.MRPBILL_LEVEL}<" & bLvl & " AND "
    BOMSql = BOMSql & "{MrpbTable.MRPBILL_LEVEL} >= 0.00 and not ({PartTable.PALEVEL} in [6, 5])"
    
    If optMbe(0).Value = True Then
       BOMSql = BOMSql & " AND {PartTable.PAMAKEBUY} ='M'"
       sTemp = sTemp & " Make,"
    ElseIf optMbe(1).Value = True Then
       BOMSql = BOMSql & " AND {PartTable.PAMAKEBUY}='B'"
       sTemp = sTemp & " Buy,"
    ElseIf optMbe(2).Value = True Then
       BOMSql = BOMSql & " AND {PartTable.PAMAKEBUY}='E'"
        sTemp = sTemp & " Make or Buy,"
    End If
    If Right(sTemp, 1) = "," Then sTemp = Left(sTemp, Len(sTemp) - 1)
    aFormulaValue.Add CStr("'" & sTemp & "'")
    
    'Build Includes3 Formula
    sTemp = ""
    If optDsc.Value = vbChecked Then sTemp = sTemp & "Include Extended Descriptions,"
    If OptCmt.Value = vbChecked Then sTemp = sTemp & "Include BOM Comments"
    If Right(sTemp, 1) = "," Then sTemp = Left(sTemp, Len(sTemp) - 1)
    aFormulaValue.Add CStr("'" & sTemp & "'")
    
    
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & sCutoffDate & "'")
    aFormulaValue.Add optExc.Value
    If optReportDate(0).Value = True Then aFormulaValue.Add CStr("'0'") Else aFormulaValue.Add CStr("'1'")
    aFormulaValue.Add CStr(optDsc)
    aFormulaValue.Add CStr(OptCmt)


    sSOTypes = ""
    For i = 0 To 25
        If optTyp(i).Value = 1 Then sSOTypes = sSOTypes & "'" & Chr(65 + i) & "',"
    Next i
    If Len(sSOTypes) > 0 Then
        aFormulaName.Add "SalesOrderTypes"
        aFormulaValue.Add Chr(34) & "(" & Left(sSOTypes, Len(sSOTypes) - 1) & ")" & Chr(34)
        
    End If
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{PartTable.PARTREF} = '" & Compress(cmbPrt) & "'"
   
   cCRViewer.SetReportSelectionFormula sSql
   
    cCRViewer.SetSubRptSelFormula "bom.rpt", BOMSql
   
   ' report parameter
   aRptPara.Add CStr(Compress(cmbPrt))
   aRptPara.Add CStr(sCutoffDate)
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   
   ' set the sub sql variable pass the sub report name
   cCRViewer.SetSubRptDBParameters "ManufacturingOrders", aRptPara, aRptParaType
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
    ' print the copies
   
   
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
   Dim iTmp As Integer
   
   On Error Resume Next
   For b = 0 To 25
      sOptions = sOptions & Trim$(optTyp(b).Value)
   Next
   
   sOptions = sOptions & Trim(Val(optExc.Value))
   sOptions = sOptions & Trim(Val(optDsc.Value))
   sOptions = sOptions & Trim(Val(OptCmt.Value))
   For iTmp = 0 To 1
     If optReportDate(iTmp).Value = True Then sOptions = sOptions & LTrim(str(iTmp))
   Next iTmp
   For iTmp = 0 To 3
     If optMbe(iTmp).Value = True Then sOptions = sOptions & LTrim(str(iTmp))
   Next iTmp
   sOptions = sOptions & Left(txtEnd & Space(10), 10)
   sOptions = sOptions & cmbLvl
   sOptions = sOptions & Trim(Val(optMODet.Value))
   
   SaveSetting "Esi2000", "EsiProd", "mrp06", Trim(sOptions)
   
'   SaveSetting "Esi2000", "EsiProd", "mrp06a", Trim(Val(optExc.Value))
   
End Sub

Private Sub GetOptions()
   Dim b As Byte
   Dim sOptions As String
   Dim sExclude As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "mrp06", Trim(sOptions))
   If Len(Trim(sOptions)) > 0 Then
      For b = 0 To 25
         optTyp(b).Value = Val(Mid$(sOptions, b + 1, 1))
      Next
   
      If Len(sOptions) > 26 And Mid(sOptions, 27, 1) = 1 Then optExc.Value = 1 Else optExc.Value = 0
      If Len(sOptions) > 27 And Mid(sOptions, 28, 1) = 1 Then optDsc.Value = 1 Else optDsc.Value = 0
      If Len(sOptions) > 28 And Mid(sOptions, 29, 1) = 1 Then OptCmt.Value = 1 Else OptCmt.Value = 0
      If Len(sOptions) > 29 Then optReportDate(Val(Mid(sOptions, 30, 1))).Value = True
      If Len(sOptions) > 30 Then optMbe(Val(Mid(sOptions, 31, 1))).Value = True
      If Len(sOptions) > 31 Then txtEnd = Trim(Mid(sOptions, 32, 10))
      If Len(sOptions) > 41 Then cmbLvl = Mid(sOptions, 42, 1)
   
      If Len(sOptions) > 42 And Mid(sOptions, 43, 1) = 1 Then optMODet.Value = 1 Else optMODet.Value = 0
   
   End If

'   sExclude = GetSetting("Esi2000", "EsiProd", "mrp06a", Trim(sExclude))
'   If sExclude = "" Then
'      If ES_CUSTOM = "INTCOA" Then optExc.Value = vbChecked
'   Else
'      optExc.Value = Val(sExclude)
'   End If
   
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
   If lblDsc.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Part Number.", vbInformation, Caption
   Else
      GetBill
      PrintReport
   End If
   
End Sub


Private Sub optExc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   GetPart
   
   If lblDsc.ForeColor = ES_RED Then
     MsgBox "Requires A Valid Part Number.", vbInformation, Caption
   Else
     GetBill
     PrintReport
   End If
   
   
   
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




'Not using recursion here to keep the levels straight and
'make it easy to read

Private Sub GetBill()
   Dim RdoBom As ADODB.Recordset
   MouseCursor 13
   iOrder = 0
   On Error GoTo DiaErr1
   sProcName = "getbill"
   sSql = "truncate table MrpbTable"
   clsADOCon.ExecuteSQL sSql
   sSql = "INSERT INTO MrpbTable (MRPBILL_PARTREF," _
          & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
          & "VALUES('" & Compress(cmbPrt) & "','" _
          & "',0)"
   clsADOCon.ExecuteSQL sSql
   
   sBomRev = GetBomRev(Compress(cmbPrt))
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM BmplTable " _
          & "WHERE (BMASSYPART='" & Compress(cmbPrt) & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            sProcName = "getbomrev"
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',1)"
            clsADOCon.ExecuteSQL sIns
            GetNextBillLevel2 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   'PrintReport
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub GetNextBillLevel2(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel2"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',2)"
            clsADOCon.ExecuteSQL sIns
            GetNextBillLevel3 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub


Private Sub GetNextBillLevel3(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel3"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',3)"
            clsADOCon.ExecuteSQL sIns
            GetNextBillLevel4 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub

Private Sub GetNextBillLevel4(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel4"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',4)"
            clsADOCon.ExecuteSQL sIns
            GetNextBillLevel5 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub


Private Sub GetNextBillLevel5(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel5"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',5)"
            ' MM 6/19/2010 - Missing record update
            clsADOCon.ExecuteSQL sIns
            
            GetNextBillLevel6 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub


Private Sub GetNextBillLevel6(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel6"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',6)"
            clsADOCon.ExecuteSQL sIns
            GetNextBillLevel7 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub

Private Sub GetNextBillLevel7(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel7"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',7)"
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub



Private Function GetBomRev(sPartNumber) As String
   Dim RdoRev As ADODB.Recordset
   sProcName = "getbomrev"
   sSql = "SELECT PARTREF,PABOMREV FROM PartTable " _
          & "WHERE PARTREF='" & sPartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRev, ES_FORWARD)
   If bSqlRows Then
      With RdoRev
         GetBomRev = "" & Trim(!PABOMREV)
         ClearResultSet RdoRev
      End With
   Else
      GetBomRev = ""
   End If
   Set RdoRev = Nothing
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

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   cmbPrt = txtPrt
End Sub

