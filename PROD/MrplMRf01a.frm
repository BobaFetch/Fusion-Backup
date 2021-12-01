VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form MrplMRf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate A New MRP"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   840
      Picture         =   "MrplMRf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Display The Report"
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CheckBox chkZeroBalances 
      Height          =   255
      Left            =   3360
      TabIndex        =   72
      ToolTipText     =   "Sets The MRP To Zero For Negative Balances (Does Not Affect Inventory)"
      Top             =   3060
      Width           =   3375
   End
   Begin VB.ComboBox cmbCutoff 
      Height          =   315
      Left            =   1440
      TabIndex        =   70
      Tag             =   "4"
      Top             =   705
      Width           =   1575
   End
   Begin VB.TextBox txtTolerance 
      Height          =   315
      Left            =   5760
      MaxLength       =   4
      TabIndex        =   68
      Text            =   "9999"
      Top             =   2160
      Width           =   510
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MrplMRf01a.frx":017E
      Style           =   1  'Graphical
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   1440
      Picture         =   "MrplMRf01a.frx":092C
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Print The Report"
      Top             =   4320
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "&All"
      Height          =   275
      Left            =   2400
      TabIndex        =   0
      ToolTipText     =   "Mark All Sales Order Types"
      Top             =   1245
      Width           =   735
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "C&lear"
      Height          =   275
      Left            =   3240
      TabIndex        =   1
      ToolTipText     =   "Clear All Types"
      Top             =   1245
      Width           =   735
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   64
      Top             =   1860
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   63
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   62
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   960
      TabIndex        =   61
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   1200
      TabIndex        =   60
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   1440
      TabIndex        =   59
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   6
      Left            =   1680
      TabIndex        =   58
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   7
      Left            =   1920
      TabIndex        =   57
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   8
      Left            =   2160
      TabIndex        =   56
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   9
      Left            =   2400
      TabIndex        =   55
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   10
      Left            =   2640
      TabIndex        =   54
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   11
      Left            =   2880
      TabIndex        =   53
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   12
      Left            =   3120
      TabIndex        =   52
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   13
      Left            =   3360
      TabIndex        =   51
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   14
      Left            =   3600
      TabIndex        =   50
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   15
      Left            =   3840
      TabIndex        =   49
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   16
      Left            =   4080
      TabIndex        =   48
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   17
      Left            =   4320
      TabIndex        =   47
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   18
      Left            =   4560
      TabIndex        =   46
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   19
      Left            =   4800
      TabIndex        =   45
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   20
      Left            =   5040
      TabIndex        =   44
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   21
      Left            =   5280
      TabIndex        =   43
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   22
      Left            =   5520
      TabIndex        =   42
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   23
      Left            =   5760
      TabIndex        =   41
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   24
      Left            =   6000
      TabIndex        =   40
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   25
      Left            =   6240
      TabIndex        =   39
      Top             =   1860
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "MrplMRf01a.frx":0AB6
      Height          =   350
      Left            =   6840
      Picture         =   "MrplMRf01a.frx":0F90
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "View Last Log (Requires A Text Viewer)"
      Top             =   960
      Width           =   360
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "&Generate"
      Height          =   315
      Left            =   6360
      TabIndex        =   26
      ToolTipText     =   "Generate A New MRP"
      Top             =   600
      Width           =   875
   End
   Begin VB.CheckBox chkNegToZero 
      Caption         =   "(Highly Recommended)"
      Height          =   255
      Left            =   3360
      TabIndex        =   24
      ToolTipText     =   "Sets The MRP To Zero For Negative Balances (Does Not Affect Inventory)"
      Top             =   2700
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6960
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4035
      FormDesignWidth =   7395
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   3480
      TabIndex        =   67
      Top             =   3600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Create beginning balances for zero QOH"
      Height          =   315
      Index           =   6
      Left            =   240
      TabIndex        =   73
      Top             =   3060
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cutoff Date"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   71
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Tolerance in calendar days.  How many days may the QOH be negative before the deficiency is exploded?"
      Height          =   495
      Left            =   240
      TabIndex        =   69
      Top             =   2160
      Width           =   5355
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   38
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   37
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   36
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   35
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   34
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   33
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   32
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   31
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   30
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      Height          =   255
      Index           =   9
      Left            =   2400
      TabIndex        =   29
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      Height          =   255
      Index           =   10
      Left            =   2640
      TabIndex        =   27
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      Height          =   255
      Index           =   11
      Left            =   2880
      TabIndex        =   25
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   23
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      Height          =   255
      Index           =   13
      Left            =   3360
      TabIndex        =   22
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      Height          =   255
      Index           =   14
      Left            =   3600
      TabIndex        =   21
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   255
      Index           =   15
      Left            =   3840
      TabIndex        =   20
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      Height          =   255
      Index           =   16
      Left            =   4080
      TabIndex        =   19
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   255
      Index           =   17
      Left            =   4320
      TabIndex        =   18
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   255
      Index           =   18
      Left            =   4560
      TabIndex        =   17
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      Height          =   255
      Index           =   19
      Left            =   4800
      TabIndex        =   16
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      Height          =   255
      Index           =   20
      Left            =   5040
      TabIndex        =   15
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      Height          =   255
      Index           =   21
      Left            =   5280
      TabIndex        =   14
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      Height          =   255
      Index           =   22
      Left            =   5520
      TabIndex        =   13
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   23
      Left            =   5760
      TabIndex        =   12
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   255
      Index           =   24
      Left            =   6000
      TabIndex        =   11
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      Height          =   255
      Index           =   25
      Left            =   6240
      TabIndex        =   10
      Top             =   1620
      Width           =   165
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Sales Order Types:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   9
      Top             =   1245
      Width           =   2535
   End
   Begin VB.Label lblUsr 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblMrp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last MRP"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setting Parameters."
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Set MRP For Negative Inventory To Zero"
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   2700
      Width           =   3075
   End
End
Attribute VB_Name = "MrplMRf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Option Explicit
Dim AdoCon2 As ClassFusionADO
Dim rdoFetch As ADODB.Recordset
Dim rdoInsert As ADODB.Recordset
Dim RdoQoh As ADODB.Recordset

Dim bOnLoad As Byte
Dim bNextLvl As Byte

Dim iCurrSoItem As Integer
Dim iDefPurLead As Integer
Dim iDefManFlow As Integer
Dim iLogNumber As Integer

Dim lCOUNTER As Long
Dim lCurrRun As Long
Dim lCurrSo As Long

'Counters
Dim lLogRows As Long
Dim lPartCount As Long
Dim lSoItemRows As Long
Dim lMoRows As Long
Dim lPoItemRows As Long
Dim lPickRows As Long
Dim lBillRows As Long

Dim cRunqty As Currency
Dim cSoQty As Currency
Dim cAdjQuantity As Currency

Dim sCurrMo As String
Dim sCurrSoItRev As String
Dim sReplaceStr As String
Dim sMOStatus As String
Dim sSalesOrders As String
Dim sType As String

Dim vReqdDate As Date
Public vEndDate As Variant

'Dim sLogNote(44, 2)

'Column updates
Dim RdoCol As ADODB.Recordset


Private ErrorCount As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'Private Function zCheckErrors() As Byte
'   zCheckErrors = 0
'   For Each ER In rdoErrors
'      If Left(ER.Description, 5) = "22003" Then
'         zCheckErrors = 1
'      End If
'   Next ER
'
'End Function

Private Function zGetConstraint(sDescription As String) As String
   Dim bByte As Byte
   zGetConstraint = ""
   bByte = InStr(1, sDescription, "DF_")
   If bByte > 0 Then
      zGetConstraint = Mid$(sDescription, bByte, Len(sDescription))
      bByte = InStr(1, zGetConstraint, "'")
      If bByte > 0 Then zGetConstraint = Left$(zGetConstraint, bByte - 1)
   Else
      bByte = InStr(1, sDescription, "DEFZERO")
      If bByte > 0 Then
         sSql = "sp_unbindefault '" & RdoCol.Fields(2) & "." & RdoCol.Fields(3) & "'"
         clsADOCon.ExecuteSql sSql
      End If
   End If
   
   
End Function

Private Function ManufacturingOrders(Optional PreEmpt As Boolean = True) As Boolean
   Dim n As Integer
   
   'counters
   Dim A As Byte
   Dim b As Currency
   Dim C As Currency
   
   Dim cQuantity As Currency
   Dim vDate As Variant
   
   If (PreEmpt) Then DoEvents
   b = 100 / lMoRows
   prg1.Value = 5
   z1(1).Visible = True
'   z1(1).Caption = "Step 4: Manufacturing Orders"
'   z1(1).Refresh
   'On Error Resume Next
   
   Log "Start Manufacturing Orders " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   sProcName = "ManufacturingOrders"
   sSql = "Qry_GetMRPMOrders '" & vEndDate & "'"   'NOT COMPLETE, CANCELED OR CLOSED

   Set rdoFetch = AdoCon2.GetRecordSet(sSql, ES_FORWARD)

   If Not rdoFetch.BOF And Not rdoFetch.EOF Then
      bSqlRows = 1
   Else
      bSqlRows = 0
   End If
   If bSqlRows Then
      MouseCursor 11
      sSql = "SELECT MRP_ROW,MRP_CATAGORY,MRP_PARTREF,MRP_PARTNUM," _
             & "MRP_PARTDESC,MRP_PARTLEVEL,MRP_PARTUOM," _
             & "MRP_PARTPRODCODE,MRP_PARTCLASS,MRP_PARTQTYRQD," _
             & "MRP_PARTDATERQD,MRP_ACTIONDATE,MRP_PARTSORTDATE,MRP_MOREF,MRP_MONUM," _
             & "MRP_MODESC,MRP_MORUNNO,MRP_COMMENT,MRP_TYPE,MRP_USER " _
             & "FROM MrplTable"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoInsert, ES_DYNAMIC)
      With rdoFetch
         Err = 0
         clsADOCon.ADOErrNum = 0
         'clsAdoCon.begintrans
         C = 5
         Do Until .EOF
'Debug.Print lCOUNTER & ": " & !runref & " run " & !RunNo
            A = A + 1
            C = C + b
            If C >= 95 Then C = 95
            If A >= 5 Then
               prg1.Value = Int(C)
               A = 0
            End If
            cQuantity = Format(!quantityLeft, "#0.0000")
            If Not IsNull(!runSched) Then
               'vDate = Format(!RUNSCHED, "mm/dd/yy 06:00")
               vDate = Format(!runSched, "mm/dd/yy")
            Else
               'vDate = Format(ES_SYSDATE, "mm/dd/yy 06:00")
               vDate = Format(ES_SYSDATE, "mm/dd/yy")
               Log "  ** Run Scheduled Date Problem MO Number " & Trim(!PartNum) & " Run " & !Runno
            End If
            lCOUNTER = lCOUNTER + 1
            sReplaceStr = ReplaceString("" & Trim(!PADESC))
            rdoInsert.AddNew
            'RdoInsert!MRP_ROW = lCOUNTER
            Select Case !RUNSTATUS
            Case "SC", "RL"
               rdoInsert!MRP_TYPE = MRPTYPE_MoNoPicklist
               rdoInsert!MRP_CATAGORY = "MO without picklist"
            Case Else
               rdoInsert!MRP_TYPE = MRPTYPE_MoWithPicklist
               rdoInsert!MRP_CATAGORY = "MO with picklist"
            End Select
            rdoInsert!MRP_PARTREF = "" & Trim(!PartRef)
            rdoInsert!MRP_PARTNUM = "" & Trim(!PartNum)
            rdoInsert!MRP_PARTDESC = sReplaceStr
            rdoInsert!MRP_PARTLEVEL = !PALEVEL
            rdoInsert!MRP_PARTUOM = "" & Trim(!PAUNITS)
            rdoInsert!MRP_PARTPRODCODE = "" & Trim(!PAPRODCODE)
            rdoInsert!MRP_PARTCLASS = "" & Trim(!PACLASS)
            rdoInsert!MRP_PARTQTYRQD = cQuantity
            rdoInsert!MRP_PARTDATERQD = vDate
            rdoInsert!MRP_ACTIONDATE = vDate
            rdoInsert!MRP_PARTSORTDATE = vDate
            rdoInsert!MRP_USER = sInitials
            rdoInsert!MRP_MOREF = "" & Compress(!PartNum)
            rdoInsert!MRP_MONUM = "" & Trim(!PartNum)
            rdoInsert!MRP_MODESC = sReplaceStr
            rdoInsert!MRP_MORUNNO = !Runno
            rdoInsert!MRP_COMMENT = "MO " & Trim(!PartNum) & " Run " & !Runno _
                                    & " Status: " & Trim(!RUNSTATUS)
            rdoInsert!MRP_ACTIONDATE = !ActionDate
            rdoInsert.Update
            If (PreEmpt) Then DoEvents
            cAdjQuantity = cQuantity
            UpdateQoh cAdjQuantity, Compress(!PartNum), 3, True
            lLogRows = lLogRows + 1
            If (PreEmpt) Then DoEvents
            .MoveNext
         Loop
         ClearResultSet rdoFetch
         'clsAdoCon.CommitTrans
         A = 0
         Do While Not rdoInsert.EOF
            A = A + 1
            If A > 5 Then Exit Do
         Loop
         If lLogRows > 500 Then WriteTransactionLog
      End With
   End If
   
   Log "End Manufacturing Orders   " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   Log "  "
   prg1.Value = 100
   
End Function

Private Sub cmdAll_Click()
   Dim b As Byte
   For b = 0 To 24
      optTyp(b).Value = vbChecked
   Next
   optTyp(b).Value = vbChecked
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdGen_Click()
   On Error Resume Next
   vEndDate = Format(cmbCutoff, "mm/dd/yy 23:59")
   If Err > 0 Then vEndDate = Format(ES_SYSDATE, "mm/dd/yy 23:59")
   GenerateMrp
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4501
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdNone_Click()
   Dim b As Byte
   For b = 0 To 24
      optTyp(b).Value = vbUnchecked
   Next
   optTyp(b).Value = vbUnchecked
   
End Sub

Private Sub cmdVew_Click()
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   
   sCustomReport = GetCustomReport("mrplog")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

    cCRViewer.CRViewerSize Me
    cCRViewer.SetDbTableConnection

   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
'
'   SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'   sCustomReport = GetCustomReport("mrplog")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      'zMakeParaTable      'moved to database.bas
      'zCheckStoredProcs
      'zCheckTableColumns
      'CheckLogTable
      'RdoCon.QueryTimeout = 1000    'allow really long timeouts
      GetMrpDefaults
      GetLastMrp
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetSettings
   GetOptions
  ' sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMQTYREQD," _
  '        & "BMCONVERSION,PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," _
  '        & "PACLASS,PAPRODCODE,PABOMREV,PAFLOWTIME,PALEADTIME " _
  '        & "FROM BmplTable,PartTable WHERE " _
  '        & "(BMPARTREF=PARTREF AND BMASSYPART= ? AND BMREV=? " _
  '        & "AND PALEVEL<5)"
  ' Set rdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   SaveSettings
   FormUnload
   'Set rdoQry = Nothing
   Set rdoFetch = Nothing
   Set rdoInsert = Nothing
   Set RdoQoh = Nothing
   Set RdoCol = Nothing
   Set AdoCon2 = Nothing
   On Error Resume Next
   Set MrplMRf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   Dim horizon As Integer, iList As Integer
   'Dim l As Long
   Dim S As String
   
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'   s = GetSetting("Esi2000", "EsiProd", "MRPRange", "undefined")
'   If s = "undefined" Then
'      SaveSetting "Esi2000", "EsiProd", "MRPRange", "365"
'      horizon = 365
'   Else
'      horizon = CInt(s)
'   End If
'   If horizon = 0 Then horizon = 30
'   If ES_CUSTOM = "ESINORTH" Then
'      cmbcutoff = Format(Now + 420, "mm/dd/yyyy")
'   Else
'      cmbcutoff = Format(Now + horizon, "mm/dd/yyyy")
'   End If
   'l = GetSetting("Esi2000", "EsiProd", "MRPMos", l)
   For iList = 0 To 25
      optTyp(iList).TabIndex = iList + 3
   Next
   'optTyp(iList).TabIndex = iList + 3
   
End Sub

Private Sub chknegtozero_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub GenerateMrp()
   Dim b As Boolean
   Dim bResponse As Byte
   Dim sMsg As String
   Dim vDate As Variant
   
   On Error GoTo DiaErr1
'   TruncateReportLog
   
   b = CheckSalesOrders()
   If b = 0 Then
      MsgBox "Requires That At Least One Sales Order Type Be Selected.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   sMsg = "Do You Want To Remove The Current  " & vbCr _
          & "(If Any) MRP And Create a New One?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      sMsg = "Are You Certain That The System Is Quiet  " & vbCr _
             & "And Other Users Are Not Affecting Results?.."
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         TruncateReportLog
         lLogRows = 0
         lCOUNTER = 0
         cmdCan.Enabled = False
         MouseCursor 11
         prg1.Visible = True
         prg1.Value = 10
         z1(1).Visible = True
         z1(1).Caption = "Setting Parameters."
         z1(1).Refresh
         Sleep 100
         'On Error Resume Next
         'On Err GoTo DiaErr1
         lSoItemRows = zGetSOItemRows()
         If lSoItemRows = 0 Then lSoItemRows = 300
         
         lMoRows = zGetMORows()
         If lMoRows = 0 Then lMoRows = 100
         
         lPoItemRows = zGetPOItemRows()
         If lPoItemRows = 0 Then lPoItemRows = 100
         
         lPickRows = zGetMOPickRows()
         If lPickRows = 0 Then lPickRows = 100
         
         lBillRows = zGetSchedMORows()
         If lBillRows = 0 Then lBillRows = 100
         'b = OpenLog()
         'iLogNumber = 1
         
         ErrorCount = 0
         
         Log "** Flags Scheduled Date Problems (Null or no date)"
         Log " "
         Log "Start MRP                 " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
         Log " "
         Log "Default Manufacturing Flow Time Is " & iDefManFlow & " Days"
         Log "Default Purchasing Lead Time Is " & iDefPurLead & " Days"
         If iDefManFlow = 0 Or iDefPurLead = 0 Then
            Log "** One Or Both Defaults Are Zero (Company Setting) "
         Else
            Log " "
         End If
         
         OpenSecondCon
         
         sSql = "TRUNCATE TABLE MrplTable"
         clsADOCon.ExecuteSql sSql
         sSql = "DBCC CHECKIDENT ('MrplTable', RESEED, 1)"
         clsADOCon.ExecuteSql sSql
         
         sSql = "TRUNCATE TABLE MrppTable"
         clsADOCon.ExecuteSql sSql
         'Sleep 1000
         
         cmdVew.Enabled = False
         cmdGen.Enabled = False
         
         ' *** Start to fill. In order of requirements ***
         
         'get inventory beginning balances
         z1(1).Caption = "Step 1: Beginning Balance"
         b = BeginningBalance()
         
         'get safety stock
         z1(1).Caption = "Step 2: Safety Stock"
         b = SafetyStock()
         
         'get inventory ins for unreceived PO items
         z1(1).Caption = "Step 3: Purchase Orders"
         b = PurchaseOrderItems()
         
         'Get inventory ins for incomplete MO items
         z1(1).Caption = "Step 4: Manufacturing Orders"
         b = ManufacturingOrders()
         
         'Get inventory outs for unshipped SO items
         z1(1).Caption = "Step 5: Sales Orders"
         b = SalesOrderItems()
         
         'Get MO unpicked pick list items
         z1(1).Caption = "Step 6: MO Picks"
         b = MOPicks()
         
         'Explode SC and RL MOs without picklists one level
         z1(1).Caption = "Step 7: Explode SC and RL MO's one level"
         ExplodeMoPicks
         
         'keep adding action items until no more are required
         Dim Pass As Integer
         Pass = 1
'         Do While AddActionItems(Pass)             '7, 9, 11 ...
'            If Not ExplodeActionItems(Pass) Then   '8, 10, 12 ...
'               Exit Do
'            End If
'
'            Pass = Pass + 1
'         Loop

         Do While True
            z1(1).Caption = "Step " & 6 + 2 * Pass & ": Add action items"
            If Not AddActionItems(Pass) Then             '8, 10, 12 ...
               Exit Do
            End If
            z1(1).Caption = "Step " & 7 + 2 * Pass & ": Explode action items"
            If Not ExplodeActionItems(Pass) Then   '9, 11, 12 ...
               Exit Do
            End If
            
            Pass = Pass + 1
            If Pass > 7 Then
               Exit Do
            End If
            
         Loop

         'close connection stuff no longer needed
         If Not rdoInsert Is Nothing Then
            'If (rdoInsert.State = adStateOpen) Then rdoInsert.Close
            Set rdoInsert = Nothing
         End If
         If Not rdoFetch Is Nothing Then
            If Not rdoFetch.ActiveConnection Is Nothing Then
               If (rdoFetch.State = adStateOpen) Then rdoFetch.Close
            End If
            Set rdoFetch = Nothing
         End If
         'clsAdoCon.CommitTrans
         
         'Set shortage indicator where end quantity < 0.
         'Also set buyers for parts.
         b = MarkShortages()
         Log "End MRP                   " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
         
         'Clean empty rows
         sSql = "DELETE FROM MrplTable WHERE MRP_PARTREF=''"
         clsADOCon.ExecuteSql sSql
         
         sSql = "UPDATE Table2 SET Table2.MRP_PAMAKEBUY=MrplTable.MRP_PAMAKEBUY " _
                & "FROM MrplTable,MrplTable as Table2 WHERE " _
                & "MrplTable.MRP_PARTREF = Table2.MRP_PARTREF AND " _
                & "MrplTable.MRP_TYPE = " & MRPTYPE_BeginningBalance
         clsADOCon.ExecuteSql sSql
         
         MouseCursor 0
         cmdCan.Enabled = True
         If ErrorCount > 0 Then
            MsgBox "MRP Completed with " & ErrorCount & " errors.  " & Format(ES_SYSDATE, "mm/dd/yy hh:mm AM/PM") & ".  See log.", _
               vbInformation, Caption
         Else
            MsgBox "MRP Completed " & Format(ES_SYSDATE, "mm/dd/yy hh:mm AM/PM") & ".  See log.", _
               vbInformation, Caption
         End If
         
         'clsAdoCon.CommitTrans
         sSql = "UPDATE MrpdTable SET MRP_CREATEDATE='" & Format(ES_SYSDATE, "mm/dd/yy hh:mm AM/PM") & "'," _
                & "MRP_THROUGHDATE='" & cmbCutoff & "',MRP_CREATEDBY='" & sInitials & "' " _
                & "WHERE MRP_ROW=1"
         clsADOCon.ExecuteSql sSql
         
         ' Display the last MRP date.
         GetLastMrp
         'Set the Sort Date to the min date recorded (Garretts 1950)
         sSql = "SELECT MIN(MRP_PARTDATERQD) AS StartDate FROM MrplTable"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoQoh, ES_FORWARD)
         If bSqlRows Then
            With RdoQoh
               If Not IsNull(!startDate) Then
                  vDate = Format(!startDate - 2, "mm/dd/yy")
                  sSql = "UPDATE MrplTable SET MRP_PARTSORTDATE='" _
                         & vDate & "' WHERE MRP_TYPE=" & MRPTYPE_BeginningBalance      '1
                  clsADOCon.ExecuteSql sSql
               End If
               .Cancel
            End With
         End If
         
         'WriteReportLog
         cmdVew.Enabled = True
         cmdVew.SetFocus
         cmdGen.Enabled = True
         prg1.Value = 0
         prg1.Visible = False
         z1(1) = ""
'         rdoCon2.Close
'         rdoEnv2.Close
      Else
         CancelTrans
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "GenerateMrp"
   prg1.Visible = False
   z1(1) = ""
   cmdCan.Enabled = True
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function AddActionItems(Pass As Integer, Optional PreEmpt As Boolean = True) As Boolean
   'returns False if no action items were required
   
   Dim poCount As Long
   Dim moCount As Long
   Dim tolerance As Integer
   Dim sBuyer As String
   tolerance = CInt(Me.txtTolerance)
   
   On Error GoTo DiaErr1
'   z1(1).Caption = "Step " & 6 + 2 * Pass & ": Add action items"
   Log "Start adding action items pass " & Pass & " " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   If Pass > 8 Then
      Log "MRP explosion was limited to 8 levels."
      Exit Function
   End If
   
'   'get MRP entries where balance goes negative at any time
'   sSql = "select vb.*, PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
'      & "PAQOH,PACLASS,PAPRODCODE,PAFLOWTIME,PALEADTIME," & vbCrLf _
'      & "PAMAKEBUY from viewMrpBalances vb" & vbCrLf _
'      & "join PartTable part on vb.MRP_PARTNUM = part.PARTREF" & vbCrLf _
'      & "where vb.MRP_PARTNUM in (select distinct MRP_PARTNUM from viewMRPBalances" & vbCrLf _
'      & "where Balance < 0 )" & vbCrLf _
'      & "order by vb.SortOrder"
      
   'faster to just get all rows in sorted order
   sSql = "select mrp.*,PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
      & "PAQOH,PACLASS,PAPRODCODE,PAFLOWTIME,PALEADTIME," & vbCrLf _
      & "PAMAKEBUY,MRP_SOTEXT,MRP_SOITEM, MRP_SOREV from MrplTable mrp" & vbCrLf _
      & "join PartTable part on mrp.MRP_PARTREF = part.PARTREF" & vbCrLf _
      & "order by MRP_PARTNUM, MRP_PARTSORTDATE, MRP_TYPE, MRP_ROW"
      
   If clsADOCon.GetDataSet(sSql, rdoFetch, ES_FORWARD) Then
      sSql = "SELECT MRP_ROW,MRP_CATAGORY,MRP_PARTREF,MRP_PARTNUM," _
             & "MRP_PARTDESC,MRP_PARTLEVEL,MRP_PARTUOM," _
             & "MRP_PARTPRODCODE,MRP_PARTCLASS,MRP_PARTQTYRQD,MRP_PARTDATERQD," _
             & "MRP_PARTSORTDATE,MRP_TYPE,MRP_COMMENT,MRP_USER," _
             & "MRP_PAMAKEBUY,MRP_PARENTROW,MRP_ACTIONDATE,MRP_POBUYER, MRP_SOTEXT,MRP_SOITEM, MRP_SOREV FROM MrplTable"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoInsert, ES_DYNAMIC)
      With rdoFetch
      
         Dim PartRef As String, priorPartRef As String '
         Dim partBalance As Currency
         Dim bRecNew As Boolean
         bRecNew = False
         
         Do Until .EOF
            PartRef = Trim(!MRP_PARTREF)
            
            'same part as prior row?
            If priorPartRef = PartRef Then
               partBalance = partBalance + !MRP_PARTQTYRQD
               'if negative for longer than tolerance period, add action item
               
               If bRecNew = True Then 'rdoInsert.Status = adRecNew
                  ' if there is no purchase number or MO then create an action item
                  ' If the tolerence is set check to see if we are good with int he tolerence and
                  ' go negative with out creating an actionitem.
                  If (!MRP_PARTQTYRQD < 0) And DateDiff("d", rdoInsert!MRP_PARTDATERQD, !MRP_PARTDATERQD) > tolerance And _
                    (rdoInsert!MRP_TYPE = MRPTYPE_PoActionItem Or rdoInsert!MRP_TYPE = MRPTYPE_MoActionItem) Then
                       If rdoInsert!MRP_TYPE = MRPTYPE_PoActionItem Then
                          poCount = poCount + 1
                       Else
                          moCount = moCount + 1
                       End If
                       rdoInsert.Update
                       bRecNew = False
                       ' Added to stop the MRP qty adding up
                       partBalance = !MRP_PARTQTYRQD
                       If (PreEmpt) Then DoEvents
                  Else
                    If DateDiff("d", rdoInsert!MRP_PARTDATERQD, !MRP_PARTDATERQD) > tolerance Then
                       partBalance = partBalance + rdoInsert!MRP_PARTQTYRQD
                       If rdoInsert!MRP_TYPE = MRPTYPE_PoActionItem Then
                          poCount = poCount + 1
                       Else
                          moCount = moCount + 1
                       End If
                       rdoInsert.Update
                       bRecNew = False
                       If (PreEmpt) Then DoEvents
                    
                    'same tolerance period.  update quantity to last one in period if negative
                    'if positive, cancel update since no action item is required
                    Else
                       ' MM I think this should be not there as we partbalace in the top
                       ' MM partBalance = partBalance + !MRP_PARTQTYRQD
                       If partBalance >= 0 Then
                          rdoInsert.CancelUpdate
                          bRecNew = False
                       Else
                          rdoInsert!MRP_PARTQTYRQD = -partBalance
                       End If
                    End If
                 End If
               End If
            
            'different part.  If prior part ended with negative balance, output it
            Else
               If bRecNew = True Then 'adRecNew
                  If rdoInsert!MRP_TYPE = MRPTYPE_PoActionItem Then
                     poCount = poCount + 1
                  Else
                     moCount = moCount + 1
                  End If
                  rdoInsert.Update
                  ' MM Added
                  bRecNew = False
                  
                  If (PreEmpt) Then DoEvents
               End If
               partBalance = !MRP_PARTQTYRQD
               priorPartRef = PartRef
            End If
            
            'if there's no current record being added and the balance is negative
            'add a potential record now
            If bRecNew <> True And partBalance < 0 Then
               rdoInsert.AddNew
               bRecNew = True
               If Trim(!PAMAKEBUY) = "B" Then
                  rdoInsert!MRP_TYPE = MRPTYPE_PoActionItem
                  rdoInsert!MRP_CATAGORY = "*PO action item"
                  If !PALEADTIME = 0 Then
                     rdoInsert!MRP_ACTIONDATE = DateAdd("d", -iDefPurLead, !MRP_PARTDATERQD)
                  Else
                     rdoInsert!MRP_ACTIONDATE = DateAdd("d", -!PALEADTIME, !MRP_PARTDATERQD)
                  End If
                  rdoInsert!MRP_COMMENT = "*Issue PO by " _
                     & Format(rdoInsert!MRP_ACTIONDATE, "mm/dd/yy")
               Else
                  rdoInsert!MRP_TYPE = MRPTYPE_MoActionItem
                  rdoInsert!MRP_CATAGORY = "*MO action item"
                  If !PAFLOWTIME = 0 Then
                     rdoInsert!MRP_ACTIONDATE = DateAdd("d", -iDefManFlow, !MRP_PARTDATERQD)
                  Else
                     rdoInsert!MRP_ACTIONDATE = DateAdd("d", -!PAFLOWTIME, !MRP_PARTDATERQD)
                  End If
                  rdoInsert!MRP_COMMENT = "*Schedule MO to start by " _
                     & Format(rdoInsert!MRP_ACTIONDATE, "mm/dd/yy")
               End If
               ' Get PO Action Buyer
               sBuyer = GetPoActionBuyer(Trim(!PartRef))
               
               rdoInsert!MRP_PARTREF = "" & Trim(!PartRef)
               rdoInsert!MRP_PARTNUM = "" & Trim(!PartNum)
               rdoInsert!MRP_PARTDESC = ReplaceString("" & Trim(!PADESC))
               rdoInsert!MRP_PARTLEVEL = !PALEVEL
               rdoInsert!MRP_PARTUOM = "" & Trim(!PAUNITS)
               rdoInsert!MRP_PARTPRODCODE = "" & Trim(!PAPRODCODE)
               rdoInsert!MRP_PARTCLASS = "" & Trim(!PACLASS)
               rdoInsert!MRP_PARTQTYRQD = -partBalance
               rdoInsert!MRP_PARTDATERQD = Format(!MRP_PARTDATERQD, "mm/dd/yy")
               rdoInsert!MRP_PARTSORTDATE = rdoInsert!MRP_PARTDATERQD
               rdoInsert!MRP_USER = sInitials
               rdoInsert!MRP_PAMAKEBUY = "" & Trim(!PAMAKEBUY)
               rdoInsert!MRP_PARENTROW = !MRP_ROW
               
               rdoInsert!MRP_SOTEXT = !MRP_SOTEXT
               rdoInsert!MRP_SOITEM = !MRP_SOITEM
               rdoInsert!MRP_SOREV = !MRP_SOREV
               
               rdoInsert!MRP_POBUYER = sBuyer
            End If
            If (PreEmpt) Then DoEvents
            .MoveNext
         Loop
               
         'end of data.  output last row if part ends with negative quantity
         If bRecNew = True Then 'adRecNew
            If rdoInsert!MRP_TYPE = MRPTYPE_PoActionItem Then
               poCount = poCount + 1
            Else
               moCount = moCount + 1
            End If
            rdoInsert.Update
            bRecNew = False
            If (PreEmpt) Then DoEvents
         End If
         ClearResultSet rdoFetch
         WriteTransactionLog
         rdoFetch.Close
      End With
   End If
   
   Log "      PO Action items added: " & poCount
   Log "      MO Action items added: " & moCount
   Log "End adding action items pass " & Pass & " " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   Log " "
   
   If poCount > 0 Or moCount > 0 Then
      AddActionItems = True
   Else
      AddActionItems = False
   End If
   Exit Function
   
DiaErr1:
   ErrorCount = ErrorCount + 1
   sProcName = "AddActionItems"
   prg1.Visible = False
   z1(1) = ""
   cmdCan.Enabled = True
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Log "* Error in AddActionItems: " & Err.Description
   DoModuleErrors Me
   
End Function

Private Function ExplodeActionItems(Pass As Integer, Optional PreEmpt As Boolean = True) As Boolean
   'returns False if no action items were exploded
   Dim sSalesOrder As String
   Dim sTempSO, sTempRev As String
 
   Dim explCount As Long
   Dim tolerance As Integer
   tolerance = CInt(Me.txtTolerance)
   
   On Error GoTo DiaErr1
'   z1(1).Caption = "Step " & 7 + 2 * Pass & ": Add action items"
   Log "Start exploding action items pass " & Pass & " " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
'   If Pass > 7 Then
'      Exit Function
'   End If
          
   'explode unexploded mo action items one level
   'only explode items not previously exploded

   sSql = "SELECT rtrim(BMASSYPART) as ParentPart, rtrim(BMPARTREF) as ChildPart," & vbCrLf _
      & "BMQTYREQD,BMCONVERSION,BMADDER,BMSETUP,BMPHANTOM," & vbCrLf _
      & "PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
      & "PACLASS,PAPRODCODE,PABOMREV,PAFLOWTIME,PALEADTIME,PAMAKEBUY," & vbCrLf _
      & "MRP_ROW, MRP_PARTQTYRQD, MRP_PARTDATERQD,MRP_ACTIONDATE," & vbCrLf _
      & "cast(((BMQTYREQD + BMADDER) " & vbCrLf _
      & "/ case when BMCONVERSION=0 then 1 else BMCONVERSION end)" & vbCrLf _
      & "* MRP_PARTQTYRQD + BMSETUP as decimal(12,3)) as Quantity, " & vbCrLf _
      & "MRP_SOTEXT,MRP_SOITEM, MRP_SOREV " & vbCrLf _
      & " from MrplTable mrp" & vbCrLf _
      & "join BmplTable bom on BMASSYPART=MRP_PARTREF" & vbCrLf _
      & "join PartTable part on BMPARTREF = PARTREF and PALEVEL<5" & vbCrLf _
      & "where MRP_TYPE=" & MRPTYPE_MoActionItem & " and MRP_ROW not in (select MRP_PARENTROW from MrplTable where MRP_PARENTROW is not null)" & vbCrLf _
      & "order by PARTREF, BMSEQUENCE"
          
    
   If clsADOCon.GetDataSet(sSql, rdoFetch, ES_FORWARD) Then
      sSql = "SELECT MRP_ROW,MRP_CATAGORY,MRP_PARTREF,MRP_PARTNUM," _
             & "MRP_PARTDESC,MRP_PARTLEVEL,MRP_PARTUOM," _
             & "MRP_PARTPRODCODE,MRP_PARTCLASS,MRP_PARTQTYRQD,MRP_PARTDATERQD," _
             & "MRP_PARTSORTDATE,MRP_TYPE,MRP_COMMENT,MRP_USER," _
             & "MRP_PAMAKEBUY,MRP_PARENTROW, MRP_ACTIONDATE FROM MrplTable"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoInsert, ES_DYNAMIC)
      With rdoFetch
      
         Dim PartRef As String, priorPartRef As String, ParentPartRef As String
         Dim partBalance As Currency
         
         Do Until .EOF
            PartRef = Trim(!ChildPart)
            ParentPartRef = Trim(!ParentPart)
            
'Debug.Print partRef & " qty " & !Quantity & " " & Format(!MRP_ACTIONDATE, "mm/dd/yy") & " row " & !MRP_ROW
               
            sTempSO = "" & Trim(!MRP_SOTEXT)
            sTempRev = "" & Trim(str(!MRP_SOITEM))
            If sTempRev <> "" Then sTempRev = "-" + sTempRev
            rdoInsert.AddNew

            explCount = explCount + 1
            
            sSalesOrder = ""
            
            If !BMPHANTOM = 1 Then
               rdoInsert!MRP_TYPE = MRPTYPE_PhantomExplosion
               rdoInsert!MRP_CATAGORY = "Phantom Expl pass " & Pass
            Else
               rdoInsert!MRP_TYPE = MRPTYPE_Explosion
               rdoInsert!MRP_CATAGORY = "Explosion pass " & Pass
               If Pass = 1 And sTempSO <> "" Then
                'If Trim(rdoInsert!MRP_SOITEM) = "" Then sTempRev = "" Else sTempRev = "-" & Trim(rdoInsert!MRP_SOITEM)
                sSalesOrder = " ( SO: " & Trim(sTempSO) & sTempRev & ")"
               End If
            End If
            rdoInsert!MRP_COMMENT = "Explode " & !ParentPart & sSalesOrder
    
            rdoInsert!MRP_PARTREF = PartRef
            rdoInsert!MRP_PARTNUM = "" & Trim(!PartNum)
            rdoInsert!MRP_PARTDESC = ReplaceString("" & Trim(!PADESC))
            rdoInsert!MRP_PARTLEVEL = !PALEVEL
            rdoInsert!MRP_PARTUOM = "" & Trim(!PAUNITS)
            rdoInsert!MRP_PARTPRODCODE = "" & Trim(!PAPRODCODE)
            rdoInsert!MRP_PARTCLASS = "" & Trim(!PACLASS)
            rdoInsert!MRP_PARTQTYRQD = -!Quantity
            rdoInsert!MRP_PARTDATERQD = !MRP_ACTIONDATE
            rdoInsert!MRP_ACTIONDATE = !MRP_ACTIONDATE
            rdoInsert!MRP_PARTSORTDATE = rdoInsert!MRP_PARTDATERQD
            rdoInsert!MRP_USER = sInitials
            rdoInsert!MRP_PAMAKEBUY = "" & Trim(!PAMAKEBUY)
            rdoInsert!MRP_PARENTROW = !MRP_ROW
            rdoInsert.Update
            If (PreEmpt) Then DoEvents
            .MoveNext
         Loop
               
         ClearResultSet rdoFetch
         WriteTransactionLog
         rdoFetch.Close
      End With
   End If
   
   Log "      Explosion items added: " & explCount
   Log "End exploding action items pass " & Pass & " " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   Log " "
   
   If explCount > 0 Then
      ExplodeActionItems = True
   Else
      ExplodeActionItems = False
   End If
   Exit Function
   
DiaErr1:
   ErrorCount = ErrorCount + 1
   sProcName = "ExplodeActionItems"
   prg1.Visible = False
   z1(1) = ""
   cmdCan.Enabled = True
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Log "* Error in ExplodeActionItems: " & Err.Description
   DoModuleErrors Me
   
End Function

Private Function ExplodeMoPicks(Optional PreEmpt As Boolean = True) As Boolean
   'returns False if no SC or RL MOs were exploded
   
   Dim explCount As Long
   Dim tolerance As Integer
   tolerance = CInt(Me.txtTolerance)
   
   On Error GoTo DiaErr1
'   z1(1).Caption = "Step 7: Explode SC and RL MO's one level"
   Log "Start exploding SC and RL MOs " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
          
   'explode SC and RL MOs one level
   sSql = "SELECT rtrim(BMASSYPART) as ParentPart, rtrim(BMPARTREF) as ChildPart," & vbCrLf _
      & "BMQTYREQD,BMCONVERSION,BMADDER,BMSETUP, BMPHANTOM," & vbCrLf _
      & "PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
      & "PACLASS,PAPRODCODE,PABOMREV,PAFLOWTIME,PALEADTIME,PAMAKEBUY," & vbCrLf _
      & "MRP_ROW, MRP_PARTQTYRQD, MRP_PARTDATERQD,MRP_ACTIONDATE,MRP_MORUNNO," & vbCrLf _
      & "cast(((BMQTYREQD + BMADDER) " & vbCrLf _
      & "/ case when BMCONVERSION=0 then 1 else BMCONVERSION end)" & vbCrLf _
      & "* MRP_PARTQTYRQD + BMSETUP as decimal(12,3)) as Quantity" & vbCrLf _
      & "from MrplTable mrp" & vbCrLf _
      & "join BmplTable bom on BMASSYPART=MRP_PARTREF" & vbCrLf _
      & "join PartTable part on BMPARTREF = PARTREF and PALEVEL<5" & vbCrLf _
      & "where MRP_TYPE=" & MRPTYPE_MoNoPicklist & vbCrLf _
      & "order by PARTREF, BMSEQUENCE"
          
'      & "where MRP_TYPE=" & MRPTYPE_MoNoPicklist & " and MRP_ROW not in (select MRP_PARENTROW from MrplTable where MRP_PARENTROW is not null)" & vbCrLf _

   If clsADOCon.GetDataSet(sSql, rdoFetch, ES_FORWARD) Then
      sSql = "SELECT MRP_ROW,MRP_CATAGORY,MRP_PARTREF,MRP_PARTNUM," _
             & "MRP_PARTDESC,MRP_PARTLEVEL,MRP_PARTUOM," _
             & "MRP_PARTPRODCODE,MRP_PARTCLASS,MRP_PARTQTYRQD,MRP_PARTDATERQD," _
             & "MRP_PARTSORTDATE,MRP_TYPE,MRP_COMMENT,MRP_USER," _
             & "MRP_PAMAKEBUY,MRP_PARENTROW, MRP_ACTIONDATE FROM MrplTable"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoInsert, ES_DYNAMIC)
      With rdoFetch
      
         Dim PartRef As String, priorPartRef As String '
         Dim partBalance As Currency
         
         Do Until .EOF
            PartRef = Trim(!ChildPart)
'Debug.Print partRef & " qty " & !Quantity & " " & Format(!MRP_ACTIONDATE, "mm/dd/yy") & " row " & !MRP_ROW
               
            rdoInsert.AddNew

            explCount = explCount + 1
            If !BMPHANTOM = 1 Then
               rdoInsert!MRP_TYPE = MRPTYPE_PhantomPick
               rdoInsert!MRP_CATAGORY = "Phantom Pick"
            Else
               rdoInsert!MRP_TYPE = MRPTYPE_UnschedPick
               rdoInsert!MRP_CATAGORY = "Unsched MO Pick"
            End If
            rdoInsert!MRP_COMMENT = "Explode MO " & !ParentPart & " run " & !MRP_MORUNNO
            rdoInsert!MRP_PARTREF = PartRef
            rdoInsert!MRP_PARTNUM = "" & Trim(!PartNum)
            rdoInsert!MRP_PARTDESC = ReplaceString("" & Trim(!PADESC))
            rdoInsert!MRP_PARTLEVEL = !PALEVEL
            rdoInsert!MRP_PARTUOM = "" & Trim(!PAUNITS)
            rdoInsert!MRP_PARTPRODCODE = "" & Trim(!PAPRODCODE)
            rdoInsert!MRP_PARTCLASS = "" & Trim(!PACLASS)
            rdoInsert!MRP_PARTQTYRQD = -!Quantity
            rdoInsert!MRP_PARTDATERQD = !MRP_ACTIONDATE
            rdoInsert!MRP_ACTIONDATE = !MRP_ACTIONDATE
            rdoInsert!MRP_PARTSORTDATE = rdoInsert!MRP_PARTDATERQD
            rdoInsert!MRP_USER = sInitials
            rdoInsert!MRP_PAMAKEBUY = "" & Trim(!PAMAKEBUY)
            rdoInsert!MRP_PARENTROW = !MRP_ROW
            rdoInsert.Update
            If (PreEmpt) Then DoEvents
            .MoveNext
         Loop
               
         ClearResultSet rdoFetch
         WriteTransactionLog
         rdoFetch.Close
      End With
   End If
   
   Log "      Explosion items added: " & explCount
   Log "End exploding SC and RL MOs " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   Log " "
   
   If explCount > 0 Then
      ExplodeMoPicks = True
   Else
      ExplodeMoPicks = False
   End If
   Exit Function
   
DiaErr1:
   ErrorCount = ErrorCount + 1
   sProcName = "ExplodeMoPicks"
   prg1.Visible = False
   z1(1) = ""
   cmdCan.Enabled = True
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Log "* Error in ExplodeActionItems: " & Err.Description
   DoModuleErrors Me
   
End Function


Private Function BeginningBalance(Optional PreEmpt As Boolean = True) As Boolean
   'Get the beginning balance. Set to zero for MRP if requested
   On Error GoTo DiaErr1
   
   Dim iProg As Integer
   Dim n As Integer
   Dim iNeg As Integer
   
   Dim iMissingFlow As Integer
   Dim iMissingLead As Integer
   
   Dim cQuantity As Currency
   Dim sStar As String
   
   'Counters
   Dim A As Integer
   Dim b As Currency
   Dim C As Currency
   
   MouseCursor 11
'   z1(1).Caption = "Step 1: Beginning Balance"
'   z1(1).Refresh
   
   lPartCount = zGetMRPPartCount()
   If lPartCount > 0 Then
      b = 100 / lPartCount
   Else
      b = 0
   End If
   prg1.Value = 5
   z1(1).Visible = True
   z1(1).Refresh
   
   sProcName = "BeginningBalance"
   
   Log "Start Beginning Inventory " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," _
      & "PAQOH,PACLASS,PAPRODCODE,PAFLOWTIME,PALEADTIME," _
      & "PAMAKEBUY" & vbCrLf _
      & "FROM PartTable" & vbCrLf _
      & "WHERE PALEVEL < 5 AND PATOOL = 0 " & vbCrLf _
      & " AND PAINACTIVE = 0 AND PAOBSOLETE = 0 "

   If Me.chkZeroBalances.Value = vbUnchecked Then
      sSql = sSql & "and PAQOH <> 0" & vbCrLf
   Else
      sSql = sSql & "and PAQOH IS NOT NULL" & vbCrLf
   End If
   sSql = sSql & "ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoFetch, ES_FORWARD)
   If bSqlRows Then
      sSql = "SELECT MRP_ROW,MRP_CATAGORY,MRP_PARTREF,MRP_PARTNUM," _
             & "MRP_PARTDESC,MRP_PARTLEVEL,MRP_PARTUOM," _
             & "MRP_PARTPRODCODE,MRP_PARTCLASS,MRP_PARTQTYRQD,MRP_PARTDATERQD," _
             & "MRP_PARTSORTDATE,MRP_TYPE,MRP_COMMENT,MRP_USER," _
             & "MRP_PAMAKEBUY FROM MrplTable"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoInsert, ES_KEYSET)
      
      With rdoFetch
         Do Until .EOF
            A = A + 1
            C = C + b
            If C > 95 Then C = 95
            If A > 4 Then
               prg1.Value = Int(C)
               A = 0
            End If
            sStar = " "
            If !PAQOH < 0 Then iNeg = iNeg + 1
            If chkNegToZero.Value = vbChecked Then
               If !PAQOH < 0 Then
                  cQuantity = 0
                  sStar = "* (Actual Negative)"
               Else
                  cQuantity = !PAQOH
               End If
            Else
               cQuantity = !PAQOH
            End If
            If Trim(!PAMAKEBUY) = "B" Then
            Else
               If !PAFLOWTIME = 0 Then iMissingFlow = iMissingFlow + 1
            End If
            sReplaceStr = ReplaceString("" & Trim(!PADESC))
            
            lCOUNTER = lCOUNTER + 1
            rdoInsert.AddNew
            'RdoInsert!MRP_ROW = lCOUNTER
            rdoInsert!MRP_TYPE = MRPTYPE_BeginningBalance
            rdoInsert!MRP_CATAGORY = "Beg MRP Inventory " & Left(sStar, 1)
            rdoInsert!MRP_PARTREF = "" & Trim(!PartRef)
            rdoInsert!MRP_PARTNUM = "" & Trim(!PartNum)
            rdoInsert!MRP_PARTDESC = sReplaceStr
            rdoInsert!MRP_PARTLEVEL = !PALEVEL
            rdoInsert!MRP_PARTUOM = "" & Trim(!PAUNITS)
            rdoInsert!MRP_PARTPRODCODE = "" & Trim(!PAPRODCODE)
            rdoInsert!MRP_PARTCLASS = "" & Trim(!PACLASS)
            rdoInsert!MRP_PARTQTYRQD = cQuantity
            rdoInsert!MRP_PARTDATERQD = Format(ES_SYSDATE, "mm/dd/yy")
            rdoInsert!MRP_PARTSORTDATE = "01/01/2000" 'Always first in line
            rdoInsert!MRP_COMMENT = "Current MRP Quantity On Hand " & sStar
            rdoInsert!MRP_USER = sInitials
            rdoInsert!MRP_PAMAKEBUY = "" & Trim(!PAMAKEBUY)
            rdoInsert.Update
            If (PreEmpt) Then DoEvents
            
            
            sSql = "INSERT INTO MrppTable (MRP_PARTREF,MRP_PARTQOH) " _
                   & "VALUES('" & Trim(!PartRef) & "'," & cQuantity & ")"
            clsADOCon.ExecuteSql sSql
            lLogRows = lLogRows + 1
            If (PreEmpt) Then DoEvents
            
            .MoveNext
         Loop
         ClearResultSet rdoFetch
         lPartCount = lCOUNTER
         iProg = 0
         Do While Not rdoInsert.EOF
            iProg = iProg + 1
            If iProg > 5 Then Exit Do
         Loop
         If lLogRows > 500 Then WriteTransactionLog
         rdoFetch.Close
         
         AdjustExcluPartLocation
      End With
   End If
   
   Log "      Items Counted With Negative Inventory " & str$(iNeg)
   Log "      Buy Items With No Lead Time  " & iMissingLead
   Log "      Make Items With No Flow Time  " & iMissingFlow
   Log "End Beginning Inventory " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   Log " "
   prg1.Value = 100
   
   Exit Function

DiaErr1:

   
   sProcName = "BeginningBalance"
   prg1.Visible = False
   z1(1) = ""
   cmdCan.Enabled = True
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function AdjustExcluPartLocation() As Boolean

      sSql = "UPDATE mrplTable SET MRP_PARTQTYRQD = MRP_PARTQTYRQD - sumLotQty FROM " _
              & " (SELECT SUM(lotremainingqty) as sumLotQty, lotpartref FROM lohdtable, LoLcTable " _
                    & " WHERE lotlocation = LOTEXLOCATION and LOTLOCATION <> ''" _
                    & " AND lotremainingqty > 0 GROUP BY lotpartref ) as f " _
            & " WHERE LotPartRef = mrp_Partref "
      
      clsADOCon.ExecuteSql sSql 'rdExecDirect
       
      sSql = "UPDATE MrppTable SET MRP_PARTQOH = MRP_PARTQOH - sumLotQty FROM " _
             & " (SELECT SUM(lotremainingqty) as sumLotQty, lotpartref from lohdtable, LoLcTable " _
                     & "WHERE lotlocation = LOTEXLOCATION and LOTLOCATION <> ''" _
                  & "AND lotremainingqty > 0 GROUP BY lotpartref ) as f " _
            & " WHERE LotPartRef = mrp_Partref"

      clsADOCon.ExecuteSql sSql 'rdExecDirect

End Function


Private Function SafetyStock(Optional PreEmpt As Boolean = True) As Boolean

   Dim currentDate As Date
   Dim curSortDate As Date
   
   Log "Start Safety Stock" & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   
   
   ' Changed on 7/14/2009
   currentDate = GetServerDate()
   curSortDate = "1/1/2008"
   sSql = "insert into MrplTable" & vbCrLf _
      & "(MRP_CATAGORY,MRP_PARTREF,MRP_PARTNUM," & vbCrLf _
      & "MRP_PARTDESC,MRP_PARTLEVEL,MRP_PARTUOM," & vbCrLf _
      & "MRP_PARTPRODCODE,MRP_PARTCLASS,MRP_PARTQTYRQD,MRP_PARTDATERQD," & vbCrLf _
      & "MRP_PARTSORTDATE,MRP_TYPE,MRP_COMMENT,MRP_USER,MRP_PAMAKEBUY)" & vbCrLf _
      & "SELECT 'Safety Stock',PARTREF,PARTNUM," & vbCrLf _
      & "PADESC,PALEVEL,PAUNITS," & vbCrLf _
      & "PAPRODCODE,PACLASS,-PASAFETY,DateAdd(day,(CASE WHEN pamakeBuy = 'B' THEN PALeadTime" & vbCrLf _
      & " WHEN PAMAKEBUY = 'M' THEN PAFLOWTime ELSE 0 END), GETDATE())," & vbCrLf _
      & "'" & curSortDate & "'," & MRPTYPE_SafetyStock & ",'','" & sInitials & "',PAMAKEBUY" & vbCrLf _
      & "FROM PartTable" & vbCrLf _
      & "WHERE PALEVEL < 5 AND PATOOL = 0 AND PASAFETY <> 0 AND PAINACTIVE = 0 AND PAOBSOLETE = 0 " & vbCrLf
   

   Debug.Print sSql
   
   clsADOCon.ExecuteSql sSql
   
   Log "End Safety Stock (" & clsADOCon.RowsAffected & " items) " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   Log "  "
   
   SafetyStock = True
End Function

Private Function PurchaseOrderItems(Optional PreEmpt As Boolean = True) As Boolean
   'Get Purchase Orders items
   Dim n As Integer
   
   'counters
   Dim A As Byte
   Dim b As Currency
   Dim C As Currency
   
   Dim cQuantity As Currency
   Dim sVendor As String
   Dim sBuyer As String
   Dim vDate As Variant
   
   MouseCursor 11
   If (PreEmpt) Then DoEvents
   b = 100 / lPoItemRows
   prg1.Value = 5
   z1(1).Visible = True
'   z1(1).Caption = "Step 3: Purchase Orders"
'   z1(1).Refresh
   
   If (PreEmpt) Then DoEvents
   
   Log "Start Purchase Order Items " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   
   sProcName = "purchaseorde"
   ' 12/16/2009
   ' if all the parts are rejected on a purchase item don't include
   ' IATYPE_PoRejectItem - 39
   sSql = "SELECT PINUMBER,PIITEM,PIREV,PITYPE,PIPART," _
          & "PIPDATE,PIPQTY,PIRUNPART,PIRUNNO,PIRUNOPNO," _
          & "PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," _
          & "PACLASS,PAPRODCODE FROM PoitTable,PartTable WHERE " _
          & "(PIPART=PARTREF AND PIAQTY=0 AND PITYPE<>16 AND " _
          & "PITYPE<>17 AND PITYPE<>39 AND PALEVEL<5) AND PIPDATE<='" & vEndDate & "' " _
          & "ORDER BY PIPDATE"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoFetch, ES_FORWARD)
   If bSqlRows Then
      sSql = "SELECT MRP_ROW,MRP_CATAGORY,MRP_PARTREF,MRP_PARTNUM," _
             & "MRP_PARTDESC,MRP_PARTLEVEL,MRP_PARTUOM," _
             & "MRP_PARTPRODCODE,MRP_PARTCLASS,MRP_PARTQTYRQD," _
             & "MRP_PARTDATERQD,MRP_PARTSORTDATE,MRP_USER,MRP_POVENDOR," _
             & "MRP_POBUYER,MRP_PONUM,MRP_POTEXT,MRP_POITEM,MRP_POREV," _
             & "MRP_TYPE,MRP_COMMENT " _
             & "FROM MrplTable"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoInsert, ES_DYNAMIC)
      With rdoFetch
         C = 5
         Do Until .EOF
            A = A + 1
            C = C + b
            If C >= 95 Then C = 95
            If A = 5 Then
               prg1.Value = Int(C)
               A = 0
            End If
            If Not IsNull(!PIPDATE) Then
               vDate = Format(!PIPDATE, "mm/dd/yy 06:00")
            Else
               vDate = Format(ES_SYSDATE, "mm/dd/yy 06:00")
               
               
               Log "  * Scheduled Date Problem PO Number " & Format$(!PINUMBER, "000000") & "-" _
                        & !PIITEM & !PIREV & " " & sVendor
            End If
            cQuantity = !PIPQTY
            cAdjQuantity = cQuantity
            If cQuantity > 0 Then
                ' Depending on the PITYPE = 5 - we will get the buyer name
               sVendor = GetPoVendor(!PINUMBER, sBuyer)
               lCOUNTER = lCOUNTER + 1
               sReplaceStr = ReplaceString("" & Trim(!PADESC))
               rdoInsert.AddNew
               'RdoInsert!MRP_ROW = lCOUNTER
               rdoInsert!MRP_TYPE = MRPTYPE_PoItem
               rdoInsert!MRP_CATAGORY = "Purchase Order "
               rdoInsert!MRP_PARTREF = "" & Trim(!PartRef)
               rdoInsert!MRP_PARTNUM = "" & Trim(!PartNum)
               rdoInsert!MRP_PARTDESC = sReplaceStr
               rdoInsert!MRP_PARTLEVEL = !PALEVEL
               rdoInsert!MRP_PARTUOM = "" & Trim(!PAUNITS)
               rdoInsert!MRP_PARTPRODCODE = "" & Trim(!PAPRODCODE)
               rdoInsert!MRP_PARTCLASS = "" & Trim(!PACLASS)
               rdoInsert!MRP_PARTQTYRQD = cQuantity
               rdoInsert!MRP_PARTDATERQD = vDate
               rdoInsert!MRP_PARTSORTDATE = vDate
               rdoInsert!MRP_USER = sInitials
               rdoInsert!MRP_POVENDOR = sVendor
               rdoInsert!MRP_POBUYER = sBuyer
               rdoInsert!MRP_PONUM = !PINUMBER
               rdoInsert!MRP_POITEM = !PIITEM
               rdoInsert!MRP_POREV = !PIREV
               rdoInsert!MRP_POTEXT = Format$(!PINUMBER, "000000")
               rdoInsert!MRP_COMMENT = "PO Number " & Format$(!PINUMBER, "000000") & "-" & !PIITEM & !PIREV & " " & sVendor
               rdoInsert.Update
               If (PreEmpt) Then DoEvents
               cAdjQuantity = cQuantity
               UpdateQoh cAdjQuantity, Compress(!PartNum), 2, True
               lLogRows = lLogRows + 1
            End If
            If (PreEmpt) Then DoEvents
            .MoveNext
         Loop
         ClearResultSet rdoFetch
         A = 0
         Do While Not rdoInsert.EOF
            A = A + 1
            If A > 5 Then Exit Do
         Loop
         If lLogRows > 500 Then WriteTransactionLog
         rdoFetch.Close
      End With
   End If
   
   
   Log "End Purchase Order Items   " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   
   
   Log " "
   prg1.Value = 100
   
End Function

'Collect Sales Order Items

Private Function SalesOrderItems(Optional PreEmpt As Boolean = True) As Boolean
   Dim cMrpQoh As Currency
   Dim cQuantity As Currency
   Dim sCust As String
   Dim sSoType As String
   Dim vDate As Variant
   
   'counters
   Dim A As Byte
   Dim b As Currency
   Dim C As Currency
   
   b = 100 / lSoItemRows
   MouseCursor 11
   If (PreEmpt) Then DoEvents
   z1(1).Visible = True
'   z1(1).Caption = "Step 4: Sales Orders"
'   z1(1).Refresh
   prg1.Value = 5
   
   
   Log "Start Sales Order Items   " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   sProcName = "salesorderitems"
   sSql = "SELECT SONUMBER,SOTYPE,ITSO,ITNUMBER," & vbCrLf _
          & "ITREV,ITPART,ITQTY,ITSCHED,ITACTUAL," & vbCrLf _
          & "PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
          & "PACLASS,PAPRODCODE FROM SohdTable,SoitTable,PartTable" & vbCrLf _
          & "WHERE SONUMBER=ITSO AND ITPART=PARTREF AND ITCANCELED=0 " & vbCrLf _
          & "AND ITACTUAL IS NULL AND ITSCHED<='" & vEndDate & "' " & vbCrLf
   If sSalesOrders <> "ALL" Then sSql = sSql & sSalesOrders
   sSql = sSql & " ORDER BY ITSCHED"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoFetch, ES_FORWARD)
   If bSqlRows Then
      sSql = "SELECT MRP_ROW,MRP_CATAGORY,MRP_PARTREF,MRP_PARTNUM," _
             & "MRP_PARTDESC,MRP_PARTLEVEL,MRP_PARTUOM," _
             & "MRP_PARTPRODCODE,MRP_PARTCLASS,MRP_PARTQTYRQD," _
             & "MRP_PARTDATERQD,MRP_PARTSORTDATE,MRP_SOCUST," _
             & "MRP_SONUM,MRP_SOITEM,MRP_SOREV,MRP_SOTEXT," _
             & "MRP_COMMENT,MRP_TYPE,MRP_USER " _
             & "FROM MrplTable"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoInsert, ES_DYNAMIC)
      With rdoFetch
         C = 5
         Do Until .EOF
            A = A + 1
            C = C + b
            If C >= 95 Then C = 95
            If A = 5 Then
               prg1.Value = Int(C)
               A = 0
            End If
            '11/16/04 trap bogus conversions
            If Not IsNull(!ITSCHED) Then
               'vDate = Format(!ITSCHED, "mm/dd/yy 18:00")
               vDate = Format(!ITSCHED, "mm/dd/yy")
            Else
               Log "   ** Scheduled Date Problem SO Number; " & sSoType & Format$(!ITSO, "; 0; ") & " - " & !ITNUMBER & !itrev & "; " & sCust
               'vDate = Format(ES_SYSDATE, "mm/dd/yy 18:00")
               vDate = Format(ES_SYSDATE, "mm/dd/yy")
            End If
            sCust = GetSoCustomer(!ITSO, sSoType)
            sReplaceStr = ReplaceString("" & Trim(!PADESC))
            lCOUNTER = lCOUNTER + 1
            'cMRPQoh = GetMRPQoh("" & Trim(!PartRef))
            cQuantity = !ITQTY
            cQuantity = cQuantity - (2 * cQuantity)
            ' cAdjQuantity = cQuantity
            If cQuantity < 0 Then
               rdoInsert.AddNew
               'RdoInsert!MRP_ROW = lCOUNTER
               rdoInsert!MRP_TYPE = MRPTYPE_SoItem
               rdoInsert!MRP_CATAGORY = "Sales Order "
               rdoInsert!MRP_PARTREF = "" & Trim(!PartRef)
               rdoInsert!MRP_PARTNUM = "" & Trim(!PartNum)
               rdoInsert!MRP_PARTDESC = sReplaceStr
               rdoInsert!MRP_PARTLEVEL = !PALEVEL
               rdoInsert!MRP_PARTUOM = "" & Trim(!PAUNITS)
               rdoInsert!MRP_PARTPRODCODE = "" & Trim(!PAPRODCODE)
               rdoInsert!MRP_PARTCLASS = "" & Trim(!PACLASS)
               rdoInsert!MRP_PARTQTYRQD = cQuantity
               rdoInsert!MRP_PARTDATERQD = vDate
               rdoInsert!MRP_PARTSORTDATE = vDate
               rdoInsert!MRP_USER = sInitials
               rdoInsert!MRP_SOCUST = sCust
               rdoInsert!MRP_SONUM = !ITSO
               rdoInsert!MRP_SOITEM = !ITNUMBER
               rdoInsert!MRP_SOREV = !itrev
               rdoInsert!MRP_SOTEXT = sSoType & Format$(!ITSO, SO_NUM_FORMAT)
               rdoInsert!MRP_COMMENT = "SO Number " & sSoType & Format$(!ITSO, SO_NUM_FORMAT) & "-" & !ITNUMBER & !itrev & " " & sCust
               rdoInsert.Update
               If (PreEmpt) Then DoEvents
            End If
            ' UpdateQoh cadjQuantity, Compress(!PARTNUM), 4, True
            .MoveNext
            lLogRows = lLogRows + 1
            If (PreEmpt) Then DoEvents
         Loop
         ClearResultSet rdoFetch
         A = 0
         Do While Not rdoInsert.EOF
            A = A + 1
            If A > 5 Then Exit Do
         Loop
         If lLogRows > 500 Then WriteTransactionLog
         rdoFetch.Close
      End With
   End If
   
   Log "End Sales Order Items     " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   Log "  "
   
   prg1.Value = 100
   
   
End Function

Private Function MOPicks(Optional PreEmpt As Boolean = True) As Boolean
   'The first level of Picks from the Pick Table

   Dim bByte As Byte
   'Dim n As Long
   
   Dim cMrpQoh As Currency
   Dim cQuantity As Currency
   Dim sMoPart As String
   Dim sMoDesc As String
   Dim sDesc As String
   Dim vDate As Variant
   Dim cPQty As Currency
   Dim cAQty As Currency
   
   'counters
   Dim A As Byte
   Dim b As Currency
   Dim C As Currency
   
   MouseCursor 11
   If (PreEmpt) Then DoEvents
'   z1(1).Caption = "Step 5: MO Picks                "
'   z1(1).Refresh
   prg1.Value = 5
   
   b = 100 / lPickRows
   
   
   Log "MO Picks                 " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   sProcName = "mopicks"
   ''sSql = "Qry_GetMRPMOPicks '" & vEndDate & "'"
   ''    sSql = "SELECT PKPARTREF,PKMOPART,PKMORUN,PKTYPE,PKPDATE," _
   ''        & "PKPQTY,PKAQTY,PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," _
   ''        & "PACLASS,PAPRODCODE FROM MopkTable,PartTable WHERE " _
   ''        & "(PKPARTREF=PARTREF AND PKAQTY=0 AND PKTYPE<>12) " _
   ''        & "AND PKPDATE<='" & vEndDate & "' "
   
   'revised 9/12/08 to exclude unpicked items on completed MO's and closed MO's
   '(MO's can be completed with unpicked items)
   sSql = "SELECT PKPARTREF,PKMOPART,PKMORUN,PKTYPE,PKPDATE," & vbCrLf _
       & "PKPQTY,PKAQTY,PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS," & vbCrLf _
       & "PACLASS,PAPRODCODE" & vbCrLf _
       & "FROM MopkTable " & vbCrLf _
       & "JOIN PartTable ON PKPARTREF = PARTREF" & vbCrLf _
       & "JOIN RunsTable ON PKMOPART = RUNREF AND PKMORUN = RUNNO" & vbCrLf _
       & "WHERE PKAQTY = 0 AND PKTYPE <> 12 " & vbCrLf _
       & "AND PKPDATE <= '" & vEndDate & "'" & vbCrLf _
       & "AND RUNSTATUS NOT IN ('CO','CL')"
   Set rdoFetch = AdoCon2.GetRecordSet(sSql, ES_FORWARD)
   If Not rdoFetch.BOF And Not rdoFetch.EOF Then
      bSqlRows = 1
   Else
      bSqlRows = 0
   End If
   If bSqlRows Then
      sSql = "SELECT MRP_ROW,MRP_CATAGORY,MRP_PARTREF,MRP_PARTNUM," _
             & "MRP_PARTDESC,MRP_PARTLEVEL,MRP_PARTUOM,MRP_PARTUSEDON," _
             & "MRP_PARTPRODCODE,MRP_PARTCLASS,MRP_PARTQTYRQD," _
             & "MRP_PARTDATERQD,MRP_PARTSORTDATE,MRP_MOREF,MRP_MONUM," _
             & "MRP_MODESC,MRP_MORUNNO,MRP_COMMENT,MRP_TYPE,MRP_USER " _
             & "FROM MrplTable "
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoInsert, ES_DYNAMIC)
      With rdoFetch
         C = 5
         Do Until .EOF
            A = A + 1
            C = C + b
            If C >= 95 Then C = 95
            If A = 5 Then
               prg1.Value = Int(C)
               A = 0
            End If
            If !PKPQTY > 0 Then
               sMoPart = "" & Trim(!PKMOPART)
               bByte = GetMoInfo(sMoPart, !PKMORUN, sMoDesc)
               sMoDesc = ReplaceString("" & Trim(!PADESC))
               sDesc = ReplaceString("" & Trim(!PADESC))
               If Not IsNull(!PKPDATE) Then
                  vDate = Format(!PKPDATE, "mm/dd/yy")
               Else
                  Log "  ** Scheduled Pick Date Problem MO Pick " & Trim(sMoPart) _
                           & " Run " & Format$(!PKMORUN)
                  vDate = Format(ES_SYSDATE, "mm/dd/yy")
               End If
               lCOUNTER = lCOUNTER + 1
               
               cPQty = !PKPQTY
               cAQty = !PKAQTY
               
               cQuantity = !PKPQTY
               
               If (cAQty > cPQty) Then
                  ' Just add the difference and will be P-qty to A-qty
                  cQuantity = cQuantity + (cAQty - cPQty)
               End If
               
               cQuantity = cQuantity - (2 * cQuantity)
               If cQuantity < 0 Then
                  rdoInsert.AddNew
                  rdoInsert!MRP_TYPE = MRPTYPE_MoPick
                  rdoInsert!MRP_CATAGORY = "MO Material Pick"
                  rdoInsert!MRP_PARTREF = "" & Trim(!PartRef)
                  rdoInsert!MRP_PARTNUM = "" & Trim(!PartNum)
                  rdoInsert!MRP_PARTDESC = "" & Trim(!PADESC)
                  rdoInsert!MRP_PARTLEVEL = !PALEVEL
                  rdoInsert!MRP_PARTUOM = "" & Trim(!PAUNITS)
                  rdoInsert!MRP_PARTUSEDON = sMoPart
                  rdoInsert!MRP_PARTPRODCODE = "" & Trim(!PAPRODCODE)
                  rdoInsert!MRP_PARTCLASS = "" & Trim(!PACLASS)
                  rdoInsert!MRP_PARTQTYRQD = cQuantity
                  rdoInsert!MRP_PARTDATERQD = Format(vDate, "mm/dd/yy")
                  rdoInsert!MRP_PARTSORTDATE = Format(vDate, "mm/dd/yy")
                  rdoInsert!MRP_USER = sInitials
                  rdoInsert!MRP_MOREF = "" & Compress(sMoPart)
                  rdoInsert!MRP_MONUM = "" & sMoPart
                  rdoInsert!MRP_MODESC = "" & sMoDesc
                  rdoInsert!MRP_MORUNNO = !PKMORUN
                  rdoInsert!MRP_COMMENT = "Pick For MO Number " & Trim(sMoPart) & " Run " & Format$(!PKMORUN)
                  rdoInsert.Update
                  If (PreEmpt) Then DoEvents
               End If
               lLogRows = lLogRows + 1
            End If
            .MoveNext
            'n = n + 1
         Loop
         ClearResultSet rdoFetch
         A = 0
         Do While Not rdoInsert.EOF
            A = A + 1
            If A > 5 Then Exit Do
         Loop
         If lLogRows > 500 Then WriteTransactionLog
         rdoFetch.Close
      End With
   End If
   Log "End MO Picks             " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   Log "  "
   prg1.Value = 100
   
End Function

Private Function GetPoVendor(lPoNum As Long, SPoBuyer As String) As String
    Dim RdoVnd As ADODB.Recordset
   
    sSql = "Qry_GetMRPVendor " & lPoNum
    
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd, ES_FORWARD)
    If bSqlRows Then
        With RdoVnd
           GetPoVendor = "" & Trim(!VENICKNAME)
           SPoBuyer = "" & Trim(!POBUYER)
           ClearResultSet RdoVnd
        End With
    Else
        GetPoVendor = ""
        SPoBuyer = ""
    End If
   
   Set RdoVnd = Nothing
   
End Function

Private Function GetPoActionBuyer(sPARef As String) As String
    Dim RdoBuyer As ADODB.Recordset
    Dim sPartBuyer, sProdCodeBuyer As String
    sSql = "Qry_GetMRPSysBuyer '" & sPARef & "'"
    
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoBuyer, ES_FORWARD)
    If bSqlRows Then
        With RdoBuyer
           sPartBuyer = "" & Trim(!PABUYER)
           sProdCodeBuyer = "" & Trim(!PRODCODEBUYER)
           
           If sPartBuyer <> "" Then
            GetPoActionBuyer = sPartBuyer
           ElseIf sProdCodeBuyer <> "" Then
            GetPoActionBuyer = sProdCodeBuyer
           Else
            GetPoActionBuyer = ""
            
           End If
           
           RdoBuyer.Close
        End With
    Else
        GetPoActionBuyer = ""
    End If
   
   Set RdoBuyer = Nothing
   
End Function


Private Sub optTyp_GotFocus(Index As Integer)
   Dim b As Byte
   lblAlp(Index).BorderStyle = vbFixedSingle
   
End Sub


Private Sub optTyp_LostFocus(Index As Integer)
   Dim b As Byte
   For b = 0 To 24
      lblAlp(b).BorderStyle = 0
   Next
   lblAlp(b).BorderStyle = 0
   
End Sub


Private Sub cmbcutoff_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub cmbcutoff_LostFocus()
   cmbCutoff = CheckDateEx(cmbCutoff)
   
End Sub

Private Function GetSoCustomer(lSonum As Long, sText As String) As String
   Dim RdoVnd As ADODB.Recordset
   sSql = "SELECT SONUMBER,SOTYPE,SOCUST,CUREF,CUNICKNAME FROM " _
          & "SohdTable,CustTable WHERE (SOCUST=CUREF AND " _
          & "SONUMBER=" & lSonum & ") "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd, ES_FORWARD)
   If bSqlRows Then
      With RdoVnd
         GetSoCustomer = "" & Trim(!CUNICKNAME)
         sText = "" & Trim(!SOTYPE)
         ClearResultSet RdoVnd
      End With
   Else
      GetSoCustomer = ""
      sText = ""
   End If
   Set RdoVnd = Nothing
   
End Function

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "MRPSo", sOptions)
   If Len(Trim(sOptions)) > 0 Then
      For iList = 0 To 24
         optTyp(iList).Value = Val(Mid$(sOptions, iList + 1, 1))
      Next
      optTyp(iList).Value = Val(Mid$(sOptions, iList + 1, 1))
   End If
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim iList As Integer
   Dim l1 As Long
   Dim l2 As Long
   On Error Resume Next
   l1 = DateValue(cmbCutoff)
   l2 = DateValue(Format(ES_SYSDATE, "mm/dd/yy"))
   iList = l1 - l2
   For iList = 0 To 24
      sOptions = sOptions & Trim$(str$(Val(optTyp(iList).Value)))
   Next
   sOptions = sOptions & Trim$(str$(Val(optTyp(iList).Value)))
   SaveSetting "Esi2000", "EsiProd", "MRPSo", sOptions
   
End Sub

Private Function GetMoInfo(sMoNumber As String, lRunNum As Long, sDescription As String) As Byte
   Dim RdoRun As ADODB.Recordset
   sSql = "SELECT RUNREF,RUNNO,PARTREF,PARTNUM,PADESC " _
          & "FROM RunsTable,PartTable WHERE (RUNREF=PARTREF " _
          & "AND RUNREF='" & sMoNumber & "' AND RUNNO=" & lRunNum & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         sMoNumber = "" & Trim(!PartNum)
         sDescription = "" & Trim(!PADESC)
         ClearResultSet RdoRun
      End With
   Else
      sMoNumber = ""
      sDescription = ""
   End If
   Set RdoRun = Nothing
   
End Function


Private Sub GetMrpDefaults()
   Dim RdoShp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetLeadTimes"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         iDefPurLead = Format(!DEFLEADTIME, "##0")
         iDefManFlow = Format(!DEFFLOWTIME, "##0")
         ClearResultSet RdoShp
      End With
   Else
      iDefPurLead = 30
      iDefManFlow = 30
   End If
   Set RdoShp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getmrpdef"
   On Error GoTo 0
   iDefPurLead = 30
   iDefManFlow = 30
   
End Sub

Private Sub OpenSecondCon()
   On Error Resume Next
   Dim strConStr2 As String
   Dim ErrNum As Long
   Dim ErrDesc As String
   
   sSaAdmin = Trim(GetSysLogon(True))
   sSaPassword = Trim(GetSysLogon(False))
   Set AdoCon2 = New ClassFusionADO
   
   strConStr2 = "Driver={SQL Server};Provider='sqloledb';UID=" & sSaAdmin & ";PWD=" & _
            sSaPassword & ";SERVER=" & sserver & ";DATABASE=" & sDataBase & ";"
   
   If AdoCon2.OpenConnection(strConStr2, ErrNum, ErrDesc) = False Then
     MsgBox "An error occured while trying to do second connect to the specified database:" & Chr(13) & Chr(13) & _
            "Error Number = " & CStr(ErrNum) & Chr(13) & _
            "Error Description = " & ErrDesc, vbOKOnly + vbExclamation, "  DB Connection Error"
     Set AdoCon2 = Nothing
   End If
   
End Sub


Private Function MarkShortages(Optional PreEmpt As Boolean = True) As Boolean
   Dim RdoShort As ADODB.Recordset
   Dim A As Integer
   Dim n As Integer
   Dim b As Currency
   Dim C As Currency
   
   MouseCursor 11
   If (PreEmpt) Then DoEvents
   'On Error Resume Next
   WriteTransactionLog
   
   On Err GoTo DiaErr1
   sProcName = "MarkShortages"
   prg1.Value = 5
   z1(1).Caption = "Marking Exceptions"
   If (PreEmpt) Then DoEvents
   
   Log "Start Exceptions           " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   
   sSql = "DELETE from MrplTable WHERE MRP_PARTDATERQD > '" & vEndDate & "'"
   clsADOCon.ExecuteSql sSql
   
   sSql = "SELECT MRP_PARTREF,ISNULL(SUM(MRP_PARTQTYRQD),0) FROM MrplTable " _
          & "GROUP BY MRP_PARTREF"
   Set RdoShort = AdoCon2.GetRecordSet(sSql, ES_FORWARD)
   If Not RdoShort.BOF And Not RdoShort.EOF Then
      bSqlRows = 1
   Else
      bSqlRows = 0
   End If
   If bSqlRows Then
      z1(1).Caption = "Marking Exceptions"
      prg1.Value = 10
      z1(1).Refresh
      If lPartCount = 0 Then lPartCount = 500
      b = 100 / lPartCount
      With RdoShort
         'On Error Resume Next
         Do Until .EOF
            A = A + 1
            C = C + b
            If A = 5 Then
               A = 0
'Debug.Print Int(C)
               If Int(C) <= 100 Then
                  prg1.Value = Int(C)
               End If
            End If
'Debug.Print "* " & .fields(1) & " " & Trim(!MRP_PARTREF)
'If Trim(!MRP_PARTREF) = "RH1BULAC120V" Then
'   Debug.Print "found it"
'End If
            If .Fields(1) < 0 Then
               sSql = "UPDATE MrplTable SET MRP_SHORTAGE=1 " _
                      & "WHERE MRP_PARTREF='" & "" & Trim(!MRP_PARTREF) & "' "
               AdoCon2.ExecuteSql sSql
            End If
            .MoveNext
'If ILOGROWS >= 2500 Then
'   Debug.Print lLogRows
'End If
            lLogRows = lLogRows + 1
            If (PreEmpt) Then DoEvents
         Loop
         ClearResultSet RdoShort
         If lLogRows > 500 Then WriteTransactionLog
      End With
   End If
   A = 0
   n = 0
   
   sSql = "SELECT DISTINCT MRP_PARTREF,MRP_POBUYER FROM MrplTable " _
          & "WHERE MRP_POBUYER<>''"
   Set RdoShort = AdoCon2.GetRecordSet(sSql, ES_FORWARD)
   If Not RdoShort.BOF And Not RdoShort.EOF Then
      bSqlRows = 1
   Else
      bSqlRows = 0
   End If
   If bSqlRows Then
      MouseCursor 11
      A = 5
      z1(1).Visible = True
      z1(1).Refresh
      b = 100 / lPartCount
      With RdoShort
         Err = 0
         Do Until .EOF
            A = A + 1
            C = C + b
            If A = 5 Then
               A = 0
               prg1.Value = Int(C)
            End If
            sSql = "UPDATE MrplTable SET MRP_POBUYER='" _
                   & Trim(!MRP_POBUYER) & "' " _
                   & "WHERE MRP_PARTREF='" & Trim(!MRP_PARTREF) & "' "
            AdoCon2.ExecuteSql sSql
            .MoveNext
            lLogRows = lLogRows + 1
         Loop
         ClearResultSet RdoShort
         If lLogRows > 500 Then WriteTransactionLog
      End With
   End If
   
   
   Log "End Exceptions             " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   
   
   Log " "
   prg1.Value = 100
   Set RdoShort = Nothing
   Exit Function
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   MouseCursor 0
   cmdCan.Enabled = True
   
End Function

Private Function CheckSalesOrders() As Byte
   Dim b As Byte
   Dim iList As Integer
   sProcName = "checksalesor"
'   For iList = 0 To 25
'      If optTyp(iList).Value = vbChecked Then
'         b = b + 1
'         If b = 1 Then
'            sSalesOrders = sSalesOrders & "SOTYPE='" & Chr$(iList + 65) & "'"
'         Else
'            sSalesOrders = sSalesOrders & " OR SOTYPE='" & Chr$(iList + 65) & "'"
'         End If
'      End If
'   Next
   
   For iList = 0 To 25
      If optTyp(iList).Value = vbChecked Then
         b = b + 1
         If b = 1 Then
            sSalesOrders = "and SOTYPE in ('" & Chr$(iList + 65) & "'"
         Else
            sSalesOrders = sSalesOrders & ",'" & Chr$(iList + 65) & "'"
         End If
      End If
   Next
   sSalesOrders = sSalesOrders & ")" & vbCrLf
   
   
   If b = 26 Then
      sSalesOrders = "ALL"
   'Else
   '   sSalesOrders = " AND (" & sSalesOrders & ")"
   End If
   CheckSalesOrders = b
   
End Function

Private Sub WriteTransactionLog()
   'clsAdoCon.CommitTrans
   lLogRows = 0
   'clsAdoCon.begintrans
   
End Sub


Private Function GetMRPQoh(MrpPart As String) As Currency
   Sleep 100
   sSql = "Qry_GetMRPPartQoh '" & MrpPart & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoQoh, ES_FORWARD)
   If bSqlRows Then
      With RdoQoh
         If Not IsNull(!MRP_PARTQOH) Then
            GetMRPQoh = !MRP_PARTQOH
         Else
            GetMRPQoh = 0
         End If
         ClearResultSet RdoQoh
      End With
   Else
      GetMRPQoh = 0
   End If
   If GetMRPQoh < 0 Then GetMRPQoh = 0
   Exit Function
   
DiaErr1:
   Err = 0
   GetMRPQoh = 0
   
End Function


'3/21/05 MrpType added for testing only
'        Tested using GARMAN's data
'5/2/06 Added stored procedures (See Queries)

Private Sub UpdateQoh(Quantity As Currency, PartNumber As String, mrpType As Byte, _
                      AddParts As Boolean)
   Dim sQuery As String
   Sleep 50
   If Not AddParts Then
      sQuery = "Qry_UpdateMRppTableNEG " & str$(Abs(Quantity)) & ",'" & Trim(PartNumber) & "'"
   Else
      sQuery = "Qry_UpdateMRppTablePOS " & str$(Abs(Quantity)) & ",'" & Trim(PartNumber) & "'"
   End If
   clsADOCon.ExecuteSql sQuery
   
   'so the above code was irrelevant.  the quantity was always set = 0 below
   'sQuery = "Qry_UpdateMRppTableZERO '" & Trim(PartNumber) & "'"
   'clsADOCon.ExecuteSQL sQuery
   'Sleep 50
   
End Sub

'Public Sub CheckLogTable()
'   On Error Resume Next
'   sSql = "SELECT LOG_NUMBER FROM EsReportMrpLog"
'   clsAdoCon.ExecuteSQL sSql
'   If Err > 0 Then
'      Err.Clear
'      sSql = "CREATE TABLE EsReportMrpLog (" _
'             & "LOG_NUMBER SMALLINT NULL DEFAULT(0)," _
'             & "LOG_TEXT VARCHAR(60) NULL DEFAULT(''))"
'      clsAdoCon.ExecuteSQL sSql
'      If Err = 0 Then
'         sSql = "CREATE UNIQUE CLUSTERED INDEX LogIndex ON " _
'                & "EsReportMrpLog(LOG_NUMBER) WITH FILLFACTOR = 80"
'         clsAdoCon.ExecuteSQL sSql
'      End If
'   End If
'
'End Sub

'Private Sub WriteReportLog()
'   Dim iList As Integer
'   On Error Resume Next
'   sSql = "TRUNCATE TABLE EsReportMrpLog"
'   clsAdoCon.ExecuteSQL sSql
'
'   For iList = 1 To iLogNumber
'      sSql = "INSERT INTO EsReportMrpLog (LOG_NUMBER,LOG_TEXT) " _
'             & "VALUES(" & sLogNote(iList, 0) & ",'" & sLogNote(iList, 1) & "')"
'      clsAdoCon.ExecuteSQL sSql
'   Next
'
'End Sub

Private Function zGetSOItemRows() As Long
   Dim RdoCount As ADODB.Recordset
   Err.Clear
   sSql = "SELECT COUNT(ITSO) AS RowsCounted FROM SoitTable WHERE " _
          & "(ITCANCELED=0 AND ITACTUAL IS NULL) AND ITSCHED<='" & vEndDate & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCount, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoCount!RowsCounted) Then
         zGetSOItemRows = RdoCount!RowsCounted
      Else
         zGetSOItemRows = 300
      End If
      ClearResultSet RdoCount
   Else
      zGetSOItemRows = 300
   End If
   If Err > 0 Then zGetSOItemRows = 300
   Err.Clear
   Set RdoCount = Nothing
   
End Function

Private Function zGetMORows() As Long
   Dim RdoCount As ADODB.Recordset
   sSql = "SELECT COUNT(RUNREF) AS RowsCounted FROM RunsTable WHERE " _
          & "(RUNSTATUS NOT LIKE 'C%' AND RUNSCHED<='" & vEndDate & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCount, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoCount!RowsCounted) Then
         zGetMORows = RdoCount!RowsCounted
      Else
         zGetMORows = 100
      End If
      ClearResultSet RdoCount
   Else
      zGetMORows = 100
   End If
   If Err > 0 Then zGetMORows = 100
   Err.Clear
   Set RdoCount = Nothing
   
End Function

Private Function zGetPOItemRows() As Long
   Dim RdoCount As ADODB.Recordset
   sSql = "SELECT COUNT(PINUMBER) AS RowsCounted FROM PoitTable," _
          & "PartTable WHERE (PIPART=PARTREF AND PIAQTY=0 AND PITYPE<>16 AND " _
          & "PITYPE<>17 AND PALEVEL<5) AND PIPDATE<='" & vEndDate & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCount, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoCount!RowsCounted) Then
         zGetPOItemRows = RdoCount!RowsCounted
      Else
         zGetPOItemRows = 100
      End If
      ClearResultSet RdoCount
   Else
      zGetPOItemRows = 100
   End If
   If Err > 0 Then zGetPOItemRows = 100
   Err.Clear
   Set RdoCount = Nothing
   
   
End Function

Private Function zGetMOPickRows()
   Dim RdoCount As ADODB.Recordset
   sSql = "SELECT COUNT(PKPARTREF) AS RowsCounted FROM MopkTable WHERE " _
          & "(PKAQTY=0 AND PKTYPE<>12) AND PKPDATE<='" & vEndDate & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCount, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoCount!RowsCounted) Then
         zGetMOPickRows = RdoCount!RowsCounted
      Else
         zGetMOPickRows = 100
      End If
      ClearResultSet RdoCount
   Else
      zGetMOPickRows = 100
   End If
   If Err > 0 Then zGetMOPickRows = 100
   Err.Clear
   Set RdoCount = Nothing
   
End Function

Private Function zGetSchedMORows()
   Dim RdoCount As ADODB.Recordset
   sSql = "SELECT COUNT(RUNREF) AS RowsCounted FROM RunsTable WHERE " _
          & "(RUNSTATUS='SC' OR RUNSTATUS='RL') AND RUNSCHED<='" & vEndDate & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCount, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoCount!RowsCounted) Then
         zGetSchedMORows = RdoCount!RowsCounted
      Else
         zGetSchedMORows = 100
      End If
      ClearResultSet RdoCount
   Else
      zGetSchedMORows = 100
   End If
   If Err > 0 Then zGetSchedMORows = 100
   Err.Clear
   Set RdoCount = Nothing
   
End Function

Private Function zGetMRPPartCount() As Long
   Dim RdoCount As ADODB.Recordset
   Err.Clear
   sSql = "SELECT COUNT(PARTREF) AS RowsCounted FROM PartTable" & vbCrLf _
      & "WHERE PALEVEL<5 AND PATOOL=0" & vbCrLf
   If Me.chkZeroBalances.Value = vbUnchecked Then
      sSql = sSql & "AND PAQOH <> 0"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCount, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoCount!RowsCounted) Then
         zGetMRPPartCount = RdoCount!RowsCounted
      Else
         zGetMRPPartCount = 500
      End If
      ClearResultSet RdoCount
   Else
      zGetMRPPartCount = 500
   End If
   'If Err > 0 Then zGetMRPPartCount = 500
   'Err.Clear
   Set RdoCount = Nothing
   
End Function

Public Sub Log(sText As String)

   iLogNumber = iLogNumber + 1
   sText = Replace(sText, "'", "''")
   If Len(sText) > 60 Then sText = Left$(sText, 60)
   'Debug.Print iLogNumber & ": " & sText
   sSql = "INSERT INTO EsReportMrpLog (LOG_NUMBER,LOG_TEXT) " _
          & "VALUES( " & iLogNumber & ", '" & sText & "')"
   clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub TruncateReportLog()
   sSql = "TRUNCATE TABLE EsReportMrpLog"
   clsADOCon.ExecuteSql sSql
   iLogNumber = 0
End Sub

Private Sub txtTolerance_LostFocus()
   If Not IsNumeric(txtTolerance.Text) Then
      MsgBox "Tolerance must be a number"
      txtTolerance.SetFocus
   End If
End Sub

Private Sub SaveSettings()
   SaveSetting "Esi2000", "EsiProd", "MRPCutoffDate", cmbCutoff.Text
   SaveSetting "Esi2000", "EsiProd", "MRPTolerance", txtTolerance.Text
End Sub

Private Sub GetSettings()
   cmbCutoff.Text = GetSetting("Esi2000", "EsiProd", "MRPCutoffDate", Format(DateAdd("d", 365, Now), "mm/dd/yy"))
   txtTolerance.Text = GetSetting("Esi2000", "EsiProd", "MRPTolerance", "0")
End Sub

Public Sub AutoGenerateMrp(strInDate As String)
   Dim b As Boolean
   Dim bResponse As Byte
   Dim sMsg As String
   Dim vDate As Variant
   
   GetMrpDefaults
   GetLastMrp
   
   TruncateReportLog
   b = CheckSalesOrders()
   If b = 0 Then
      Exit Sub
   End If
         
   lLogRows = 0
   lCOUNTER = 0
   Sleep 100
   'On Error Resume Next
   'On Err GoTo DiaErr1
   lSoItemRows = zGetSOItemRows()
   If lSoItemRows = 0 Then lSoItemRows = 300
         
   lMoRows = zGetMORows()
   If lMoRows = 0 Then lMoRows = 100
   
   lPoItemRows = zGetPOItemRows()
   If lPoItemRows = 0 Then lPoItemRows = 100
   
   lPickRows = zGetMOPickRows()
   If lPickRows = 0 Then lPickRows = 100
   
   lBillRows = zGetSchedMORows()
   If lBillRows = 0 Then lBillRows = 100
   'b = OpenLog()
   'iLogNumber = 1
   
   ErrorCount = 0
   
   Log "** Flags Scheduled Date Problems (Null or no date)"
   Log " "
   Log "Start MRP              " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   Log " "
   Log "Default Manufacturing Flow Time Is " & iDefManFlow & " Days"
   Log "Default Purchasing Lead Time Is " & iDefPurLead & " Days"
   If iDefManFlow = 0 Or iDefPurLead = 0 Then
      Log "** One Or Both Defaults Are Zero (Company Setting) "
   Else
      Log " "
   End If
   
   OpenSecondCon
   
   sSql = "TRUNCATE TABLE MrplTable"
   clsADOCon.ExecuteSql sSql
   sSql = "DBCC CHECKIDENT ('MrplTable', RESEED, 1)"
   clsADOCon.ExecuteSql sSql
   
   sSql = "TRUNCATE TABLE MrppTable"
   clsADOCon.ExecuteSql sSql
   'Sleep 1000
   
   'cmdVew.Enabled = False
   'cmdGen.Enabled = False
   prg1.Visible = True
   prg1.Value = 10
   z1(1).Visible = True
   z1(1).Caption = "Setting Parameters."
   z1(1).Refresh
   Sleep 100
   
   ' *** Start to fill. In order of requirements ***
   z1(1).Caption = "Step 1: Beginning Balance"
   b = BeginningBalance(False)
   z1(1).Refresh
   'get safety stock
   z1(1).Caption = "Step 2: Safety Stock"
   b = SafetyStock(False)
   z1(1).Refresh
   'get inventory ins for unreceived PO items
   z1(1).Caption = "Step 3: Purchase Orders"
   b = PurchaseOrderItems(False)
   z1(1).Refresh
   'Get inventory ins for incomplete MO items
   z1(1).Caption = "Step 4: Manufacturing Orders"
   b = ManufacturingOrders(False)
   z1(1).Refresh
   
   'Get inventory outs for unshipped SO items
   z1(1).Caption = "Step 5: Sales Orders"
   b = SalesOrderItems(False)
   z1(1).Refresh
   
   'Get MO unpicked pick list items
   z1(1).Caption = "Step 6: MO Picks"
   b = MOPicks(False)
   z1(1).Refresh
   
   'Explode SC and RL MOs without picklists one level
   z1(1).Caption = "Step 7: Explode SC and RL MO's one level"
   ExplodeMoPicks (False)
   z1(1).Refresh
   
   'keep adding action items until no more are required
   Dim Pass As Integer
   Pass = 1
   Do While True
      z1(1).Caption = "Step " & 6 + 2 * Pass & ": Add action items"
      If Not AddActionItems(Pass, False) Then             '8, 10, 12 ...
         Exit Do
      End If
      z1(1).Refresh
      
      z1(1).Caption = "Step " & 7 + 2 * Pass & ": Explode action items"
      If Not ExplodeActionItems(Pass, False) Then   '9, 11, 12 ...
         Exit Do
      End If
      z1(1).Refresh
      
      Pass = Pass + 1
      If Pass > 7 Then
         Exit Do
      End If
      
   Loop

   'close connection stuff no longer needed
   If Not rdoInsert Is Nothing Then
      'rdoInsert.Close
      Set rdoInsert = Nothing
   End If
   If Not rdoFetch Is Nothing Then
      If Not rdoFetch.ActiveConnection Is Nothing Then
         If (rdoFetch.State = adStateOpen) Then rdoFetch.Close
      End If
      Set rdoFetch = Nothing
   End If
   'RdoCon.CommitTrans
   
   'Set shortage indicator where end quantity < 0.
   'Also set buyers for parts.
   b = MarkShortages(False)
   
   'Clean empty rows
   sSql = "DELETE FROM MrplTable WHERE MRP_PARTREF=''"
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE Table2 SET Table2.MRP_PAMAKEBUY=MrplTable.MRP_PAMAKEBUY " _
          & "FROM MrplTable,MrplTable as Table2 WHERE " _
          & "MrplTable.MRP_PARTREF = Table2.MRP_PARTREF AND " _
          & "MrplTable.MRP_TYPE = " & MRPTYPE_BeginningBalance
   clsADOCon.ExecuteSql sSql
   
   MouseCursor 0
   If ErrorCount > 0 Then
      Log "MRP Completed with " & ErrorCount & " errors.  " & Format(ES_SYSDATE, "mm/dd/yy hh:mm AM/PM") & ".  See log."
   Else
      Log "MRP Completed " & Format(ES_SYSDATE, "mm/dd/yy hh:mm AM/PM") & ".  See log."
   End If
   
   Log "End MRP                   " & Format$(Now, "mm/dd/yy hh:mm AM/PM")
   
   'RdoCon.CommitTrans
   sSql = "UPDATE MrpdTable SET MRP_CREATEDATE='" & Format(ES_SYSDATE, "mm/dd/yy hh:mm AM/PM") & "'," _
          & "MRP_THROUGHDATE='" & cmbCutoff & "',MRP_CREATEDBY='" & sInitials & "' " _
          & "WHERE MRP_ROW=1"
   clsADOCon.ExecuteSql sSql
   
   'Set the Sort Date to the min date recorded (Garretts 1950)
   sSql = "SELECT MIN(MRP_PARTDATERQD) AS StartDate FROM MrplTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoQoh, ES_FORWARD)
   If bSqlRows Then
      With RdoQoh
         If Not IsNull(!startDate) Then
            vDate = Format(!startDate - 2, "mm/dd/yy")
            sSql = "UPDATE MrplTable SET MRP_PARTSORTDATE='" _
                   & vDate & "' WHERE MRP_TYPE=" & MRPTYPE_BeginningBalance      '1
            clsADOCon.ExecuteSql sSql
         End If
         .Cancel
      End With
   End If
   
   prg1.Value = 0
   prg1.Visible = False
   z1(1) = ""
   z1(1).Refresh
 
End Sub
