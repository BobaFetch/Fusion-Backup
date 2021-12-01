VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form diaGLe07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Numbers For Parts"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter a New Part or Select From List (30 chars)"
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6600
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   16
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLe07a.frx":0000
      PictureDn       =   "diaGLe07a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   5640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5700
      FormDesignWidth =   7560
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   4455
      Left            =   60
      TabIndex        =   22
      Top             =   1200
      Width           =   7400
      _ExtentX        =   13044
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      TabCaption(0)   =   "General Accounts"
      TabPicture(0)   =   "diaGLe07a.frx":028C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "z1(38)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTcg"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTrv"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblInv"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDis"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblRev"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "z1(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "z1(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "z1(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "z1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "z1(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "z1(9)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblRej"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtTcg"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtTrv"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtInv"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtDis"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtRev"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtRej"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Inventory/Expense"
      TabPicture(1)   =   "diaGLe07a.frx":02A8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "z1(45)"
      Tab(1).Control(1)=   "z1(44)"
      Tab(1).Control(2)=   "z1(43)"
      Tab(1).Control(3)=   "z1(42)"
      Tab(1).Control(4)=   "lblWla"
      Tab(1).Control(5)=   "lblWma"
      Tab(1).Control(6)=   "lblWoh"
      Tab(1).Control(7)=   "lblWex"
      Tab(1).Control(8)=   "z1(39)"
      Tab(1).Control(9)=   "z1(46)"
      Tab(1).Control(10)=   "lblGla"
      Tab(1).Control(11)=   "z1(47)"
      Tab(1).Control(12)=   "z1(48)"
      Tab(1).Control(13)=   "z1(49)"
      Tab(1).Control(14)=   "z1(50)"
      Tab(1).Control(15)=   "lblGma"
      Tab(1).Control(16)=   "lblGoh"
      Tab(1).Control(17)=   "lblGex"
      Tab(1).Control(18)=   "txtWla"
      Tab(1).Control(19)=   "txtWma"
      Tab(1).Control(20)=   "txtWoh"
      Tab(1).Control(21)=   "txtWex"
      Tab(1).Control(22)=   "txtGla"
      Tab(1).Control(23)=   "txtGma"
      Tab(1).Control(24)=   "txtGoh"
      Tab(1).Control(25)=   "txtGex"
      Tab(1).ControlCount=   26
      Begin VB.ComboBox txtRej 
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ComboBox txtRev 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox txtDis 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox txtInv 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox txtTrv 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox txtTcg 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox txtGex 
         Height          =   315
         Left            =   -72840
         TabIndex        =   14
         Tag             =   "3"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.ComboBox txtGoh 
         Height          =   315
         Left            =   -72840
         TabIndex        =   13
         Tag             =   "3"
         Top             =   3480
         Width           =   1935
      End
      Begin VB.ComboBox txtGma 
         Height          =   315
         Left            =   -72840
         TabIndex        =   12
         Tag             =   "3"
         Top             =   3120
         Width           =   1935
      End
      Begin VB.ComboBox txtGla 
         Height          =   315
         Left            =   -72840
         TabIndex        =   11
         Tag             =   "3"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ComboBox txtWex 
         Height          =   315
         Left            =   -72840
         TabIndex        =   10
         Tag             =   "3"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox txtWoh 
         Height          =   315
         Left            =   -72840
         TabIndex        =   9
         Tag             =   "3"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox txtWma 
         Height          =   315
         Left            =   -72840
         TabIndex        =   8
         Tag             =   "3"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox txtWla 
         Height          =   315
         Left            =   -72840
         TabIndex        =   7
         Tag             =   "3"
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblRej 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   53
         Top             =   2760
         Width           =   2400
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rejection Tag Acct"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   52
         Top             =   2760
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Revenue Account"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   51
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Account"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   50
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory/Expense Acct"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   49
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Revenue Acct"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   48
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer CGS Acct"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   47
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label lblRev 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   46
         Top             =   960
         Width           =   2400
      End
      Begin VB.Label lblDis 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   45
         Top             =   1320
         Width           =   2400
      End
      Begin VB.Label lblInv 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   44
         Top             =   1680
         Width           =   2400
      End
      Begin VB.Label lblTrv 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   43
         Top             =   2040
         Width           =   2400
      End
      Begin VB.Label lblTcg 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   42
         Top             =   2400
         Width           =   2400
      End
      Begin VB.Label lblGex 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   41
         Top             =   3840
         Width           =   2400
      End
      Begin VB.Label lblGoh 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   40
         Top             =   3480
         Width           =   2400
      End
      Begin VB.Label lblGma 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   39
         Top             =   3120
         Width           =   2400
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   255
         Index           =   50
         Left            =   -74640
         TabIndex        =   38
         Top             =   2760
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Account"
         Height          =   255
         Index           =   49
         Left            =   -74640
         TabIndex        =   37
         Top             =   3120
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead Account"
         Height          =   255
         Index           =   48
         Left            =   -74640
         TabIndex        =   36
         Top             =   3480
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Account"
         Height          =   255
         Index           =   47
         Left            =   -74640
         TabIndex        =   35
         Top             =   3840
         Width           =   1800
      End
      Begin VB.Label lblGla 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   34
         Top             =   2760
         Width           =   2400
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Of Goods Sold Accounts:"
         Height          =   255
         Index           =   46
         Left            =   -74760
         TabIndex        =   33
         Top             =   2400
         Width           =   3075
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory/Expense Accounts:"
         Height          =   255
         Index           =   39
         Left            =   -74760
         TabIndex        =   32
         Top             =   600
         Width           =   3075
      End
      Begin VB.Label lblWex 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   31
         Top             =   2040
         Width           =   2400
      End
      Begin VB.Label lblWoh 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   30
         Top             =   1680
         Width           =   2400
      End
      Begin VB.Label lblWma 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   29
         Top             =   1320
         Width           =   2400
      End
      Begin VB.Label lblWla 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -70800
         TabIndex        =   28
         Top             =   960
         Width           =   2400
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Account"
         Height          =   255
         Index           =   42
         Left            =   -74640
         TabIndex        =   27
         Top             =   2040
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead Account"
         Height          =   255
         Index           =   43
         Left            =   -74640
         TabIndex        =   26
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Account"
         Height          =   255
         Index           =   44
         Left            =   -74640
         TabIndex        =   25
         Top             =   1320
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   255
         Index           =   45
         Left            =   -74640
         TabIndex        =   24
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "General Accounts:"
         Height          =   255
         Index           =   38
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   3075
      End
   End
   Begin Threed.SSFrame fra2 
      Height          =   30
      Left            =   60
      TabIndex        =   54
      Top             =   1080
      Width           =   7400
      _Version        =   65536
      _ExtentX        =   13053
      _ExtentY        =   53
      _StockProps     =   14
      ForeColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCde 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6000
      TabIndex        =   56
      Top             =   720
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   55
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6000
      TabIndex        =   19
      Top             =   360
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   18
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   17
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "diaGLe07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

Dim RdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim rdoAct As ADODB.Recordset

Dim bCanceled As Byte
Dim bGoodPart As Byte
Dim bNoAccts As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   bGoodPart = GetPart()
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCanceled Then Exit Sub
   bGoodPart = GetPart()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillAccounts
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   Tab1.Tab = 0
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL," _
          & "PAPRODCODE,PAACCTNO,PAREVACCT,PADISACCT," _
          & "PATFRREVACCT,PATFRCGSACCT,PAREJACCT," _
          & "PAINVLABACCT,PAINVMATACCT,PAINVOHDACCT,PAINVEXPACCT," _
          & "PACGSLABACCT,PACGSMATACCT,PACGSOHDACCT,PACGSEXPACCT " _
          & "FROM PartTable WHERE PARTREF= ? "
   Set RdoQry = New ADODB.Command
   RdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   
   RdoQry.parameters.Append AdoParameter1
   
   bOnLoad = True
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   Set AdoParameter1 = Nothing
   Set rdoAct = Nothing
   Set RdoQry = Nothing
   Set diaGLe07a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Public Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim i As Integer
   On Error GoTo DiaErr1
   sSql = "Qry_FillSortedParts"
   cmbPrt.Clear
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         sPassedPart = Trim(cUR.CurrentPart)
         If Len(sPassedPart) > 0 Then
            cmbPrt = cUR.CurrentPart
         Else
            cmbPrt = "" & Trim(!PARTNUM)
         End If
         Do Until .EOF
            'cmbPrt.AddItem "" & Trim(!PARTNUM)
            AddComboStr cmbPrt.hwnd, "" & Trim(!PARTNUM)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoCmb = Nothing
   bGoodPart = GetPart()
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetPart() As Byte
   On Error GoTo DiaErr1
   RdoQry.parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(rdoAct, RdoQry, ES_KEYSET, True)
   If bSqlRows Then
      With rdoAct
         cmbPrt = "" & Trim(!PARTNUM)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = Format(!PALEVEL, "#")
         lblCde = "" & Trim(!PAPRODCODE)
         txtRev = "" & Trim(!PAREVACCT)
         GetAccount txtRev, "txtRev"
         
         txtDis = "" & Trim(!PADISACCT)
         GetAccount txtDis, "txtDis"
         
         txtInv = "" & Trim(!PAACCTNO)
         GetAccount txtInv, "txtInv"
         
         txtTrv = "" & Trim(!PATFRREVACCT)
         GetAccount txtTrv, "txtTrv"
         
         txtTcg = "" & Trim(!PATFRCGSACCT)
         GetAccount txtTcg, "txtTcg"
         
         txtRej = "" & Trim(!PAREJACCT)
         GetAccount txtRej, "txtRej"
         
         '10/7/99
         txtWla = "" & Trim(!PAINVLABACCT)
         GetAccount txtWla, "txtWla"
         
         txtWex = "" & Trim(!PAINVEXPACCT)
         GetAccount txtWex, "txtWex"
         
         txtWma = "" & Trim(!PAINVMATACCT)
         GetAccount txtWma, "txtWma"
         
         txtWoh = "" & Trim(!PAINVOHDACCT)
         GetAccount txtWoh, "txtWoh"
         
         txtGla = "" & Trim(!PACGSLABACCT)
         GetAccount txtGla, "txtGla"
         
         txtGex = "" & Trim(!PACGSEXPACCT)
         GetAccount txtGex, "txtGex"
         
         txtGma = "" & Trim(!PACGSMATACCT)
         GetAccount txtGma, "txtGma"
         
         txtGoh = "" & Trim(!PACGSOHDACCT)
         GetAccount txtGoh, "txtGoh"
         
         .Cancel
      End With
      CheckPartType
      Tab1.enabled = True
      GetPart = True
   Else
      Tab1.enabled = False
      lblDsc = "*** No Current Part ***"
      lblLvl = "0"
      txtRev = ""
      lblRev = ""
      txtDis = ""
      lblDis = ""
      txtInv = ""
      lblInv = ""
      txtTrv = ""
      lblTrv = ""
      txtTcg = ""
      lblTcg = ""
      GetPart = False
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Sub FillAccounts()
   Dim i As Integer
   Dim RdoGlm As ADODB.Recordset
   On Error GoTo DiaErr1
   i = -1
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      With RdoGlm
         Do Until .EOF
            i = i + 1
            AddComboStr txtRev.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtDis.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtInv.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtTrv.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtTcg.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtRej.hwnd, "" & Trim(!GLACCTNO)
            
            AddComboStr txtWla.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWma.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWex.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWoh.hwnd, "" & Trim(!GLACCTNO)
            
            AddComboStr txtGla.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGma.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGex.hwnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGoh.hwnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      bNoAccts = True
      CloseBoxes
   End If
   Set RdoGlm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillaccou"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   bNoAccts = True
   CloseBoxes
   
End Sub

Public Sub GetAccount(sAccount As String, sBox As String)
   Dim RdoGlm As ADODB.Recordset
   On Error GoTo DiaErr1
   If bNoAccts Then Exit Sub
   sAccount = Compress(sAccount)
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR FROM GlacTable " _
          & "WHERE GLACCTREF='" & sAccount & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm)
   If bSqlRows Then
      With RdoGlm
         Select Case sBox
            Case "txtRev"
               txtRev = "" & Trim(!GLACCTNO)
               lblRev = "" & Trim(!GLDESCR)
            Case "txtDis"
               txtDis = "" & Trim(!GLACCTNO)
               lblDis = "" & Trim(!GLDESCR)
            Case "txtInv"
               txtInv = "" & Trim(!GLACCTNO)
               lblInv = "" & Trim(!GLDESCR)
            Case "txtTrv"
               txtTrv = "" & Trim(!GLACCTNO)
               lblTrv = "" & Trim(!GLDESCR)
            Case "txtTcg"
               txtTcg = "" & Trim(!GLACCTNO)
               lblTcg = "" & Trim(!GLDESCR)
            Case "txtRej"
               txtRej = "" & Trim(!GLACCTNO)
               lblRej = "" & Trim(!GLDESCR)
               'Inv/Exp
            Case "txtWla"
               txtWla = "" & Trim(!GLACCTNO)
               lblWla = "" & Trim(!GLDESCR)
            Case "txtWex"
               txtWex = "" & Trim(!GLACCTNO)
               lblWex = "" & Trim(!GLDESCR)
            Case "txtWma"
               txtWma = "" & Trim(!GLACCTNO)
               lblWma = "" & Trim(!GLDESCR)
            Case "txtWoh"
               txtWoh = "" & Trim(!GLACCTNO)
               lblWoh = "" & Trim(!GLDESCR)
               'CGS
            Case "txtGla"
               txtGla = "" & Trim(!GLACCTNO)
               lblGla = "" & Trim(!GLDESCR)
            Case "txtGex"
               txtGex = "" & Trim(!GLACCTNO)
               lblGex = "" & Trim(!GLDESCR)
            Case "txtGma"
               txtGma = "" & Trim(!GLACCTNO)
               lblGma = "" & Trim(!GLDESCR)
            Case "txtGoh"
               txtGoh = "" & Trim(!GLACCTNO)
               lblGoh = "" & Trim(!GLDESCR)
         End Select
         .Cancel
      End With
   Else
      Select Case sBox
         Case "txtRev"
            txtRev = ""
            lblRev = ""
         Case "txtDis"
            txtDis = ""
            lblDis = ""
         Case "txtInv"
            txtInv = ""
            lblInv = ""
         Case "txtTrv"
            txtTrv = ""
            lblTrv = ""
         Case "txtTcg"
            txtTcg = ""
            lblTcg = ""
      End Select
   End If
   Set RdoGlm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getaccoun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub CloseBoxes()
   txtRev.enabled = False
   txtDis.enabled = False
   txtInv.enabled = False
   txtTrv.enabled = False
   txtTcg.enabled = False
   
   txtWla.enabled = False
   txtWma.enabled = False
   txtWex.enabled = False
   txtWoh.enabled = False
   
   txtGla.enabled = False
   txtGma.enabled = False
   txtGex.enabled = False
   txtGoh.enabled = False
   MsgBox "There Are No Accounts Established.", _
      vbInformation, Caption
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 10) = "*** No Cur" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub




Private Sub tab1_GotFocus()
   On Error Resume Next
   If Tab1.Tab = 0 Then txtRev.SetFocus Else txtWla.SetFocus
   
End Sub


Private Sub txtDis_Click()
   GetAccount txtDis, "txtDis"
   
End Sub

Private Sub txtDis_LostFocus()
   txtDis = CheckLen(txtDis, 12)
   GetAccount txtDis, "txtDis"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PADISACCT = Compress(txtDis)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtGex_Click()
   GetAccount txtGex, "txtGex"
   
End Sub


Private Sub txtGex_LostFocus()
   txtGex = CheckLen(txtGex, 12)
   GetAccount txtGex, "txtGex"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PACGSEXPACCT = Compress(txtGex)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtGla_Click()
   GetAccount txtGla, "txtGla"
   
End Sub


Private Sub txtGla_LostFocus()
   txtGla = CheckLen(txtGla, 12)
   GetAccount txtGla, "txtGla"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PACGSLABACCT = Compress(txtGla)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtGma_Click()
   GetAccount txtGma, "txtGma"
   
End Sub

Private Sub txtGma_LostFocus()
   txtGma = CheckLen(txtGma, 12)
   GetAccount txtGma, "txtGma"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PACGSMATACCT = Compress(txtGma)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtGoh_Click()
   GetAccount txtGoh, "txtGoh"
   
End Sub


Private Sub txtGoh_LostFocus()
   txtGoh = CheckLen(txtGoh, 12)
   GetAccount txtGoh, "txtGoh"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PACGSOHDACCT = Compress(txtGoh)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtInv_Click()
   GetAccount txtInv, "txtInv"
   
End Sub

Private Sub txtInv_LostFocus()
   txtInv = CheckLen(txtInv, 12)
   GetAccount txtInv, "txtInv"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PAACCTNO = Compress(txtInv)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtRej_Click()
   GetAccount txtRej, "txtRej"
   
End Sub


Private Sub txtRej_LostFocus()
   txtRej = CheckLen(txtRej, 12)
   GetAccount txtRej, "txtRej"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PAREJACCT = Compress(txtRej)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtRev_Click()
   GetAccount txtRev, "txtRev"
   
End Sub

Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 12)
   GetAccount txtRev, "txtRev"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PAREVACCT = Compress(txtRev)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtTcg_Click()
   GetAccount txtTcg, "txtTcg"
   
End Sub

Private Sub txtTcg_LostFocus()
   txtTcg = CheckLen(txtTcg, 12)
   GetAccount txtTcg, "txtTcg"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PATFRCGSACCT = Compress(txtTcg)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtTrv_Click()
   GetAccount txtTrv, "txtTrv"
   
End Sub

Private Sub txtTrv_LostFocus()
   txtTrv = CheckLen(txtTrv, 12)
   GetAccount txtTrv, "txtTrv"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PATFRREVACCT = Compress(txtTrv)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtWex_Click()
   GetAccount txtWex, "txtWex"
   
End Sub


Private Sub txtWex_LostFocus()
   txtWex = CheckLen(txtWex, 12)
   GetAccount txtWex, "txtWex"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PAINVEXPACCT = Compress(txtWex)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtWla_Click()
   GetAccount txtWla, "txtWla"
   
End Sub


Private Sub txtWla_LostFocus()
   txtWla = CheckLen(txtWla, 12)
   GetAccount txtWla, "txtWla"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PAINVLABACCT = Compress(txtWla)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtWma_Click()
   GetAccount txtWma, "txtWma"
   
End Sub


Private Sub txtWma_LostFocus()
   txtWma = CheckLen(txtWma, 12)
   GetAccount txtWma, "txtWma"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PAINVMATACCT = Compress(txtWma)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub


Private Sub txtWoh_Click()
   GetAccount txtWoh, "txtWoh"
   
End Sub


Private Sub txtWoh_LostFocus()
   txtWoh = CheckLen(txtWoh, 12)
   GetAccount txtWoh, "txtWoh"
   If bGoodPart Then
      On Error Resume Next
      rdoAct!PAINVOHDACCT = Compress(txtWoh)
      rdoAct.Update
      If Err > 0 Then ValidateEdit Me
   End If
   
End Sub



Public Sub CheckPartType()
   On Error Resume Next
   Select Case Val(lblLvl)
      Case 4
         txtWma.enabled = True
         txtGma.enabled = True
         
         txtWla.enabled = False
         txtWex.enabled = False
         txtWoh.enabled = False
         
         txtGla.enabled = False
         txtGex.enabled = False
         txtGoh.enabled = False
      Case 5, 6, 7
         txtWex.enabled = True
         txtGex.enabled = True
         
         txtWla.enabled = False
         txtWma.enabled = False
         txtWoh.enabled = False
         
         txtGla.enabled = False
         txtGma.enabled = False
         txtGoh.enabled = False
      Case Else
         txtWla.enabled = True
         txtWma.enabled = True
         txtWoh.enabled = True
         txtWex.enabled = True
         
         txtGla.enabled = True
         txtGma.enabled = True
         txtGoh.enabled = True
         txtGex.enabled = True
   End Select
   
End Sub
