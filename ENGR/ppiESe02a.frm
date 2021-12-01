VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ppiESe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Full Estimate"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3502
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   5280
      Top             =   120
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ppiESe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   70
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtDue 
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Tag             =   "4"
      ToolTipText     =   "Due Date"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "N&ew"
      Height          =   285
      Left            =   5880
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   "Add A New Estimate"
      Top             =   480
      Width           =   1000
   End
   Begin VB.CheckBox optFrom 
      Caption         =   "from"
      Height          =   255
      Left            =   720
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDis 
      Caption         =   "&Discounts"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   64
      TabStop         =   0   'False
      ToolTipText     =   "Volume Discounts"
      Top             =   1140
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CommandButton cmdOth 
      Caption         =   "O&thers"
      Height          =   300
      Left            =   2880
      TabIndex        =   13
      ToolTipText     =   "Other Estimating Features"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCom 
      Caption         =   "Unc&omplete"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   "Mark This Bid As Complete"
      Top             =   800
      Width           =   1000
   End
   Begin VB.CheckBox optSle 
      Caption         =   "View Sales"
      Height          =   255
      Left            =   0
      TabIndex        =   46
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View Parts"
      Height          =   255
      Left            =   1800
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optPart 
      Caption         =   "Parts"
      Height          =   195
      Left            =   1200
      TabIndex        =   44
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Unit Costs"
      Top             =   1080
      Width           =   915
   End
   Begin VB.TextBox txtPrc 
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Locked TextBoxes"
      Top             =   4800
      Width           =   915
   End
   Begin VB.TextBox txtPrf 
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "Locked TextBoxes"
      Top             =   4440
      Width           =   915
   End
   Begin VB.TextBox txtGna 
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   2880
      TabIndex        =   15
      Tag             =   "1"
      ToolTipText     =   "Locked TextBoxes"
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox txtScr 
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   2880
      TabIndex        =   14
      Tag             =   "1"
      ToolTipText     =   "Locked TextBoxes"
      Top             =   3720
      Width           =   912
   End
   Begin VB.CommandButton cmdSrv 
      Caption         =   "S&ervices"
      Height          =   300
      Left            =   2880
      MaskColor       =   &H8000000F&
      TabIndex        =   12
      ToolTipText     =   "Services And Subcontracting"
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdMat 
      Caption         =   "&Material"
      Enabled         =   0   'False
      Height          =   300
      Left            =   2880
      MaskColor       =   &H8000000F&
      TabIndex        =   17
      ToolTipText     =   "Material-Parts List"
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdLbr 
      Caption         =   "&Labor"
      Height          =   300
      Left            =   2880
      MaskColor       =   &H8000000F&
      TabIndex        =   11
      ToolTipText     =   "Routing Costs"
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   120
      TabIndex        =   29
      Top             =   1440
      Width           =   6735
   End
   Begin VB.CheckBox optCom 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "ppiESe02a.frx":07AE
      Height          =   315
      Left            =   4800
      Picture         =   "ppiESe02a.frx":0AF0
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number Find A Part Number (F4 At Part Number)"
      Top             =   1560
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Enter A Part Or Click The ""Look Up"""
      Top             =   1560
      Width           =   3075
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Tag             =   "3"
      ToolTipText     =   "Select A Customer"
      Top             =   1920
      Width           =   1555
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "ppiESe02a.frx":0E32
      Height          =   315
      Left            =   6120
      Picture         =   "ppiESe02a.frx":130C
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Show Sales Orders For Parts"
      Top             =   1560
      Width           =   350
   End
   Begin VB.ComboBox cmbRfq 
      Height          =   315
      Left            =   4080
      TabIndex        =   10
      Tag             =   "3"
      ToolTipText     =   "Select RFQ From List"
      Top             =   1920
      Width           =   2040
   End
   Begin VB.CommandButton cmdPrt 
      Height          =   315
      Left            =   6480
      Picture         =   "ppiESe02a.frx":17E6
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "New Part Numbers Find A Part Number (F2 At Part Number)"
      Top             =   1560
      Width           =   350
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Class A-Z"
      Top             =   1080
      Width           =   495
   End
   Begin VB.ComboBox cmbBid 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Select Or Enter A Bid Number (Contains Full Bids)"
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "4"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   1000
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   5400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5685
      FormDesignWidth =   6990
   End
   Begin VB.Label txtRte 
      Caption         =   "Rate"
      Height          =   252
      Left            =   5520
      TabIndex        =   73
      Top             =   6000
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Label Estimator 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1200
      TabIndex        =   72
      Top             =   5280
      Width           =   2292
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimator"
      Height          =   252
      Index           =   11
      Left            =   240
      TabIndex        =   71
      Top             =   5280
      Width           =   1332
   End
   Begin VB.Label lblRouting 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   69
      Top             =   6000
      Visible         =   0   'False
      Width           =   3492
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1680
      TabIndex        =   68
      Top             =   2280
      Width           =   3852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bid Due"
      Height          =   255
      Index           =   34
      Left            =   3360
      TabIndex        =   67
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblEstimator 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2880
      TabIndex        =   63
      ToolTipText     =   "Labor Hours"
      Top             =   6000
      Visible         =   0   'False
      Width           =   2412
   End
   Begin VB.Label lblUnitServices 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6000
      TabIndex        =   62
      ToolTipText     =   "Total Unit Services"
      Top             =   3000
      Width           =   852
   End
   Begin VB.Label lblRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   5160
      TabIndex        =   61
      ToolTipText     =   "Labor Hours"
      Top             =   5640
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label lblFohRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   4200
      TabIndex        =   60
      ToolTipText     =   "Labor Hours"
      Top             =   5640
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label lblOther 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6000
      TabIndex        =   59
      ToolTipText     =   "Total Other Costs"
      Top             =   3360
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Costs"
      Height          =   252
      Index           =   10
      Left            =   240
      TabIndex        =   58
      Top             =   3360
      Width           =   1932
   End
   Begin VB.Label PALEVEL 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2880
      TabIndex        =   57
      ToolTipText     =   "Material"
      Top             =   5640
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label lblBidTot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6000
      TabIndex        =   55
      ToolTipText     =   "Total Of This Estimate"
      Top             =   4800
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Total"
      Height          =   252
      Index           =   9
      Left            =   4200
      TabIndex        =   54
      Top             =   4800
      Width           =   1332
   End
   Begin VB.Label lblPrf 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6000
      TabIndex        =   53
      ToolTipText     =   "Unit Profit"
      Top             =   4440
      Width           =   852
   End
   Begin VB.Label lblGNA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6000
      TabIndex        =   52
      ToolTipText     =   "Unit General And Administration"
      Top             =   4080
      Width           =   852
   End
   Begin VB.Label lblScrap 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6000
      TabIndex        =   51
      ToolTipText     =   "Unit Scrap"
      Top             =   3720
      Width           =   852
   End
   Begin VB.Label lblFoh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   5040
      TabIndex        =   50
      ToolTipText     =   "Labor Overhead"
      Top             =   2640
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label lblTotMat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   288
      Left            =   6000
      TabIndex        =   49
      ToolTipText     =   "Total Unit Material"
      Top             =   6600
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label lblBurden 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   288
      Left            =   5040
      TabIndex        =   48
      ToolTipText     =   "Material Burden"
      Top             =   6600
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label LblLabor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6000
      TabIndex        =   47
      ToolTipText     =   "Total Unit Labor"
      Top             =   2640
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   43
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   42
      Top             =   4800
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Profit"
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   41
      Top             =   4440
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   252
      Index           =   15
      Left            =   3960
      TabIndex        =   40
      Top             =   4440
      Width           =   252
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   252
      Index           =   27
      Left            =   3960
      TabIndex        =   39
      Top             =   4080
      Width           =   252
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "G&A (General And Admin)"
      Height          =   252
      Index           =   28
      Left            =   240
      TabIndex        =   38
      Top             =   4080
      UseMnemonic     =   0   'False
      Width           =   2412
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reduction (Scrap)"
      Height          =   252
      Index           =   17
      Left            =   240
      TabIndex        =   37
      Top             =   3720
      Width           =   2172
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   252
      Index           =   18
      Left            =   3960
      TabIndex        =   36
      Top             =   3720
      Width           =   252
   End
   Begin VB.Label lblTotServices 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   4080
      TabIndex        =   35
      ToolTipText     =   "Total Services"
      Top             =   3000
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label lblMaterial 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   34
      ToolTipText     =   "Material"
      Top             =   6600
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label lblHours 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   4080
      TabIndex        =   33
      ToolTipText     =   "Unit Labor Hours"
      Top             =   2640
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Services"
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   32
      Top             =   3000
      Width           =   1932
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Material (Parts List)"
      Enabled         =   0   'False
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   31
      Top             =   6600
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Labor (Routing)"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   30
      Top             =   2640
      Width           =   1932
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RFQ"
      Height          =   255
      Index           =   25
      Left            =   3360
      TabIndex        =   28
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   26
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Number"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bid Date"
      Height          =   255
      Index           =   26
      Left            =   240
      TabIndex        =   23
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Estimate"
      Height          =   255
      Index           =   31
      Left            =   240
      TabIndex        =   22
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblNxt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   21
      Top             =   360
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete"
      Height          =   255
      Index           =   35
      Left            =   3360
      TabIndex        =   20
      ToolTipText     =   "Bid Is Marked Complete"
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "ppiESe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/8/06 Added rates to new bid
'4/11/06 Added Select a Routing
'4/14/06 Fixed Bid sort
'4/17/06 See CreateIndex (ppi only)
'5/23/06 Changed default date to ES_SYSDATE
'8/30/06 Added sCurrEstimator and label Estimator
'8/31/06 Fixed RFQ (wrong keyset)
Option Explicit

Dim bBomLevel As Byte
Dim bCanceled As Byte
Dim bFromNew As Byte
Dim bOnLoad As Byte
Dim bOpenKey As Byte
Dim bGoodCust As Byte
Dim bGoodPart As Byte
Dim bGoodRout As Byte

Dim iBids As Integer

'General defaults
Dim cFoh As Currency
Dim cGna As Currency
Dim cProfit As Currency
Dim cOldQty As Currency
Dim cScrap As Currency

'Bid stuff
Dim cBFoh As Currency
Dim cMBurden As Currency
Dim cBGna As Currency
Dim cBprofit As Currency
Dim cRate As Currency
Dim cBScrap As Currency

Dim cBGnaRate As Currency
Dim cBProfitRate As Currency
Dim cBScrapRate As Currency

Dim sBomRev As String
Dim sEstimator As String
Dim sBEstimator As String
Dim sRouting As String

Dim lBids(300) As Long

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
'ppi only

Private Sub CreateIndex()
   Dim RdoIndex As ADODB.Recordset
   Dim bByte As Byte
   
   On Error Resume Next
   sSql = "sp_helpindex 'EsrtTable'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoIndex, ES_FORWARD)
   If bSqlRows Then
      With RdoIndex
         Do Until .EOF
            If "" & Trim(!index_name) = "BidNotes" Then bByte = 1
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   If bByte = 0 Then
      sSql = "CREATE INDEX BidNotes ON " _
             & "EsrtTable(BIDFORMULANOTES) WITH FILLFACTOR = 80"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   End If
   
End Sub

Private Sub GetTheLabor()
   Dim RdoLabor As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT SUM(BIDRTELABOR) AS BidLabor FROM EsrtTable WHERE BIDRTEREF=" _
          & Val(cmbBid)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLabor, ES_FORWARD)
   If bSqlRows Then
      With RdoLabor
         ppiESe02a.lblLabor = Format(!Bidlabor, "####0.00")
         .Cancel
      End With
   Else
      ppiESe02a.lblLabor = "0.00"
   End If
   
End Sub

'10/9/06

Private Sub GetRates()
   On Error Resume Next
   GetEstimatingRates cMBurden, cFoh, cRate, cGna, cProfit, cScrap
   '    sSql = "SELECT EstMatlBurden,EstFactoryOverHead," _
   '        & "EstGenAdmnExp,EstProfitOfSale,EstUseWCOverhead," _
   '        & "EstLaborRate,EstScrapRate FROM Preferences WHERE PreRecord=1"
   '    bSqlRows = clsADOCon.GetDataSet(sSql,RdoPar, ES_FORWARD)
   '        If bSqlRows Then
   '            With RdoPar
   '                txtGna = Format(!EstGenAdmnExp, "####0.00")
   '                cGna = Format(!EstGenAdmnExp / 100, "####0.00")
   '                cProfit = Format(!EstProfitOfSale / 100, "####0.00")
   '                txtScr = Format(!EstScrapRate, "####0.00")
   '                txtPrf = Format(cProfit * 100, "####0.00")
   '                cScrap = Format(!EstScrapRate / 100, "####0.00")
   '
   '            End With
   '        End If
   
   'Defaults
   cBGnaRate = cGna
   cBProfitRate = cProfit
   cBScrapRate = cScrap
   
   sEstimator = GetSetting("Esi2000", "EsiEngr", "Estimator", sEstimator)
   Exit Sub
   
DiaErr1:
   sProcName = "getrates"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub cmbBid_Click()
   bGoodBid = GetThisPPIFullBid(True)
   
End Sub


Private Sub cmbBid_LostFocus()
   cmbBid = CheckLen(cmbBid, 6)
   cmbBid = Format(Abs(Val(cmbBid)), "000000")
   If bCanceled Then Exit Sub
   bGoodBid = GetThisPPIFullBid(False)
   If bGoodBid = 0 Then
      MsgBox "That Bid Has Not Been Recorded. Select New To Add.", _
         vbInformation, Caption
      cmbBid = cmbBid.List(0)
   End If
   
End Sub


Private Sub cmbCls_Change()
   If Len(cmbCls) > 1 Then cmbCls = Left(cmbCls, 1)
   
End Sub

Private Sub cmbCls_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbCls = CheckLen(cmbCls, 1)
   For iList = 0 To cmbCls.ListCount - 1
      If cmbCls.List(iList) = cmbCls Then b = 1
   Next
   If b = 0 Then cmbCls = "Q"
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDPRE = Trim(cmbCls)
      RdoFull.Update
   End If
   
End Sub


Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
   FillCustomerRFQs Me, cmbCst, True
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   bGoodCust = GetBidCustomer(Me, cmbCst)
   If bGoodCust Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDCUST = Compress(cmbCst)
      If bGoodPart Then
         RdoFull!BIDLOCKED = 0
      Else
         RdoFull!BIDLOCKED = 1
      End If
      RdoFull.Update
   End If
   '        If bGoodCust = 0 Or bGoodPart = 0 Then
   '            MsgBox "Bids Without A Valid Customer And Valid " _
   '                & "Part Number Will Not Be Saved.", _
   '                    vbInformation, Caption
   '        End If
   FillCustomerRFQs Me, cmbCst, True
   On Error Resume Next
   If Trim(RdoFull!BIDRFQ) <> "NONE" Then cmbRfq = Trim(RdoFull!BIDRFQ)
   
End Sub


Private Sub cmbRfq_Click()
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDRFQ = cmbRfq
      RdoFull.Update
   End If
   
End Sub

Private Sub cmbRfq_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbRfq = CheckLen(cmbRfq, 14)
   If Trim(cmbRfq) = "" Then cmbRfq = "NONE"
   If cmbRfq <> "NONE" Then
      For iList = 0 To cmbRfq.ListCount - 1
         If cmbRfq = cmbRfq.List(iList) Then bByte = 1
      Next
      If bByte = 0 Then
         MsgBox "RFQ " & cmbRfq & " Has Not Been Recorded.", _
            vbInformation, Caption
         cmbRfq = "NONE"
      End If
   End If
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDRFQ = cmbRfq
      RdoFull.Update
   End If
   
End Sub


Private Sub cmdCan_Click()
   Dim bByte As Byte
   bByte = CheckBidEntries(bGoodPart, bGoodCust)
   Unload Me
   
End Sub

Private Sub cmdCom_Click()
   cmdCom_MouseDown 0, 0, 0, 0
   
End Sub

Private Sub cmdCom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim bResponse As Byte
   Dim sMsg As String
   
   If cmdCom.Caption = "C&omplete" Then
      sMsg = "Do You Wish To Mark This Bid As Completed?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         On Error Resume Next
         sSql = "UPDATE EstiTable SET BIDCOMPLETE=1, " _
                & "BIDCOMPLETED='" & Format(ES_SYSDATE, "mm/dd/yy") & "' " _
                & "WHERE BIDREF=" & Val(cmbBid) & " "
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         If Err = 0 Then
            MsgBox "Bid Was Successfully Marked Complete.", _
               vbInformation, Caption
            bGoodBid = GetThisPPIFullBid()
         Else
            MsgBox "Bid Not Was Successfully Marked Complete.", _
               vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   Else
      sMsg = "Do You Wish To Mark This Bid As Incomplete?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         On Error Resume Next
         sSql = "UPDATE EstiTable SET BIDCOMPLETE=0, " _
                & "BIDCOMPLETED=Null " _
                & "WHERE BIDREF=" & Val(cmbBid) & " "
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         If Err = 0 Then
            MsgBox "Bid Was Successfully Marked Incomplete.", _
               vbInformation, Caption
            bGoodBid = GetThisPPIFullBid()
         Else
            MsgBox "Bid Not Was Successfully Marked Incomplete.", _
               vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   End If
   
End Sub


Private Sub cmdDis_Click()
   optFrom.value = vbChecked
   'EstiESe01c.optfrom.Value = vbChecked
   'EstiESe01c.Show
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   optVew.value = vbChecked
   ViewParts.Show
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3502
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


'Should have enough options here

Private Sub cmdLbr_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   bResponse = CheckBidEntries(bGoodPart, bGoodCust)
   If bResponse = 1 Then Exit Sub
   
   bGoodRout = GetBidRouting()
   If bGoodRout = 1 Then
      ppiESe02b.lblBid = cmbBid
      ppiESe02b.Show
   Else
      bGoodRout = GetPartRouting()
      If bGoodRout = 0 Then
         'See if they want to use one
         sMsg = "This Part Does Not Have A Routing Would " & vbCrLf _
                & "You Like To Select An Existing Routing?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            EstiEsf04a.optFrom = vbChecked
            EstiEsf04a.Show 1
            If sRouting <> "" Then
               CopyARouting
               Exit Sub
            End If
         End If
         sMsg = "This Part Does Not Have A Routing." & vbCrLf _
                & "Would You Like To Create One?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            ppiESe02b.lblBid = cmbBid
            ppiESe02b.Show
         Else
            CancelTrans
         End If
      Else
         sMsg = "This Part Has A Routing." & vbCrLf _
                & "Would You Like To Copy It?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            CopyARouting
         Else
            sMsg = "Make A New Bid Only Routing?"
            bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
            If bResponse = vbYes Then
               ppiESe02b.lblBid = cmbBid
               ppiESe02b.Show
            Else
               CancelTrans
            End If
         End If
      End If
   End If
   
End Sub

Private Sub cmdNew_Click()
   ppiESe02h.Show
   
End Sub

Private Sub cmdOth_Click()
   Dim bByte As Byte
   bByte = CheckBidEntries(bGoodPart, bGoodCust)
   If bByte = 1 Then Exit Sub
   
   On Error Resume Next
   ppiESe02e.txtHoles = Format(RdoFull!BIDHOLES, "##0")
   ppiESe02e.txtHoleCost = Format(RdoFull!BIDHOLESCOST, "##0.00")
   ppiESe02e.lblHoles = Format(RdoFull!BIDHOLESTOTAL, "##0.00")
   ppiESe02e.txtMask = Format(RdoFull!BIDMASKING, "##0.00")
   ppiESe02e.txtPackage = Format(RdoFull!BIDPACKAGING, "##0.00")
   
   ppiESe02e.txtFst = Format(RdoFull!BIDFIRSTDELIVERY, "mm/dd/yy")
   ppiESe02e.txtDue = Format(RdoFull!BIDDUE, "mm/dd/yy")
   ppiESe02e.txtByr = "" & Trim(RdoFull!BIDBUYER)
   ppiESe02e.txtEst = sCurrEstimator
   ppiESe02e.txtCmt = "" & Trim(RdoFull!BIDCOMMENT)
   ppiESe02e.Show
   
End Sub

Private Sub cmdPrt_Click()
   optPart.value = vbChecked
   ppiESe01p.optFull.value = vbChecked
   ppiESe01p.txtPrt = txtPrt
   ppiESe01p.Show
   
End Sub

Private Sub cmdSrv_Click()
   Dim bByte As Byte
   bByte = CheckBidEntries(bGoodPart, bGoodCust)
   If bByte = 1 Then Exit Sub
   
   ppiESe02d.lblBid = cmbBid
   ppiESe02d.Show
   
End Sub

Private Sub cmdVew_Click()
   optSle.value = vbChecked
   ViewSales.lblPrt = txtPrt
   ViewSales.Show
   
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   sCurrEstimator = GetSetting("Esi2000", "EsiEngr", "Estimator", sCurrEstimator)
   If bOnLoad Then
      CreateIndex '* 4/17/06 Remove after update
      OpenBoxes True
      GetNextBid Me
      GetRates
      FillCustomers
      FillCustomerRFQs Me, cmbCst, True
      FillCombo
      bOnLoad = 0
   Else
      If optFrom.value = vbChecked Then
         'Unload EstiESe01c
         On Error Resume Next
         optFrom.value = vbUnchecked
         txtPrt.SetFocus
      End If
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   lblBidTot.ForeColor = ES_BLUE
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim iList As Integer
   MouseCursor 13
   Dim bResponse As Byte
   If Trim(cmbCst) = "" Or Trim(txtPrt) = "" Then
      bResponse = MsgBox("The Current Estimate Is Missing a Part Number And/Or " & vbCrLf _
                  & "Customer And Will Be Removed. Do You Still Wish To Quit?", _
                  ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         Cancel = True
      Else
         bResponse = GetBidRouting()
         If bResponse = 0 Then
            sSql = "DELETE FROM EstiTable WHERE BIDREF=" & Val(cmbBid) & " "
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
         End If
      End If
   Else
      If optPart.value = vbChecked Then Unload ppiESe01p
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   RdoFull.Close
   Set ppiESe02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   Dim bScrap As Byte
   Dim bGna As Byte
   Dim bProfit As Byte
   
   bOnLoad = 1
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   For b = 65 To 88
      cmbCls.AddItem Chr$(b)
   Next
   cmbCls = "Q"
   cmbCls.AddItem Chr$(b)
   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   txtDte.ToolTipText = "The Date Of The Current Bid. Default Today For New Bids."
   txtDue = Format(ES_SYSDATE + 10, "mm/dd/yy")
   GetEstimatingPermissions bScrap, bGna, bProfit
   If bScrap = 0 Then
      txtScr.BackColor = Es_TextDisabled
      txtScr.Locked = True
   End If
   If bGna = 0 Then
      txtGna.BackColor = Es_TextDisabled
      txtGna.Locked = True
   End If
   
   If bProfit = 0 Then
      txtPrf.BackColor = Es_TextDisabled
      txtPrf.Locked = True
   End If
   txtPrc.BackColor = Es_TextDisabled
   txtPrc.BackColor = Es_TextDisabled
   
End Sub

Private Sub FillCombo()
   cmbBid.Clear
   On Error GoTo DiaErr1
   FillEstimateCombo Me, "FULL"
   If cmbBid.ListCount > 0 Then
      cmbBid = cmbBid.List(0)
      bGoodBid = GetThisPPIFullBid(True)
   Else
      cmbBid = lblNxt
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetThisPPIFullBid(Optional bHideMsg As Boolean) As Byte
   Dim b As Byte
   Dim lBidNo As Long
   
   lBidNo = Val(cmbBid)
   bGoodBid = 0
   sRouting = ""
   lblRouting = ""
   GetThisPPIFullBid = 0
   GetRates
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM EstiTable WHERE BIDREF=" & lBidNo & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFull, ES_KEYSET)
   If bSqlRows Then
      bOpenKey = 1
      With RdoFull
         If !BIDCANCELED = 1 Then
            ClearResultSet RdoFull
            OpenBoxes False
            bCanceled = 1
            Exit Function
         End If
         If !BIDACCEPTED = 1 Then
            If Not bHideMsg Then
               MsgBox "Bid " & cmbBid & " Was Accepted And Cannot Be Edited.", _
                  vbInformation
            End If
            ClearResultSet RdoFull
            OpenBoxes False
            bCanceled = 1
            Exit Function
         End If
         bCanceled = 0
         If !BIDCOMPLETE = 1 Then
            OpenBoxes False
            cmdCom.Enabled = True
            cmdCom.Caption = "Unc&omplete"
            cmdCom.ToolTipText = "Remove Complete Flag From Bid"
            optCom.value = vbChecked
         Else
            OpenBoxes True
            cmdCom.Caption = "C&omplete"
            cmdCom.ToolTipText = "Mark Bid As Completed"
            cmdCom.Enabled = True
            txtPrt.Enabled = True
            cmbCst.Enabled = True
            txtQty.Enabled = True
            optCom.value = vbUnchecked
         End If
         If Trim(!BidClass) = "QWIK" Then
            MsgBox "Estimate Is A Qwik Bid.  Edit From The Qwik Bid Area.", _
               vbInformation, Caption
            ClearResultSet RdoFull
            OpenBoxes False
            bCanceled = 1
            Exit Function
         End If
         cmbCls = "" & Trim(!BIDPRE)
         txtPrt = "" & Trim(!BidPart)
         bGoodPart = GetBidPart(Me)
         If "" & Trim(!BIDCUST) <> "" Then cmbCst = "" & Trim(!BIDCUST)
         bGoodCust = GetBidCustomer(Me, cmbCst)
         txtDte = Format(!BIDDATE, "mm/dd/yy")
         If Not IsNull(!BIDDUE) Then
            txtDue = Format(!BIDDUE, "mm/dd/yy")
         Else
            txtDue = Format(!BIDDATE + 10, "mm/dd/yy")
         End If
         If !BidQuantity > 0 Then
            txtQty = Format(!BidQuantity, "####0.00")
         Else
            txtQty = "1.000"
         End If
         cOldQty = Val(txtQty)
         lblHours = Format(!BIDHOURS, "####0.00")
         lblRate = Format(!BIDRATE, "####0.00")
         lblFohRate = Format(!BIDFOHRATE, "####0.00")
         lblLabor = Format(!BIDTOTLABOR, "####0.00")
         
         If !BIDGNARATE = 0 Then cBGnaRate = cGna _
                          Else cBGnaRate = !BIDGNARATE
         txtGna = Format(cBGnaRate * 100, "####0.00")
         
         If !BIDPROFITRATE = 0 Then cBProfitRate = cProfit _
                             Else cBProfitRate = !BIDPROFITRATE
         txtPrf = Format(cBProfitRate * 100, "####0.00")
         
         If !BIDPROFIT = 0 Then cBprofit = cProfit _
                         Else cBprofit = !BIDPROFIT
         
         If !BIDSCRAPRATE = 0 Then cBScrapRate = cScrap _
                            Else cBScrapRate = !BIDSCRAPRATE
         txtScr = Format(cBScrapRate * 100, "####0.00")
         
         If Trim(!BIDESTIMATOR) = "" Then
            sBEstimator = sEstimator
         Else
            sBEstimator = "" & Trim(!BIDESTIMATOR)
         End If
         Estimator = sBEstimator
         lblEstimator = sBEstimator
         'lblScrap = Format(!BIDSCRAP, "####0.00")
         'lblFoh = Format(!BIDFOH, "####0.00")
         'lblMaterial = Format(!BIDMATL, "####0.00")
         'lblBurden = Format(!BIDBURDEN, "####0.00")
         'lblTotMat = Format(!BIDTOTMATL, "####0.00")
         lblTotServices = Format(!BIDOSP, "####0.00")
         lblOther = Format(!BIDHOLESTOTAL + !BIDMASKING + !BIDPACKAGING, "#####0.00")
         lblGNA = Format(!BIDGNA, "####0.00")
         FindCustomer Me, cmbCst
         FillCustomerRFQs Me, cmbCst, True
         cmbRfq = "" & Trim(!BIDRFQ)
         If Trim(cmbRfq) = "" Then cmbRfq = "NONE"
         txtPrc = Format(!BIDUNITPRICE, "#####0.00")
         
         On Error Resume Next
         'fix joins
         For b = 1 To 5
            sSql = "INSERT INTO EsosTable (BIDOSREF,BIDOSROW) " _
                   & "VALUES('" & Val(cmbBid) & "'," _
                   & b & ")"
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            If Err > 0 Then Exit For
         Next
         
         'b = GetBidServices(ppiESe02a)
         b = GetBidServices(ppiESe02a, CCur("0" & ppiESe02a.txtQty))
         UnitPrice
      End With
      GetThisPPIFullBid = 1
   Else
      'Reset defaults
      cBGnaRate = cGna
      cBProfitRate = cProfit
      cBScrapRate = cScrap
      cBprofit = 0
      GetThisPPIFullBid = 0
   End If
   bOpenKey = 0
   Exit Function
   
DiaErr1:
   sProcName = "getthisbid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub lblBurden_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull.Update
   End If
   
End Sub

Private Sub lblGNA_Change()
   Dim cGna As Currency
   Dim cGnr As Currency
   cGna = Format(Val(lblGNA), "####0.00")
   cGnr = Format(Val(txtGna) / 100, "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDGNA = Format(cGna, "####0.00")
      RdoFull!BIDGNARATE = Format(cGnr, "####0.00")
      RdoFull.Update
   End If
   
End Sub

Private Sub lblHours_Change()
   '
End Sub

Private Sub LblLabor_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      'RdoFull!BIDRATE = Format(Val(lblRate), "####0.00")
      RdoFull!BIDTOTLABOR = Format(Val(lblLabor), "####0.00")
      'RdoFull!BIDHOURS = Format(Val(lblHours), "####0.00")
      RdoFull!BIDFOH = Format(Val(lblFoh), "####0.00")
      RdoFull!BIDFOHRATE = Format(Val(lblFohRate), "####0.00")
      RdoFull.Update
      UnitPrice
   End If
   
End Sub

Private Sub lblMaterial_Change()
   '
End Sub

Private Sub lblOther_Change()
   UnitPrice
   
End Sub

Private Sub lblPrf_Change()
   If bOpenKey = 1 Then Exit Sub
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDPROFIT = Format(Val(lblPrf), "####0.00")
      RdoFull!BIDPROFITRATE = Format(Val(txtPrf) / 100)
      RdoFull.Update
   End If
   
End Sub

Private Sub lblRouting_Change()
   sRouting = lblRouting
   
End Sub

Private Sub lblScrap_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDSCRAP = Format(Val(lblScrap), "####0.00")
      RdoFull!BIDSCRAPRATE = Format(Val(txtScr) / 100, "####0.00")
      RdoFull.Update
   End If
   
End Sub

Private Sub lblTotMat_Change()
   Dim cMat As Currency
   Dim cBur As Currency
   Dim cTot As Currency
   Dim cBbd As Currency
   If bGoodBid Then
      On Error Resume Next
      cTot = Format(Val(lblTotMat), "####0.00")
      cMat = Format(Val(lblMaterial), "####0.00")
      cBur = Format(Val(lblBurden), "####0.00")
      'RdoFull.Edit
      RdoFull!BIDTOTMATL = Format(cTot, "####0.00")
      RdoFull!BIDMATL = Format(cMat, "####0.00")
      RdoFull!BIDBURDEN = Format(cBur, "####0.00")
      If cMat > 0 Then
         cBbd = cBur / cMat
         RdoFull!BIDBURDENRATE = Format(cBbd, "####0.00")
      Else
         RdoFull!BIDBURDENRATE = 0
      End If
      If Val(lblTotMat) > 0 Then
         RdoFull!BIDMATLDESC = "Total From Parts List"
      Else
         RdoFull!BIDMATLDESC = "No Parts List"
      End If
      RdoFull.Update
      UnitPrice
   End If
   
End Sub

Private Sub lblTotServices_Change()
   '
End Sub

Private Sub lblUnitServices_Change()
   lblUnitServices = Format(Val(lblUnitServices), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDOSP = Val(lblUnitServices)
      If Val(lblUnitServices) > 0 Then
         RdoFull!BIDOSPDESC = "Total Entries"
      Else
         RdoFull!BIDOSPDESC = "No Services Entered"
      End If
      RdoFull.Update
      UnitPrice
   End If
   
End Sub

Private Sub optVew_Click()
   If optVew.value = vbUnchecked Then txtPrt.SetFocus
   
End Sub


Private Sub Timer1_Timer()
   GetNextBid Me
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDDATE = Format(txtDte, "mm/dd/yy")
      RdoFull.Update
   End If
   
End Sub


Private Sub txtDue_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDue_LostFocus()
   txtDue = CheckDate(txtDue)
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDDUE = Format(txtDue, "mm/dd/yy")
      RdoFull.Update
   End If
   
End Sub


Private Sub txtGna_LostFocus()
   txtGna = Format(Abs(Val(txtGna)), "#####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDGNARATE = Format(Val(txtGna) / 100)
      RdoFull.Update
   End If
   UnitPrice
   
End Sub


Private Sub txtPrc_Change()
   If bOpenKey = 1 Or bOnLoad = 1 Then Exit Sub
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDUNITPRICE = Format(Val(txtPrc), "####0.00")
      RdoFull.Update
      UpdateDiscounts
   End If
   
End Sub

Private Sub txtPrf_LostFocus()
   txtPrf = CheckLen(txtPrf, 7)
   txtPrf = Format(Abs(Val(txtPrf)), "##0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDPROFITRATE = Format(Val(txtPrf) / 100)
      RdoFull.Update
   End If
   UnitPrice
   
End Sub


Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      optVew.value = vbChecked
      ViewParts.Show
   ElseIf KeyCode = vbKeyF2 Then
      optPart.value = vbChecked
      cmbBid.Enabled = False
      ppiESe01p.optFull.value = vbUnchecked
      ppiESe01p.txtPrt = txtPrt
      ppiESe01p.Show
      
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   bGoodPart = GetBidPart(Me)
   If bGoodPart Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BidPart = Compress(txtPrt)
      If bGoodCust Then
         RdoFull!BIDLOCKED = 0
      Else
         RdoFull!BIDLOCKED = 1
      End If
      RdoFull.Update
   Else
      '   MsgBox "Bids Without A Valid PartNumber Will Not Be Saved.", _
      '       vbInformation, Caption
   End If
   
End Sub


Private Sub OpenBoxes(bOpen As Boolean)
   Dim iList As Integer
   lblHours = "0.00"
   lblBurden = "0.00"
   lblLabor = "0.00"
   lblMaterial = "0.00"
   lblUnitServices = "0.00"
   lblTotServices = "0.00"
   lblTotMat = "0.00"
   lblOther = "0.00"
   lblFoh = "0.00"
   lblScrap = "0.00"
   lblGNA = "0.00"
   lblPrf = "0.00"
   txtPrf = "0.00"
   txtPrc = "0.00"
   lblBidTot = "0.00"
   If bOpen Then
      For iList = 0 To Controls.Count - 1
         If TypeOf Controls(iList) Is TextBox Or TypeOf Controls(iList) Is ComboBox _
                              Or TypeOf Controls(iList) Is CommandButton Or TypeOf Controls(iList) _
                              Is ComboBox Then
            Controls(iList).Enabled = True
         End If
      Next
   Else
      For iList = 0 To Controls.Count - 1
         If TypeOf Controls(iList) Is TextBox Or TypeOf Controls(iList) Is ComboBox _
                              Or TypeOf Controls(iList) Is CommandButton Or TypeOf Controls(iList) _
                              Is ComboBox Then
            Controls(iList).Enabled = False
         End If
      Next
   End If
   cmdCan.Enabled = True
   cmbBid.Enabled = True
   cmdNew.Enabled = True
   cmdVew.Enabled = True
   
End Sub

Private Sub UnitPrice()
   Dim cQuantity As Currency
   Dim cUnitLbr As Currency
   Dim cUnitMtl As Currency
   Dim cUnitOsp As Currency
   Dim cUnitPrc As Currency
   Dim cOthers As Currency
   Dim cScrap As Currency
   Dim cGenAdm As Currency
   Dim cProfit As Currency
   
   cBGnaRate = Val(txtGna) / 100
   cBProfitRate = Val(txtPrf) / 100
   cScrap = Val(txtScr)
   
   cQuantity = cOldQty
   If cQuantity = 0 Then cQuantity = 1
   cUnitLbr = Val(lblLabor)
   cUnitMtl = Val(lblTotMat)
   cOthers = Val(lblOther) 'Unit Cost For PROPLA
   cUnitOsp = Val(lblUnitServices)
   cScrap = (cUnitLbr + cUnitMtl + cUnitOsp) * (cScrap / 100)
   lblScrap = Format(cScrap, "####0.00")
   cUnitPrc = cUnitLbr + cUnitMtl + cUnitOsp + cScrap
   cGenAdm = cUnitPrc * cBGnaRate
   lblGNA = Format(cGenAdm, "####0.00")
   cUnitPrc = cUnitPrc + cGenAdm
   cProfit = cUnitPrc * cBProfitRate
   lblPrf = Format(cProfit, "####0.00")
   cUnitPrc = cUnitPrc + cProfit + cOthers
   txtPrc = Format(cUnitPrc, "####0.00")
   TotalBid
   
   
End Sub


Private Function GetPartRouting() As Byte
   Dim RdoRte As ADODB.Recordset
   sSql = "SELECT PARTREF,PAROUTING FROM PartTable Where PARTREF='" _
          & Compress(txtPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         If Trim(!PAROUTING) = "" Then
            GetPartRouting = 0
            sRouting = ""
            lblRouting = ""
         Else
            sRouting = "" & Trim(!PAROUTING)
            lblRouting = sRouting
            GetPartRouting = 1
         End If
         ClearResultSet RdoRte
      End With
   Else
      GetPartRouting = 0
   End If
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getbidrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub CopyARouting()
   Dim RdoRte As ADODB.Recordset
   Dim sMsg As String
   On Error GoTo DiaErr1

'this blows up if there are multiple wc's with the same name
'   sSql = "SELECT OPREF,OPNO,OPSHOP,OPCENTER,OPSETUP,OPUNIT," _
'          & "OPQHRS,OPMHRS,WCNREF,WCNOHPCT,WCNESTRATE FROM " _
'          & "RtopTable,WcntTable WHERE (OPCENTER=WCNREF AND OPREF='" _
'          & sRouting & "') ORDER BY OPNO"
   
   sSql = "SELECT OPREF,OPNO,OPSHOP,OPCENTER,OPSETUP,OPUNIT," & vbCrLf _
          & "OPQHRS,OPMHRS,WCNREF,WCNOHPCT,WCNESTRATE" & vbCrLf _
          & "FROM RtopTable" & vbCrLf _
          & "JOIN WcntTable ON OPCENTER=WCNREF AND OPSHOP=WCNSHOP" & vbCrLf _
          & "WHERE OPREF='" & sRouting & "'" & vbCrLf _
          & "ORDER BY OPNO"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      MouseCursor 13
      With RdoRte
         'On Error Resume Next
         clsADOCon.BeginTrans
         Do Until .EOF
            sSql = "INSERT INTO EsrtTable (BIDRTEREF," _
                   & "BIDRTEOPNO,BIDRTESHOP,BIDRTECENTER," _
                   & "BIDRTESETUP,BIDRTEUNIT,BIDRTEQHRS," _
                   & "BIDRTEMHRS,BIDRTERATE,BIDRTEFOHRATE) VALUES(" _
                   & Val(cmbBid) & "," _
                   & str$(!OPNO) & ",'" _
                   & Trim(!OPSHOP) & "','" _
                   & Trim(!OPCENTER) & "'," _
                   & Format$(!OPSETUP) & "," _
                   & Format$(!OPUNIT) & "," _
                   & Format$(!OPQHRS) & "," _
                   & Format$(!OPMHRS) & "," _
                   & Format$(!WCNESTRATE) & "," _
                   & Format$(!WCNOHPCT / 100) & ")"
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            .MoveNext
         Loop
         MouseCursor 0
'         If Err = 0 Then
            clsADOCon.CommitTrans
            Sleep 500
            MsgBox "Your Routing Is Ready. Reselect Labor.", _
               vbInformation, Caption
'         Else
'            clsADOCon.RollbackTrans
'            MsgBox "Could Not Copy The Routing.", _
'               vbExclamation, Caption
'         End If
         ClearResultSet RdoRte
      End With
   End If
   Set RdoRte = Nothing
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "copyrouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Function GetPartsList()
   Dim RdoPls As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT BMASSYPART,BMREV FROM BmplTable WHERE " _
          & "BMASSYPART='" & Compress(txtPrt) & "' AND BMREV='" _
          & sBomRev & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_FORWARD)
   If bSqlRows Then
      GetPartsList = 1
      ClearResultSet RdoPls
   Else
      GetPartsList = 0
   End If
   Set RdoPls = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpartslist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

'Changed 8/11/03
'Added BOMLEVEL 8/1/03

Private Sub CopyPartsList()
   Dim RdoPls As ADODB.Recordset
   Dim b As Byte
   Dim sPartNumber As String
   Dim sPartRev As String
   On Error GoTo 0
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMPARTREV,BMSEQUENCE,BMQTYREQD," _
          & "BMUNITS,BMCONVERSION,BMADDER,BMSETUP,BMCOMT,BMESTLABOR," _
          & "BMESTLABOROH FROM BmplTable WHERE BMASSYPART='" & Compress(txtPrt) _
          & "' AND BMREV='" & sBomRev & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_FORWARD)
   If bSqlRows Then
      MouseCursor 13
      With RdoPls
         On Error Resume Next
         Do Until .EOF
            b = b + 1
            bBomLevel = 1
            sPartNumber = "" & Trim(!BMPARTREF)
            sPartRev = "" & Trim(!BMPARTREV)
            sSql = "INSERT INTO EsbmTable (BIDBOMREF,BIDBOMASSYPART," _
                   & "BIDBOMSEQUENCE,BIDBOMPARTREF,BIDBOMLEVEL,BIDBOMQTYREQD,BIDBOMUNITS," _
                   & "BIDBOMCONVERSION,BIDBOMADDER,BIDBOMSETUP," _
                   & "BIDBOMCOMT,BIDBOMLABOR,BIDBOMLABOROH) VALUES(" _
                   & Val(cmbBid) & ",'" _
                   & Compress(txtPrt) & "'," _
                   & b & ",'" _
                   & Trim(!BMPARTREF) & "'," _
                   & bBomLevel & "," _
                   & Trim(!BMQTYREQD) & ",'" _
                   & Trim(!BMUNITS) & "'," _
                   & Trim(!BMCONVERSION) & "," _
                   & Trim(!BMADDER) & "," _
                   & Trim(!BMSETUP) & ",'" _
                   & Trim(!BMCOMT) & "'," _
                   & Trim(!BMESTLABOR) & "," _
                   & Trim(!BMESTLABOROH) & ")"
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            GetNextBomLevel sPartNumber, sPartRev
            .MoveNext
         Loop
         ClearResultSet RdoPls
         MouseCursor 0
         If Err = 0 Then
            clsADOCon.CommitTrans
            MsgBox "Your Bill Of Material Is Ready. Reselect Material.", _
               vbInformation, Caption
         Else
            MsgBox "Could Not Copy Some Of The Bill " & vbCrLf _
               & "Possibly A Duplicate Or Circle." & vbCrLf _
               & "Double Check Your Bill Of Material.", _
               vbExclamation, Caption
         End If
      End With
   End If
   Set RdoPls = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "copypartslist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtQty_LostFocus()
   Dim b As Byte
   txtQty = CheckLen(txtQty, 10)
   txtQty = Format(Abs(Val(txtQty)), "####0.00")
   If Val(txtQty) = 0 Then txtQty = "1.000"
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BidQuantity = Val(txtQty)
      RdoFull.Update
      If cOldQty <> Val(txtQty) Then
         GetTheLabor
         'b = GetBidServices(ppiESe02a)
         b = GetBidServices(ppiESe02a, CCur("0" & ppiESe02a.txtQty))
      End If
   End If
   cOldQty = Val(txtQty)
   UnitPrice
   
End Sub



Private Sub TotalBid()
   Dim cBidTot As Currency
   cBidTot = (Val(txtQty) * Format(Val(txtPrc), "####0.00"))
   lblBidTot = Format(cBidTot, "##,###,##0.00")
   
End Sub

Private Sub UpdateDiscounts()
   Dim cPer As Currency
   Dim cPrc As Currency
   On Error Resume Next
   bGoodBid = RefreshKeySet
   If bGoodBid = 1 Then
      With RdoFull
         '.Edit
         cPer = (!BIDQTYDISC1 / 100)
         cPrc = (1 - cPer) * Val(txtPrc)
         !BIDQTYPRICE1 = cPrc
         
         cPer = (!BIDQTYDISC2 / 100)
         cPrc = (1 - cPer) * Val(txtPrc)
         !BIDQTYPRICE2 = cPrc
         
         cPer = (!BIDQTYDISC3 / 100)
         cPrc = (1 - cPer) * Val(txtPrc)
         !BIDQTYPRICE3 = cPrc
         
         cPer = (!BIDQTYDISC4 / 100)
         cPrc = (1 - cPer) * Val(txtPrc)
         !BIDQTYPRICE4 = cPrc
         
         cPer = (!BIDQTYDISC5 / 100)
         cPrc = (1 - cPer) * Val(txtPrc)
         !BIDQTYPRICE5 = cPrc
         
         cPer = (!BIDQTYDISC6 / 100)
         cPrc = (1 - cPer) * Val(txtPrc)
         !BIDQTYPRICE6 = cPrc
         .Update
      End With
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "updatedisc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'8/7/03 - Adjusts for lose of KeySet when adding bom items

Private Function RefreshKeySet() As Byte
   On Error Resume Next
   sSql = "SELECT * FROM EstiTable WHERE BIDREF=" & Val(cmbBid) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFull, ES_KEYSET)
   If bSqlRows Then
      RefreshKeySet = 1
   Else
      RefreshKeySet = 0
   End If
   
End Function

'ensueing levels 8/11/03
'From Copy Parts List

Private Sub GetNextBomLevel(BomPartRef As String, BomPartRev As String)
   Dim RdoBm2 As ADODB.Recordset
   Dim b As Byte
   Dim sPartNumber As String
   Dim sPartRev As String
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMPARTREV,BMSEQUENCE,BMQTYREQD," _
          & "BMUNITS,BMCONVERSION,BMADDER,BMSETUP,BMCOMT,BMESTLABOR," _
          & "BMESTLABOROH FROM BmplTable WHERE BMASSYPART='" & BomPartRef _
          & "' AND BMREV='" & BomPartRev & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBm2, ES_FORWARD)
   If bSqlRows Then
      With RdoBm2
         bBomLevel = bBomLevel + 1
         Do Until .EOF
            sPartNumber = "" & Trim(!BMPARTREF)
            sPartRev = "" & Trim(!BMPARTREV)
            sSql = "INSERT INTO EsbmTable (BIDBOMREF,BIDBOMASSYPART," _
                   & "BIDBOMSEQUENCE,BIDBOMPARTREF,BIDBOMLEVEL,BIDBOMQTYREQD,BIDBOMUNITS," _
                   & "BIDBOMCONVERSION,BIDBOMADDER,BIDBOMSETUP," _
                   & "BIDBOMCOMT,BIDBOMLABOR,BIDBOMLABOROH) VALUES(" _
                   & Val(cmbBid) & ",'" _
                   & BomPartRef & "'," _
                   & b & ",'" _
                   & Trim(!BMPARTREF) & "'," _
                   & bBomLevel & "," _
                   & Trim(!BMQTYREQD) & ",'" _
                   & Trim(!BMUNITS) & "'," _
                   & Trim(!BMCONVERSION) & "," _
                   & Trim(!BMADDER) & "," _
                   & Trim(!BMSETUP) & ",'" _
                   & Trim(!BMCOMT) & "'," _
                   & Trim(!BMESTLABOR) & "," _
                   & Trim(!BMESTLABOROH) & ")"
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            GetNextBomLevel sPartNumber, sPartRev
            .MoveNext
         Loop
         ClearResultSet RdoBm2
      End With
   End If
   
End Sub

Private Sub txtScr_LostFocus()
   txtScr = CheckLen(txtScr, 7)
   txtScr = Format(Abs(Val(txtScr)), "###0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDSCRAPRATE = Format(Val(txtScr) / 100)
      RdoFull.Update
   End If
   UnitPrice
   
End Sub
