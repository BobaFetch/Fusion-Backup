VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form diaSCf02a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy Last Invoiced Cost To Standard"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7380
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.CheckBox optRpt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   76
      Top             =   3240
      Width           =   735
   End
   Begin VB.ComboBox txtThr 
      Height          =   315
      Left            =   3240
      TabIndex        =   67
      Tag             =   "4"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox txtFrm 
      Height          =   315
      Left            =   1320
      TabIndex        =   65
      Tag             =   "4"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox optVew 
      Height          =   255
      Left            =   4680
      TabIndex        =   54
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   320
      Left            =   4200
      Picture         =   "diaSCf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Show BOM Structure"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox cmbprt1 
      Height          =   285
      Left            =   1320
      TabIndex        =   51
      Tag             =   "3"
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   49
      ToolTipText     =   "Update Costs"
      Top             =   480
      Width           =   875
   End
   Begin VB.CheckBox optStd 
      Caption         =   "___"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   46
      Top             =   3000
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CheckBox optPrp 
      Caption         =   "___"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   45
      Top             =   2760
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.Frame frmCriteria 
      Caption         =   "Update By Invoice \  Part Criteria "
      Height          =   3555
      Left            =   240
      TabIndex        =   19
      Top             =   5520
      Width           =   6435
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   4920
         TabIndex        =   43
         Top             =   2640
         Width           =   1335
         Begin VB.CommandButton cmdUpd2 
            Caption         =   "Update"
            Height          =   315
            Left            =   240
            TabIndex        =   44
            ToolTipText     =   "Print Report And Update Costs"
            Top             =   420
            Width           =   875
         End
         Begin VB.CommandButton optDis 
            Height          =   330
            Left            =   120
            Picture         =   "diaSCf02a.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Display The Report"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton optPrn 
            Height          =   330
            Left            =   720
            Picture         =   "diaSCf02a.frx":04C0
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Print The Report"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.ComboBox cmbCde 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "diaSCf02a.frx":064A
         Left            =   3720
         List            =   "diaSCf02a.frx":064C
         Sorted          =   -1  'True
         TabIndex        =   4
         Tag             =   "3"
         ToolTipText     =   "Enter/Revise Product Code (6 Char)"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cmbCls 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtVar 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Tag             =   "1"
         Top             =   2640
         Width           =   615
      End
      Begin VB.CheckBox optTyp 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   9
         Top             =   1980
         Width           =   555
      End
      Begin VB.CheckBox optExt 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   3180
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox optDsc 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   2940
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox optStd2 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   4140
         TabIndex        =   11
         Top             =   2340
         Width           =   735
      End
      Begin VB.CheckBox optPrp2 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   2340
         Width           =   735
      End
      Begin VB.CheckBox optTyp 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   8
         Top             =   1980
         Width           =   555
      End
      Begin VB.CheckBox optTyp 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Index           =   1
         Left            =   2220
         TabIndex        =   7
         Top             =   1980
         Width           =   615
      End
      Begin VB.CheckBox optTyp 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   6
         Top             =   1980
         Width           =   555
      End
      Begin VB.ComboBox cmbMBE 
         Height          =   315
         Left            =   4260
         TabIndex        =   5
         Tag             =   "4"
         Top             =   1620
         Width           =   615
      End
      Begin VB.ComboBox txtEnd 
         Height          =   315
         Left            =   4020
         TabIndex        =   1
         Tag             =   "4"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox optBOM 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2940
         TabIndex        =   2
         Top             =   840
         Value           =   1  'Checked
         Width           =   555
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   " (Blank For All)"
         Height          =   285
         Index           =   26
         Left            =   5160
         TabIndex        =   40
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   " (Blank For All)"
         Height          =   285
         Index           =   25
         Left            =   5040
         TabIndex        =   39
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         Height          =   255
         Index           =   27
         Left            =   2520
         TabIndex        =   38
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Class"
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   " (Only Type 5's On Them!)"
         Height          =   405
         Index           =   23
         Left            =   3960
         TabIndex        =   36
         Top             =   840
         Width           =   2145
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Include Part Type 3's With No BOM? "
         Height          =   285
         Index           =   22
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   2745
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "%   "
         Height          =   285
         Index           =   10
         Left            =   3480
         TabIndex        =   34
         Top             =   2700
         Width           =   285
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Flag Variances Greater Than:"
         Height          =   285
         Index           =   8
         Left            =   360
         TabIndex        =   33
         Top             =   2700
         Width           =   2145
      End
      Begin VB.Label zTyp 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         Height          =   255
         Index           =   3
         Left            =   3660
         TabIndex        =   32
         Top             =   1980
         Width           =   165
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Include Extended Descriptions?"
         Height          =   285
         Index           =   20
         Left            =   360
         TabIndex        =   31
         Top             =   3180
         Width           =   2385
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Include Part Description?"
         Height          =   285
         Index           =   19
         Left            =   360
         TabIndex        =   30
         Top             =   2940
         Width           =   2025
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Cost?"
         Height          =   285
         Index           =   14
         Left            =   2880
         TabIndex        =   29
         Top             =   2340
         Width           =   1185
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proposed Cost?"
         Height          =   285
         Index           =   13
         Left            =   840
         TabIndex        =   28
         Top             =   2340
         Width           =   1185
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Update:"
         Height          =   285
         Index           =   12
         Left            =   120
         TabIndex        =   27
         Top             =   2340
         Width           =   945
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   " (Blank For All)"
         Height          =   285
         Index           =   11
         Left            =   5040
         TabIndex        =   26
         Top             =   1620
         Width           =   1185
      End
      Begin VB.Label zTyp 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   2820
         TabIndex        =   25
         Top             =   1980
         Width           =   165
      End
      Begin VB.Label zTyp 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   24
         Top             =   1980
         Width           =   180
      End
      Begin VB.Label zTyp 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   1260
         TabIndex        =   23
         Top             =   1980
         Width           =   180
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Types:"
         Height          =   285
         Index           =   9
         Left            =   360
         TabIndex        =   22
         Top             =   1980
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Update Following Part Types If Make\Buy Field Is: "
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   1620
         Width           =   3735
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   " Ending "
         Height          =   255
         Index           =   5
         Left            =   3360
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6360
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3600
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4590
      FormDesignWidth =   7380
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaSCf02a.frx":064E
      PictureDn       =   "diaSCf02a.frx":0794
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   41
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaSCf02a.frx":08DA
      PictureDn       =   "diaSCf02a.frx":0A20
   End
   Begin ComctlLib.ProgressBar prg1 
      Height          =   255
      Left            =   240
      TabIndex        =   68
      Top             =   4080
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Report"
      Height          =   285
      Index           =   31
      Left            =   240
      TabIndex        =   75
      Top             =   3240
      Width           =   1665
   End
   Begin VB.Label lblFnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5640
      TabIndex        =   74
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Found"
      Height          =   255
      Index           =   30
      Left            =   4920
      TabIndex        =   73
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   72
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Of"
      Height          =   285
      Index           =   29
      Left            =   2400
      TabIndex        =   71
      Top             =   3720
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblRec 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   70
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Updating Part"
      Height          =   285
      Index           =   28
      Left            =   240
      TabIndex        =   69
      Top             =   3720
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   " Ending "
      Height          =   255
      Index           =   21
      Left            =   2520
      TabIndex        =   66
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   255
      Index           =   18
      Left            =   240
      TabIndex        =   64
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Or"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   63
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblStd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5640
      TabIndex        =   62
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current"
      Height          =   285
      Index           =   3
      Left            =   5640
      TabIndex        =   61
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label lblPrp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4440
      TabIndex        =   60
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proposed"
      Height          =   285
      Index           =   6
      Left            =   4440
      TabIndex        =   59
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblUnt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3240
      TabIndex        =   58
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   57
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   56
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   55
      Top             =   1320
      Width           =   1665
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   52
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   50
      Top             =   600
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Proposed"
      Height          =   285
      Index           =   16
      Left            =   240
      TabIndex        =   48
      Top             =   2760
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Standard"
      Height          =   285
      Index           =   17
      Left            =   240
      TabIndex        =   47
      Top             =   3000
      Width           =   1665
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   42
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaSCf02a"
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

'*********************************************************************************
' diaSCp06a
'
' Notes:
'   *WHAT THIS FUNCTION DOES:
'       If updating by ivoice and part criteria:
'            This function grabs all of the last invoices for the parts that match
'            the criteria specified (type, bom, etc.) Then, if that invoice date is
'            within the specified daterange it will update the std\prop. cost from
'            that invoice unit cost.
'        If  updating by part no.
'             Simply grabs the very last invoice for that part and updates costs.
' Created: 03/05/04 (JCW)
' Revisions:
' 03/26/04 (JCW)- Added printing functionality from form
' 09/20/04 (nth) Completely reworked
'
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sMsg As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub


Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      MouseCursor 0
      cmdHlp = False
   End If
End Sub

Private Sub cmdUpd_Click()
   If Val(lblFnd) > 0 Then
      BatchUpdate
   Else
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      UpdateCost cmbPrt, lblUnt, optPrp, optStd
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         optPrp.Caption = lblUnt.Caption
         If optStd Then
            lblStd = lblUnt
         End If
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         sMsg = "Cannot Update Cost."
         MsgBox sMsg, vbInformation, Caption
      End If
   End If
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillPartCombo cmbPrt
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   GetOptions
   sCurrForm = Caption
   bOnLoad = True
   bCancel = False
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaSCp02a = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub cmbPrt_LostFocus()
   If Not bCancel Then
      cmbPrt = CheckLen(cmbPrt, 30)
      If Len(Trim(cmbPrt)) Then
         GetPart cmbPrt
         txtFrm.enabled = False
         txtThr.enabled = False
      Else
         txtFrm.enabled = True
         txtThr.enabled = True
      End If
   End If
End Sub

Private Sub cmbPrt_Click()
   If Not bCancel Then
      cmbPrt = CheckLen(cmbPrt, 30)
      If Len(Trim(cmbPrt)) Then
         GetPart cmbPrt
         txtFrm.enabled = False
         txtThr.enabled = False
      Else
         txtFrm.enabled = True
         txtThr.enabled = True
      End If
   End If
End Sub


Private Sub cmdFnd_Click()
   optVew.Value = vbChecked
   ViewParts.Show
End Sub

Private Sub optPrp_Click()
   If optPrp = vbUnchecked Then
      optStd = vbUnchecked
      cmdUpd.enabled = False
   Else
      cmdUpd.enabled = True
   End If
End Sub

Private Sub optStd_Click()
   If optStd = vbChecked Then
      optPrp = vbChecked
      cmdUpd.enabled = True
   Else
      If optPrp = vbUnchecked Then
         cmdUpd.enabled = False
      End If
   End If
End Sub

Private Sub UpdateCost(sPart As String, _
                       cCost As Currency, _
                       bProposed As Byte, _
                       bStandard As Byte)
   Dim sTemp As String
   sProcName = "updateco"
   sPart = Compress(sPart)
   sTemp = "UPDATE PartTable SET "
   If bProposed Then
      sTemp = sTemp & "PALEVMATL = " & cCost & ",PATOTMATL = " & cCost
   End If
   If bStandard Then
      sTemp = sTemp & ",PAREVDATE='" & Format(ES_SYSDATE, "mm/dd/yy") _
              & "',PASTDCOST = " & cCost
      UpdatePrevious sPart
   End If
   sTemp = sTemp & " WHERE PartRef = '" & sPart & "'"
   sSql = sTemp
   clsADOCon.ExecuteSQL sSql
End Sub

Private Sub UpdatePrevious(sPart As String)
   sProcName = "updatepr"
   sSql = "UPDATE PartTable SET " _
          & "PAPREVSTDCOST = PASTDCOST," _
          & "PAPREVHRS = PATOTHRS," _
          & "PAPREVLABOR = PATOTLABOR," _
          & "PAPREVMATL = PATOTMATL," _
          & "PAPREVEXP = PATOTEXP," _
          & "PAPREVOH = PATOTOH " _
          & "WHERE PARTREF = '" & sPart & "'"
   clsADOCon.ExecuteSQL sSql
End Sub

Private Sub GetPart(sPart As String)
   Dim RdoPrt As ADODB.Recordset
   Dim sInvoice As String
   Dim cUnit As Currency
   On Error GoTo DiaErr1
   sProcName = "getpart"
   sPart = Compress(sPart)
   sSql = "SELECT PASTDCOST,(PATOTHRS+PATOTLABOR+PATOTMATL+PATOTEXP+PATOTOH)," _
          & "PADESC FROM PartTable WHERE PARTREF = '" & sPart & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
   If bSqlRows Then
      With RdoPrt
         lblStd = Format(.Fields(0), CURRENCYMASK)
         lblPrp = Format(.Fields(1), CURRENCYMASK)
         lblDsc.ForeColor = Me.ForeColor
         lblDsc = "" & Trim(.Fields(2))
         .Cancel
      End With
      LastInvoiceCost sPart, cUnit, sInvoice
      lblUnt = Format(cUnit, CURRENCYMASK)
      lblInv = sInvoice
      cmdUpd.enabled = True
      optPrp.enabled = True
      optStd.enabled = True
   Else
      lblDsc.ForeColor = ES_RED
      lblDsc = "*** No Invoices Found ***"
      lblInv = ""
      lblStd = ""
      lblUnt = ""
      lblPrp = ""
      cmdUpd.enabled = False
      optPrp.enabled = False
      optStd.enabled = False
   End If
   Set RdoPrt = Nothing
   Exit Sub
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub LastInvoiceCost(ByVal sPart As String, _
                            ByRef cCost As Currency, _
                            ByRef sInvoice As String)
   Dim rdoCst As ADODB.Recordset
   sProcName = "lastinvo"
   sPart = Compress(sPart)
   
   sSql = "SELECT     ViitTable.VITCOST, VihdTable.VINO " _
          & "FROM         PoitTable INNER JOIN " _
          & "ViitTable ON PoitTable.PINUMBER = ViitTable.VITPO AND PoitTable.PIRELEASE = ViitTable.VITPORELEASE AND " _
          & "PoitTable.PIITEM = ViitTable.VITPOITEM AND PoitTable.PIREV = ViitTable.VITPOITEMREV INNER JOIN " _
          & "VihdTable ON ViitTable.VITNO = VihdTable.VINO " _
          & "WHERE     (PoitTable.PIPART = '" & sPart & "') AND (VihdTable.VIDATE = " _
          & "(SELECT     MAX(VIDATE) " _
          & "FROM          PoitTable INNER JOIN " _
          & "ViitTable ON PINUMBER = VITPO AND PIRELEASE = VITPORELEASE AND PIITEM = VITPOITEM AND PIREV = VITPOITEMREV INNER JOIN " _
          & "VihdTable ON VITNO = VINO " _
          & "WHERE      (PIPART = '" & sPart & "'))) "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst)
   If bSqlRows Then
      With rdoCst
         cCost = CCur(.Fields(0))
         sInvoice = "" & Trim(.Fields(1))
         .Cancel
      End With
   End If
   Set rdoCst = Nothing
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub

Private Sub txtFrm_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtFrm_LostFocus()
   If Trim(txtFrm) <> "" Then
      txtFrm = CheckDate(txtFrm)
      
   End If
End Sub

Private Sub txtThr_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtThr_LostFocus()
   If Trim(txtThr) <> "" Then
      txtThr = CheckDate(txtThr)
      lblFnd = NumberOfParts
      If Val(lblFnd) > 0 Then
         cmdUpd.enabled = True
         optRpt.enabled = True
         optStd.enabled = True
         optPrp.enabled = True
      Else
         cmdUpd.enabled = False
         optRpt.enabled = False
         optStd.enabled = False
         optPrp.enabled = False
      End If
   End If
End Sub

Private Function NumberOfParts() As Integer
   Dim RdoPrt As ADODB.Recordset
   sProcName = "numberof"
   sSql = "SELECT COUNT(DISTINCT PIPART) FROM VihdTable INNER JOIN " _
          & "ViitTable ON VINO = VITNO INNER JOIN PoitTable ON VITPO = PINUMBER " _
          & "AND VITPORELEASE = PIRELEASE AND VITPOITEM = PIITEM And VITPOITEMREV " _
          & "=PIREV WHERE VIDATE >= '" & txtFrm & "' AND VIDATE <= '" & txtThr & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
   If bSqlRows Then
      With RdoPrt
         NumberOfParts = .Fields(0)
         .Cancel
      End With
   End If
   Set RdoPrt = Nothing
End Function

Private Sub BatchUpdate()
   Dim RdoPrt As ADODB.Recordset
   Dim sInvoice As String
   Dim cCost As Currency
   Dim i As Single
   Dim K As Integer
   MouseCursor 13
   sProcName = "batchup"
   sSql = "SELECT DISTINCT PIPART FROM VihdTable INNER JOIN " _
          & "ViitTable ON VINO = VITNO INNER JOIN PoitTable ON VITPO = PINUMBER " _
          & "AND VITPORELEASE = PIRELEASE AND VITPOITEM = PIITEM And VITPOITEMREV " _
          & "=PIREV WHERE VIDATE >= '" & txtFrm & "' AND VIDATE <= '" & txtThr & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
   If bSqlRows Then
      prg1.max = 100
      prg1.Value = 0
      prg1.Visible = True
      lblRec.Visible = True
      lblRec = lblFnd
      lblCount.Visible = True
      lblCount = 0
      z1(28).Visible = True
      z1(29).Visible = True
      DoEvents
      i = 100 / lblFnd
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      With RdoPrt
         While Not .EOF
            prg1.Value = prg1.Value + i
            K = K + 1
            lblCount = K
            lblCount.Refresh
            LastInvoiceCost .Fields(0), cCost, sInvoice
            If sInvoice <> "" Then
               UpdateCost .Fields(0), cCost, optPrp, optStd
            End If
            sProcName = "batchup"
            .MoveNext
         Wend
         .Cancel
      End With
      clsADOCon.CommitTrans
   Else
      sMsg = "No Invoiced Parts Found."
      MsgBox sMsg, vbInformation, Caption
   End If
   Set RdoPrt = Nothing
   prg1.Visible = False
   prg1.Value = 0
   lblRec.Visible = False
   lblRec = ""
   lblCount.Visible = False
   lblCount = ""
   z1(28).Visible = False
   z1(29).Visible = False
   
   If optRpt Then
      optDis = True
      'PrintReport
   End If
   MouseCursor 0
   Exit Sub
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport()
   Dim sCustomReport As String
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("finsc02")
   sSql = "{VihdTable.VIDTRECD} >=#" & txtFrm & _
          "# AND {VihdTable.VIDTRECD} <= #" & txtThr & "#"
   MdiSect.crw.SelectionFormula = sSql
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
