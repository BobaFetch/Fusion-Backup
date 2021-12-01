VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form RoutRTe01b 
   Caption         =   "Routing Operations"
   ClientHeight    =   10110
   ClientLeft      =   1845
   ClientTop       =   1650
   ClientWidth     =   11895
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10110
   ScaleWidth      =   11895
   Begin VB.ComboBox cboFillPn 
      Height          =   315
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   48
      Tag             =   "dont auto select"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   6360
      Picture         =   "RoutRTe01b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Print The Report"
      Top             =   9600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdCpyCom 
      Caption         =   "&Copy Op Cmts"
      Height          =   555
      Left            =   8520
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "Show the Routing"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTe01b.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox lblLst 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      Top             =   9600
      Width           =   3075
   End
   Begin VB.TextBox lblCurList 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      ToolTipText     =   "Click To View The Tool List"
      Top             =   9240
      Width           =   3075
   End
   Begin VB.CommandButton cmdLst 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10320
      Picture         =   "RoutRTe01b.frx":0938
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Previous Entry"
      Top             =   8040
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdNxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10845
      Picture         =   "RoutRTe01b.frx":0A6E
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Next Entry"
      Top             =   8040
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.TextBox txtSvc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   12
      Tag             =   "1"
      ToolTipText     =   "Unit Price"
      Top             =   8880
      Width           =   945
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "RoutRTe01b.frx":0BA4
      DownPicture     =   "RoutRTe01b.frx":1516
      Height          =   350
      Left            =   10080
      Picture         =   "RoutRTe01b.frx":1E88
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Standard Comments"
      Top             =   4680
      Width           =   350
   End
   Begin VB.CheckBox optLib 
      Caption         =   "Library"
      Height          =   375
      Left            =   240
      TabIndex        =   37
      Top             =   11520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame z2 
      Height          =   2355
      Index           =   0
      Left            =   10440
      TabIndex        =   31
      Top             =   720
      Width           =   1005
      Begin VB.CommandButton cmdShw 
         Caption         =   "&Show"
         Height          =   315
         Left            =   90
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Show the Routing"
         Top             =   1980
         Width           =   850
      End
      Begin VB.CommandButton cmdFil 
         Caption         =   "&Library"
         Height          =   315
         Left            =   90
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Fill From Library"
         Top             =   1620
         Width           =   850
      End
      Begin VB.CommandButton cmdAut 
         Caption         =   "&Auto No"
         Height          =   315
         Left            =   90
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Renumber Operations"
         Top             =   1260
         Width           =   850
      End
      Begin VB.CommandButton cmdOrd 
         Caption         =   "&Reorder"
         Height          =   315
         Left            =   90
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Resort Operations"
         Top             =   900
         Width           =   850
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   90
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Delete This Operation"
         Top             =   540
         Width           =   850
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   315
         Left            =   90
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "New Operation"
         Top             =   180
         Width           =   850
      End
   End
   Begin VB.ComboBox cmbPrt 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1140
      TabIndex        =   11
      Tag             =   "3"
      ToolTipText     =   "Select Service Part From List"
      Top             =   8460
      Width           =   3345
   End
   Begin VB.CheckBox optSrv 
      Alignment       =   1  'Right Justify
      Caption         =   "Service Op?"
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Top             =   4320
      Width           =   1300
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "Pick Op?"
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Top             =   3960
      Width           =   1300
   End
   Begin VB.ComboBox cmbJmp 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "8"
      ToolTipText     =   "Jump To Operation"
      Top             =   11640
      Width           =   2445
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   10440
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   1005
   End
   Begin VB.TextBox txtCmt 
      Height          =   3585
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Tag             =   "9"
      Text            =   "RoutRTe01b.frx":248A
      ToolTipText     =   "Comment (5120 Chars Max)"
      Top             =   4680
      Width           =   9015
   End
   Begin VB.TextBox txtUnt 
      Height          =   285
      Left            =   3120
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "Unit Or Cycle Hours"
      Top             =   4320
      Width           =   825
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Left            =   1020
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Setup Hours"
      Top             =   4320
      Width           =   825
   End
   Begin VB.TextBox txtMdy 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Move Hours"
      Top             =   3960
      Width           =   825
   End
   Begin VB.TextBox txtQdy 
      Height          =   285
      Left            =   1020
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Hours In Queue"
      Top             =   3960
      Width           =   825
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Matching Work Center To Shop"
      Top             =   3600
      Width           =   1815
   End
   Begin VB.ComboBox cmbShp 
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Shop"
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtOpn 
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Operation (000 Format)"
      Top             =   3600
      Width           =   555
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   9240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   10110
      FormDesignWidth =   11895
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Click To Select Or Scroll And Press Enter"
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
   End
   Begin VB.Label lblFiller 
      Caption         =   "Filler Material"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   49
      Top             =   3330
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool List"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   44
      Top             =   9240
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Unit Cost"
      Height          =   375
      Index           =   7
      Left            =   4440
      TabIndex        =   39
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   36
      ToolTipText     =   "Routing Date As 08/08/97,08 08 97 or 08-08-97"
      Top             =   5940
      Width           =   615
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6000
      TabIndex        =   35
      Top             =   8520
      Width           =   405
   End
   Begin VB.Label lblUpd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5760
      TabIndex        =   34
      Top             =   11760
      Width           =   1575
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1140
      TabIndex        =   33
      Top             =   8820
      Width           =   3135
   End
   Begin VB.Label lblRout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5760
      TabIndex        =   30
      Top             =   11640
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Part No"
      Height          =   375
      Index           =   11
      Left            =   255
      TabIndex        =   29
      Top             =   8460
      Width           =   645
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   285
      Index           =   10
      Left            =   135
      TabIndex        =   28
      Top             =   4680
      Width           =   1320
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit/Cy Hrs"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   27
      Top             =   4320
      Width           =   1005
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Hrs"
      Height          =   285
      Index           =   8
      Left            =   135
      TabIndex        =   26
      Top             =   4320
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Move Hrs"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   25
      Top             =   3960
      Width           =   885
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Queue Hrs"
      Height          =   285
      Index           =   5
      Left            =   135
      TabIndex        =   24
      Top             =   3960
      Width           =   1035
   End
   Begin VB.Label z1 
      Caption         =   "Work Center                     "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   23
      Top             =   3330
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop                               "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1020
      TabIndex        =   22
      Top             =   3330
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Op No "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   135
      TabIndex        =   21
      Top             =   3330
      Width           =   675
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jump"
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   20
      Top             =   11640
      Width           =   705
   End
End
Attribute VB_Name = "RoutRTe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'4/7/04 New selection grid
'7/8/04 Smoothed grid scrolling
'5/20/05 Fixed CrLf in Grid comments
'11/2/05 Corrected initial Time Format
'7/20/06 Revised OPNO/Grid (change the OPNO and RtpcTable)
'10/26/06 Added Index to Grid_Click to align Next/Last buttons.
'2/5/07 Changed Grid <> keys to .TopRow/refined FillWorkCenters 7.2.0
'3/7/07 Fixed the Grid Header 7.2.2
Option Explicit
Dim AdoStm As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim ADOParameter2 As ADODB.Parameter

Dim bCancel As Byte
Dim bFromGrid As Byte
Dim bNewOp As Byte
Dim bOnLoad As Byte

Dim iIndex As Integer
Dim iOldOpn As Integer
Dim iTotalOps As Integer
Dim iCurrentOp As Integer

Dim sOldShop As String
Dim sOldCenter As String
Dim sComments As String

Dim iOperations(300) As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   ES_TimeFormat = GetTimeFormat()
   txtSvc = "0.000"
   lblCurList.BackColor = BackColor
   lblLst.BackColor = BackColor
   
End Sub


Private Sub cboFillPn_LostFocus()
   UpdateOp    'BBS Added for Ticket #43840 (2/3/2011)
End Sub

Private Sub cmbJmp_Click()
   If Left(cmbJmp, 3) <> txtOpn Then
      UpdateOp
      iIndex = cmbJmp.ListIndex + 1
      txtOpn = Left(cmbJmp, 3)
      GetThisOp
   End If
   
End Sub

Private Sub cmbPrt_Click()
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   FindPart
   If cmbPrt <> "NONE" Or cmbPrt <> "" Then optSrv.Value = vbChecked
   
End Sub

Private Sub cmbShp_Click()
   If sOldShop <> cmbShp Then FillWorkCenters
   Grd.Col = 1
   If iIndex > 0 Then Grd.Text = cmbShp
   'GetOperations
End Sub



Private Sub cmbShp_LostFocus()
   cmbShp = CheckLen(cmbShp, 12)
   
   UpdateOp    'BBS Added for Ticket #43840 (2/3/2011)

   If sOldShop <> cmbShp Then FillWorkCenters
   cmbJmp.Enabled = True
   z2(0).Enabled = True
   'Grd.Col = 1
   'Grd.Text = cmbShp
   'GetOperations
End Sub

Private Sub cmbWcn_Click()
   FindCenter bNewOp
   Grd.Col = 2
   Grd.Text = cmbWcn
   'GetOperations
End Sub

Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 12)
   If cmbWcn <> sOldCenter And bNewOp = 1 Then
      FindCenter 1
   Else
      FindCenter bNewOp
   End If
   sOldCenter = cmbWcn
   UpdateOp    'BBS Added for Ticket #43840 (2/3/2011)
'   Grd.Col = 2
'   Grd.Text = cmbWcn
   'GetOperations
   bNewOp = 0
   cmbJmp.Enabled = True
   z2(0).Enabled = True
   
End Sub


Private Sub FindCenter(bNewOne As Byte)
   Dim AdoWcn As ADODB.Recordset
   Dim sCenter As String
   Dim sShop As String
   
   sCenter = Compress(cmbWcn)
   On Error Resume Next
   If Len(sCenter) > 0 Then
      If bNewOne Then
         sShop = cmbShp
         sSql = "SELECT WCNREF,WCNNUM,WCNQHRS,WCNMHRS,WCNSUHRS," & _
                    " WCNUNITHRS,WCNSERVICE FROM WcntTable " & _
                 " WHERE WCNREF='" & sCenter & "'"
         If (sShop <> "") Then
           sSql = sSql & " AND WCNSHOP = '" & sShop & "'"
         End If
      
      Else
         sSql = "Qry_GetWorkCenter '" & sCenter & "'"
      End If
      
      bSqlRows = clsADOCon.GetDataSet(sSql, AdoWcn, ES_STATIC)
      If bSqlRows Then
         With AdoWcn
            cmbWcn = "" & Trim(!WCNNUM)
            If bNewOne = 1 Then
               txtMdy = Format(!WCNMHRS, ES_QuantityDataFormat)
               txtQdy = Format(!WCNQHRS, ES_QuantityDataFormat)
               If !WCNSERVICE = 0 Then
                  txtSet = Format(!WCNSUHRS, ES_QuantityDataFormat)
                  txtUnt = Format(!WCNUNITHRS, ES_TimeFormat)
               Else
                  txtSet = Format(0, ES_QuantityDataFormat)
                  txtUnt = Format(0, ES_TimeFormat)
                  optSrv.Value = vbChecked
                  cmbPrt.Enabled = True
                  txtSvc.Enabled = True
               End If
            End If
         End With
      Else
         MsgBox "Work Center Wasn't Found.", vbExclamation, Caption
         cmbWcn = ""
      End If
      On Error Resume Next
   End If
   Set AdoWcn = Nothing
   
End Sub

Private Sub cmdAut_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Auto Renumber Operations?", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      Width = Width + 10
   Else
      GetOperations
      AutoNumber
   End If
   
End Sub

Private Sub cmdCan_Click()
   RoutRTe01a.optOps.Value = vbUnchecked
   UpdateOp
   Unload Me
   
End Sub


Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtCmt.SetFocus
      SysComments.lblListIndex = 6
      SysComments.Show
      cmdComments = False
   End If
   
End Sub

Private Sub cmdCpyCom_Click()
   RoutRTf08a.Show
End Sub

Private Sub cmdDel_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Delete Operation " & txtOpn & "?", ES_NOQUESTION, Caption)
   If bResponse = vbNo Then
      Width = Width + 10
      Exit Sub
   End If
   On Error GoTo DiaErr1
   sSql = "DELETE FROM RtopTable WHERE OPREF='" & lblRout & "' AND OPNO=" & Val(txtOpn) & " "
   clsADOCon.ExecuteSql sSql
   SysMsg "Operation Deleted", True, Me
   GetOperations
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   MsgBox CurrError.Description & vbCrLf & "Couldn't Delete Op.", vbExclamation, Caption
   
End Sub



Private Sub cmdFil_Click()
   MouseCursor 13
   bNewOp = 0
   RoutRTe01e.lblGridRow = Grd.Row
   optLib.Value = vbChecked
   RoutRTe01e.Show
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3170
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdLst_Click()
   Dim A As Integer
   bNewOp = 0
   UpdateOp
   iIndex = iIndex - 1
   If iIndex <= 1 Then iIndex = 1
   txtOpn = Format(iOperations(iIndex), "000")
   On Error Resume Next
   Grd.Col = 0
   Grd.Row = iIndex
   GetThisOp
   If Grd.Row > 6 Then Grd.TopRow = Grd.TopRow - 1
   
End Sub

Private Sub cmdNew_Click()
   UpdateOp
   AddOperation
   
End Sub

Private Sub cmdNxt_Click()
   Dim A As Integer
   bNewOp = 0
   UpdateOp
   iIndex = iIndex + 1
   If iIndex >= iTotalOps Then iIndex = iTotalOps
   txtOpn = Format(iOperations(iIndex), "000")
   On Error Resume Next
   Grd.Col = 0
   Grd.Row = iIndex
   GetThisOp
   If Grd.Row > 7 Then Grd.TopRow = Grd.TopRow + 1
   '    If Grd.Row < 6 Then
   '        Grd.TopRow = 1
   '    Else
   '        a = Grd.Row Mod 4
   '        If a = 2 Then Grd.TopRow = Grd.Row - 1
   '    End If
   
End Sub

Private Sub cmdOrd_Click()
   GetOperations
   
End Sub



Private Sub cmdShw_Click()
   MouseCursor 13
   UpdateOp
   PrintReport
   'DONE: Change to CR11
'   Dim A As Integer
'   Dim b As Integer
'   Dim iScreenHeight As Integer
'   Dim iScreenWidth As Integer
'
'   'clear any report variables
'   For b = 0 To 60
'      MDISect.Crw.Formulas(b) = ""
'      MDISect.Crw.SectionFormat(b) = ""
'      MDISect.Crw.SectionFont(b) = ""
'   Next
'   GetCrystalConnect
'   A = Screen.TwipsPerPixelX
'   b = Screen.TwipsPerPixelY
'   MDISect.Crw.WindowTop = 2050 / b
'   MDISect.Crw.WindowHeight = (MDISect.Height / b) - (2550 / b)
'   MDISect.Crw.WindowLeft = 2170 / A
'   MDISect.Crw.WindowWidth = (MDISect.Width / A) - (2400 / A)
'
'   On Error GoTo DiaErr1
'   MDISect.Crw.ReportFileName = sReportPath & "engrt01.rpt"
'   MDISect.Crw.SelectionFormula = "{RthdTable.RTREF}='" & lblRout & "' "
'   MDISect.Crw.WindowTitle = Caption
'   MDISect.Crw.Action = 0
'   MDISect.Crw.PageZoom (80)
'   MouseCursor 0
'   If bNoCrystal Then
'      SendKeys "% R", True
'      bNoCrystal = False
'   End If
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   MouseCursor 0
   MsgBox CurrError.Description & vbCrLf & "Couldn't Show Report.", vbExclamation, Caption
   
End Sub

Private Sub PrintReport()
   Dim sRout As String
   sRout = Compress(lblRout)
   
   optPrn.Value = False
   MouseCursor 13
   On Error GoTo DiaErr1
   
    Dim sCustomReport As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    Dim strIncludes As String
    Dim strRequestBy As String
   
    sCustomReport = GetCustomReport("engrt01")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    'aFormulaName.Add "ShowPartDesc"
    aFormulaName.Add "ShowCmt"
    aFormulaName.Add "ShowToolList"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")

    strRequestBy = "'Requested By: " & sInitials & "'"
    aFormulaValue.Add CStr(strRequestBy)
   
    'aFormulaValue.Add CStr("1")
    aFormulaValue.Add CStr("1")
    aFormulaValue.Add CStr("1")
   
   sSql = "{RthdTable.RTREF}='" & sRout & "' "
   
    ' Set Formula values
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.SetReportDistinctRecords True
    ' set the report Selection
    cCRViewer.SetReportSelectionFormula (sSql)
    'cCRViewer.CRViewerSize Me
    
    ' Set report parameter
    cCRViewer.SetDbTableConnection


    cCRViewer.OpenCrystalReportObject Me, aFormulaName

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aRptParaType
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


Private Sub Form_Activate()
   If bOnLoad Then
      cmdComments.Enabled = True
      FillCombos
      If (EnabledSelLot = 1) Then
       lblFiller.Visible = True
       cboFillPn.Visible = True
       
       FillFillerMaterial
      End If
      
      GetOperations
      
      
      bOnLoad = 0
      
   End If
   If optLib.Value = vbChecked Then
      Unload RoutRTe01e
      optLib.Value = vbUnchecked
   End If
   MouseCursor 0
   
End Sub
Private Sub FillFillerMaterial()
   sSql = "select distinct PARTREF from partTable where Paprodcode = 'FILL'"
   Dim rs As ADODB.Recordset
   cboFillPn.Clear
   AddComboStr cboFillPn.hwnd, ""
   Set rs = clsADOCon.GetRecordSet(sSql, ES_STATIC)
   If Not rs.BOF And Not rs.EOF Then
      With rs
         Do Until .EOF
            AddComboStr cboFillPn.hwnd, "" & Trim(.Fields(0))
            .MoveNext
         Loop
         ClearResultSet rs
      End With
   End If
   If cboFillPn.ListCount <> 0 Then
      bSqlRows = 1
      cboFillPn.ListIndex = 0
   Else
      bSqlRows = 0
   End If
   sSql = ""
   Set rs = Nothing
End Sub

Public Function EnabledSelLot()
   Dim rdo As ADODB.Recordset
   Dim companyAccount As String
   
   sSql = "select ISNULL(COLOTATPOM, 0) as COLOTATPOM from ComnTable"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      EnabledSelLot = rdo!COLOTATPOM
      rdo.Close
   Else
      EnabledSelLot = 0
   End If
   
End Function

Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiEngr", "width", Me.Width
   SaveSetting "Esi2000", "EsiEngr", "height", Me.Height
   
End Sub


Private Sub GetOptions()
   Dim strWidth As String
   Dim strHeight As String
   strWidth = Trim(GetSetting("Esi2000", "EsiEngr", "width", strWidth))
   strHeight = Trim(GetSetting("Esi2000", "EsiEngr", "height", strHeight))
   Me.Top = 0
   If strWidth = "" And strHeight = "" Then
      Me.Width = 12015
      Me.Height = 11445
   Else
      Me.Width = strWidth
      Me.Height = strHeight
   End If
   ReSize1.Enabled = True
   On Error Resume Next
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   Move RoutRTe01a.Left + 400, RoutRTe01a.Top + 600
   FormatControls
   GetOptions
   sCurrForm = "Routings"
   On Error Resume Next
   cUR.CurrentShop = GetSetting("Esi2000", "Current", "Shop", cUR.CurrentShop)
   sSql = "SELECT OPREF,OPNO,OPSHOP,OPCENTER,OPSETUP,OPUNIT," _
          & "OPQHRS,OPMHRS,OPSERVICE,OPPICKOP,OPSERVPART,OPSVCUNIT," _
          & "OPTOOLLIST,OPCOMT,OPFILLREF FROM RtopTable WHERE OPREF= ? AND OPNO= ?"

   Set AdoStm = New ADODB.Command
   AdoStm.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.Size = 30
   AdoStm.Parameters.Append AdoParameter1
   
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.Type = adSmallInt
   AdoStm.Parameters.Append ADOParameter2
   
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 0
      
      .Row = 0
      .Col = 0
      .Text = "Op No"
      .ColWidth(0) = 750
      .Col = 1
      .Text = "Shop"
      .ColWidth(1) = 1500
      .Col = 2
      .Text = "Work Center"
      .ColWidth(2) = 1500
      .Col = 3
      .Text = "Comment"
      .ColWidth(3) = 6000
      .Col = 0
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   Hide
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   RoutRTe01a.Width = RoutRTe01a.Width + 20
   On Error Resume Next
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set AdoStm = Nothing
   
   Set RoutRTe01b = Nothing
   
End Sub


Private Sub grd_Click()
   UpdateOp True
   Grd.Col = 0
   'iIndex = Grd.Row
   txtOpn = Grd.Text
   GetThisOp
   
End Sub

Private Sub Grd_GotFocus()
   Grd.Col = 0
   
End Sub


Private Sub Grd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      UpdateOp True
      Grd.Col = 0
      txtOpn = Grd.Text
      GetThisOp
   End If
   
End Sub


Private Sub lblCurList_Click()
   If Trim(lblCurList) <> "" Then
      ViewTool.lblLst = lblCurList
      ViewTool.Show
   Else
      MsgBox "There Is No Tool List Assigned"
   End If
End Sub


Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optLib_Click()
   'never visible-unloads fill
   
End Sub

Private Sub optPck_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPck_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub


Private Sub optPck_LostFocus()
   cmbJmp.Enabled = True
   z2(0).Enabled = True
   
End Sub

Private Sub optSrv_Click()
   If optSrv.Value = vbChecked Then
      cmbPrt.Enabled = True
      txtSvc.Enabled = True
   Else
      cmbPrt.Enabled = False
      cmbPrt = "NONE"
      lblDsc = ""
      txtSvc.Enabled = False
      txtSvc = "0.000"
   End If
   
End Sub



Private Sub FillCombos()
   On Error GoTo DiaErr1
   sSql = "Qry_FillShops"
   LoadComboBox cmbShp
   If bSqlRows Then
      If cUR.CurrentShop <> "" Then
         cmbShp = cUR.CurrentShop
      Else
         cmbShp = cmbShp.List(0)
      End If
   End If
   
   cmbPrt = "None"
   sSql = "Qry_FillRoutingPT7"
   LoadComboBox cmbPrt
   FillWorkCenters
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'10/17/03 Added error trap for AWI SQL 6.5

Private Sub FillWorkCenters()
   Dim AdoWcn As ADODB.Recordset
   Dim bByte As Byte
   Dim iList As Integer
   Dim sCurCenter As String
   sCurCenter = cmbWcn
   cmbWcn.Clear
   On Error GoTo DiaErr1
   sSql = "Qry_FillWorkCenters '" & Compress(cmbShp) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoWcn, ES_FORWARD)
   If bSqlRows Then
      With AdoWcn
         Do Until .EOF
            AddComboStr cmbWcn.hwnd, "" & Trim(!WCNNUM)
            .MoveNext
         Loop
      End With
      ClearResultSet AdoWcn
   End If
   cmbWcn = sCurCenter
   If cmbWcn.ListCount > 0 Then
      For iList = 0 To cmbWcn.ListCount - 1
         If Trim(cmbWcn.List(iList)) = Trim(cmbWcn) Then bByte = 1
      Next
      If bByte = 0 Then cmbWcn = cmbWcn.List(0)
      Grd.Col = 2
      If iIndex > 0 Then Grd.Text = cmbWcn
   End If
   sOldShop = cmbShp
   Set AdoWcn = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillworkc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetOperations()
   Dim AdoRes2 As ADODB.Recordset
   Dim iRows As Integer
   Dim sShop As String
   Dim sCenter As String
   Dim sString As String
   Erase iOperations
   cmbJmp.Clear
   Grd.Rows = 2
   iTotalOps = 0
   On Error Resume Next
   If iAutoIncr <= 0 Then iAutoIncr = 10
   sSql = "SELECT OPREF,OPNO,OPSHOP,OPCENTER,OPCOMT FROM RtopTable WHERE OPREF='" & sPassedRout & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoRes2, ES_STATIC)
   If bSqlRows Then
      With AdoRes2
         txtOpn = Format(!opNo, "000")
         Do Until .EOF
            iTotalOps = iTotalOps + 1
            iOperations(iTotalOps) = !opNo
            sComments = "" & Trim(!OPCOMT)
            sComments = TrimComment(sComments)
            cmbJmp.AddItem Format(!opNo, "000") & " " & sComments
            iRows = iRows + 1
            If iRows > 1 Then Grd.Rows = Grd.Rows + 1
            Grd.Row = iRows
            Grd.Col = 0
            Grd.Text = Format(!opNo, "000")
            Grd.Col = 1
            sShop = GetRoutShop("" & Trim(!OPSHOP))
            Grd.Text = sShop
            Grd.Col = 2
            sCenter = GetRoutCenter("" & Trim(!OPCENTER))
            Grd.Text = sCenter
            Grd.Col = 3
            
            sString = "" & Trim(Left(!OPCOMT, 20))
            sString = Replace(sString, vbCrLf, " ")
            Grd.Text = sString
            .MoveNext
         Loop
         ClearResultSet AdoRes2
      End With
      Grd.Row = 1
      Grd.Col = 0
      cmbJmp.ListIndex = 0
      iIndex = 1
      GetThisOp
   Else
      AdoRes2.Close
      Grd.Row = 1
      Grd.Col = 0
      txtOpn = Format(iAutoIncr, "000")
      Grd.Text = txtOpn
      Grd.Col = 1
      Grd.Text = cmbShp
      Grd.Col = 2
      Grd.Text = cmbWcn
      Grd.Col = 0
      sSql = "INSERT INTO RtopTable (OPREF,OPNO,OPSHOP,OPCENTER) " _
             & "VALUES('" & sPassedRout & "'," & Val(txtOpn) _
             & ",'" & Compress(cmbShp) & "','" & Compress(cmbWcn) & "')"
      clsADOCon.ExecuteSql sSql
      txtQdy = "0.000"
      txtMdy = "0.000"
      txtSet = "0.000"
      txtUnt = Format(0, ES_TimeFormat)
      cmbJmp = txtOpn
      cmbJmp.AddItem txtOpn
      iTotalOps = 1
      iIndex = 1
      iOperations(1) = Val(txtOpn)
      bNewOp = 1
   End If
   Set AdoRes2 = Nothing
   
End Sub


Private Function TrimComment(sComment As String)
   Dim n As String
   On Error GoTo DiaErr1
   If Len(sComment) > 0 Then
      sComment = Replace(sComment, vbCrLf, " ")
      sComment = Replace(sComment, vbCr, " ")
      sComment = Replace(sComment, vbLf, " ")
      If Len(sComment) > 20 Then sComment = Left(sComment, 20)
   End If
   TrimComment = sComment
   Exit Function
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   On Error GoTo 0
   
End Function


Private Sub UpdateOp(Optional HideNotice As Boolean)
   MouseCursor 11
   Dim sShop As String
   Dim sCenter As String
   Dim sService As String
   'Dim sSavecomments As String
   Dim sFillerPart As String
   
   sShop = Compress(cmbShp)
   sCenter = Compress(cmbWcn)
   sFillerPart = Compress(cboFillPn)
   
   If sShop = "" Then
      MsgBox "Requires A Valid Shop.", vbExclamation, Caption
      MouseCursor vbNormal
      Exit Sub
   End If
   If sCenter = "" Then
      MsgBox "Requires A Valid Work Center.", vbExclamation, Caption
      MouseCursor vbNormal
      Exit Sub
   End If
   If optSrv.Value = vbChecked Then
      sService = Compress(cmbPrt)
   Else
      sService = ""
   End If
   If Not HideNotice Then lblUpd = "Updating."
   lblUpd.Refresh
   'sSavecomments = "" & Trim(txtCmt)
   'sSavecomments = SqlString(txtCmt)
   On Error Resume Next
   sSql = "UPDATE RtopTable SET " _
          & "OPSHOP='" & sShop & "', " _
          & "OPCENTER='" & sCenter & "'," _
          & "OPSETUP=" & Format(Val(txtSet), ES_QuantityDataFormat) & "," _
          & "OPUNIT=" & Format(Val(txtUnt), ES_QuantityDataFormat) & "," _
          & "OPQHRS=" & Format(Val(txtQdy), ES_QuantityDataFormat) & "," _
          & "OPMHRS=" & Format(Val(txtMdy), ES_QuantityDataFormat) & "," _
          & "OPSERVICE=" & optSrv.Value & ", " _
          & "OPPICKOP=" & optPck.Value & ", " _
          & "OPSERVPART='" & sService & "'," _
          & "OPSVCUNIT=" & Val(txtSvc) & "," _
          & "OPCOMT='" & SqlString(txtCmt) & "', " _
          & "OPFILLREF='" & sFillerPart & "' " _
          & "WHERE OPREF='" & lblRout & "' AND OPNO=" & Val(txtOpn) & " "
   clsADOCon.ExecuteSql sSql
   MouseCursor vbNormal
   Sleep 100
   RoutRTe01a.GetOperationTimes
   lblUpd = ""
   lblUpd.Refresh
   
End Sub

Private Sub GetThisOp()
   Dim AdoRes2 As ADODB.Recordset
   Dim A As Integer
   On Error Resume Next
   

   AdoStm.Parameters(0).Value = sPassedRout
   AdoStm.Parameters(1).Value = Val(txtOpn)
   bSqlRows = clsADOCon.GetQuerySet(AdoRes2, AdoStm, ES_STATIC, True, 1)
   If bSqlRows Then
      With AdoRes2
         cmbShp = "" & Trim(!OPSHOP)
         FindShop cmbShp
         cmbWcn = "" & Trim(!OPCENTER)
         ' 4/20/2009 Should pass the NewOp value
         'FindCenter 0
         FindCenter bNewOp
         bFromGrid = 1
         If sOldShop <> cmbShp Then FillWorkCenters
         txtSet = Format(!OPSETUP, ES_QuantityDataFormat)
         txtUnt = Format(!OPUNIT, ES_TimeFormat)
         txtQdy = Format(!OPQHRS, ES_QuantityDataFormat)
         txtMdy = Format(!OPMHRS, ES_QuantityDataFormat)
         optSrv.Value = !OPSERVICE
         optPck.Value = !OPPICKOP
         cmbPrt = "" & Trim(!OPSERVPART)
         txtCmt = "" & Trim(!OPCOMT)
         txtSvc = Format(!OPSVCUNIT, ES_QuantityDataFormat)
         lblCurList = "" & Trim(!OPTOOLLIST)
         If lblCurList <> "" Then lblCurList = FindToolList(lblCurList, lblLst) _
            Else lblLst = ""
         If Right(txtCmt, 1) = vbLf And Right(txtCmt, 1) = vbCr Then
            If Len(txtCmt) > 1 Then
               A = Len(txtCmt)
               txtCmt = (Left$(txtCmt, A - 1))
            Else
               txtCmt = ""
            End If
         End If
         cboFillPn = "" & Trim(!OPFILLREF)
      End With
      '4/20/2009
      'bNewOp = 0
      sOldShop = cmbShp
      sOldCenter = cmbWcn
      FindPart
   End If
   Grd.Col = 0
   Grd.SetFocus
   Set AdoRes2 = Nothing
   
End Sub

Private Sub optSrv_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optSrv_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub optSrv_LostFocus()
   cmbJmp.Enabled = True
   z2(0).Enabled = True
   
End Sub

Private Sub txtCmt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub


Private Sub txtCmt_LostFocus()
   Dim sString As String
   txtCmt = CheckLen(txtCmt, 5120)
   If Len(txtCmt) Then txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   'txtCmt = ParseComment(txtCmt, False)
   z2(0).Enabled = True
   sString = Left(txtCmt, 20)
   sString = Replace(sString, vbCrLf, " ")
   Grd.Col = 3
   Grd.Text = sString
   On Error Resume Next
   sSql = "UPDATE RtopTable SET OPCOMT='" & SqlString(txtCmt) & "' " _
          & "WHERE OPREF='" & lblRout & "' AND OPNO=" & Val(txtOpn) & " "
   clsADOCon.ExecuteSql sSql
   
End Sub


Private Sub txtMdy_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtMdy_LostFocus()
   txtMdy = CheckLen(txtMdy, 7)
   txtMdy = Format(Abs(Val(txtMdy)), ES_QuantityDataFormat)
   cmbJmp.Enabled = True
   z2(0).Enabled = True
   
End Sub

Private Sub txtOpn_Click()
   iOldOpn = Val(Left(txtOpn, 3))
   
End Sub

Private Sub txtOpn_GotFocus()
   bCancel = 0
   iCurrentOp = Val(txtOpn)
   iOldOpn = Abs(Val(txtOpn))
   
End Sub


Private Sub txtOpn_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtOpn_LostFocus()
   Dim iList As Integer
   Dim bByte As Byte
   If bCancel = 1 Then Exit Sub
   txtOpn = CheckLen(txtOpn, 3)
   If Len(txtOpn) = 0 Then txtOpn = Format(iCurrentOp, "000")
   txtOpn = Format(Abs(Val(txtOpn)), "000")
   If iOldOpn <> Val(txtOpn) Then
      bByte = False
      For iList = 0 To cmbJmp.ListCount - 1
         'get op # from list item
'         Dim searchString As String
'         Dim numString As String
'         Dim j As Integer
'         searchString = cmbJmp.List(iList)
'         numString = "0"
'         For j = 1 To Len(searchString)
'            If IsNumeric(Mid(searchString, j, 1)) Then
'               numString = numString & Mid(searchString, j, 1)
'            Else
'               Exit For
'            End If
'         Next
      
         If Val(txtOpn) = Val(Replace(cmbJmp.List(iList), " ", "*")) Then
         'If Val(txtOpn) = Val(numString) Then
            If Not bNewOp Then
               MsgBox "Operation Exists.", vbInformation, Caption
               txtOpn = Format(iOldOpn, "000")
               bByte = True
            End If
            Exit For
         End If
      Next
      If Not bByte Then
         On Error GoTo DiaErr1
         sSql = "UPDATE RtopTable SET OPNO=" & txtOpn & " WHERE OPREF='" & lblRout & "' AND OPNO=" & str(iOldOpn) & ""
         clsADOCon.ExecuteSql sSql
         
         sSql = "UPDATE RtpcTable SET OPNO=" & txtOpn & " WHERE OPREF='" & lblRout & "' AND OPNO=" & str(iOldOpn) & ""
         clsADOCon.ExecuteSql sSql
         Grd.Col = 0
         Grd.Text = txtOpn
         
         
         'On Error Resume Next
         For iList = 0 To cmbJmp.ListCount - 1
            'If Val(cmbJmp.List(iList)) = iOldOpn Then cmbJmp.RemoveItem iList
            If Val(txtOpn) = Val(Replace(cmbJmp.List(iList), " ", "*")) Then
               cmbJmp.RemoveItem iList
            End If
         Next
         iOperations(iIndex) = Val(txtOpn)
         sComments = Trim(Left(txtCmt, 20))
         sComments = TrimComment(sComments)
         cmbJmp = txtOpn & " " & sComments
         cmbJmp.AddItem txtOpn & " " & sComments
         'On Error Resume Next
         cmbShp.SetFocus
      End If
   End If
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   MsgBox CurrError.Description & " Couldn't Change Operation.", vbInformation, Caption
   On Error GoTo 0
   
End Sub

Private Sub txtQdy_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtQdy_LostFocus()
   txtQdy = CheckLen(txtQdy, 7)
   txtQdy = Format(Abs(Val(txtQdy)), ES_QuantityDataFormat)
   'cmbJmp.Enabled = True
   z2(0).Enabled = True
   
End Sub

Private Sub txtSet_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtSet_LostFocus()
   txtSet = CheckLen(txtSet, 7)
   txtSet = Format(Abs(Val(txtSet)), ES_QuantityDataFormat)
   'cmbJmp.Enabled = True
   z2(0).Enabled = True
   
End Sub

Private Sub txtSvc_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub


Private Sub txtSvc_LostFocus()
   txtSvc = CheckLen(txtSvc, 9)
   txtSvc = Format(Abs(Val(txtSvc)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtUnt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtUnt_LostFocus()
   txtUnt = CheckLen(txtUnt, 8)
   txtUnt = Format(Abs(Val(txtUnt)), ES_TimeFormat)
   'cmbJmp.Enabled = True
   z2(0).Enabled = True
   
End Sub



Private Sub AddOperation()
   Dim sShop As String
   Dim sCenter As String
   Dim AdoRes2 As ADODB.Recordset
   lblUpd = "Adding Item."
   lblUpd.Refresh
   
   sShop = Compress(cmbShp)
   sCenter = Compress(cmbWcn)
   txtOpn = Format(iOperations(iTotalOps) + iAutoIncr, "000")
   
   ' On Error GoTo DiaErr1
   On Error GoTo 0
   sSql = "INSERT INTO RtopTable (OPREF,OPNO,OPSHOP,OPCENTER) " _
          & "VALUES('" & sPassedRout & "'," & Trim(txtOpn) & "," _
          & "'" & sShop & "','" & sCenter & "')"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.RowsAffected = 0 Then
      lblUpd = ""
      lblUpd.Refresh
      MsgBox "Couldn't Add The Operation", vbExclamation, Caption
      Exit Sub
   End If
   
   txtQdy = "0.000"
   txtMdy = "0.000"
   txtSet = "0.000"
   txtUnt = ES_TimeFormat
   cmbPrt = ""
   cboFillPn = ""
   'cmbJmp.Enabled = False
   z2(0).Enabled = False
   
   optSrv.Value = vbUnchecked
   optPck.Value = vbUnchecked
   cmbPrt.Enabled = False
   txtSvc.Enabled = False
   'cmbJmp.AddItem txtOpn
   Grd.Rows = Grd.Rows + 1
   Grd.Row = Grd.Rows - 1
   If Grd.Row > 5 Then Grd.TopRow = Grd.Row - 4
   Grd.Col = 0
   Grd.Text = txtOpn
   iTotalOps = iTotalOps + 1
   iOperations(iTotalOps) = Val(txtOpn)
   bNewOp = 1
   iIndex = iTotalOps
   lblUpd = ""
   lblUpd.Refresh
   SysMsg "Operation " & txtOpn & " Added.", True, Me
   GetThisOp
   txtOpn.SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "addoperati"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub AutoNumber()
   Dim iList As Integer
   Dim iNewNumber As Integer
   Dim iNewOps(300, 3) As Integer
   
   On Error GoTo DiaErr1
   cmdCan.Enabled = False
   MouseCursor 11
   For iList = 0 To cmbJmp.ListCount - 1
      iNewOps(iList, 0) = (iList + 1052)
      iNewOps(iList, 1) = Val(Left(cmbJmp.List(iList), 3))
   Next
   clsADOCon.BeginTrans
   For iList = 0 To cmbJmp.ListCount - 1
      sSql = "UPDATE RtopTable SET OPNO=" & str(iNewOps(iList, 0)) & " WHERE OPREF='" & sPassedRout & "' AND OPNO=" & str(iNewOps(iList, 1))
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE RtpcTable SET OPNO=" & str(iNewOps(iList, 0)) & " WHERE OPREF='" & sPassedRout & "' AND OPNO=" & str(iNewOps(iList, 1))
      clsADOCon.ExecuteSql sSql
   Next
   clsADOCon.CommitTrans
   
   clsADOCon.BeginTrans
   iNewNumber = 0
   For iList = 0 To cmbJmp.ListCount - 1
      iNewNumber = iNewNumber + iAutoIncr
      sSql = "UPDATE RtopTable SET OPNO=" & str(iNewNumber) & " WHERE OPREF='" & sPassedRout & "' AND OPNO=" & str(iNewOps(iList, 0))
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE RtpcTable SET OPNO=" & str(iNewNumber) & " WHERE OPREF='" & sPassedRout & "' AND OPNO=" & str(iNewOps(iList, 0))
      clsADOCon.ExecuteSql sSql
   Next
   clsADOCon.CommitTrans
   MouseCursor 0
   SysMsg "Auto Numbering Complete.", True, Me
   GetOperations
   MouseCursor 0
   On Error Resume Next
   cmdCan.Enabled = True
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   On Error Resume Next
   MouseCursor 0
   clsADOCon.RollbackTrans
   cmdCan.Enabled = True
   MsgBox CurrError.Description & " Can't Reorganize Operations.", vbExclamation, Caption
   
End Sub

Private Sub z1_Click(Index As Integer)
   On Error Resume Next
   cmbJmp.SetFocus
   
End Sub



Private Sub NextOp()
   Dim iRow As Integer
   bCancel = 1
   iRow = Grd.Row
   iRow = iRow + 1
   If iRow > Grd.Rows - 1 Then iRow = Grd.Rows - 1
   Grd.Row = iRow
   Grd.Col = 0
   txtOpn = Grd.Text
   GetThisOp
   
End Sub

Private Sub LastOp()
   Dim iRow As Integer
   bCancel = 1
   iRow = Grd.Row
   iRow = iRow - 1
   If iRow < 1 Then iRow = 1
   Grd.Row = iRow
   Grd.Col = 0
   txtOpn = Grd.Text
   GetThisOp
   
End Sub
