VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form ShopSHe02c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise MO Information, Comments"
   ClientHeight    =   6435
   ClientLeft      =   1845
   ClientTop       =   1635
   ClientWidth     =   8805
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6435
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdStatCode 
      DisabledPicture =   "ShopSHe02c.frx":0000
      DownPicture     =   "ShopSHe02c.frx":0972
      Height          =   350
      Left            =   8280
      MaskColor       =   &H8000000F&
      Picture         =   "ShopSHe02c.frx":12E4
      Style           =   1  'Graphical
      TabIndex        =   46
      TabStop         =   0   'False
      ToolTipText     =   "Add Internal Status Code"
      Top             =   2820
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe02c.frx":1773
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
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      Top             =   6000
      Width           =   3105
   End
   Begin VB.TextBox lblCurList 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1020
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "2"
      Text            =   " "
      ToolTipText     =   "Click To View The Tool List"
      Top             =   5640
      Width           =   3105
   End
   Begin VB.CommandButton cmdLst 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      Picture         =   "ShopSHe02c.frx":1F21
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Previous Entry"
      Top             =   4860
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdNxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      Picture         =   "ShopSHe02c.frx":2057
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Next Entry"
      Top             =   4860
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "ShopSHe02c.frx":218D
      DownPicture     =   "ShopSHe02c.frx":2AFF
      Height          =   350
      Left            =   7920
      Picture         =   "ShopSHe02c.frx":3471
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Standard Comments"
      Top             =   2820
      Width           =   350
   End
   Begin VB.CheckBox optLib 
      Caption         =   "Library"
      Height          =   255
      Left            =   5400
      TabIndex        =   38
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optCom 
      Caption         =   "Completed"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4780
      TabIndex        =   37
      Top             =   2160
      Width           =   1200
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Left            =   6000
      TabIndex        =   35
      Tag             =   "1"
      Top             =   4860
      Width           =   975
   End
   Begin VB.Frame z2 
      Height          =   2055
      Index           =   0
      Left            =   7740
      TabIndex        =   30
      Top             =   480
      Width           =   1005
      Begin VB.CommandButton cmdFil 
         Caption         =   "&Library"
         Height          =   315
         Left            =   90
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Fill From Library"
         Top             =   1260
         Width           =   850
      End
      Begin VB.CommandButton cmdAut 
         Caption         =   "&Auto No"
         Height          =   315
         Left            =   90
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Renumber Operations"
         Top             =   1620
         Width           =   850
      End
      Begin VB.CommandButton cmdOrd 
         Caption         =   "&Reorder"
         Height          =   315
         Left            =   90
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Resort Operations"
         Top             =   900
         Width           =   850
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   90
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Delete This Operation"
         Top             =   540
         Width           =   850
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   315
         Left            =   90
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "New Operation"
         Top             =   180
         Width           =   850
      End
   End
   Begin VB.ComboBox cmbPrt 
      Enabled         =   0   'False
      Height          =   288
      Left            =   1020
      TabIndex        =   11
      Tag             =   "3"
      ToolTipText     =   "Select Service Part From List"
      Top             =   4860
      Width           =   3345
   End
   Begin VB.CheckBox optSrv 
      Alignment       =   1  'Right Justify
      Caption         =   "Service Op?"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3780
      TabIndex        =   9
      Top             =   2880
      Width           =   1185
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "Pick Op?"
      Height          =   285
      Left            =   3780
      TabIndex        =   6
      Top             =   2520
      Width           =   1185
   End
   Begin VB.ComboBox cmbJmp 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   1500
      Sorted          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "8"
      ToolTipText     =   "Jump To Operation"
      Top             =   6840
      Width           =   2445
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7725
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   45
      Width           =   1005
   End
   Begin VB.TextBox txtCmt 
      Height          =   1545
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Tag             =   "9"
      Text            =   "ShopSHe02c.frx":3A73
      Top             =   3240
      Width           =   7575
   End
   Begin VB.TextBox txtUnt 
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "Unit Or Cycle Time"
      Top             =   2880
      Width           =   825
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Left            =   1020
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Setup Hours"
      Top             =   2880
      Width           =   825
   End
   Begin VB.TextBox txtMdy 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Hours To Move"
      Top             =   2520
      Width           =   825
   End
   Begin VB.TextBox txtQdy 
      Height          =   285
      Left            =   1020
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Hours In Queue"
      Top             =   2520
      Width           =   825
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   288
      Left            =   2880
      TabIndex        =   3
      Tag             =   "3"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   1020
      TabIndex        =   2
      Tag             =   "8"
      ToolTipText     =   "Select Shop From List"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtOpn 
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Tag             =   "1"
      Top             =   2160
      Width           =   555
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   5760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6435
      FormDesignWidth =   8805
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Click To Select Or Scroll And Press Enter"
      Top             =   0
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
   End
   Begin VB.Label lblRow 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6360
      TabIndex        =   47
      Top             =   2580
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool List"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   12
      Left            =   120
      TabIndex        =   44
      Top             =   5640
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sv Unit Cost"
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   36
      Top             =   4860
      Width           =   1245
   End
   Begin VB.Label lblMon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MO Number"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1680
      TabIndex        =   34
      Top             =   6600
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.Label lblRun 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Run"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   4080
      TabIndex        =   33
      Top             =   6600
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblUpd 
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   120
      TabIndex        =   32
      Top             =   5280
      Width           =   1572
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1020
      TabIndex        =   31
      Top             =   5160
      Width           =   3108
   End
   Begin VB.Label lblRout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Routing"
      ForeColor       =   &H80000008&
      Height          =   288
      Left            =   120
      TabIndex        =   29
      Top             =   6480
      Visible         =   0   'False
      Width           =   1368
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Part No"
      Height          =   372
      Index           =   11
      Left            =   252
      TabIndex        =   28
      Top             =   4860
      Width           =   648
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   288
      Index           =   10
      Left            =   120
      TabIndex        =   27
      Top             =   3120
      Width           =   1320
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Hrs"
      Height          =   288
      Index           =   9
      Left            =   2040
      TabIndex        =   26
      Top             =   2880
      Width           =   888
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Hrs"
      Height          =   288
      Index           =   8
      Left            =   132
      TabIndex        =   25
      Top             =   2880
      Width           =   1188
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Move Hrs"
      Height          =   288
      Index           =   6
      Left            =   2040
      TabIndex        =   24
      Top             =   2520
      Width           =   1008
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Queue Hrs"
      Height          =   288
      Index           =   5
      Left            =   132
      TabIndex        =   23
      Top             =   2520
      Width           =   1152
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
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
      Height          =   288
      Index           =   3
      Left            =   2880
      TabIndex        =   22
      Top             =   1896
      Width           =   1812
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
      Height          =   288
      Index           =   2
      Left            =   1020
      TabIndex        =   21
      Top             =   1896
      Width           =   1812
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
      Height          =   288
      Index           =   1
      Left            =   132
      TabIndex        =   20
      Top             =   1896
      Width           =   672
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jump"
      Height          =   288
      Index           =   0
      Left            =   840
      TabIndex        =   19
      Top             =   6720
      Width           =   708
   End
End
Attribute VB_Name = "ShopSHe02c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'9/23/03 Added Reschedule PopUp
'7/8/04 Smoothed grid scrolling
'3/7/05 Added OPPICKOP to UpdateOp
'5/20/05 Fixed CrLf in Grid comments
'11/02/05 Corrected initial Time Format
'2/5/07 Changed Grid <> keys to .TopRow/Refined FillWorkCenters 7.2.4
Option Explicit
Dim AdoStm As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim ADOParameter2 As ADODB.Parameter
Dim AdoParameter3 As ADODB.Parameter


Dim bNewOp As Byte
Dim bOnLoad As Byte
Dim bShowBox As Byte
Dim bTimeChg As Byte

Dim iCurrentOp As Integer
Dim iIndex As Integer
Dim iOldOpn As Integer
Dim iTotalOps As Integer

Dim sOldShop As String
Dim sOldCenter As String
Dim sComments As String
Dim sMoNumber As String

Dim iOperations(300) As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub NextOp()
   Dim iRow As Integer
   iRow = Grd.row
   iRow = iRow + 1
   If iRow > Grd.Rows - 1 Then iRow = Grd.Rows - 1
   Grd.row = iRow
   Grd.Col = 0
   txtOpn = Grd.Text
   GetThisOp
   
End Sub

Private Sub LastOp()
   Dim iRow As Integer
   iRow = Grd.row
   iRow = iRow - 1
   If iRow < 1 Then iRow = 1
   Grd.row = iRow
   Grd.Col = 0
   txtOpn = Grd.Text
   GetThisOp
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   ES_TimeFormat = GetTimeFormat()
   lblCurList.BackColor = BackColor
   lblLst.BackColor = BackColor
   
End Sub

Private Sub cmbJmp_Click()
   If cmbJmp <> txtOpn Then
      iIndex = cmbJmp.ListIndex + 1
      txtOpn = Left(cmbJmp, 3)
      GetThisOp
   End If
   
End Sub

Private Sub cmbPrt_Click()
   If Trim(cmbPrt) <> "NONE" Then cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) = 0 Then cmbPrt = "NONE"
   If Trim(cmbPrt) <> "NONE" Then
      optSrv.Value = vbChecked
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   End If
   
   
End Sub

Private Sub cmbShp_Click()
   If cmbShp <> sOldShop Then FillWorkCenters
   Grd.Col = 1
   Grd.Text = cmbShp
   
End Sub



Private Sub cmbShp_LostFocus()
   cmbShp = CheckLen(cmbShp, 30)
   If sOldShop <> cmbShp Then FillWorkCenters
   cmbJmp.Enabled = True
   z2(0).Enabled = True
'   Grd.Col = 1
'   Grd.Text = cmbShp
   Grd.TextMatrix(CInt(lblRow), 1) = cmbShp

   
End Sub

Private Sub cmbWcn_Click()
   FindCenter cmbWcn, bNewOp
   Grd.Col = 2
   Grd.Text = cmbWcn
   
End Sub

Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 12)
   If cmbWcn <> sOldCenter And bNewOp = 1 Then
      FindCenter cmbWcn, 1
   Else
      FindCenter cmbWcn, bNewOp
   End If
   sOldCenter = cmbWcn
   bNewOp = 0
   cmbJmp.Enabled = True
   z2(0).Enabled = True
'   Grd.Col = 2
'   Grd.Text = cmbWcn
   Grd.TextMatrix(CInt(lblRow), 2) = cmbWcn
    
End Sub


Private Sub FindCenter(sGetCenter As String, bNewOne As Byte)
   Dim RdoWcn As ADODB.Recordset
   Dim sShop  As String
   
   sGetCenter = Compress(sGetCenter)
   On Error GoTo DiaErr1
   If Len(sGetCenter) > 0 Then
      If bNewOne Then
         sShop = cmbShp
         sSql = "SELECT WCNREF,WCNNUM,WCNQHRS,WCNMHRS," & _
                    " WCNSUHRS,WCNUNITHRS,WCNSERVICE " & _
                "FROM WcntTable WHERE WCNREF='" & sGetCenter & "'"
          
          If (sShop <> "") Then
            sSql = sSql & " AND WCNSHOP = '" & sShop & "'"
          End If
      
      Else
         sSql = "Qry_GetWorkCenter '" & sGetCenter & "'"
      End If
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoWcn, ES_FORWARD)
      If bSqlRows Then
         With RdoWcn
            cmbWcn = "" & Trim(!WCNNUM)
            If bNewOne = 1 Then
               txtMdy = Format(!WCNMHRS, "##0.000")
               txtQdy = Format(!WCNQHRS, "##0.000")
               If !WCNSERVICE = 0 Then
                  txtSet = Format(!WCNSUHRS, "##0.000")
                  txtUnt = Format(!WCNUNITHRS, ES_TimeFormat)
                  cmbPrt.Enabled = True
               End If
            End If
            ClearResultSet RdoWcn
         End With
      Else
         MsgBox "Work Center Wasn't Found.", vbExclamation, Caption
         cmbWcn = ""
      End If
   End If
   Set RdoWcn = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "findcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
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
   Dim bResponse As Byte
   Dim sMsg As String
   
   If bShowBox = 1 Then
      If bTimeChg Then
         sMsg = "The Times May Have Changed, " & vbCr _
                & "Do You Wish To Reshedule The MO Now?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, ShopSHe02a.Caption)
         If bResponse = vbYes Then
            ShopSHe02a.optSrv.Value = vbChecked
            ShopSHe02b.Show
         End If
      End If
   End If
   Unload Me
   
End Sub


Private Sub cmdComments_Click()
   If cmdComments Then
      '6/7/2009 - Enabled the Comment textbox to write the comments
      txtCmt.Enabled = True
      txtCmt.SetFocus
      SysComments.lblListIndex = 6
      SysComments.Show
      cmdComments = False
   End If
   
End Sub

'fixed 2/20/01

Private Sub cmdDel_Click()
   Dim RdoISMo As ADODB.Recordset
   Dim bResponse As Byte
   Dim strEmpNo As String
   Dim strIsMOStart As String
   
   sSql = "SELECT ISEmployee,  CONVERT(varchar(12), ISMOSTART, 101) AS ISMOSTART " _
         & " FROM IstcTable WHERE (ISMO ='" & Compress(lblMon) & "' " _
          & "AND ISRUN =" & Val(lblRun) & " AND ISOP = " & Val(txtOpn) & ")"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoISMo, ES_STATIC)
   If bSqlRows Then
      With RdoISMo
         strEmpNo = "" & Trim(!ISEmployee)
         strIsMOStart = "" & Trim(!ISMOSTART)
      End With
      
      MsgBox "Employee No:" & strEmpNo & " has logged in to the operation. " & _
               "Can not cancel this MO Operation.", vbInformation, Caption
      
      Set RdoISMo = Nothing
      Exit Sub
   End If
      
   bResponse = MsgBox("Delete Operation " & txtOpn & "?", ES_NOQUESTION, Caption)
   If bResponse = vbNo Then
      Width = Width + 10
      Exit Sub
   End If
   On Error GoTo diaDelopErr1
   sSql = "DELETE FROM RnopTable WHERE (OPREF='" & Compress(lblMon) & "' " _
          & "AND OPRUN=" & Val(lblRun) & " AND OPNO = " & Val(txtOpn) & ")"
   clsADOCon.ExecuteSql sSql
   SysMsg "Operation Deleted", True, Me
   GetOperations
   Exit Sub
   
diaDelopErr1:
   CurrError.Description = Err.Description
   Resume diaDelopErr2
diaDelopErr2:
   MsgBox CurrError.Description & vbCr & "Couldn't Delete Op.", vbExclamation, Caption
   
End Sub




Private Sub cmdFil_Click()
   Dim bResponse As Byte
   Dim iList As Integer
   
   bResponse = MsgBox("Fill Operation " & txtOpn & " From Library?", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      Width = Width + 10
   Else
      MouseCursor 13
      optLib.Value = vbChecked
      ShopSHe02g.Show
      iList = iIndex - 1
      sComments = Trim(txtCmt)
      sComments = TrimComment(sComments)
      cmbJmp.List(iList) = Left(cmbJmp.List(iList), 3) & " " & sComments
      cmbJmp = cmbJmp.List(iList)
   End If
   
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
   UpdateOp
   iIndex = iIndex - 1
   If iIndex <= 1 Then iIndex = 1
   txtOpn = Format(iOperations(iIndex), "000")
   On Error Resume Next
   Grd.Col = 0
   Grd.row = iIndex
   bNewOp = 0
   GetThisOp
   If Grd.row > 6 Then Grd.TopRow = Grd.TopRow - 1
   '    If grd.Row < 6 Then
   '        grd.TopRow = 1
   '    Else
   '        a = grd.Row Mod 4
   '        If a = 2 Then grd.TopRow = grd.Row - 4
   '    End If
   
End Sub

Private Sub cmdNew_Click()
   UpdateOp
   bNewOp = 1
   AddOperation
   
End Sub

Private Sub cmdNxt_Click()
   Dim A As Integer
   UpdateOp
   iIndex = iIndex + 1
   If iIndex >= iTotalOps Then iIndex = iTotalOps
   txtOpn = Format(iOperations(iIndex), "000")
   On Error Resume Next
   Grd.Col = 0
   Grd.row = iIndex
   bNewOp = 0
   GetThisOp
   If Grd.row > 7 Then Grd.TopRow = Grd.TopRow + 1
   '    If grd.Row < 6 Then
   '        grd.TopRow = 1
   '    Else
   '        a = grd.Row Mod 4
   '        If a = 2 Then grd.TopRow = grd.Row - 1
   '    End If
   
End Sub

Private Sub cmdOrd_Click()
   UpdateOp
   GetOperations
   
End Sub




Private Sub cmdStatCode_Click()
    StatusCode.lblSCTypeRef = "MO Part"
    StatusCode.txtSCTRef = lblMon
    StatusCode.LableRef1 = "Run"
    StatusCode.lblSCTRef1 = lblRun
    StatusCode.lblSCTRef2 = txtOpn.Text
    StatusCode.lblStatType = "MOI"
    StatusCode.lblSysCommIndex = 8 ' The index in the Sys Comment "MO Comments"
    StatusCode.txtCurUser = cUR.CurrentUser
    
    StatusCode.Show

End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      cmdComments.Enabled = True
      GetRoutingIncrementDefault
      Grd.row = 1
      FillCombos
      GetOperations
      bOnLoad = 0
   End If
   If optLib.Value = vbChecked Then
      Unload ShopSHe02g
      optLib.Value = vbUnchecked
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move 400, 400
   FormatControls
   cUR.CurrentShop = GetSetting("Esi2000", "Current", "Shop", cUR.CurrentShop)
   
   'sSql = "SELECT * FROM RnopTable WHERE OPREF= ? AND OPRUN= ? AND OPNO= ?"
   sSql = "SELECT ro.*, wc.WCNNUM FROM RnopTable ro" & vbCrLf _
   & "join WcntTable wc on wc.WCNREF = ro.OPCENTER" & vbCrLf _
   & "WHERE OPREF= ? AND OPRUN= ? AND OPNO= ?"
   Set AdoStm = New ADODB.Command
   AdoStm.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.Type = adInteger
   
   Set AdoParameter3 = New ADODB.Parameter
   AdoParameter3.Type = adSmallInt
   
   AdoStm.Parameters.Append AdoParameter1
   AdoStm.Parameters.Append ADOParameter2
   AdoStm.Parameters.Append AdoParameter3
   

   lblRun = ShopSHe02a.cmbRun
   lblMon = Compress(ShopSHe02a.cmbPrt)
   sMoNumber = Compress(lblMon)
   
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .row = 0
      .Col = 0
      .Text = "Op No"
      .ColWidth(0) = 600
      .Col = 1
      .Text = "Shop"
      .ColWidth(1) = 600
      .Col = 2
      .Text = "Wk Ctr"
      .ColWidth(2) = 650
      .Col = 3
      .Text = "Comment"
      .ColWidth(3) = Grd.Width - 2300     ' leave room for scrollbar
      .Col = 0
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   UpdateOp
   ShopSHe02a.optRte.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set AdoParameter3 = Nothing
   Set AdoStm = Nothing
   Set ShopSHe02c = Nothing
   
End Sub


Private Sub Grd_Click()
   UpdateOp True
   Grd.Col = 0
   txtOpn = Grd.Text
   ' The bNewOp = 0 is already set in UpdateOp
   bNewOp = 0
   GetThisOp
   
End Sub

Private Sub Grd_GotFocus()
   Grd.Col = 0
   
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
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
   If Trim(cmbPrt) <> "" Then
      If Left(lblDsc, 8) = "*** Part" Then
         If Trim(cmbPrt) <> "NONE" Then
            lblDsc.ForeColor = ES_RED
         Else
            lblDsc = ""
         End If
      Else
         lblDsc.ForeColor = Es_TextForeColor
      End If
   Else
      If Left(lblDsc, 8) = "*** Part" Then lblDsc = ""
      lblDsc.ForeColor = Es_TextForeColor
   End If
End Sub

Private Sub optLib_Click()
   'never visible-unloads fill
   
End Sub

Private Sub optPck_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
   
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
   Else
      cmbPrt.Enabled = False
      cmbPrt = "NONE"
      lblDsc = ""
   End If
   
End Sub



Private Sub FillCombos()
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "Qry_FillShops"
   LoadComboBox cmbShp
   'LoadComboBoxAndSelect cmbShp
   
   If bSqlRows Then
      If cUR.CurrentShop <> "" Then
         cmbShp = cUR.CurrentShop
      Else
         If cmbShp.ListCount > 0 Then cmbShp = cmbShp.List(0)
      End If
   End If
   
   sSql = "SELECT PARTREF,PARTNUM,PALEVEL FROM PartTable WHERE PALEVEL=7 " _
          & "ORDER BY PARTREF"
   LoadComboBox cmbPrt
   MouseCursor 0
   FillWorkCenters
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillWorkCenters()
   Dim RdoWcn As ADODB.Recordset
   Dim bByte As Byte
   Dim iList As Integer
   Dim sCurCenter As String
   sCurCenter = cmbWcn
   cmbWcn.Clear
   On Error GoTo DiaErr1
   sSql = "Qry_FillWorkCenters '" & Compress(cmbShp) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWcn, ES_FORWARD)
   If bSqlRows Then
      With RdoWcn
         Do Until .EOF
            AddComboStr cmbWcn.hwnd, "" & Trim(!WCNNUM)
            .MoveNext
         Loop
      End With
      ClearResultSet RdoWcn
   End If
   cmbWcn = sCurCenter
   If cmbWcn.ListCount > 0 Then
      bByte = 0
      For iList = 0 To cmbWcn.ListCount - 1
         If Trim(cmbWcn.List(iList)) = Trim(cmbWcn) Then bByte = 1
      Next
      If bByte = 0 Then
         cmbWcn = cmbWcn.List(0)
      End If
      Grd.Col = 2
      Grd.Text = cmbWcn
   End If
   sOldShop = cmbShp
   Set RdoWcn = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillworkcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetOperations()
   Dim RdoRes2 As ADODB.Recordset
   Dim iRows As Integer
   Dim sShop As String
   Dim sCenter As String
   Dim sString As String
   
   Erase iOperations
   Grd.Rows = 2
   cmbJmp.Clear
   iTotalOps = 0
   iRows = 0 ' first column is header
   On Error GoTo DiaErr1
   If iAutoIncr <= 0 Then iAutoIncr = 10
   sSql = "SELECT OPREF,OPRUN,OPNO,OPSHOP,OPCENTER,OPCOMT FROM RnopTable WHERE OPREF='" & sMoNumber & "' " _
          & "AND OPRUN=" & Val(lblRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRes2, ES_STATIC)
   If bSqlRows Then
      With RdoRes2
         txtOpn = Format(!opNo, "000")
         Do Until .EOF
            On Error Resume Next
            iTotalOps = iTotalOps + 1
            iOperations(iTotalOps) = !opNo
            sComments = "" & Trim(!OPCOMT)
            sComments = TrimComment(sComments)
            cmbJmp.AddItem Format(!opNo, "000") & " " & sComments
            iRows = iRows + 1
            If iRows > 1 Then Grd.Rows = Grd.Rows + 1
            Grd.row = iRows
            Grd.Col = 0
            Grd.Text = Format(!opNo, "000")
            Grd.Col = 1
            sShop = GetRoutShop("" & Trim(!OPSHOP))
            Grd.Text = sShop
            Grd.Col = 2
            sCenter = GetRoutCenter("" & Trim(!OPCENTER))
            Grd.Text = sCenter
            Grd.Col = 3
            sString = "" & Trim(Left(!OPCOMT, 50))
            sString = Replace(sString, vbCrLf, " ")
            Grd.Text = sString
            .MoveNext
         Loop
         ClearResultSet RdoRes2
      End With
      cmbJmp.ListIndex = 0
      iIndex = 1
      Grd.row = 1
      Grd.Col = 0
      GetThisOp
   Else
      On Error Resume Next
      RdoRes2.Close
      txtOpn = Format(iAutoIncr, "000")
      sSql = "INSERT INTO RnopTable (OPREF,OPNO) " _
             & "VALUES('" & sPassedRout & "'," & Trim(txtOpn) & ")"
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
   Set RdoRes2 = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getoperations"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
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


Private Sub UpdateOp(Optional NoNotice As Boolean)
   Dim sShop As String
   Dim sCenter As String
   Dim sService As String
   Dim sSavecomments As String
   
   On Error GoTo DiaErr1
   sShop = Compress(cmbShp)
   sCenter = Compress(cmbWcn)
   If Len(sShop) = 0 Then
      MsgBox "Requires A Valid Shop.", vbExclamation, Caption
      Exit Sub
   End If
   If Len(sCenter) = 0 Then
      MsgBox "Requires A Valid Work Center.", vbExclamation, Caption
      Exit Sub
   End If
   If optSrv.Value = 1 Then
      sService = Compress(cmbPrt)
   Else
      sService = ""
   End If
   MouseCursor 13
   If Not NoNotice Then lblUpd = "Updating."
   lblUpd.Refresh
   sSavecomments = "" & Trim(txtCmt)
   On Error Resume Next
   sSql = "UPDATE RnopTable SET " _
          & "OPSHOP='" & sShop & "', " _
          & "OPCENTER='" & sCenter & "'," _
          & "OPSUHRS=" & Val(txtSet) & "," _
          & "OPUNITHRS=" & Val(txtUnt) & "," _
          & "OPQHRS=" & Val(txtQdy) & "," _
          & "OPMHRS=" & Val(txtMdy) & "," _
          & "OPSERVPART='" & sService & "'," _
          & "OPSVCUNIT=" & Val(txtCst) & "," _
          & "OPPICKOP=" & optPck.Value & "," _
          & "OPCOMT='" & txtCmt & "' " _
          & "WHERE OPREF='" & sMoNumber & "' " _
          & "AND OPRUN=" & Val(lblRun) & " AND OPNO=" & Val(txtOpn) & " "
   clsADOCon.ExecuteSql sSql
   bNewOp = 0
   Sleep 100
   MouseCursor 0
   lblUpd = ""
   lblUpd.Refresh
   Exit Sub
   
DiaErr1:
   sProcName = "updateop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetThisOp()
   On Error Resume Next
   Dim RdoRes2 As ADODB.Recordset
   
   If bTimeChg = 1 Then bShowBox = 1
   AdoStm.Parameters(0).Value = sMoNumber
   AdoStm.Parameters(1).Value = Val(lblRun)
   AdoStm.Parameters(2).Value = Val(txtOpn)

   bSqlRows = clsADOCon.GetQuerySet(RdoRes2, AdoStm)
'   Set RdoRes2 = RdoStm.OpenResultset(rdOpenStatic, rdConcurReadOnly)
   If Not RdoRes2.BOF And Not RdoRes2.EOF Then
      With RdoRes2
         cmbShp = "" & Trim(!OPSHOP)
'         cmbWcn = "" & Trim(!OPCENTER)
         cmbWcn = "" & Trim(!WCNNUM)
         If cmbShp <> sOldShop Then FillWorkCenters
         txtSet = Format(!OPSUHRS, "##0.000")
         txtUnt = Format(!OPUNITHRS, ES_TimeFormat)
         txtQdy = Format(!OPQHRS, "##0.000")
         txtMdy = Format(!OPMHRS, "##0.000")
         cmbPrt = "" & Trim(!OPSERVPART)
         If Len(Trim(cmbPrt)) > 0 And Trim(cmbPrt) <> "NONE" Then
            optSrv.Value = vbChecked
            cmbPrt.Enabled = True
         Else
            optSrv.Value = vbUnchecked
            cmbPrt.Enabled = False
         End If
         optPck.Value = !OPPICKOP
         txtCst = Format(!OPSVCUNIT, ES_QuantityDataFormat)
         txtCmt = "" & Trim(!OPCOMT)
         lblCurList = "" & Trim(!OPTOOLLIST)
         If lblCurList <> "" Then lblCurList = FindToolList(lblCurList, lblLst) _
            Else lblLst = ""
         If !OPCOMPLETE Then
            cmbShp.Enabled = False
            cmbWcn.Enabled = False
            txtSet.Enabled = False
            txtUnt.Enabled = False
            txtQdy.Enabled = False
            txtMdy.Enabled = False
            txtCst.Enabled = False
            cmbPrt.Enabled = False
            txtCmt.Enabled = False
            cmdDel.Enabled = False
            optPck.Enabled = False
            optSrv.Enabled = False
            optCom.Value = vbChecked
         Else
            cmbShp.Enabled = True
            cmbWcn.Enabled = True
            txtSet.Enabled = True
            txtUnt.Enabled = True
            txtQdy.Enabled = True
            txtMdy.Enabled = True
            txtCst.Enabled = True
            cmbPrt.Enabled = True
            txtCmt.Enabled = True
            cmdDel.Enabled = True
            optPck.Enabled = True
            optSrv.Enabled = True
            optCom.Value = vbUnchecked
         End If
         bTimeChg = 0
         ClearResultSet RdoRes2
      End With
      ' 4/21/2004
      'bNewOp = 0
      FindShop
      FindCenter cmbWcn, bNewOp
      If Trim(cmbPrt) <> "NONE" Then cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
      On Error Resume Next
      Grd.SetFocus
   End If
   If bOnLoad = 1 Then bTimeChg = 0
   sOldCenter = cmbWcn
   sOldShop = cmbShp
   RdoRes2.Close
   Set RdoRes2 = Nothing
   lblRow = Grd.row
   
End Sub

Private Sub optSrv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
   
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
   z2(0).Enabled = True
   On Error Resume Next
   'txtCmt = ParseComment(txtCmt, False)
   sSql = "UPDATE RnopTable SET " _
          & "OPCOMT='" & SqlString(txtCmt) & "' " _
          & "WHERE OPREF='" & sMoNumber & "' " _
          & "AND OPRUN=" & Val(lblRun) & " AND OPNO=" & Val(txtOpn) & " "
   clsADOCon.ExecuteSql sSql
   sString = Left(txtCmt, 20)
   sString = Replace(sString, vbCrLf, " ")
'   Grd.Col = 3
'   Grd.Text = sString
   Grd.TextMatrix(CInt(lblRow), 3) = sString
   Grd.Col = 0
   
End Sub


Private Sub txtCst_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub


Private Sub txtCst_LostFocus()
   txtCst = CheckLen(txtCst, 8)
   txtCst = Format(Abs(Val(txtCst)), "###0.000")
   
End Sub


Private Sub txtMdy_Change()
   bTimeChg = 1
   
End Sub

Private Sub txtMdy_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtMdy_LostFocus()
   txtMdy = CheckLen(txtMdy, 7)
   txtMdy = Format(Abs(Val(txtMdy)), "##0.000")
   cmbJmp.Enabled = True
   z2(0).Enabled = True
   
End Sub

Private Sub txtOpn_Click()
   iOldOpn = Val(Left(txtOpn, 3))
   
End Sub

Private Sub txtOpn_GotFocus()
   iCurrentOp = Val(txtOpn)
   iOldOpn = Abs(Val(txtOpn))
   
End Sub


Private Sub txtOpn_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtOpn_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   
   txtOpn = CheckLen(txtOpn, 3)
   If Len(txtOpn) = 0 Then txtOpn = Format(iCurrentOp, "000")
   txtOpn = Format(Abs(Val(txtOpn)), "000")
   If iOldOpn <> Val(txtOpn) Then
      bByte = 0
      For iList = 0 To cmbJmp.ListCount - 1
         If Val(txtOpn) = Val(Left(cmbJmp.List(iList), 3)) Then
            'If bNewOp = 0 Then
               MsgBox "Operation Exists.", vbInformation, Caption
               txtOpn = Format(iOldOpn, "000")
               bByte = 1
            'End If
            Exit For
         End If
      Next
      If bByte = 0 Then
         On Error GoTo diaOpsOpnErr1
         sSql = "UPDATE RnopTable SET OPNO=" & txtOpn & " WHERE (OPREF='" & lblMon & "' AND " _
                & "OPRUN=" & lblRun & " AND OPNO=" & str(iOldOpn) & ")"
         clsADOCon.ExecuteSql sSql
         
         sSql = "UPDATE runsTable set RUNOPCUR = MinOPno FROM " _
               & "(select MIN(OPNO) as MinOpno,  opref, Oprun FROM rnopTable " _
                  & "where OPCOMPLETE = 0 AND opref = '" & lblMon & "' and " _
                     & " OPRUN=" & lblRun & " GROUP BY opref, Oprun) as f  " _
                  & " WHERE f.OPREF = runref and f.OPRUN = runno " _
                  & " And runref = '" & lblMon & "' AND runno = " & lblRun _
                  & " AND MinOPno <> RUNOPCUR"
   
         clsADOCon.ExecuteSql sSql
         
         On Error Resume Next
         For iList = 0 To cmbJmp.ListCount - 1
            If Val(Left(cmbJmp.List(iList), 3)) = iOldOpn Then cmbJmp.RemoveItem iList
         Next
         iOperations(iIndex) = Val(txtOpn)
         cmbJmp = txtOpn
         cmbJmp.AddItem txtOpn
         On Error Resume Next
         cmbShp.SetFocus
      End If
   End If
   Exit Sub
   
diaOpsOpnErr1:
   CurrError.Description = Err.Description
   Resume diaOpsOpnErr2
diaOpsOpnErr2:
   MsgBox CurrError.Description & " Couldn't Change Operation.", vbInformation, Caption
   On Error GoTo 0
   
End Sub

Private Sub txtQdy_Change()
   bTimeChg = 1
   
End Sub

Private Sub txtQdy_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtQdy_LostFocus()
   txtQdy = CheckLen(txtQdy, 7)
   txtQdy = Format(Abs(Val(txtQdy)), "##0.000")
   cmbJmp.Enabled = True
   z2(0).Enabled = True
   
End Sub

Private Sub txtSet_Change()
   bTimeChg = 1
   
End Sub

Private Sub txtSet_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtSet_LostFocus()
   txtSet = CheckLen(txtSet, 7)
   txtSet = Format(Abs(Val(txtSet)), "##0.000")
   cmbJmp.Enabled = True
   z2(0).Enabled = True
   
End Sub

Private Sub txtUnt_Change()
   bTimeChg = 1
   
End Sub

Private Sub txtUnt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then LastOp
   If KeyCode = vbKeyPageDown Then NextOp
   
End Sub

Private Sub txtUnt_LostFocus()
   txtUnt = CheckLen(txtUnt, 8)
   txtUnt = Format(Abs(Val(txtUnt)), ES_TimeFormat)
   cmbJmp.Enabled = True
   z2(0).Enabled = True
   
End Sub



Private Sub AddOperation()
   Dim sShop As String
   Dim sCenter As String
   Dim RdoRes2 As ADODB.Recordset
   
   On Error GoTo DiaErr1
   lblUpd = "Adding Item."
   lblUpd.Refresh
   
   sShop = Compress(cmbShp)
   sCenter = Compress(cmbWcn)
   txtOpn = Format(iOperations(iTotalOps) + iAutoIncr, "000")
   
   On Error Resume Next
   sSql = "INSERT INTO RnopTable (OPREF,OPRUN,OPNO,OPSHOP,OPCENTER) " _
          & "VALUES('" & sMoNumber & "'," & Val(lblRun) & "," & Val(txtOpn) & "," _
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
   txtUnt = Format(0, ES_TimeFormat)
   cmbPrt = ""
   cmbJmp.Enabled = False
   z2(0).Enabled = False
   
   optSrv.Value = vbUnchecked
   optPck.Value = vbUnchecked
   cmbPrt.Enabled = False
   cmbJmp.AddItem txtOpn
   Grd.Rows = Grd.Rows + 1
   Grd.row = Grd.Rows - 1
   If Grd.row > 5 Then Grd.TopRow = Grd.row - 4
   Grd.Col = 0
   Grd.Text = txtOpn
   iTotalOps = iTotalOps + 1
   iOperations(iTotalOps) = Val(txtOpn)
   bNewOp = 1
   iIndex = iTotalOps
   lblUpd = ""
   lblUpd.Refresh
   SysMsg "Operation " & txtOpn & " Added.", True, Me
   Set RdoRes2 = Nothing
   GetThisOp
   Exit Sub
   
DiaErr1:
   sProcName = "addoperation"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub AutoNumber()
   Dim iList As Integer
   Dim iNewNumber As Integer
   Dim iNewOps(300, 3) As Integer
   
   On Error GoTo RopsAutoErr1
   cmdCan.Enabled = False
   MouseCursor 11
   For iList = 0 To cmbJmp.ListCount - 1
      iNewOps(iList, 0) = (iList + 1052)
      iNewOps(iList, 1) = Val(Left(cmbJmp.List(iList), 3))
   Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   For iList = 0 To cmbJmp.ListCount - 1
      sSql = "UPDATE RnopTable SET OPNO=" & str(iNewOps(iList, 0)) _
             & " WHERE OPREF='" & sMoNumber & "' AND OPRUN=" & Val(lblRun) & " AND OPNO=" & iNewOps(iList, 1)
      clsADOCon.ExecuteSql sSql
   Next
   clsADOCon.CommitTrans
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   iNewNumber = 0
   For iList = 0 To cmbJmp.ListCount - 1
      iNewNumber = iNewNumber + iAutoIncr
      sSql = "UPDATE RnopTable SET OPNO=" & str(iNewNumber) & " WHERE OPREF='" & sMoNumber & "' " _
             & "AND OPRUN=" & Val(lblRun) & " AND OPNO=" & str(iNewOps(iList, 0))
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
   
RopsAutoErr1:
   sProcName = "autonumber"
   CurrError.Description = Err.Description
   Resume RopsAutoErr2
RopsAutoErr2:
   On Error Resume Next
   MouseCursor 0
   clsADOCon.RollbackTrans
   cmdCan.Enabled = True
   MsgBox CurrError.Description & " Can't Reorganize Operations.", vbExclamation, Caption
   DoModuleErrors Me
   
End Sub

Private Sub z1_Click(Index As Integer)
   On Error Resume Next
   cmbJmp.SetFocus
   
End Sub
