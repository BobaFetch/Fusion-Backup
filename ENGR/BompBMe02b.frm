VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form BompBMe02b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parts List Items"
   ClientHeight    =   6225
   ClientLeft      =   2475
   ClientTop       =   1290
   ClientWidth     =   7830
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMe02b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdNxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7150
      Picture         =   "BompBMe02b.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "Next Entry"
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLst 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6620
      Picture         =   "BompBMe02b.frx":08E4
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Previous Entry"
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.ComboBox cmbRev 
      Height          =   315
      Left            =   4680
      Sorted          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Revision (Blank For Default)"
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton optBills 
      Caption         =   "From Bills"
      Height          =   255
      Left            =   480
      TabIndex        =   49
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "BompBMe02b.frx":0A1A
      DownPicture     =   "BompBMe02b.frx":138C
      Height          =   350
      Left            =   6240
      Picture         =   "BompBMe02b.frx":1CFE
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Standard Comments"
      Top             =   3480
      Width           =   350
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   360
      TabIndex        =   47
      Top             =   4560
      Width           =   7335
   End
   Begin VB.TextBox txtLab 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Tag             =   "1"
      ToolTipText     =   "Total Accumulated Labor Cost For This Level"
      Top             =   5160
      Width           =   1035
   End
   Begin VB.TextBox txtLabOh 
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Tag             =   "1"
      ToolTipText     =   "Factory Overhead Rate"
      Top             =   5160
      Width           =   1035
   End
   Begin VB.TextBox txtMat 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Tag             =   "1"
      ToolTipText     =   "Total Material Costs For This Level"
      Top             =   5520
      Width           =   1035
   End
   Begin VB.TextBox txtMatbr 
      Height          =   285
      Left            =   4560
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "Material Burden Percentage"
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "BompBMe02b.frx":2300
      Height          =   350
      Left            =   7200
      Picture         =   "BompBMe02b.frx":27DA
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "Show Parts List"
      Top             =   1680
      Width           =   350
   End
   Begin VB.Frame Z2 
      Height          =   1215
      Index           =   0
      Left            =   6720
      TabIndex        =   35
      Top             =   2160
      Width           =   1005
      Begin VB.CommandButton cmdOrd 
         Caption         =   "&Reorder"
         Height          =   315
         Left            =   80
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Resort After Sequence Change"
         Top             =   840
         Width           =   850
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   80
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Delete This Item From The Parts List"
         Top             =   480
         Width           =   850
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   315
         Left            =   80
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Add a Part To Parts List"
         Top             =   140
         Width           =   850
      End
   End
   Begin VB.TextBox txtRef 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   10
      Top             =   6360
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.ComboBox cmbJmp 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "8"
      Top             =   6240
      Width           =   3320
   End
   Begin VB.TextBox txtCmt 
      Height          =   1150
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Tag             =   "9"
      ToolTipText     =   "Comments (2048 Chars Max)"
      Top             =   3360
      Width           =   4335
   End
   Begin VB.CheckBox optPhn 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4740
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtSup 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "Use for Operation Testing"
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtAdr 
      Height          =   285
      Left            =   4740
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Wasted (cut off)"
      Top             =   2520
      Width           =   915
   End
   Begin VB.TextBox txtCvt 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Units Conversion (Feet to Inches = 12.000)"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtBum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Unit of Measure for Parts List"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   5760
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Quantity Used"
      Top             =   1800
      Width           =   915
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1800
      Width           =   3320
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Sort Sequence (Otherwise Part Number)"
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6720
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6960
      Top             =   5520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6225
      FormDesignWidth =   7830
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Click To Select Or Scroll And Press Enter"
      Top             =   60
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   19
      Left            =   5640
      TabIndex        =   51
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   18
      Left            =   5640
      TabIndex        =   50
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimating Costs For This Level.  Should Not Include Lower Level Costs"
      Height          =   255
      Index           =   17
      Left            =   360
      TabIndex        =   48
      Top             =   4800
      Width           =   6135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Cost"
      Height          =   255
      Index           =   16
      Left            =   360
      TabIndex        =   46
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Overhead"
      Height          =   255
      Index           =   15
      Left            =   3000
      TabIndex        =   45
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Cost"
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   44
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Burden"
      Height          =   255
      Index           =   13
      Left            =   3000
      TabIndex        =   43
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type "
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
      Index           =   12
      Left            =   4200
      TabIndex        =   41
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4200
      TabIndex        =   40
      Top             =   1800
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
      Left            =   360
      TabIndex        =   39
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reference:"
      Height          =   255
      Index           =   11
      Left            =   4320
      TabIndex        =   34
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblLvl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   33
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jump"
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   32
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   30
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblPls 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   29
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   840
      TabIndex        =   28
      Top             =   2160
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   27
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phantom:  "
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   26
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Qty:"
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   25
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Units Wasted:"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   24
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Convert:   "
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   23
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Um     "
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
      Index           =   4
      Left            =   6720
      TabIndex        =   22
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity       "
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
      Index           =   3
      Left            =   5760
      TabIndex        =   21
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev             "
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
      Index           =   2
      Left            =   4680
      TabIndex        =   20
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                       "
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
      Index           =   1
      Left            =   840
      TabIndex        =   19
      Top             =   1560
      Width           =   3270
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seq     "
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
      Index           =   0
      Left            =   360
      TabIndex        =   18
      Top             =   1560
      Width           =   375
   End
End
Attribute VB_Name = "BompBMe02b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'
'Reference hidden and added Est Costs 10/10/02
'7/8/04 Smoothed grid scrolling
'9/1/04 omit tools
Option Explicit
'Dim RdoStm As rdoQuery
Dim AdoCmdObj As ADODB.Command
Dim RdoBom As ADODB.Recordset

Dim bClosing As Byte
Dim bFromBills As Byte
Dim bGoodPart As Byte
Dim bGoodRev As Byte
Dim bNewItem As Byte
Dim bOnLoad As Byte

Dim iIndex As Integer
Dim iTotalItems As Integer
Dim sOldPart As String
Dim sOldRev As String
Dim sUonPart As String
Dim sBillParts(300, 3) As String
'0 for Sequence, 1 Compressed Part, 2 Desc

Private partNumberEnteringComboBox As String

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbJmp_Click()
   On Error Resume Next
   UpdateThisItem
   iIndex = cmbJmp.ListIndex + 1
   GetThisItem
   
End Sub


Private Sub cmbJmp_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub cmbPrt_Click()
   bGoodPart = GetPart(False, True)
   If bGoodPart Then
      grd.col = 2
      grd.Text = cmbPrt
   End If
   
End Sub

Private Sub cmbPrt_GotFocus()
   partNumberEnteringComboBox = Compress(cmbPrt)
End Sub

Private Sub cmbPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub cmbPrt_LostFocus()
   
   If Compress(cmbPrt) = partNumberEnteringComboBox Then
      Exit Sub
   End If
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   Dim A As Integer
   Dim iList As Integer
   Dim sNewPart As String
   Dim sDesc As String
   cmbPrt = CheckLen(cmbPrt, 30)
   
   bGoodPart = GetPart(False)
   If bGoodPart = 0 Then
      MsgBox "Part Wasn't Found or Wrong Type.", vbExclamation, Caption
      cmbPrt = sOldPart
      partNumberEnteringComboBox = sOldPart
      FindPart
      Exit Sub
   End If
'   If sOldPart <> cmbPrt Then
'      If cmbPrt = sUonPart Then
'         MsgBox "A Part May Not Be Used On Itself.", vbExclamation, Caption
'         cmbPrt = sOldPart
'         cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
'         Exit Sub
'      End If
'      For iList = 0 To cmbJmp.ListCount - 1
'         A = InStr(cmbJmp, Chr(160))
'         If A > 0 Then
'            If cmbPrt = Left(cmbJmp.List(iList), A - 1) Then
'               MsgBox "Can't Use That Part Number Twice.", vbExclamation, Caption
'               If Len(sOldPart) > 0 Then cmbPrt = sOldPart
'               FindPart
'               bGoodPart = 0
'               Exit For
'            End If
'         End If
'      Next
'      If sOldPart <> cmbPrt Then FillBomhRev cmbPrt
   
   Dim compressedPart As String, oldPart As String
   oldPart = Compress(sOldPart)
   compressedPart = Compress(cmbPrt)
   If oldPart <> compressedPart Then
      If compressedPart = sUonPart Then
         MsgBox "A Part May Not Be Used On Itself.", vbExclamation, Caption
         cmbPrt = sOldPart
         cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
         partNumberEnteringComboBox = sOldPart
         Exit Sub
      End If
      For iList = 0 To cmbJmp.ListCount - 1
         A = InStr(Compress(cmbJmp), Chr(160))
         Dim itemLength As Integer
         itemLength = InStr(Compress(cmbJmp.list(iList)), Chr(160))

         If A > 1 And itemLength > 1 Then
            If compressedPart = Left(Compress(cmbJmp.list(iList)), itemLength - 1) Then
               MsgBox "Can't Use That Part Number Twice.", vbExclamation, Caption
               If Len(oldPart) > 0 Then cmbPrt = oldPart
               FindPart
               bGoodPart = 0
               Exit For
            End If
         End If
      Next
      If sOldPart <> cmbPrt Then
         'sOldPart = cmbPrt
         FillBomhRev cmbPrt
      End If
   
   End If
   If bGoodPart = 1 And cmbPrt <> sOldPart Then
      sNewPart = Compress(cmbPrt)
      sDesc = GetDesc(cmbPrt)
      cmbJmp = cmbPrt & Chr(160) & lblDsc
      cmbRev = ""
   End If
   On Error Resume Next
   sDesc = GetDesc(cmbPrt)
   cmbJmp.list(iIndex - 1) = cmbPrt & Chr(160) & lblDsc
   cmbJmp = cmbPrt & Chr(160) & lblDsc
   grd.col = 2
   grd.Text = cmbPrt
   
End Sub

Private Sub cmbRev_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   bGoodRev = GetRevision()
   If Not bGoodRev Then
      MsgBox "That Revision Wasn't Found.", vbExclamation, Caption
      cmbRev = sOldRev
   End If
   sOldRev = cmbRev
   
End Sub


Private Sub cmdCan_Click()
   bClosing = True
   If Trim(cmbPrt) = "" Then
      sSql = "DELETE FROM BmplTable WHERE BMASSYPART='" & lblPls & "'" _
             & "AND BMREV='" & lblRev & "' AND BMPARTNUM=''"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
   Else
      UpdateThisItem
   End If
   If bFromBills = 0 Then BompBMe02a.optPls.Value = vbUnchecked
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdCan_Click
   
End Sub


Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtCmt.SetFocus
      SysComments.lblListIndex = 9
      SysComments.Show
      cmdComments = False
   End If
   
End Sub

Private Sub cmdDel_Click()
   Dim bResponse As Byte
   Dim sCurrPart As String
   Dim sMsg As String
   
   On Error Resume Next
   sMsg = "Do You Really Want To Delete " & vbCrLf _
          & "Part Number " & cmbPrt & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      sCurrPart = Compress(cmbPrt)
      sSql = "DELETE FROM BmplTable WHERE (BMASSYPART='" & lblPls & "' " _
             & "AND BMREV='" & lblRev & "' " _
             & "AND BMPARTREF='" & sCurrPart & "')"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
      If clsADOCon.RowsAffected > 0 Then
         MsgBox "Item Was Deleted.", vbInformation, Caption
         GetItems
      Else
         MsgBox "Couldn't Delete Selected Item.", vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3102
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdLst_Click()
   Dim A As Integer
   UpdateThisItem
   iIndex = iIndex - 1
   If iIndex < 1 Then iIndex = 1
   grd.col = 0
   grd.row = iIndex
   GetThisItem
   If grd.row < 6 Then
      grd.TopRow = 1
   Else
      A = grd.row Mod 4
      If A = 2 Then grd.TopRow = grd.row - 4
   End If
   
End Sub

Private Sub cmdNew_Click()
   UpdateThisItem
   AddItem
   
End Sub

Private Sub cmdNxt_Click()
   Dim A As Integer
   UpdateThisItem
   iIndex = iIndex + 1
   If iIndex > iTotalItems Then iIndex = iTotalItems
   grd.col = 0
   grd.row = iIndex
   GetThisItem
   If grd.row < 6 Then
      grd.TopRow = 1
   Else
      A = grd.row Mod 4
      If A = 2 Then grd.TopRow = grd.row - 1
   End If
   
End Sub


Private Sub cmdOrd_Click()
   UpdateThisItem
   GetItems
   
End Sub


Private Sub cmdVew_Click()
   Dim iList As Integer
   Dim iCol As Integer
   Dim iRows As Integer
   Dim RdoVew As ADODB.Recordset
   If Val(txtQty) = 0 Then txtQty = Format(1, ES_QuantityDataFormat)
   If bGoodPart = 1 Then UpdateThisItem
   On Error Resume Next
   
   sSql = "SELECT BMASSYPART,BMPARTREF,BMPARTNUM,BMREV,BMSEQUENCE," _
          & "BMQTYREQD,BMUNITS,BMCONVERSION FROM BmplTable WHERE " _
          & "BMASSYPART='" & lblPls & "' AND BMREV='" & lblRev & "' " _
          & "ORDER BY BMSEQUENCE,BMPARTNUM"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVew)
   MouseCursor 13
   iRows = 10
   With BompBMview.grd
      .Rows = iRows
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      If Screen.Width > 9999 Then
         .ColWidth(0) = 400 * 1.25
         .ColWidth(1) = 1550 * 1.25
         .ColWidth(2) = 900 * 1.25
         .ColWidth(3) = 800 * 1.25
         .ColWidth(4) = 800 * 1.25
      Else
         .ColWidth(0) = 400
         .ColWidth(1) = 1550
         .ColWidth(2) = 900
         .ColWidth(3) = 800
         .ColWidth(4) = 800
      End If
   End With
   If bSqlRows Then
      With RdoVew
         Do Until .EOF
            iList = iList + 1
            iRows = iRows + 1
            BompBMview.grd.Rows = iRows
            BompBMview.grd.row = iRows - 11
            
            BompBMview.grd.col = 0
            BompBMview.grd = "" & str(!BMSEQUENCE)
            
            BompBMview.grd.col = 1
            BompBMview.grd = "" & Trim(!BMPARTNUM)
            
            BompBMview.grd.col = 2
            BompBMview.grd = Format(!BMQTYREQD, ES_QuantityDataFormat)
            
            BompBMview.grd.col = 3
            BompBMview.grd = "" & Trim(!BMUNITS)
            
            BompBMview.grd.col = 4
            BompBMview.grd = Format(!BMCONVERSION, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoVew
      End With
      If iList > 9 Then BompBMview.grd.Rows = iList + 1
   End If
   MouseCursor 0
   Set RdoVew = Nothing
   BompBMview.Show
   On Error GoTo 0
   
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      cmdComments.Enabled = True
      GetItems
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim sPls As String
   Dim b As Byte
   Move BompBMe02a.Left + 100, BompBMe02a.Top + 500
   FormatControls
   sUonPart = BompBMe02a.cmbPls
   sPls = Compress(BompBMe02a.cmbPls)
   lblPls = sPls
   lblRev = Trim(BompBMe02a.cmbRev)
   lblLvl = Trim(BompBMe02a.lblLvl)
   Caption = Caption & " For " & BompBMe02a.cmbPls
   sCurrForm = "Parts List"
   FillPartCombo
   sSql = "SELECT * FROM BmplTable WHERE BMASSYPART='" & Compress(BompBMe02a.cmbPls) & "' " _
          & "AND BMREV='" & Trim(BompBMe02a.cmbRev) & "' AND " _
          & "BMPARTREF= ? "
   
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim prmBMPtr As ADODB.Parameter
   Set prmBMPtr = New ADODB.Parameter
   prmBMPtr.Type = adChar
   prmBMPtr.Size = 30
   AdoCmdObj.Parameters.Append prmBMPtr

   'Set RdoStm = RdoCon.CreateQuery("", sSql)
   ' TODO: Set on Recordset
   'RdoStm.MaxRows = 1
   
   
   With grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .row = 0
      .col = 0
      .Text = "Seq"
      .ColWidth(0) = 650
      .col = 1
      .Text = "Rev"
      .ColWidth(1) = 400
      .col = 2
      .Text = "Part Number"
      .ColWidth(2) = 3500
      .col = 3
      .Text = "Quantity"
      .ColWidth(3) = 950
      .col = 0
   End With
   bClosing = 0
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   If bFromBills = 1 Then
      sCurrForm = "Bills Of Material"
      BompBMe01a.optRefresh.Value = vbChecked
      BompBMe01a.cmdQuit.Enabled = True
      BompBMe01a.cmdAdd.Enabled = True
      BompBMe01a.cmdEdit.Enabled = True
      BompBMe01a.cmdCut.Enabled = True
      BompBMe01a.cmdCut.Enabled = True
      BompBMe01a.cmdCopy.Enabled = True
      BompBMe01a.cmdDelete.Enabled = True
   End If
   sSql = "DELETE FROM BmplTable WHERE BMASSYPART='" & Trim(lblPls) & "' " _
          & "AND BMREV='" & Trim(lblRev) & "' " _
          & "AND (BMPARTREF='' OR BMPARTREF='NONE' OR BMPARTREF IS NULL)"
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   If bFromBills = 0 Then BompBMe02a.cmdPls.Enabled = True
   'RdoStm.Close
   Set AdoCmdObj = Nothing
   Set RdoBom = Nothing
   Set BompBMe02b = Nothing
   
End Sub



Private Sub GetItems()
   Dim iList As Integer
   Dim iRows As Integer
   Dim sDesc As String
   cmbJmp.Clear
   grd.Rows = 2
   Erase sBillParts
   iList = 0
   On Error GoTo DiaErr1
   sSql = "SELECT BMASSYPART,BMPARTREF,BMPARTNUM,BMREV,BMQTYREQD,BMSEQUENCE " _
          & "FROM BmplTable WHERE (BMASSYPART='" & lblPls & "' " _
          & "AND BMREV='" & BompBMe02a.cmbRev & "') " _
          & "ORDER BY BMSEQUENCE,BMPARTNUM"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_STATIC)
   If bSqlRows Then
      With RdoBom
         Do Until .EOF
            iList = iList + 1
            sBillParts(iList, 2) = GetDesc(!BMPARTNUM)
            cmbJmp.AddItem "" & Trim(!BMPARTNUM) & Chr(160) & sBillParts(iList, 2)
Debug.Print iList & ": " & Trim(!BMPARTNUM) & " * " & sBillParts(iList, 2)
            sBillParts(iList, 0) = "" & Trim(str(!BMSEQUENCE))
            sBillParts(iList, 1) = "" & Trim(!BMPARTREF)
            iRows = iRows + 1
            If iRows > 1 Then grd.Rows = grd.Rows + 1
            grd.row = iRows
            grd.col = 0
            grd.Text = Format(!BMSEQUENCE, "##0")
            grd.col = 1
            grd.Text = "" & Trim(!BMREV)
            grd.col = 2
            grd.Text = "" & Trim(!BMPARTNUM)
            grd.col = 3
            grd.Text = Format(!BMQTYREQD, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
      grd.col = 0
      On Error Resume Next
      RdoBom.Close
      iTotalItems = iList
      'iIndex = 1
      iIndex = iList '1 MM
      
      grd.row = iList
      grd.TopRow = iList
      
      GetThisItem
   Else
      bNewItem = 1
      bGoodPart = GetPart(False)
      AddItem
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetThisItem()
   On Error GoTo DiaErr1
   AdoCmdObj.Parameters(0) = sBillParts(iIndex, 1)
   'RdoStm(0) = sBillParts(iIndex, 1)
   bSqlRows = clsADOCon.GetQuerySet(RdoBom, AdoCmdObj, ES_STATIC, True, 1)
   If bSqlRows Then
      With RdoBom
         txtSeq = Format(!BMSEQUENCE, "##0")
         cmbPrt = "" & Trim(!BMPARTNUM)
         cmbRev = "" & Trim(!BMPARTREV)
         txtQty = Format(!BMQTYREQD, ES_QuantityDataFormat)
         txtBum = "" & Trim(!BMUNITS)
         txtCvt = Format(!BMCONVERSION, ES_QuantityDataFormat)
         txtAdr = Format(!BMADDER, ES_QuantityDataFormat)
         txtSup = Format(!BMSETUP, ES_QuantityDataFormat)
         optPhn.Value = !BMPHANTOM
         txtRef = "" & Trim(!BMREFERENCE)
         txtCmt = "" & !BMCOMT
         '10/10/02
         txtLab = Format(!BMESTLABOR, ES_QuantityDataFormat)
         txtLabOh = Format(!BMESTLABOROH, ES_QuantityDataFormat)
         txtMat = Format(!BMESTMATERIAL, ES_QuantityDataFormat)
         txtMatbr = Format(!BMESTMATERIALBRD, ES_QuantityDataFormat)
         
         If txtBum = "" Then txtBum = "EA"
      End With
      On Error Resume Next
      bNewItem = 0
      bGoodPart = GetPart(True)
      sOldPart = cmbPrt
      sOldRev = cmbRev
      'txtSeq.SetFocus
      grd.row = iIndex
      grd.SetFocus
      sBillParts(iIndex, 2) = GetDesc(cmbPrt)
      cmbJmp = cmbPrt & Chr(160) & sBillParts(iIndex, 2)
   Else
      txtSeq = "0"
      cmbPrt = ""
      cmbRev = ""
      txtQty = "0.000"
      txtBum = ""
      txtCvt = "0.000"
      txtAdr = "0.000"
      txtSup = "0.000"
      optPhn.Value = vbUnchecked
      txtRef = ""
      txtCmt = ""
      FindPart
   End If
   RdoBom.Close
   Exit Sub
   
DiaErr1:
   sProcName = "getthisit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub grd_Click()
   UpdateThisItem
   iIndex = grd.row
   GetThisItem
   
End Sub

Private Sub Grd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      UpdateThisItem
      iIndex = grd.row
      GetThisItem
   End If
   
End Sub


Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optPhn_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPhn_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub optPhn_LostFocus()
   If Val(txtQty) > 0 Then
      cmbJmp.Enabled = True
      cmdNew.Enabled = True
      cmdVew.Enabled = True
      cmdOrd.Enabled = True
   End If
   
End Sub

Private Sub txtAdr_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 9)
   txtAdr = Format(Abs(Val(txtAdr)), ES_QuantityDataFormat)
   If Val(txtQty) > 0 Then
      cmbJmp.Enabled = True
      cmdNew.Enabled = True
      cmdVew.Enabled = True
      cmdOrd.Enabled = True
   End If
   
End Sub


Private Sub txtBum_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtBum_LostFocus()
   txtBum = CheckLen(txtBum, 2)
   If txtBum = "" Then txtBum = "EA"
   If Val(txtQty) > 0 Then
      cmbJmp.Enabled = True
      cmdNew.Enabled = True
      cmdVew.Enabled = True
      cmdOrd.Enabled = True
   End If
   
End Sub

Private Sub txtCmt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   If Val(txtQty) > 0 Then
      cmbJmp.Enabled = True
      cmdNew.Enabled = True
      cmdVew.Enabled = True
      cmdOrd.Enabled = True
   End If
   sSql = "UPDATE BmplTable SET " _
          & "BMCOMT='" & Trim(txtCmt) & "' " _
          & "WHERE BMASSYPART='" & lblPls & "' " _
          & "AND BMREV='" & lblRev & "' AND BMPARTNUM='" & sOldPart & "' "
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
End Sub

Private Sub txtCvt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtCvt_LostFocus()
   txtCvt = CheckLen(txtCvt, 9)
   txtCvt = Format(Abs(Val(txtCvt)), ES_QuantityDataFormat)
   If Val(txtQty) > 0 Then
      cmbJmp.Enabled = True
      cmdNew.Enabled = True
      cmdVew.Enabled = True
      cmdOrd.Enabled = True
   End If
   
End Sub

Private Sub txtLab_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtLab_LostFocus()
   txtLab = CheckLen(txtLab, 9)
   txtLab = Format(Abs(Val(txtLab)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtLabOh_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtLabOh_LostFocus()
   txtLabOh = CheckLen(txtLabOh, 9)
   txtLabOh = Format(Abs(Val(txtLabOh)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtMat_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtMat_LostFocus()
   txtMat = CheckLen(txtMat, 9)
   txtMat = Format(Abs(Val(txtMat)), ES_QuantityDataFormat)
   
End Sub



Private Sub txtMatbr_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtMatbr_LostFocus()
   txtMatbr = CheckLen(txtMatbr, 9)
   txtMatbr = Format(Abs(Val(txtMatbr)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtQty_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   If Val(txtQty) > 0 Then
      cmbJmp.Enabled = True
      cmdNew.Enabled = True
      cmdVew.Enabled = True
      cmdOrd.Enabled = True
   End If
   grd.col = 3
   grd.Text = txtQty
   grd.col = 0
   
End Sub

Private Sub txtRef_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtRef_LostFocus()
   txtRef = CheckLen(txtRef, 16)
   
End Sub

Private Sub txtSeq_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtSeq_LostFocus()
   txtSeq = CheckLen(txtSeq, 3)
   txtSeq = Format$(Abs(Val(txtSeq)), "##0")
   sBillParts(iIndex, 0) = txtSeq
   grd.col = 0
   grd.Text = txtSeq
   
End Sub

Private Sub txtSup_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtSup_LostFocus()
   txtSup = CheckLen(txtSup, 9)
   txtSup = Format(Abs(Val(txtSup)), ES_QuantityDataFormat)
   If Val(txtQty) > 0 Then
      cmbJmp.Enabled = True
      cmdNew.Enabled = True
      cmdVew.Enabled = True
      cmdOrd.Enabled = True
   End If
   
End Sub

Private Sub z1_Click(Index As Integer)
   On Error Resume Next
   cmbJmp.SetFocus
   
End Sub



Private Sub FillPartCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PALEVEL,PATOOL FROM PartTable " _
          & "WHERE (PALEVEL BETWEEN " & lblLvl & " AND 5 " _
          & "AND PATOOL=0) AND PARTREF<>'" & lblPls _
          & "' AND PAINACTIVE = 0 AND PAOBSOLETE = 0  ORDER BY PARTREF"
   LoadComboBox cmbPrt
   Exit Sub
   
DiaErr1:
   sProcName = "fillpartco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub UpdateThisItem()
   Dim sPartNumber As String
   sPartNumber = Compress(cmbPrt)
   On Error GoTo BitmsUi1
   If Val(txtQty) = 0 Then
      If Not bClosing Then
         MsgBox "Requires A Valid Quantity." & vbCrLf _
            & "Item Couldn't Be Updated.", vbExclamation, Caption
      End If
      Exit Sub
   End If
   If bGoodPart = 0 Then
      If Not bClosing Then
         MsgBox "Requires A Valid Part Number." & vbCrLf _
            & "Item Couldn't Be Updated.", vbExclamation, Caption
      End If
      Exit Sub
   End If
   lblUpd = "Updating."
   lblUpd.Refresh
   sSql = "UPDATE BmplTable SET " _
          & "BMPARTREF='" & sPartNumber & "'," _
          & "BMPARTNUM='" & cmbPrt & "'," _
          & "BMPARTREV='" & cmbRev & "'," _
          & "BMQTYREQD=" & Format(Val(txtQty), ES_QuantityDataFormat) & "," _
          & "BMUNITS='" & txtBum & "', " _
          & "BMCONVERSION=" & Format(Val(txtCvt), ES_QuantityDataFormat) & "," _
          & "BMSEQUENCE=" & Val(txtSeq) & ", " _
          & "BMADDER=" & Format(Val(txtAdr), ES_QuantityDataFormat) & "," _
          & "BMSETUP=" & Format(Val(txtSup), ES_QuantityDataFormat) & "," _
          & "BMPHANTOM=" & str(optPhn.Value) & ", " _
          & "BMREFERENCE='" & txtRef & "'," _
          & "BMCOMT='" & Trim(txtCmt) & "'," _
          & "BMESTLABOR=" & Format(Val(txtLab), ES_QuantityDataFormat) & "," _
          & "BMESTLABOROH=" & Format(Val(txtLabOh), ES_QuantityDataFormat) & "," _
          & "BMESTMATERIAL=" & Format(Val(txtMat), ES_QuantityDataFormat) & "," _
          & "BMESTMATERIALBRD=" & Format(Val(txtMatbr), ES_QuantityDataFormat) & " " _
          & "WHERE (BMASSYPART='" & lblPls & "' " _
          & "AND BMREV='" & Trim(BompBMe02a.cmbRev) & "' AND " _
          & "BMPARTREF='" & Compress(sOldPart) & "')"
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   sBillParts(iIndex, 1) = sPartNumber
   sBillParts(iIndex, 2) = GetDesc(sPartNumber)
   Sleep 100
   cmbJmp.Enabled = True
   cmdNew.Enabled = True
   cmdVew.Enabled = True
   cmdOrd.Enabled = True
   lblUpd = ""
   lblUpd.Refresh
   bNewItem = 0
   Exit Sub
   
BitmsUi1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume BitmsUi2
BitmsUi2:
   lblUpd = ""
   lblUpd.Refresh
   MsgBox str(Err.Number) & vbCrLf & "Could Not Update.", vbExclamation, Caption
   
End Sub

Private Function GetPart(bSkipJump As Boolean, Optional NewPart As Boolean) As Byte
   Dim RdoRes2 As ADODB.Recordset
   Dim sCurrPart As String
   sCurrPart = Compress(cmbPrt)
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAPHANTOM," _
          & "PAUNITS,PABOMREV,PATOOL FROM PartTable " _
          & "WHERE (PARTREF='" & sCurrPart & "' AND " _
          & "PALEVEL BETWEEN " & lblLvl & " AND 5) AND PATOOL=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRes2)
   If bSqlRows Then
      With RdoRes2
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblTyp = Format(0 + !PALEVEL, "0")
         GetPart = 1
         If bNewItem = 1 Or NewPart Then
            optPhn.Value = !PAPHANTOM
            cmbRev = "" & !PABOMREV
            txtBum = "" & Trim(!PAUNITS)
            If txtBum = "" Then txtBum = "EA"
         End If
         If Not bSkipJump Then cmbJmp = cmbPrt & Chr(160) & lblDsc
      End With
   Else
      lblDsc = ""
      lblTyp = ""
      GetPart = 0
   End If
   Set RdoRes2 = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetRevision() As Byte
   Dim RdoRes2 As ADODB.Recordset
   Dim sCurrPart As String
   sCurrPart = Compress(cmbPrt)
   If Trim(cmbRev) = "" Then
      GetRevision = True
      Exit Function
   End If
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREF,BMHREV FROM BmhdTable " _
          & "WHERE BMHREF='" & sCurrPart & "' AND " _
          & "BMHREV='" & Trim(cmbRev) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRes2)
   If bSqlRows Then
      GetRevision = True
   Else
      cmbRev = ""
      GetRevision = False
   End If
   Set RdoRes2 = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrevis"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddItem()
   On Error GoTo DiaErr1
   
   Dim lNextSeq As Long
   lblUpd = "Adding Item."
   lblUpd.Refresh
   sSql = "DELETE FROM BmplTable WHERE (BMASSYPART='" & lblPls & "'" _
          & "AND BMREV='" & BompBMe02a.cmbRev & "' AND BMPARTNUM='')"
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   lNextSeq = GetNextSeqNum(lblPls)
   
   'sSql = "INSERT INTO BmplTable (BMASSYPART,BMREV,BMPARTREF) " _
   '       & "VALUES('" & lblPls & "','" & lblRev & "','')"
   sSql = "INSERT INTO BmplTable (BMASSYPART,BMREV,BMPARTREF, BMSEQUENCE) " _
          & "VALUES('" & lblPls & "','" & lblRev & "',''," & CStr(lNextSeq) & ")"
   
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   Sleep 500
   If clsADOCon.RowsAffected > 0 Then
      cmbJmp.AddItem "New Item                     "
      cmbJmp = "New Item"
      iTotalItems = iTotalItems + 1
      iIndex = iTotalItems
      
      ' set the text seq as the last item
      txtSeq = lNextSeq
      
      sBillParts(iIndex, 0) = txtSeq
      sBillParts(iIndex, 1) = ""
      sBillParts(iIndex, 2) = ""
      If bNewItem = 0 Then
         cmbRev = ""
         txtBum = ""
      End If
      cmbPrt = ""
      txtQty = "0.000"
      txtCvt = "0.000"
      txtAdr = "0.000"
      txtSup = "0.000"
      'txtSeq = "0"
      txtSeq = lNextSeq '"0" MM
      
      optPhn.Value = vbUnchecked
      txtRef = ""
      txtCmt = ""
      sOldPart = ""
      cmdNew.Enabled = False
      cmdNew.Enabled = False
      cmdVew.Enabled = False
      '  cmbJmp.Enabled = False
      cmdOrd.Enabled = False
      On Error Resume Next
      bNewItem = 1
      GetItems
      txtSeq.SetFocus
   Else
      MsgBox "Couldn't Add An Item.", vbInformation, Caption
   End If
   lblUpd = ""
   lblUpd.Refresh
   Exit Sub
   
DiaErr1:
   sProcName = "additem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetNextSeqNum(strAssyPart As String) As Long
   Dim RdoNextSeq As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT ISNULL(MAX(BMSEQUENCE +1),0) FROM BmplTable where BMASSYPART = '" & Compress(strAssyPart) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNextSeq)
   If bSqlRows Then
      With RdoNextSeq
         GetNextSeqNum = .Fields(0)
         ClearResultSet RdoNextSeq
      End With
   Else
      GetNextSeqNum = 0
   End If
   Set RdoNextSeq = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetNextSeqNum"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Private Function GetDesc(sPartNumber As String) As String
   Dim RdoDsc As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PADESC FROM PartTable " _
          & "WHERE PARTREF='" & Compress(sPartNumber) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDsc)
   If bSqlRows Then
      With RdoDsc
         GetDesc = .Fields(1)
         ClearResultSet RdoDsc
      End With
   Else
      GetDesc = ""
   End If
   Set RdoDsc = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getdesc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
