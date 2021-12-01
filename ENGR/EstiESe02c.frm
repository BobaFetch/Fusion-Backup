VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form EstiESe02c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimating Bill Of Material"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "EstiESe02c.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESe02c.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtAssy 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CheckBox optExp 
      Alignment       =   1  'Right Justify
      Caption         =   "Open Expanded"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   "Expands The Tree On Selection"
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.ListBox lstNodes 
      Height          =   2595
      Left            =   8000
      TabIndex        =   28
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CheckBox optPls 
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox optRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   5640
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame zFrame 
      Height          =   4575
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   7550
      Begin VB.CommandButton optPrn 
         DownPicture     =   "EstiESe02c.frx":0AB8
         Height          =   320
         Left            =   6600
         Picture         =   "EstiESe02c.frx":0C42
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Print This Form"
         Top             =   2880
         Width           =   350
      End
      Begin VB.CommandButton cmdPrt 
         Height          =   315
         Left            =   7080
         Picture         =   "EstiESe02c.frx":0DCC
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "New Part Numbers"
         Top             =   2880
         Width           =   350
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   6600
         TabIndex        =   5
         ToolTipText     =   "Add A Part To The Current Selection"
         Top             =   240
         Width           =   875
      End
      Begin VB.CommandButton cmdDelete 
         Cancel          =   -1  'True
         Caption         =   "Delete"
         Height          =   315
         Left            =   6600
         TabIndex        =   10
         ToolTipText     =   "Delete The Selected Item"
         Top             =   2040
         Width           =   875
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   315
         Left            =   6600
         TabIndex        =   8
         ToolTipText     =   "Copy From One Parts List To Another"
         Top             =   1320
         Width           =   875
      End
      Begin ComctlLib.TreeView tvw1 
         Height          =   3855
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Double Click Items For Detail"
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   6800
         _Version        =   327682
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imlSmallIcons"
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   315
         Left            =   6600
         TabIndex        =   11
         ToolTipText     =   "Refresh The List"
         Top             =   2400
         Width           =   875
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   6600
         TabIndex        =   9
         ToolTipText     =   "Paste A Copied Or Cut Selection"
         Top             =   1680
         Width           =   875
      End
      Begin VB.CommandButton cmdCut 
         Caption         =   "Cut"
         Height          =   315
         Left            =   6600
         TabIndex        =   7
         ToolTipText     =   "Cut From One Parts List Then Paste To A Selected Item"
         Top             =   960
         Width           =   875
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   315
         Left            =   6600
         TabIndex        =   6
         ToolTipText     =   "Edit The Current Selection"
         Top             =   600
         Width           =   875
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "&Quit"
         Height          =   315
         Left            =   6600
         TabIndex        =   13
         ToolTipText     =   "End You Work On This Bill"
         Top             =   3720
         Visible         =   0   'False
         Width           =   875
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Type"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   25
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lblLvl 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5880
         TabIndex        =   20
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label lblDsc 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   4200
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.ComboBox cmbRev 
      Height          =   315
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Revision "
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmbPls 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   720
      Width           =   3345
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Expand"
      Height          =   315
      Left            =   6840
      TabIndex        =   3
      ToolTipText     =   "Expand The Tree"
      Top             =   720
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   5520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6465
      FormDesignWidth =   7770
   End
   Begin VB.Label lblTopLevel 
      Height          =   255
      Left            =   4800
      TabIndex        =   35
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblUnits 
      Height          =   255
      Left            =   3720
      TabIndex        =   34
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Updating 
      BackStyle       =   0  'Transparent
      Caption         =   "Updating."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   120
      Picture         =   "EstiESe02c.frx":124E
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   480
      Picture         =   "EstiESe02c.frx":15D8
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Active 
      Height          =   210
      Left            =   3360
      Picture         =   "EstiESe02c.frx":1962
      Top             =   1440
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Bill Of Material Is Unique To This Estimate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   31
      Top             =   45
      Width           =   4215
   End
   Begin VB.Label lblBid 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   30
      Top             =   360
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   29
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label txtRev 
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label txtPls 
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   6360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblLevel 
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   0
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EstiESe02c.frx":1CEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EstiESe02c.frx":1FAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EstiESe02c.frx":2348
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EstiESe02c.frx":26E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EstiESe02c.frx":29D8
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "EstiESe02c.frx":303A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "EstiESe02c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'New 7/17/03
'10/24/03 Added "key:" to TreeView Key work around to bug
'4/19/05 Added iKey to keep TreeView from ignoring duplicates
Option Explicit
Dim AdoCmdObj1 As ADODB.Command
Dim AdoCmdObj2 As ADODB.Command
'Dim RdoQry1 As rdoQuery
'Dim RdoQry2 As rdoQuery
Dim tNode As Node

Dim bGoodPart As Byte
Dim bGoodRev As Byte
Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bCopy As Byte
Dim bCut As Byte
Dim bLevel As Byte

Dim iCounter As Integer
Dim iCurrIdx As Integer
Dim iKey As Integer

Dim sCurrPart As String

Dim sNewPart As String
Dim sNewRev As String
Dim sNewUon As String
Dim sOldPart As String
Dim sOldRev As String
Dim sOldUon As String

'Passed
Dim sCompart As String
Dim sPartNum As String
Dim sPADESC As String

'''Material Totals
''Dim cBidQuantity As Currency
''Dim cBidBurden As Currency
''Dim cBidMaterial As Currency
''Dim cBidTotMat As Currency

Dim sBillParts(700, 7) As String
'0 = Compressed Part
'1 = Revision
'2 = Description
'3 = Level
'4 = Part Number
'5 = Compressed Used On
'6 = Bom Level

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetUom() As String
   Dim RdoUom As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PAUNITS,PALEVEL FROM PartTable where " _
          & "PARTREF='" & Compress(cmbPls) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoUom, ES_FORWARD)
   If bSqlRows Then
      GetUom = "" & Trim(RdoUom!PAUNITS)
      lblTopLevel = Format(RdoUom!PALEVEL, "0")
   Else
      GetUom = ""
   End If
   Set RdoUom = Nothing
   Exit Function
   
DiaErr1:
   GetUom = ""
   lblTopLevel = 0
End Function

Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiEngr", "EstiESe02c", Trim(optExp.Value)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "EstiESe02c", sOptions)
   If Len(sOptions) > 0 Then optExp.Value = Val(sOptions) _
          Else optExp.Value = vbChecked
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   Updating.ForeColor = ES_BLUE
   
End Sub

Private Sub Active_Click()
   txtAssy_Click
   
End Sub

Private Sub cmbPls_Click()
   FillBomhRev cmbPls
   bGoodRev = 1
   
End Sub


Private Sub cmbPls_LostFocus()
   cmbPls = CheckLen(cmbPls, 30)
   bGoodRev = 1
   
End Sub


Private Sub cmbRev_Change()
   If cmbRev <> cmbRev.List(0) Then bGoodRev = 0 _
                            Else bGoodRev = 1
   
End Sub

Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   cmbRev = Compress(cmbRev)
   
End Sub


Private Sub cmdAdd_Click()
   bCopy = 0
   bCut = 0
   sNewPart = ""
   sNewRev = ""
   sNewUon = ""
   sOldPart = ""
   sOldRev = ""
   sOldUon = ""
   cmdPaste.Enabled = False
   If Val(lblLvl) > 3 Then
      MsgBox "A Part Type 4 Cannot Have Lower Level Parts.", _
         vbInformation, Caption
      Exit Sub
   End If
   If Trim(sBillParts(iCurrIdx, 5)) = "" Then
      MsgBox "You Must Select An Item To Edit Items.", _
         vbInformation, Caption
   Else
      CloseSwitches
      txtPls = sBillParts(iCurrIdx, 5)
      txtRev = sBillParts(iCurrIdx, 1)
      EstiESe02f.lblBid = lblBid
      EstiESe02f.lblAssy = sBillParts(iCurrIdx, 4)
      EstiESe02f.lblBomLevel = sBillParts(iCurrIdx, 6) + 1
      EstiESe02f.Show
   End If
   
End Sub

Private Sub cmdCan_Click()
   Dim b As Byte
   MouseCursor 11
   'did they forget something?
   For b = 0 To Forms.Count - 1
      If Forms(b).Name = "EstiESe02f" Then Unload Forms(b)
   Next
   TotalBidMatl CLng(lblBid), Compress(cmbPls), CCur("0" & EstiESe02a.txtQty)
   Unload Me
   
End Sub

Private Sub cmdCopy_Click()
   bCut = 0
   If Trim(sBillParts(iCurrIdx, 0)) = "" Then
      MsgBox "You Must Select An Item To Copy.", _
         vbInformation, Caption
      bCopy = 0
   Else
      cmdPaste.Enabled = True
      sPartNum = sBillParts(iCurrIdx, 4)
      sOldPart = sBillParts(iCurrIdx, 0)
      sOldRev = sBillParts(iCurrIdx, 1)
      sOldUon = sBillParts(iCurrIdx, 5)
      bCopy = 1
   End If
   
End Sub

Private Sub cmdCut_Click()
   bCopy = 0
   If Trim(sBillParts(iCurrIdx, 0)) = "" Then
      MsgBox "You Must Select An Item To Cut.", _
         vbInformation, Caption
      bCut = 0
   Else
      sPartNum = sBillParts(iCurrIdx, 4)
      sOldPart = sBillParts(iCurrIdx, 0)
      sOldRev = sBillParts(iCurrIdx, 1)
      sOldUon = sBillParts(iCurrIdx, 5)
      bCut = 1
      cmdPaste.Enabled = True
   End If
   
End Sub

Private Sub cmdDelete_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   bCopy = 0
   bCut = 0
   sNewPart = ""
   sNewRev = ""
   sNewUon = ""
   sOldPart = ""
   sOldRev = ""
   sOldUon = ""
   cmdPaste.Enabled = False
   If Trim(sBillParts(iCurrIdx, 0)) = "" Then
      MsgBox "You Must Select An Item To Delete.", _
         vbInformation, Caption
   Else
      If Trim(sBillParts(iCurrIdx, 5)) = "" Then
         MsgBox "No, You May Not Delete The Entire Bill.", _
            vbInformation, Caption
      Else
         sPartNum = ""
         sOldPart = sBillParts(iCurrIdx, 4)
         sOldUon = sBillParts(iCurrIdx, 5)
         sMsg = "This Function Removes The Select Item And Items Attached." & vbCrLf _
                & "This Function Cannot Be Reversed.  Please Confirm That" & vbCrLf _
                & "You Wish To Delete " & sOldPart & "."
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         If bResponse = vbYes Then
            On Error Resume Next
            Err = 0
            clsADOCon.ADOErrNum = 0
            sSql = "DELETE FROM EsbmTable WHERE (BIDBOMASSYPART='" & sOldUon & "' " _
                   & "AND BIDBOMPARTREF='" & Compress(sOldPart) & "' AND BIDBOMREF=" & Val(lblBid) & ") "
            clsADOCon.ExecuteSQL sSql ' rdExecDirect
            If clsADOCon.ADOErrNum = 0 Then
               optRefresh.Value = vbChecked
               SysMsg sOldPart & " Was Deleted.", True
            Else
               SysMsg "Couldn't Delete " & sOldPart & ".", True
            End If
         Else
            CancelTrans
         End If
      End If
   End If
   
End Sub

Private Sub cmdEdit_Click()
   bCopy = 0
   bCut = 0
   sNewPart = ""
   sNewRev = ""
   sNewUon = ""
   sOldPart = ""
   sOldRev = ""
   sOldUon = ""
   cmdPaste.Enabled = False
   If Trim(sBillParts(iCurrIdx, 0)) = "" Then
      MsgBox "You Must Select An Item To Edit.", _
         vbInformation, Caption
   Else
      CloseSwitches
      txtPls = sBillParts(iCurrIdx, 4)
      txtRev = sBillParts(iCurrIdx, 1)
      EstiESe02g.lblBid = lblBid
      EstiESe02g.lblAssy = sBillParts(iCurrIdx, 5)
      'EstiESe02g.lblRev = Trim(cmbRev)
      EstiESe02g.cmbPrt = sBillParts(iCurrIdx, 0)
      EstiESe02g.Show
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3201
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdPaste_Click()
   If sBillParts(iCurrIdx, 0) = "" Then
      MsgBox "You Must Select A Valid Paste Destination.", _
         vbInformation, Caption
   Else
      sNewPart = sBillParts(iCurrIdx, 0)
      sNewRev = sBillParts(iCurrIdx, 1)
      sNewUon = sBillParts(iCurrIdx, 5)
      If sNewUon & sNewPart = sOldUon & sOldPart Then
         MsgBox "You Can't Paste A To The Same Destination.", _
            vbInformation, Caption
         Exit Sub
      End If
      If sNewPart = sOldPart Then
         MsgBox "You Must Select A Valid Paste Destination.", _
            vbInformation, Caption
         Exit Sub
      End If
      If bCut = 1 Then
         PasteCut
      Else
         If bCopy = 1 Then PasteCopy
      End If
   End If
   
End Sub

Private Sub cmdPrt_Click()
   InvcINe01a.Show
   cmdPrt.Value = False
   
End Sub

Private Sub cmdQuit_Click()
   On Error Resume Next
   tvw1.Nodes.Clear
   bCopy = 0
   bCut = 0
   sNewPart = ""
   sNewRev = ""
   sNewUon = ""
   sOldPart = ""
   sOldRev = ""
   sOldUon = ""
   lblDsc(1) = ""
   lblLvl = ""
   txtAssy.Visible = False
   Active.Visible = False
   'cmbPls.Enabled = True
   'cmbRev.Enabled = True
   'zFrame.Enabled = False
   'CloseSwitches
   'cmdRefresh.Enabled = False
   'cmbPls.SetFocus
   
End Sub

Private Sub cmdRefresh_Click()
   cmdPaste.Enabled = False
   bCopy = 0
   bCut = 0
   sNewPart = ""
   sNewRev = ""
   sNewUon = ""
   sOldPart = ""
   sOldRev = ""
   sOldUon = ""
   optRefresh.Value = vbChecked
   
End Sub

Private Sub cmdSel_Click()
   ExpandTree
   '    If bGoodPart = 1 And bGoodRev = 1 Then
   '        On Error Resume Next
   '        zFrame.Enabled = True
   '        cmdAdd.Enabled = True
   '        cmdEdit.Enabled = True
   '        cmdQuit.Enabled = True
   '        cmdCut.Enabled = True
   '        cmdCopy.Enabled = True
   '        cmdDelete.Enabled = True
   '        cmdRefresh.Enabled = True
   '        cmbPls.Enabled = False
   '        cmbRev.Enabled = False
   '        cmdSel.Enabled = False
   '        cmdAdd.SetFocus
   '        FillTree
   '    Else
   '        MsgBox "Either The Part Or The Revision Was Not Found.", _
   '            vbInformation, Caption
   '    End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      'On Error Resume Next
      FillTree
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   GetOptions
   sSql = "SELECT BIDBOMREF,BIDBOMASSYPART,BIDBOMPARTREF,BIDBOMLEVEL," _
          & "BIDBOMSEQUENCE,BIDBOMQTYREQD,BIDBOMUNITS,BIDBOMCONVERSION " _
          & "FROM EsbmTable WHERE (BIDBOMREF= ? AND BIDBOMASSYPART= ? ) " _
          & "ORDER BY BIDBOMSEQUENCE,BIDBOMPARTREF "
   
   Set AdoCmdObj1 = New ADODB.Command
   AdoCmdObj1.CommandText = sSql
   
   Dim prmBomRef As ADODB.Parameter
   Set prmBomRef = New ADODB.Parameter
   prmBomRef.Type = adInteger
   AdoCmdObj1.Parameters.Append prmBomRef
   
   Dim prmAssPrt As ADODB.Parameter
   Set prmAssPrt = New ADODB.Parameter
   prmAssPrt.Type = adChar
   prmAssPrt.Size = 30
   AdoCmdObj1.Parameters.Append prmAssPrt
   'Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable WHERE " _
          & "PARTREF= ? "
   Set AdoCmdObj2 = New ADODB.Command
   AdoCmdObj2.CommandText = sSql
   
   Dim prmPrtRef As ADODB.Parameter
   Set prmPrtRef = New ADODB.Parameter
   prmPrtRef.Type = adChar
   prmPrtRef.Size = 30
   AdoCmdObj2.Parameters.Append prmPrtRef
   'Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   
   
   bOnLoad = 1
   
End Sub

Private Sub FillTree()
   Dim RdoBom As ADODB.Recordset
   Dim iList As Integer
   
   Dim cQtyReq As Currency
   
   Dim sPl1 As String
   Dim sPl2 As String
   Dim sRev As String
   
   MouseCursor 11
   tvw1.Nodes.Clear
   lstNodes.Clear
   Erase sBillParts
   iKey = 0
   sRev = Trim(cmbRev)
   sPl1 = "" & Compress(cmbPls)
   sPl2 = Compress(cmbPls)
   GetPartInfo (sPl1)
   lblDsc(0) = sPADESC
   tvw1.ToolTipText = "Click Items For Detail"
   optRefresh.Value = vbUnchecked
   On Error GoTo DiaErr1
   lblUnits = GetUom()
   sBillParts(0, 0) = sPl1
   sBillParts(0, 2) = "" & Trim(lblDsc(0))
   sBillParts(0, 4) = "" & Trim(cmbPls)
   sBillParts(0, 5) = sPl1
   sBillParts(0, 6) = 0
   txtAssy = cmbPls
   txtAssy.Visible = True
   Active.Visible = True
   AdoCmdObj1.Parameters(0).Value = Val(lblBid)
   AdoCmdObj1.Parameters(1).Value = Compress(cmbPls)
   bSqlRows = clsADOCon.GetQuerySet(RdoBom, AdoCmdObj1, ES_FORWARD)
   If bSqlRows Then
      'sBillParts(tNode.Index, 0) = sPl1
      'sBillParts(tNode.Index, 2) = "" & Trim(lblDsc(0))
      'sBillParts(tNode.Index, 4) = "" & Trim(cmbPls)
      'sBillParts(tNode.Index, 5) = sPl1
      'lstNodes.AddItem "" & Trim(cmbPls)
      With RdoBom
         Do Until .EOF
            bLevel = 0
            iKey = iKey + 1
            sCompart = !BIDBOMPARTREF
            GetPartInfo (sCompart)
            cQtyReq = !BIDBOMQTYREQD
            Set tNode = tvw1.Nodes.Add(, , "key:" & str$(iKey) & Trim(!BIDBOMASSYPART) & Trim(sCompart), "" & Trim(sPartNum), 2, 3)
            sBillParts(tNode.Index, 0) = "" & Trim(sCompart)
            sBillParts(tNode.Index, 2) = "" & Trim(sPADESC) _
                       & " Qty: " & Format$(cQtyReq, "#0.000") & " " & Trim(!BIDBOMUNITS)
            sBillParts(tNode.Index, 3) = GetPALevel(sCompart)
            sBillParts(tNode.Index, 4) = "" & Trim(sPartNum)
            sBillParts(tNode.Index, 5) = "" & Trim(!BIDBOMASSYPART)
            sBillParts(tNode.Index, 6) = 1      'bom level
            iList = tNode.Index
            NextBillLevel iList, Trim(sCompart), 2
            .MoveNext
         Loop
         ClearResultSet RdoBom
         If optExp.Value = vbChecked Then ExpandTree
      End With
   End If
   MouseCursor 0
   On Error Resume Next
   Set RdoBom = Nothing
   lblDsc(1) = cmbPls & " Qty: 1 " & lblUnits
   lblLvl = lblTopLevel
   Active.Picture = Chkyes.Picture
   txtAssy.SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "filltree"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bByte As Byte
   SaveOptions
   bByte = GetBidLabor(Compress(cmbPls), Val(lblBid), CCur("0" & EstiESe02a.txtQty))
   Sleep 500
   MouseCursor 0
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set AdoCmdObj1 = Nothing
   Set AdoCmdObj2 = Nothing
   Set EstiESe02c = Nothing
   
End Sub


Private Sub optPrn_Click()
   PrintForm
   optPrn.Value = False
   
End Sub

Private Sub optRefresh_Click()
   If optRefresh.Value = vbChecked Then FillTree
   
End Sub

Private Sub tvw1_Click()
   Active.Picture = Chkno.Picture
   If iCurrIdx = 0 Then
      cmdEdit.Enabled = False
      cmdCut.Enabled = False
      cmdCopy.Enabled = False
      cmdDelete.Enabled = False
   Else
      cmdEdit.Enabled = True
      cmdCut.Enabled = True
      cmdCopy.Enabled = True
      cmdDelete.Enabled = True
   End If
   If Val(sBillParts(iCurrIdx, 3)) > 3 Then cmdAdd.Enabled = False _
          Else cmdAdd.Enabled = True
   If Trim(sBillParts(iCurrIdx, 0)) = "" Then iCurrIdx = 0
   lblDsc(1) = sBillParts(iCurrIdx, 2)
   lblLvl = sBillParts(iCurrIdx, 3)
   
End Sub

Private Sub tvw1_Collapse(ByVal Node As ComctlLib.Node)
   Node.Image = 1
   iCurrIdx = Node.Index
   
End Sub

Private Sub tvw1_Expand(ByVal Node As ComctlLib.Node)
   Node.Image = 4
   iCurrIdx = Node.Index
   tvw1.ToolTipText = ""
   
End Sub

Private Sub tvw1_NodeClick(ByVal Node As ComctlLib.Node)
   sCurrPart = Compress(Node.Text)
   iCurrIdx = Node.Index
   'If Len(sCurrPart) Then NextBillLevel Node.Index, sBillParts(Node.Index, 0), sBillParts(Node.Index, 1)
   
End Sub



Private Sub NextBillLevel(iNode As Integer, sCompart As String, BomLevel As Integer)
   Dim iList As Integer
   Dim cQtyReq As Currency
   Dim RdoBm1 As ADODB.Recordset
   Static sAssy As String
   
   On Error GoTo DiaErr1
   AdoCmdObj1.Parameters(0).Value = Val(lblBid)
   AdoCmdObj1.Parameters(1).Value = sCompart
   bSqlRows = clsADOCon.GetQuerySet(RdoBm1, AdoCmdObj1, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBm1
         Do Until .EOF
            iKey = iKey + 1
            sCompart = !BIDBOMPARTREF
            GetPartInfo (sCompart)
            cQtyReq = !BIDBOMQTYREQD
            Set tNode = tvw1.Nodes.Add(iNode, tvwChild, "key:" & str$(iKey) & Trim(!BIDBOMASSYPART) & Trim(sCompart), Trim(sPartNum), 2, 3)
            If Err > 0 Then Exit Do
            If sAssy <> "" & Trim(!BIDBOMASSYPART) Then bLevel = bLevel + 1
            sAssy = "" & Trim(!BIDBOMASSYPART)
            iList = tNode.Index
            sBillParts(tNode.Index, 0) = "" & Trim(sCompart)
            sBillParts(tNode.Index, 2) = "" & Trim(sPADESC) _
                       & " Qty: " & Format$(cQtyReq, "#0.000") & " " & Trim(!BIDBOMUNITS)
            sBillParts(tNode.Index, 3) = GetPALevel(sCompart)
            sBillParts(tNode.Index, 4) = "" & Trim(sPartNum)
            sBillParts(tNode.Index, 5) = "" & Trim(!BIDBOMASSYPART)
            sBillParts(tNode.Index, 6) = BomLevel
            NextBillLevel iList, Trim(sCompart), BomLevel + 1
            .MoveNext
         Loop
         ClearResultSet RdoBm1
      End With
   End If
   Set RdoBm1 = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getnextbilll"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub

Private Sub PasteCopy()
   Dim RdoPls As ADODB.Recordset
   Dim RdoNew As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT BMASSYPART,BMPARTREF From BmplTable " _
          & "WHERE (BMASSYPART='" & sNewPart & "' AND " _
          & "BMPARTREF='" & sOldPart & "' AND BMREV='" & sNewRev & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_FORWARD)
   If bSqlRows Then
      MsgBox "You May Not Paste At " & sNewPart & " It Will " & vbCrLf _
         & "Create A Duplicate Key.", vbInformation, Caption
   Else
      sMsg = "You May Create A Circular Bill By Pasting " & vbCrLf _
             & "Are You Certain That You Want To Paste Here?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         'insert if
         sSql = "SELECT * FROM EsbmTable WHERE (BIDBOMASSYPART='" _
                & sOldUon & "' AND BIDBOMPARTREF='" & sOldPart _
                & "' AND BIDBOMREF=" & Val(lblBid) & ")"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_FORWARD)
         If bSqlRows Then
            With RdoPls
               sSql = "INSERT INTO EsbmTable (" _
                      & "BIDBOMREF,BIDBOMASSYPART,BIDBOMPARTREF," _
                      & "BIDBOMQTYREQD,BIDBOMUNITS,BIDBOMCONVERSION," _
                      & "BIDBOMSEQUENCE,BIDBOMESTUNITCOST,BIDBOMADDER," _
                      & "BIDBOMSETUP,BIDBOMCOMT,BIDBOMLABOR,BIDBOMLABOROH," _
                      & "BIDBOMLABORHRS,BIDBOMMATERIAL,BIDBOMMATERIALBRD) " _
                      & "VALUES(" & Val(lblBid) & ",'" & sNewPart & "','" & sOldPart & "'," _
                      & !BIDBOMQTYREQD & ",'" & !BIDBOMUNITS & "'," _
                      & !BIDBOMCONVERSION & "," & !BIDBOMSEQUENCE & "," & !BIDBOMESTUNITCOST & "," _
                      & !BIDBOMADDER & "," & !BIDBOMSETUP & ",'" & Trim(!BIDBOMCOMT) & "'," _
                      & !BIDBOMLABOR & "," & !BIDBOMLABOROH & "," & !BIDBOMLABORHRS & "," _
                      & !BIDBOMMATERIAL & "," & !BIDBOMMATERIALBRD & ")"
               clsADOCon.ExecuteSQL sSql ' rdExecDirect
               cmdPaste.Enabled = False
               If Err > 0 Then
                  MsgBox "Couldn't Make The Copy", _
                     vbExclamation, Caption
               Else
                  optRefresh.Value = vbChecked
                  SysMsg "Selection Was Copied.", True
               End If
               ClearResultSet RdoPls
            End With
         End If
      Else
         CancelTrans
      End If
   End If
   Set RdoPls = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "NextBillLev"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PasteCut()
   Dim RdoPls As ADODB.Recordset
   Dim RdoNew As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT BIDBOMASSYPART,BIDBOMPARTREF From EsbmTable " _
          & "WHERE (BIDBOMASSYPART='" & sNewPart & "' AND " _
          & "BIDBOMPARTREF='" & sOldPart & "' AND BIDBOMREF=" & Val(lblBid) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_FORWARD)
   If bSqlRows Then
      MsgBox "You May Not Paste At " & sNewPart & " It Will " & vbCrLf _
         & "Create A Duplicate Key.", vbInformation, Caption
   Else
      sMsg = "You May Create A Circular Bill By Cutting And Pasting" & vbCrLf _
             & "Here Are You Certain That You Want To Paste Here?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         'insert it
         On Error Resume Next
         sSql = "SELECT * FROM EsbmTable WHERE (BIDBOMASSYPART='" _
                & sOldUon & "' AND BIDBOMPARTREF='" & sOldPart _
                & "' AND BIDBOMREF=" & Val(lblBid) & ")"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_FORWARD)
         If bSqlRows Then
            With RdoPls
               sSql = "INSERT INTO EsbmTable (" _
                      & "BIDBOMREF,BIDBOMASSYPART,BIDBOMPARTREF," _
                      & "BIDBOMQTYREQD,BIDBOMUNITS,BIDBOMCONVERSION," _
                      & "BIDBOMSEQUENCE,BIDBOMESTUNITCOST,BIDBOMADDER," _
                      & "BIDBOMSETUP,BIDBOMCOMT,BIDBOMLABOR,BIDBOMLABOROH," _
                      & "BIDBOMLABORHRS,BIDBOMMATERIAL,BIDBOMMATERIALBRD) " _
                      & "VALUES(" & Val(lblBid) & ",'" & sNewPart & "','" & sOldPart & "'," _
                      & !BIDBOMQTYREQD & ",'" & !BIDBOMUNITS & "'," _
                      & !BIDBOMCONVERSION & "," & !BIDBOMSEQUENCE & "," & !BIDBOMESTUNITCOST & "," _
                      & !BIDBOMADDER & "," & !BIDBOMSETUP & ",'" & Trim(!BIDBOMCOMT) & "'," _
                      & !BIDBOMLABOR & "," & !BIDBOMLABOROH & "," & !BIDBOMLABORHRS & "," _
                      & !BIDBOMMATERIAL & "," & !BIDBOMMATERIALBRD & ")"
               
               clsADOCon.ExecuteSQL sSql ' rdExecDirect
               cmdPaste.Enabled = False
               ClearResultSet RdoPls
               If Err > 0 Then
                  MsgBox "Couldn't Make The Cut And Paste.", _
                     vbExclamation, Caption
               Else
                  sSql = "DELETE FROM EsbmTable WHERE (BIDBOMASSYPART='" & sOldUon & "' " _
                         & "AND BIDBOMPARTREF='" & sOldPart & "' AND BIDBOMREF=" & Val(lblBid) & ")"
                  clsADOCon.ExecuteSQL sSql ' rdExecDirect
                  optRefresh.Value = vbChecked
                  SysMsg "Selection Was Cut And Copied.", True
               End If
            End With
         End If
      Else
         CancelTrans
      End If
   End If
   Set RdoPls = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "NextBillLev"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub CloseSwitches()
   Exit Sub
   cmdQuit.Enabled = False
   cmdAdd.Enabled = False
   cmdEdit.Enabled = False
   cmdCut.Enabled = False
   cmdCut.Enabled = False
   cmdCopy.Enabled = False
   cmdDelete.Enabled = False
   
End Sub

Private Sub ExpandTree()
   Dim iList As Integer
   On Error Resume Next
   If tvw1.Nodes.Count > 1 Then
      For iList = 1 To tvw1.Nodes.Count
         tvw1.Nodes(iList).Expanded = True
      Next
   End If
   
End Sub


Private Sub GetPartInfo(sPart As String)
   Dim RdoInf As ADODB.Recordset
   AdoCmdObj2.Parameters(0).Value = sPart
   bSqlRows = clsADOCon.GetQuerySet(RdoInf, AdoCmdObj2, ES_FORWARD)
   If bSqlRows Then
      With RdoInf
         sPartNum = "" & Trim(!PartNum)
         sPADESC = "" & Trim(!PADESC)
         ClearResultSet RdoInf
      End With
   Else
      sPartNum = ""
      sPADESC = ""
   End If
   
End Sub

Private Sub txtAssy_Click()
   optRefresh.Value = vbChecked
   iCurrIdx = 0
   optRefresh.Value = vbChecked
   Active.Picture = Chkyes.Picture
   cmdEdit.Enabled = False
   cmdCut.Enabled = False
   cmdCopy.Enabled = False
   cmdDelete.Enabled = False
   If Val(sBillParts(iCurrIdx, 3)) > 3 Then cmdAdd.Enabled = False _
          Else cmdAdd.Enabled = True
   lblDsc(1) = cmbPls & " Qty: 1 " & lblUnits
   lblLvl = lblTopLevel
   
End Sub



'Private Sub TotalBidMatl(EstQty As Currency)
'   Dim rdoMat As ADODB.Recordset
'   Dim cBurden As Currency
'   Dim cConvert As Currency
'   Dim cQuantity As Currency
'
'   iCounter = 0
'   cBidBurden = 0
'   cBidMaterial = 0
'   cBidTotMat = 0
'   Updating.Visible = True
'   MouseCursor 11
'   Updating.Refresh
'   On Error GoTo DiaErr1
'   sSql = "SELECT BIDBOMREF,BIDBOMASSYPART,BIDBOMPARTREF,BIDBOMQTYREQD," _
'          & "BIDBOMCONVERSION,BIDBOMSETUP,BIDBOMADDER,BIDBOMMATERIAL," _
'          & "BIDBOMMATERIALBRD,BIDBOMESTUNITCOST FROM EsbmTable " _
'          & "WHERE (BIDBOMREF=" & Val(lblBid) & " AND " _
'          & "BIDBOMASSYPART='" & Compress(cmbPls) & "')"
'   bSqlRows = clsADOCon.GetDataSet(sSql,rdoMat, ES_FORWARD)
'   If bSqlRows Then
'      With rdoMat
'         Do Until .EOF
'            sProcName = "totalbidmatl"
'            cBidQuantity = 1
'            iCounter = iCounter + 1
'            cConvert = Format(!BIDBOMCONVERSION, ES_QuantityDataFormat)
'            If cConvert = 0 Then cConvert = 1
'            cBurden = Format(!BIDBOMMATERIALBRD, ES_QuantityDataFormat)
'            If cBurden > 0 Then cBurden = cBurden / 100
'            cQuantity = cQuantity / cConvert
'            cQuantity = Format((!BIDBOMQTYREQD * cBidQuantity), ES_QuantityDataFormat)
'            cBurden = (cBurden * (!BIDBOMMATERIAL * !BIDBOMQTYREQD))
'            cBidBurden = cBidBurden + cBurden
'            cBidMaterial = cBidMaterial + ((!BIDBOMESTUNITCOST - cBurden) * cBidQuantity)
'            cBidTotMat = cBidTotMat + (!BIDBOMESTUNITCOST * cBidQuantity)
'            cBidQuantity = Format(!BIDBOMQTYREQD, ES_QuantityDataFormat)
'            TotalBidNextMatl Trim(!BIDBOMPARTREF), cBidQuantity, iCounter
'            .MoveNext
'         Loop
'         ClearResultSet rdoMat
'      End With
'   End If
'   Set rdoMat = Nothing
'   EstiESe02a.lblMaterial = Format(cBidMaterial, ES_MoneyFormat)
'   EstiESe02a.lblBurden = Format(cBidBurden, ES_MoneyFormat)
'   EstiESe02a.lblTotMat = Format(cBidTotMat, ES_MoneyFormat)
'   EstiESe02a.lblEstTotMatl = Format(cBidTotMat * EstQty, ES_MoneyFormat)
'   Exit Sub
'
'DiaErr1:
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
'Private Sub TotalBidNextMatl(AssemblyPart As String, BidQuantity As Currency, _
'                             iCounter As Integer)
'   Dim RdoNextMat As ADODB.Recordset
'   Dim cBurden As Currency
'   Dim cConvert As Currency
'   Dim cQuantity As Currency
'   sSql = "SELECT BIDBOMREF,BIDBOMASSYPART,BIDBOMPARTREF,BIDBOMQTYREQD," _
'          & "BIDBOMCONVERSION,BIDBOMSETUP,BIDBOMADDER,BIDBOMMATERIAL," _
'          & "BIDBOMMATERIALBRD,BIDBOMESTUNITCOST FROM EsbmTable " _
'          & "WHERE (BIDBOMREF=" & Val(lblBid) & " AND " _
'          & "BIDBOMASSYPART='" & AssemblyPart & "')"
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoNextMat, ES_FORWARD)
'   If bSqlRows Then
'      With RdoNextMat
'         Do Until .EOF
'            sProcName = "totalbidnextmatl"
'            iCounter = iCounter + 1
'            cConvert = Format(!BIDBOMCONVERSION, ES_QuantityDataFormat)
'            If cConvert = 0 Then cConvert = 1
'            cBurden = Format(!BIDBOMMATERIALBRD, ES_QuantityDataFormat)
'            If cBurden > 0 Then cBurden = cBurden / 100
'            cQuantity = cQuantity / cConvert
'            cQuantity = Format((!BIDBOMQTYREQD * cBidQuantity), ES_QuantityDataFormat)
'            cBurden = (cBurden * (!BIDBOMMATERIAL * !BIDBOMQTYREQD))
'            cBidBurden = cBidBurden + cBurden
'            cBidMaterial = cBidMaterial + ((!BIDBOMESTUNITCOST - cBurden) * cBidQuantity)
'            cBidTotMat = cBidTotMat + (!BIDBOMESTUNITCOST * cBidQuantity)
'            cBidQuantity = Format(!BIDBOMQTYREQD, ES_QuantityDataFormat)
'            TotalBidNextMatl Trim(!BIDBOMPARTREF), cBidQuantity, iCounter
'            .MoveNext
'         Loop
'         ClearResultSet RdoNextMat
'      End With
'   End If
'   Set RdoNextMat = Nothing
'
'End Sub
'

Private Function GetPALevel(BomPart) As String
   Dim RdoUom As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PALEVEL FROM PartTable where " _
          & "PARTREF='" & Compress(BomPart) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoUom, ES_FORWARD)
   If bSqlRows Then
      GetPALevel = Format(RdoUom!PALEVEL, "0")
   Else
      GetPALevel = 0
   End If
   Set RdoUom = Nothing
   Exit Function
   
DiaErr1:
   GetPALevel = 0
   
End Function
