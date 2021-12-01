VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form BompBMe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bill Of Material"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3201
   Icon            =   "BompBMe01a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optExp 
      Alignment       =   1  'Right Justify
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   6600
      TabIndex        =   43
      ToolTipText     =   "Expands The Tree On Selection"
      Top             =   1560
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.TextBox txtCmt 
      Height          =   1185
      Left            =   7800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Tag             =   "9"
      Text            =   "BompBMe01a.frx":030A
      ToolTipText     =   "Comment (5120 Chars Max)"
      Top             =   4800
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CheckBox chkUpChild 
      Alignment       =   1  'Right Justify
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   6840
      TabIndex        =   41
      ToolTipText     =   "Expands The Tree On Selection"
      Top             =   6480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Left            =   1320
      TabIndex        =   38
      Tag             =   "2"
      ToolTipText     =   "Approval Name"
      Top             =   1080
      Width           =   2085
   End
   Begin VB.ComboBox txtAte 
      Height          =   315
      Left            =   5280
      TabIndex        =   37
      Tag             =   "4"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMe01a.frx":0311
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
      TabIndex        =   31
      ToolTipText     =   "Top Level (Zero Level) Part Number"
      Top             =   1560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox lstNodes 
      Height          =   2595
      Left            =   8000
      TabIndex        =   29
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CheckBox optPls 
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox optRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame zFrame 
      Enabled         =   0   'False
      Height          =   5055
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   7550
      Begin VB.CommandButton cmdRel 
         Caption         =   "&Release"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   11
         ToolTipText     =   "Release (Unrelease) To Production"
         Top             =   3120
         Width           =   875
      End
      Begin VB.CommandButton cmdPhn 
         Caption         =   "&Assign"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Set This Parts List as Default for Part"
         Top             =   2760
         Width           =   875
      End
      Begin VB.CommandButton optPrn 
         DownPicture     =   "BompBMe01a.frx":0ABF
         Height          =   320
         Left            =   6600
         Picture         =   "BompBMe01a.frx":0C49
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Print This Form"
         Top             =   3840
         Width           =   350
      End
      Begin VB.CommandButton cmdPrt 
         Height          =   315
         Left            =   7080
         Picture         =   "BompBMe01a.frx":11DB
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "New Part Numbers"
         Top             =   3840
         Width           =   350
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   4
         ToolTipText     =   "Add A Part To The Current Selection"
         Top             =   600
         Width           =   875
      End
      Begin VB.CommandButton cmdDelete 
         Cancel          =   -1  'True
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   9
         ToolTipText     =   "Delete The Selected Item"
         Top             =   2400
         Width           =   875
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   7
         ToolTipText     =   "Copy From One Parts List To Another"
         Top             =   1680
         Width           =   875
      End
      Begin ComctlLib.TreeView tvw1 
         Height          =   4095
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Double Click Items For Detail"
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   7223
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
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   12
         ToolTipText     =   "Refresh The List"
         Top             =   3480
         Width           =   875
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   8
         ToolTipText     =   "Paste A Copied Or Cut Selection"
         Top             =   2040
         Width           =   875
      End
      Begin VB.CommandButton cmdCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   6
         ToolTipText     =   "Cut From One Parts List Then Paste To A Selected Item"
         Top             =   1320
         Width           =   875
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   5
         ToolTipText     =   "Edit The Current Selection"
         Top             =   960
         Width           =   875
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "&Quit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   14
         ToolTipText     =   "End You Work On This Bill"
         Top             =   240
         Width           =   875
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Type"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   26
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label lblLvl 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5880
         TabIndex        =   21
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblDsc 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   4560
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6840
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.ComboBox cmbRev 
      Height          =   315
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Revision (Add Or Revise)"
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox cmbPls 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   360
      Width           =   3345
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6840
      TabIndex        =   2
      Top             =   600
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   5880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7125
      FormDesignWidth =   7875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "App Date"
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   40
      Top             =   1110
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Approved By"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   39
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Released"
      Height          =   255
      Index           =   7
      Left            =   4800
      TabIndex        =   35
      ToolTipText     =   "Released To Production Or Not Released"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblRel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5640
      TabIndex        =   34
      ToolTipText     =   "Released To Production Or Not Released"
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblTopLevel 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2520
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUnits 
      Height          =   255
      Left            =   3840
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Active 
      Height          =   210
      Left            =   3360
      Picture         =   "BompBMe01a.frx":1676
      ToolTipText     =   "Click To Edit Top (Zero) Level"
      Top             =   1560
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   360
      Picture         =   "BompBMe01a.frx":1A00
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   120
      Picture         =   "BompBMe01a.frx":1D8A
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open Expanded"
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   30
      ToolTipText     =   "Expands The Tree On Selection"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label txtRev 
      Height          =   375
      Left            =   3000
      TabIndex        =   24
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label txtPls 
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   7200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblLevel 
      Height          =   255
      Left            =   5880
      TabIndex        =   22
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   17
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   15
      Top             =   720
      Width           =   3135
   End
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   0
      Top             =   5880
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
            Picture         =   "BompBMe01a.frx":2114
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BompBMe01a.frx":23D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BompBMe01a.frx":2770
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BompBMe01a.frx":2B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BompBMe01a.frx":2E00
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BompBMe01a.frx":3462
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "BompBMe01a"
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
'5/7/05 Added updates to BMHREVDATE column (all changes)
'5/8/05 Added Released and Assign functions to panel
'3/7/07 Add Parts - made qualifying Parts the same as in individual
'       Parts Lists
Option Explicit
Dim AdoCmdObj As ADODB.Command
Dim tNode As Node

Dim bGoodPart As Byte
Dim bGoodRev As Byte
Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bCopy As Byte
Dim bCut As Byte
Dim bLevel As Byte
Dim bBOMSec As Byte

Dim iCurrIdx As Integer
Dim iKey As Integer

Dim sCurrPart As String

Dim sPartNum As String
Dim sNewPart As String
Dim sNewRev As String
Dim sNewUon As String
Dim sOldPart As String
Dim sOldRev As String
Dim sOldUon As String

Private sBillParts(700, 6) As String
'0 = Compressed Part
'1 = Revision
'2 = Description
'3 = Level
'4 = Part Number
'5 = Compressed Used On

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiEngr", "BompBMe01a", Trim(optExp.Value)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "BompBMe01a", sOptions)
   If Len(sOptions) > 0 Then optExp.Value = Val(sOptions) _
          Else optExp.Value = vbChecked
   
End Sub

Private Function GetList() As Byte
   Dim RdoPls As ADODB.Recordset
   cmbRev.Clear
   cmbRev = ""
   
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV FROM PartTable " _
          & "WHERE PARTREF='" & Compress(cmbPls) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_FORWARD)
   If bSqlRows Then
      With RdoPls
         cmbPls = "" & Trim(!PartNum)
         lblDsc(0) = "" & Trim(!PADESC)
         lblLevel = "" & Trim(!PALEVEL)
         cmbRev = "" & Trim(!PABOMREV)
         ClearResultSet RdoPls
      End With
      GetList = 1
   Else
      lblDsc(0) = ""
      lblLevel = "0"
      
      MsgBox "Part number does not exist", vbExclamation, Caption
      'Debug.Print cmdCan.Value
      'If cmdCan.Value Then
      cmbPls.SetFocus
      'End If
      
      GetList = 0
      Exit Function
   End If
   If GetList = 1 Then
      cmdSel.Enabled = True
      FillBomhRev cmbPls
   Else
      cmdSel.Enabled = False
   End If
   lblUnits = GetUom()
   Set RdoPls = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getlist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtAssy.ForeColor = ES_BLUE
   
End Sub

Private Sub Active_Click()
   txtAssy_Click
   
End Sub

Private Sub cmbPls_Click()
   FillBomhRev cmbPls
   lblUnits = GetUom()
   bGoodRev = 1
   ' get approval names
   SetApproval
End Sub


Private Sub cmbPls_LostFocus()
   If Me.ActiveControl.Name = "cmdCan" Then
      Exit Sub
   End If
   
   If (Not ValidPartNumber(cmbPls.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPls = ""
      Exit Sub
   End If
   
   cmbPls = CheckLen(cmbPls, 30)
   bGoodPart = GetList()
   bGoodRev = 1
   
End Sub


Private Sub cmbRev_Change()
   If Len(cmbRev) > 4 Then cmbRev = Left(cmbRev, 4)
   If cmbRev <> cmbRev.List(0) Then bGoodRev = 0 _
                            Else bGoodRev = 1
   
End Sub

Private Sub cmbRev_LostFocus()
   If bGoodPart = 0 Then
      'MsgBox "Part number does not exist", vbExclamation, Caption
      Exit Sub
   End If
   cmbRev = CheckLen(cmbRev, 4)
   cmbRev = Compress(cmbRev)
   'bGoodRev = GetThisRevision()
      
   Dim b As Byte
   Dim iList As Integer
   'cmbRev = CheckLen(cmbRev, 6)
   For iList = 0 To (cmbRev.ListCount - 1)
      If cmbRev = cmbRev.List(iList) Then b = True
   Next
   
   If Not b Then
      If (bBOMSec = True) Then
         Dim bret As Boolean
         bret = CheckForBomSec(cmbRev.Text, False)
         If (bret = False) Then
            Exit Sub
         End If
         
         bGoodRev = GetThisRevision()
         ResetApproval
         Exit Sub
      
      End If
   End If
   
   bGoodRev = GetThisRevision()
   ' get approval names
   SetApproval
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
      MsgBox "Cannot Add A Part To A Part Type 4.", _
         vbInformation, Caption
      Exit Sub
   End If
   If Trim(sBillParts(iCurrIdx, 5)) = "" Then
      MsgBox "You Must Select An Item To Edit Items Too.", _
         vbInformation, Caption
   Else
      
      If ((bBOMSec = True) And (Trim(txtApp) <> "")) Then
         Dim bSec As Boolean
         Dim strLastRev As String
         
         strLastRev = cmbRev.List(cmbRev.ListCount - 1)
         bSec = CheckForBomSec(strLastRev, True)
         If (bSec = True) Then
            ' Change the Rev
            bGoodRev = GetThisRevision
            ResetApproval
         Else
            Exit Sub
         End If
      End If
      
      CloseSwitches
      txtPls = sBillParts(iCurrIdx, 4)
      ' TODO MM
      txtRev = Trim(cmbRev) 'sBillParts(iCurrIdx, 1)
      BompBMe01b.lblAssy = sBillParts(iCurrIdx, 4)
      BompBMe01b.lblRev = Trim(cmbRev) 'sBillParts(iCurrIdx, 1)
      BompBMe01b.Show
   End If
   
End Sub

Private Sub cmdCan_Click()
   
   Dim b As Byte
   'did they forget something?
   For b = 0 To Forms.Count - 1
      If Forms(b).Name = "BompBM02b" Then Unload Forms(b)
   Next
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
         sOldRev = sBillParts(iCurrIdx, 1)
         sOldUon = sBillParts(iCurrIdx, 5)
         sMsg = "This Function Removes The Select Item And Items Attached." & vbCrLf _
                & "This Function Cannot Be Reversed.  Please Confirm That" & vbCrLf _
                & "You Wish To Delete " & sOldPart & "."
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         If bResponse = vbYes Then
            On Error Resume Next
            clsADOCon.ADOErrNum = 0
            sSql = "DELETE FROM BmplTable WHERE (BMASSYPART='" & sOldUon & "' " _
                   & "AND BMPARTREF='" & Compress(sOldPart) & "' AND BMREV='" & sOldRev & "') "
            clsADOCon.ExecuteSql sSql 'rdExecDirect
            
            sSql = "UPDATE BmhdTable SET BMHREVDATE='" _
                   & Format(ES_SYSDATE, "mm/dd/yy") & "' WHERE " _
                   & "BMHREF='" & Compress(BompBMe01a.cmbPls) & "' " _
                   & "AND BMHREV='" & Trim(BompBMe01a.cmbRev) & "'"
            clsADOCon.ExecuteSql sSql 'rdExecDirect
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
      If ((bBOMSec = True) And (Trim(txtApp) <> "")) Then
         Dim bSec As Boolean
         Dim strLastRev As String
         
         strLastRev = cmbRev.List(cmbRev.ListCount - 1)
         bSec = CheckForBomSec(strLastRev, True)
         If (bSec = True) Then
            ' Change the Rev
            bGoodRev = GetThisRevision
            ResetApproval
         Else
            Exit Sub
         End If
      End If
      
      CloseSwitches
      txtPls = sBillParts(iCurrIdx, 4)
      txtRev = sBillParts(iCurrIdx, 1)
      BompBMe01c.lblAssy = sBillParts(iCurrIdx, 5)
      BompBMe01c.lblRev = Trim(cmbRev)
      BompBMe01c.cmbPrt = sBillParts(iCurrIdx, 0)
      BompBMe01c.Show
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

Private Sub cmdPhn_Click()
   Dim bResponse As Byte
   Dim sCurrPart As String
   Dim sMsg As String
   
   sMsg = "Assign As Default Parts List For" & vbCrLf _
          & cmbPls & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      sCurrPart = Compress(cmbPls)
      sSql = "UPDATE PartTable SET PABOMREV='" & cmbRev & "' " _
             & "WHERE PARTREF='" & sCurrPart & "' "
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      sSql = "UPDATE BmhdTable SET BMHPART='" & sCurrPart & "' " _
             & "WHERE BMHREF='" & sCurrPart & "' AND BMHREV='" _
             & cmbRev & "' "
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      If clsADOCon.RowsAffected > 0 Then
         SysMsg "Default Set.", True, Me
      Else
         MsgBox "Couldn't Assign.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   On Error Resume Next
   
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
   lblRel = ""
   lblDsc(1) = ""
   lblLvl = ""
   txtAssy.Visible = False
   Active.Visible = False
   cmbPls.Enabled = True
   cmbRev.Enabled = True
   zFrame.Enabled = False
   CloseSwitches
   cmdPhn.Enabled = True
   cmdRel.Enabled = True
   cmdRefresh.Enabled = False
   cmdCan.Enabled = True
   cmbPls.SetFocus
   
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

Private Sub cmdRel_Click()
   'Dim iList As Integer
   Dim sPartNumber As String
   Dim sPartRev As String
   
   sPartNumber = Compress(cmbPls)
   sPartRev = Compress(cmbRev)
   On Error GoTo DiaErr1
'   If lblRel = "N" Then iList = 1 Else iList = 0
'   If iList = 1 Then
'      sSql = "UPDATE BmhdTable SET BMHRELEASED=1,BMHRELEASEDATE='" _
'             & Format(Now, "mm/dd/yy") & "' WHERE BMHREF='" _
'             & sPartNumber & "' AND BMHREV='" & sPartRev & "' "
'   Else
'      sSql = "UPDATE BmhdTable SET BMHRELEASED=2,BMHRELEASEDATE=" _
'             & "Null WHERE BMHREF='" _
'             & sPartNumber & "' AND BMHREV='" & sPartRev & "' "
'   End If

   If lblRel = "N" Then
      sSql = "UPDATE BmhdTable SET BMHRELEASED=1,BMHRELEASEDATE='" _
             & Format(Now, "mm/dd/yy") & "' WHERE BMHREF='" _
             & sPartNumber & "' AND BMHREV='" & sPartRev & "' "
   Else
      sSql = "UPDATE BmhdTable SET BMHRELEASED=0,BMHRELEASEDATE=" _
             & "Null WHERE BMHREF='" _
             & sPartNumber & "' AND BMHREV='" & sPartRev & "' "
   End If

   clsADOCon.ExecuteSql sSql ', rdExecDirect
   'Sleep 500
   If clsADOCon.RowsAffected Then
      If lblRel = "N" Then
         lblRel = "Y"
         SysMsg "Parts Listed Was Released.", True, Me
      Else
         lblRel = "N"
         SysMsg "Parts Listed Was Unreleased.", True, Me
      End If
   Else
      MsgBox "Couldn't Find The Parts List Revision Record.", vbExclamation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "cmdrel_click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdSel_Click()
   If bGoodPart = 1 And bGoodRev = 1 Then
      On Error Resume Next
      cmdCan.Enabled = False
      zFrame.Enabled = True
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdQuit.Enabled = True
      cmdCut.Enabled = True
      cmdCopy.Enabled = True
      cmdDelete.Enabled = True
      cmdPhn.Enabled = True
      cmdRel.Enabled = True
      cmdRefresh.Enabled = True
      cmbPls.Enabled = False
      cmbRev.Enabled = False
      cmdSel.Enabled = False
      cmdAdd.SetFocus
      FillTree
   Else
      MsgBox "Either The Part Or The Revision Was Not Found.", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillPartsBelow4 cmbPls
      If cmbPls.ListCount > 0 Then bGoodPart = GetList()
      lblUnits = GetUom()
      
      bBOMSec = GetBOMSecurity()
      
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   sSql = "SELECT BMASSYPART,BMPARTREF,BMPARTNUM,BMREV," _
          & "BMPARTREV,BMSEQUENCE,BMQTYREQD,BMUNITS," _
          & "PARTREF,PARTNUM,PADESC,PALEVEL FROM BmplTable,PartTable " _
          & "WHERE BMPARTREF=PARTREF AND (BMASSYPART= ? AND " _
          & "BMREV= ? ) ORDER BY BMSEQUENCE,BMPARTREF "
   
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim prmAssPtr As ADODB.Parameter
   Set prmAssPtr = New ADODB.Parameter
   prmAssPtr.Type = adChar
   prmAssPtr.Size = 30
   AdoCmdObj.Parameters.Append prmAssPtr
   
   Dim prmBMRev As ADODB.Parameter
   Set prmBMRev = New ADODB.Parameter
   prmBMRev.Type = adChar
   prmBMRev.Size = 4
   AdoCmdObj.Parameters.Append prmBMRev
   
   bOnLoad = 1
   
End Sub



Private Sub FillTree()
   Dim RdoBom As ADODB.Recordset
   Dim bLen As Byte
   Dim A As Integer
   Dim iList As Integer
   Dim sPl1 As String
   Dim sPl2 As String
   Dim sRev As String
   
   MouseCursor 11
   tvw1.Nodes.Clear
   lstNodes.Clear
   Erase sBillParts
   iKey = 0
   sRev = Trim(cmbRev)
   sPl1 = cmbPls & " Rev: " & cmbRev
   sPl2 = "" & Compress(cmbPls)
   tvw1.ToolTipText = "Click Items For Detail"
   optRefresh.Value = vbUnchecked
   On Error GoTo DiaErr1
   sProcName = "getreleased"
   GetReleased
   sProcName = "filltree"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMPARTNUM,BMREV," _
          & "BMSEQUENCE,BMPARTREV,BMQTYREQD,BMUNITS," _
          & "PARTREF,PARTNUM,PADESC,PALEVEL FROM BmplTable,PartTable " _
          & "WHERE (BMPARTREF=PARTREF AND BMASSYPART='" & sPl2 & "' " _
          & "AND BMREV='" & Trim(cmbRev) & "') ORDER BY BMSEQUENCE,BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_STATIC)
   txtAssy = cmbPls
   txtAssy.Visible = True
   Active.Visible = True
   sBillParts(0, 0) = "" & Trim(Compress(cmbPls))
   sBillParts(0, 1) = "" & Trim(cmbRev)
   sBillParts(0, 2) = "" & Trim(lblDsc(0))
   sBillParts(0, 3) = "" & str$(lblLevel)
   sBillParts(0, 4) = "" & Compress(cmbPls)
   sBillParts(0, 5) = "" & Trim(cmbPls)
   If bSqlRows Then
      On Error Resume Next
      sBillParts(tNode.Index, 0) = ""
      sBillParts(tNode.Index, 1) = "" & Trim(cmbRev)
      sBillParts(tNode.Index, 2) = "" & Trim(lblDsc(0))
      sBillParts(tNode.Index, 3) = "" & str$(lblLevel)
      sBillParts(tNode.Index, 4) = "" & Compress(cmbPls)
      sBillParts(tNode.Index, 5) = "" & Trim(cmbPls)
      lstNodes.AddItem "" & Trim(cmbPls)
      With RdoBom
         Do Until .EOF
            iKey = iKey + 1
            bLevel = 0
            If !PALEVEL < 4 Then
               Set tNode = tvw1.Nodes.Add(, , "key:" & str$(iKey) & Trim(!BMASSYPART) & Trim(!PartRef), "" & Trim(!PartNum) & " Rev: " & Trim(!BMREV), 2, 3)
            Else
               Set tNode = tvw1.Nodes.Add(, , "key:" & str$(iKey) & Trim(!BMASSYPART) & Trim(!PartRef), "" & Trim(!PartNum) & "      " & Trim(!BMREV), 2, 3)
            End If
            sBillParts(tNode.Index, 0) = "" & Trim(!PartRef)
            sBillParts(tNode.Index, 1) = "" & Trim(!BMREV)
            sBillParts(tNode.Index, 2) = "" & Trim(!PADESC) _
                       & " Qty: " & Format$(!BMQTYREQD, ES_QuantityDataFormat) & " " & Trim(!BMUNITS)
            sBillParts(tNode.Index, 3) = "" & str$(!PALEVEL)
            sBillParts(tNode.Index, 4) = "" & Trim(!BMPARTNUM)
            sBillParts(tNode.Index, 5) = "" & Trim(!BMASSYPART)
            If !PALEVEL < 4 Then lstNodes.AddItem "" & Trim(!PartNum)
            iList = tNode.Index
            NextBillLevel iList, Trim(!PartRef), "" & Trim(!BMREV)
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
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoCmdObj = Nothing
   Set BompBMe01a = Nothing
   
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
   If iCurrIdx > 0 Then Node.Image = 4
   iCurrIdx = Node.Index
   tvw1.ToolTipText = ""
   
End Sub

Private Sub tvw1_NodeClick(ByVal Node As ComctlLib.Node)
   sCurrPart = Compress(Node.Text)
   iCurrIdx = Node.Index
   'If Len(sCurrPart) Then NextBillLevel Node.Index, sBillParts(Node.Index, 0), sBillParts(Node.Index, 1)
   
End Sub



Private Sub NextBillLevel(iNode As Integer, sPartNumber As String, sRev As String)
   Dim iList As Integer
   Dim RdoBm1 As ADODB.Recordset
   Static sAssy As String
   On Error GoTo DiaErr1
   AdoCmdObj.Parameters(0) = sPartNumber
   AdoCmdObj.Parameters(1) = sRev
   bSqlRows = clsADOCon.GetQuerySet(RdoBm1, AdoCmdObj, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBm1
         Do Until .EOF
            iKey = iKey + 1
            If !PALEVEL < 4 Then
               Set tNode = tvw1.Nodes.Add(iNode, tvwChild, "key:" & str$(iKey) & Trim(!BMASSYPART) & Trim(!PartRef), Trim(!PartNum) & " Rev: " & !BMPARTREV, 2, 3)
            Else
               Set tNode = tvw1.Nodes.Add(iNode, tvwChild, "key:" & str$(iKey) & Trim(!BMASSYPART) & Trim(!PartRef), Trim(!PartNum) & "      " & !BMPARTREV, 2, 3)
            End If
            If Err > 0 Then Exit Do
            If sAssy <> "" & Trim(!BMASSYPART) Then bLevel = bLevel + 1
            sAssy = "" & Trim(!BMASSYPART)
            iList = tNode.Index
            sBillParts(tNode.Index, 0) = "" & Trim(!PartRef)
            sBillParts(tNode.Index, 1) = "" & Trim(!BMPARTREV)
            sBillParts(tNode.Index, 2) = "" & Trim(!PADESC) _
                       & " Qty: " & Format$(!BMQTYREQD, ES_QuantityDataFormat) & " " & Trim(!BMUNITS)
            sBillParts(tNode.Index, 3) = "" & Trim(!PALEVEL)
            sBillParts(tNode.Index, 4) = "" & Trim(!BMPARTNUM)
            sBillParts(tNode.Index, 5) = "" & Trim(!BMASSYPART)
            NextBillLevel iList, Trim(!PartRef), "" & Trim(!BMREV)
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
         sSql = "SELECT * FROM BmplTable WHERE (BMASSYPART='" _
                & sOldUon & "' AND BMPARTREF='" & sOldPart _
                & "' AND BMREV='" & sOldRev & "')"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_FORWARD)
         If bSqlRows Then
            With RdoPls
               sSql = "INSERT INTO BmplTable (" _
                      & "BMASSYPART,BMPARTREF,BMPARTNUM," _
                      & "BMREV,BMQTYREQD,BMUNITS,BMCONVERSION," _
                      & "BMSEQUENCE,BMESTCOST,BMTYPE,BMADDER," _
                      & "BMPURCONV,BMREFERENCE,BMSETUP,BMPHANTOM," _
                      & "BMCOMT,BMESTLABOR,BMESTLABOROH,BMESTMATERIAL," _
                      & "BMESTMATERIALBRD) VALUES('" & sNewPart & "','" _
                      & sOldPart & "','" & sPartNum & "','" & sNewRev & "'," _
                      & !BMQTYREQD & ",'" & !BMUNITS & "'," & !BMCONVERSION & "," _
                      & !BMSEQUENCE & "," & !BMESTCOST & "," & !BMTYPE & "," _
                      & !BMADDER & "," & !BMPURCONV & ",'" & Trim(!BMREFERENCE) & "'," _
                      & !BMSETUP & "," & !BMPHANTOM & ",'" & Trim(!BMCOMT) & "'," _
                      & !BMESTLABOR & "," & !BMESTLABOROH & "," & !BMESTMATERIAL & "," _
                      & !BMESTMATERIALBRD & ")"
               clsADOCon.ExecuteSql sSql ' rdExecDirect
               
               sSql = "UPDATE BmhdTable SET BMHREVDATE='" _
                      & Format(ES_SYSDATE, "mm/dd/yy") & "' WHERE " _
                      & "BMHREF='" & Compress(BompBMe01a.cmbPls) & "' " _
                      & "AND BMHREV='" & Trim(BompBMe01a.cmbRev) & "'"
               clsADOCon.ExecuteSql sSql 'rdExecDirect
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
   sSql = "SELECT BMASSYPART,BMPARTREF From BmplTable " _
          & "WHERE (BMASSYPART='" & sNewPart & "' AND " _
          & "BMPARTREF='" & sOldPart & "' AND BMREV='" & sNewRev & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_FORWARD)
   If bSqlRows Then
      MsgBox "You May Not Paste At " & sNewPart & " It Will " & vbCrLf _
         & "Create A Duplicate Key.", vbInformation, Caption
   Else
      sMsg = "You May Create A Circular Bill By Cutting And Pasting" & vbCrLf _
             & "Here Are You Certain That You Want To Paste Here?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         'insert if
         On Error Resume Next
         sSql = "SELECT * FROM BmplTable WHERE (BMASSYPART='" _
                & sOldUon & "' AND BMPARTREF='" & sOldPart _
                & "' AND BMREV='" & sOldRev & "')"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_FORWARD)
         If bSqlRows Then
            With RdoPls
               sSql = "INSERT INTO BmplTable (" _
                      & "BMASSYPART,BMPARTREF,BMPARTNUM," _
                      & "BMREV,BMQTYREQD,BMUNITS,BMCONVERSION," _
                      & "BMSEQUENCE,BMESTCOST,BMTYPE,BMADDER," _
                      & "BMPURCONV,BMREFERENCE,BMSETUP,BMPHANTOM," _
                      & "BMCOMT,BMESTLABOR,BMESTLABOROH,BMESTMATERIAL," _
                      & "BMESTMATERIALBRD) VALUES('" & sNewPart & "','" _
                      & sOldPart & "','" & sPartNum & "','" & sNewRev & "'," _
                      & !BMQTYREQD & ",'" & !BMUNITS & "'," & !BMCONVERSION & "," _
                      & !BMSEQUENCE & "," & !BMESTCOST & "," & !BMTYPE & "," _
                      & !BMADDER & "," & !BMPURCONV & ",'" & Trim(!BMREFERENCE) & "'," _
                      & !BMSETUP & "," & !BMPHANTOM & ",'" & Trim(!BMCOMT) & "'," _
                      & !BMESTLABOR & "," & !BMESTLABOROH & "," & !BMESTMATERIAL & "," _
                      & !BMESTMATERIALBRD & ")"
               clsADOCon.ExecuteSql sSql ' rdExecDirect
               cmdPaste.Enabled = False
               ClearResultSet RdoPls
               If Err > 0 Then
                  MsgBox "Couldn't Make The Cut And Paste.", _
                     vbExclamation, Caption
               Else
                  sSql = "DELETE FROM BmplTable WHERE (BMASSYPART='" & sOldUon & "' " _
                         & "AND BMPARTREF='" & sOldPart & "' AND BMREV='" & sOldRev & "')"
                  clsADOCon.ExecuteSql sSql ' rdExecDirect
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
   cmdQuit.Enabled = False
   cmdAdd.Enabled = False
   cmdEdit.Enabled = False
   cmdCut.Enabled = False
   cmdCut.Enabled = False
   cmdCopy.Enabled = False
   cmdDelete.Enabled = False
   
End Sub

Private Function GetThisRevision() As Byte
   Dim RdoRev As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREF,BMHREV FROM BmhdTable WHERE " _
          & "(BMHREF='" & Compress(cmbPls) & "' AND BMHREV='" _
          & Compress(cmbRev) & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRev, ES_FORWARD)
   If bSqlRows Then
      GetThisRevision = 1
   Else
      GetThisRevision = 0
   End If
   
   If GetThisRevision = 0 Then
      sMsg = "That Revision Wasn't Found For The Current Part Number." & vbCrLf _
             & "Would You Like To Create It Now?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         
         sSql = "INSERT INTO BmhdTable (BMHREF,BMHPARTNO,BMHREV,BMHPART," _
                & "BMHOBSOLETE,BMHEFFECTIVE,BMHRELEASED) " _
                & "VALUES('" & Compress(cmbPls) & "','" & cmbPls & "','" & Compress(cmbRev) & "','" _
                & Compress(cmbPls) & "','" & Format(ES_SYSDATE, "mm/dd/yy") & "','" _
                & Format(ES_SYSDATE + 365, "mm/dd/yy") & "',0)"
         clsADOCon.ExecuteSql sSql ' rdExecDirect
         If clsADOCon.ADOErrNum = 0 Then
            cmbRev.AddItem cmbRev
            sMsg = "The Bill Of Material Revision Has Been Created." & vbCrLf _
                   & "No Items Have been Copied To The Parts List And " & vbCrLf _
                   & "It Has Not Been Released."
            MsgBox sMsg, vbInformation, Caption
            GetThisRevision = 1
         Else
            MsgBox "Could Not Create The Bill Of Material Revision.", _
               vbExclamation, Caption
            GetThisRevision = 0
         End If
         
      Else
         cmbRev = ""
         CancelTrans
      End If
   End If
   
   Set RdoRev = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthisrev"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
   
End Function

Private Sub ExpandTree()
   Dim iList As Integer
   For iList = 1 To tvw1.Nodes.Count - 1
      tvw1.Nodes(iList).Expanded = True
   Next
   iCurrIdx = 0
End Sub

Private Sub txtApp_LostFocus()
   
   Dim strApp As String
   strApp = txtApp.Text
   
   If (strApp = "") Then
      MsgBox "Please entry Approval Name.", vbInformation
   Else
      sSql = "UPDATE BmhdTable SET BMAPPBY='" _
             & strApp & "' WHERE " _
             & "BMHREF='" & Compress(BompBMe01a.cmbPls) & "' " _
             & "AND BMHREV='" & Trim(BompBMe01a.cmbRev) & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
End Sub

Private Sub txtAte_DropDown()
   ShowCalendar Me
End Sub


Private Sub txtAte_LostFocus()
   
   Dim strDate As String
   If Trim(txtAte) = "" Then txtAte = CheckDate(txtAte)
   
   strDate = ""
   If Len(txtAte) > 0 Then
      strDate = Format(txtAte, "mm/dd/yy")
   End If
   
   If (strDate = "") Then
      MsgBox "Please entry Approval date.", vbInformation
   Else
      sSql = "UPDATE BmhdTable SET BMAPPDATE ='" _
             & strDate & "' WHERE " _
             & "BMHREF='" & Compress(BompBMe01a.cmbPls) & "' " _
             & "AND BMHREV='" & Trim(BompBMe01a.cmbRev) & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
End Sub


Private Sub txtAssy_Click()
   optRefresh.Value = vbChecked
   iCurrIdx = 0
   Active.Picture = Chkyes.Picture
   cmdEdit.Enabled = False
   cmdCut.Enabled = False
   cmdCopy.Enabled = False
   cmdDelete.Enabled = False
   cmdAdd.Enabled = True
   lblDsc(1) = cmbPls & " Qty: 1 " & lblUnits
   lblLvl = lblTopLevel
   If Val(sBillParts(iCurrIdx, 3)) > 3 Then cmdAdd.Enabled = False _
          Else cmdAdd.Enabled = True
   
End Sub



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

Private Sub GetReleased()
   Dim RdoRel As ADODB.Recordset
   sSql = "SELECT BMHREF,BMHREV,BMHRELEASED FROM BmhdTable " _
          & "WHERE BMHREF='" & Compress(cmbPls) & "' AND " _
          & "BMHREV='" & Trim(cmbRev) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRel, ES_FORWARD)
   If bSqlRows Then _
      If RdoRel!BMHRELEASED = 1 Then lblRel = "Y" Else lblRel = "N"
   
   Set RdoRel = Nothing
   
End Sub
Private Function GetBOMSecurity()
   On Error GoTo DiaErr1
   Dim RdoBom As ADODB.Recordset
   Err = 0
   sSql = "SELECT ISNULL(COBOMSEC, 0) COBOMSEC FROM ComnTable WHERE COREF=1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      With RdoBom
         GetBOMSecurity = IIf((!COBOMSEC = 0), False, True)
         ClearResultSet RdoBom
      End With
   Else
      GetBOMSecurity = False
   End If
   Set RdoBom = Nothing
   Exit Function
   
   
DiaErr1:
   sProcName = "GetBOMSecurity"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function CheckForBomSec(strRev As String, bIncRev As Boolean)
   
   
   If (bBOMSec = True) Then
      
      Dim strNewRev As String
      strNewRev = strRev
      If (Trim(strNewRev) = "") Then strNewRev = "0"

      If (bIncRev = True) Then
         If (IsNumeric(strNewRev)) Then
            strNewRev = CStr(CDbl(strNewRev) + 1)
         Else
            strNewRev = Chr$(Asc(strNewRev) + 1)
         End If
      End If
      
      BompBMe01d.txtRev = strNewRev
      chkUpChild = 0
      BompBMe01d.Show vbModal
      
      
      If (chkUpChild = 0) Then
         cmbRev = strRev
         MsgBox "Please change the revision number to modify the Operations.", vbCritical
         CheckForBomSec = False
      Else
         ' Update the Comments
         Dim strCmt As String
         strCmt = txtCmt.Text
         UpdateComments strCmt
         ' The option is not enabled
         CheckForBomSec = True
      End If
   Else
      ' The option is not enabled
      CheckForBomSec = True
   End If
End Function

Private Function SetApproval()
   If bGoodRev Then
      On Error Resume Next
       
      Dim RdoRel As ADODB.Recordset
      sSql = "SELECT BMAPPDATE,BMAPPBY FROM BmhdTable WHERE " _
              & "BMHREF='" & Compress(BompBMe01a.cmbPls) & "' " _
              & "AND BMHREV='" & Trim(BompBMe01a.cmbRev) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRel, ES_FORWARD)
      If bSqlRows Then
         txtApp = RdoRel!BMAPPBY
         txtAte = RdoRel!BMAPPDATE
      End If
      Set RdoRel = Nothing
      
    End If
End Function

Private Function ResetApproval()
   If bGoodRev Then
      txtApp = ""
      txtAte = ""
      On Error Resume Next
       
       sSql = "UPDATE BmhdTable SET BMAPPDATE = NULL, BMAPPBY = NULL WHERE " _
              & "BMHREF='" & Compress(BompBMe01a.cmbPls) & "' " _
              & "AND BMHREV='" & Trim(BompBMe01a.cmbRev) & "'"
       clsADOCon.ExecuteSql sSql 'rdExecDirect
      
    End If
End Function

Private Function UpdateComments(strCmt As String)
   
   On Error GoTo DiaErr1
   
   sSql = "UPDATE BmhdTable SET BMREVNOTES ='" & strCmt & "' WHERE " _
          & "BMHREF='" & Compress(BompBMe01a.cmbPls) & "' " _
          & "AND BMHREV='" & Trim(BompBMe01a.cmbRev) & "'"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   Exit Function
DiaErr1:
   sProcName = "txtAte_LostFocus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


