VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form AdmnUuser2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Manager"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkHideModule 
      Caption         =   "Hide modules for users with no permissions "
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   240
      Width           =   3555
   End
   Begin VB.Frame fraUser 
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   6615
      Begin VB.CommandButton cmdPer 
         Caption         =   "&Dt Col"
         Enabled         =   0   'False
         Height          =   375
         Index           =   8
         Left            =   5800
         TabIndex        =   29
         ToolTipText     =   "Data Collection"
         Top             =   1560
         Width           =   660
      End
      Begin VB.TextBox txtNme 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Tag             =   "2"
         Top             =   180
         Width           =   3855
      End
      Begin VB.TextBox txtNik 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Tag             =   "2"
         Top             =   540
         Width           =   1935
      End
      Begin VB.TextBox txtInt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         TabIndex        =   20
         Tag             =   "3"
         Top             =   540
         Width           =   975
      End
      Begin VB.ComboBox cmbGrp 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1440
         TabIndex        =   19
         Tag             =   "8"
         ToolTipText     =   "Select User Class From List"
         Top             =   900
         Width           =   1935
      End
      Begin VB.CheckBox optAct 
         Alignment       =   1  'Right Justify
         Caption         =   "Active User?"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3600
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdPer 
         Caption         =   "&Admn"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   17
         ToolTipText     =   "Administration"
         Top             =   1560
         Width           =   660
      End
      Begin VB.CommandButton cmdPer 
         Caption         =   "&Sales"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   780
         TabIndex        =   16
         ToolTipText     =   "Customer Order Processing"
         Top             =   1560
         Width           =   660
      End
      Begin VB.CommandButton cmdPer 
         Caption         =   "&Engr"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   1500
         TabIndex        =   15
         ToolTipText     =   "Engineering"
         Top             =   1560
         Width           =   660
      End
      Begin VB.CommandButton cmdPer 
         Caption         =   "&Prod"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   2220
         TabIndex        =   14
         ToolTipText     =   "Production Control"
         Top             =   1560
         Width           =   660
      End
      Begin VB.CommandButton cmdPer 
         Caption         =   "&Invc"
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   3660
         TabIndex        =   13
         ToolTipText     =   "Inventory Control"
         Top             =   1560
         Width           =   660
      End
      Begin VB.CommandButton cmdPer 
         Caption         =   "&Fina"
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   4380
         TabIndex        =   12
         ToolTipText     =   "Financial Administration"
         Top             =   1560
         Width           =   660
      End
      Begin VB.CommandButton cmdPer 
         Caption         =   "&Qual"
         Enabled         =   0   'False
         Height          =   375
         Index           =   6
         Left            =   5100
         TabIndex        =   11
         ToolTipText     =   "Quality Assurance"
         Top             =   1560
         Width           =   660
      End
      Begin VB.CommandButton cmdPer 
         Caption         =   "&Time"
         Enabled         =   0   'False
         Height          =   375
         Index           =   7
         Left            =   2940
         TabIndex        =   10
         ToolTipText     =   "Time Management"
         Top             =   1560
         Width           =   660
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nickname"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Initials"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   25
         Top             =   540
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Class"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Section Permissions:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1260
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkShowInactive 
      Caption         =   "Show inactive users"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   2475
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnUuser2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optChg 
      Caption         =   "From Change"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox optFrm 
      Caption         =   "From New"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "C&hange"
      Enabled         =   0   'False
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      ToolTipText     =   "Change The Current User Password"
      Top             =   1320
      Width           =   840
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      ToolTipText     =   "Add A New User"
      Top             =   960
      Width           =   840
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double Click To Select A User (Or Select And Press Enter)"
      Top             =   960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   20
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorSel    =   -2147483636
      AllowBigSelection=   0   'False
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   840
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   2100
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5550
      FormDesignWidth =   6870
   End
   Begin VB.Label lblUsers 
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "AdmnUuser2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'9/9/05 Corrected opening and closing files
Dim bOnLoad As Byte
Dim iOldrec As Integer

'grid columns
Private Const COL_Name = 0
Private Const COL_Group = 1
Private Const COL_Active = 2
Private Const COL_RecNo = 3
Private Const COL_Count = 4


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub chkHideModule_Click()
    'Secure.UserZHideModule= chkHideModule.Value
    ' Set the Hide module button flag
    'Put #iFreeDbf, iCurrentRec, Secure
    Dim iHideFlg As Integer
    
    If (chkHideModule.Value) Then
        iHideFlg = 1
    Else
        iHideFlg = 0
    End If
    ' Set the hide module flag
    SetHideModule (iHideFlg)
    
End Sub

Private Sub chkShowInactive_Click()
   FormatGrid
End Sub

Private Sub cmbGrp_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   
   On Error Resume Next
   cmbGrp = Trim(cmbGrp)
   For iList = 0 To cmbGrp.ListCount - 1
      If cmbGrp = cmbGrp.List(iList) Then b = 1
   Next
   If b = 0 Then
      'Beep
      cmbGrp = cmbGrp.List(0)
   End If
   If cmbGrp = cmbGrp.List(0) Then
      SecPw.UserAdmn = 1
      If iCurrentRec >= FIRSTUSERRECORDNO Then Secure.UserLevel = 10
   Else
      SecPw.UserAdmn = 0
      If iCurrentRec >= FIRSTUSERRECORDNO Then Secure.UserLevel = 20
   End If
   Put #iFreeIdx, iCurrentRec, SecPw
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdChg_Click()
   optChg.Value = vbChecked
   AdmnUsrpw.Show
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      OpenHelpContext 30
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub cmdNew_Click()
   AdmnUnewu.optFrm.Value = vbChecked
   AdmnUnewu.Show
   'Formatgrid
End Sub

Private Sub cmdPer_Click(Index As Integer)
   MouseCursor 13
   Select Case Index
      Case 0
         AdmnUperm1.Show
      Case 1
         AdmnUperm2.Show
      Case 2
         AdmnUperm3.Show
      Case 3
         AdmnUperm4.Show
      Case 4
         AdmnUperm5.Show
      Case 5
         AdmnUperm6.Show
      Case 6
         AdmnUperm7.Show
      Case 7
         AdmnUperm8.Show 'Time Management
      Case 8
         AdmnUperm9.Show 'Data collection
   End Select
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      iOldrec = iUserIdx
      FormatGrid
      ' Set the Hide module checkbox
      chkHideModule.Value = GetHideModule
      bOnLoad = 0
      
   End If
   If optFrm.Value = vbChecked Then
      optFrm.Value = vbUnchecked
      FormatGrid
   End If
   If optChg.Value = vbChecked Then
      optChg.Value = vbUnchecked
      Unload AdmnUsrpw
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   OpenDbfFiles
   cmbGrp.AddItem "Administrators"
   cmbGrp.AddItem "Users"
   cmbGrp = cmbGrp.List(0)
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   iUserIdx = iOldrec
   On Error Resume Next
   Get #iFreeIdx, iUserIdx, SecPw
   Get #iFreeDbf, iUserIdx, Secure
   FormUnload
   Set AdmnUuser2 = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Public Sub FormatGrid()
   Dim row As Integer
   Dim iList As Integer
   
   Grid1.Clear
   Grid1.Cols = COL_Count
   Grid1.ColWidth(COL_Name) = 1900
   Grid1.ColWidth(COL_Group) = 1550
   Grid1.ColWidth(COL_Active) = 900
   Grid1.ColWidth(COL_RecNo) = 0
   Grid1.row = 0
   Grid1.TextMatrix(0, COL_Name) = "User Id"
   Grid1.TextMatrix(0, COL_Group) = "Group"
   Grid1.TextMatrix(0, COL_Active) = "Active"
   Grid1.TextMatrix(0, COL_RecNo) = "Row"
   

   Grid1.Rows = 1
   For iList = FIRSTUSERRECORDNO To LOF(iFreeDbf) \ Len(Secure)
      Get #iFreeIdx, iList, SecPw
      Get #iFreeDbf, iList, Secure
      If Me.chkShowInactive = 1 Or Secure.UserActive = 1 Then
         row = row + 1
         If row >= Grid1.Rows Then Grid1.Rows = Grid1.Rows + 1
         Grid1.TextMatrix(row, COL_Name) = SecPw.UserLcName
         Grid1.TextMatrix(row, COL_Group) = IIf(SecPw.UserAdmn = 1, "Administrator", "User")
         Grid1.TextMatrix(row, COL_Active) = IIf(Secure.UserActive = 1, "Active", "Inactive")
         Grid1.TextMatrix(row, COL_RecNo) = iList
      End If
   Next

   'lblUsers = row
   If Grid1.Rows > 1 Then
      Grid1.Sort = 1
      Grid1.Col = 0
      Grid1.row = 1
      Grid1_DblClick
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "formatgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Grid1_Click()
   GetSelectedUser
End Sub

Private Sub Grid1_DblClick()
'   If Grid1.row = 0 Then
'      If Grid1.Rows > 1 Then
'         Grid1.row = 1
'      Else
'         Exit Sub
'      End If
'   End If
'
'   iCurrentRec = Grid1.TextMatrix(Grid1.row, COL_RecNo)
'   Get #iFreeIdx, iCurrentRec, SecPw
'   Get #iFreeDbf, iCurrentRec, Secure
'   txtNme = Trim(Secure.UserName)
'   txtNik = Trim(Secure.UserNickName)
'   txtInt = Trim(Secure.UserInitials)
'   If SecPw.UserAdmn Then
'      cmbGrp = cmbGrp.List(0)
'   Else
'      cmbGrp = cmbGrp.List(1)
'   End If
'   optAct.Value = Secure.UserActive
'   ResetControls True
   
   GetSelectedUser
   
End Sub


Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      
'      If Grid1.row = 0 Then
'         If Grid1.Rows > 1 Then
'            Grid1.row = 1
'         Else
'            Exit Sub
'         End If
'      End If
'
'      iCurrentRec = Grid1.TextMatrix(Grid1.row, COL_RecNo)
'      Get #iFreeIdx, iCurrentRec, SecPw
'      Get #iFreeDbf, iCurrentRec, Secure
'      txtNme = Trim(Secure.UserName)
'      txtNik = Trim(Secure.UserNickName)
'      txtInt = Trim(Secure.UserInitials)
'      If SecPw.UserAdmn Then
'         cmbGrp = cmbGrp.List(0)
'      Else
'         cmbGrp = cmbGrp.List(1)
'      End If
'      optAct.Value = Secure.UserActive
'      ResetControls True

      GetSelectedUser
   End If
   
End Sub

Private Sub optAct_Click()
   Secure.UserActive = optAct.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optAct_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtInt_LostFocus()
   txtInt = CheckLen(txtInt, 3)
   Secure.UserInitials = Trim(txtInt)
   If Secure.UserInitials = "" Then
      MsgBox "Requires An Entry.", vbInformation, Caption
   Else
      Put #iFreeDbf, iCurrentRec, Secure
   End If
   
End Sub


Private Sub txtNik_LostFocus()
   txtNik = CheckLen(txtNik, 20)
   Secure.UserNickName = Trim(txtNik)
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub


Private Sub txtNme_LostFocus()
   txtNme = CheckLen(txtNme, 40)
   txtNme = StrCase(txtNme)
   Secure.UserName = Trim(txtNme)
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub



Private Sub ResetControls(bOpen As Boolean)
   Dim iList As Integer
   txtNme.Enabled = bOpen
   txtNik.Enabled = bOpen
   txtInt.Enabled = bOpen
   cmbGrp.Enabled = bOpen
   optAct.Enabled = bOpen
   cmdChg.Enabled = bOpen
'   If bOpen Then
'      optAct.Caption = "____"
'   Else
'      optAct.Caption = ""
'   End If
   For iList = 0 To 8
      cmdPer(iList).Enabled = bOpen
   Next
   
End Sub

Private Sub OpenDbfFiles()
   On Error Resume Next
   Close iFreeIdx
   Close iFreeDbf
   iFreeIdx = FreeFile
   Open sFilePath & "rstval.eid" For Random Shared As iFreeIdx Len = Len(SecPw)
   iFreeDbf = FreeFile
   Open sFilePath & "rstval.edd" For Random Shared As iFreeDbf Len = Len(Secure)
   
End Sub

Private Sub GetSelectedUser()
   If Grid1.row = 0 Then
      If Grid1.Rows > 1 Then
         Grid1.row = 1
      Else
         Exit Sub
      End If
   End If
   
   iCurrentRec = Grid1.TextMatrix(Grid1.row, COL_RecNo)
   Get #iFreeIdx, iCurrentRec, SecPw
   Get #iFreeDbf, iCurrentRec, Secure
   txtNme = Trim(Secure.UserName)
   txtNik = Trim(Secure.UserNickName)
   txtInt = Trim(Secure.UserInitials)
   If SecPw.UserAdmn Then
      cmbGrp = cmbGrp.List(0)
   Else
      cmbGrp = cmbGrp.List(1)
   End If
   optAct.Value = Secure.UserActive
   
   ' Hide module button options
   'chkHideModule.Value = Secure.UserZHideModule
   
   ResetControls True
   
   'don't allow editing of SysMgr to insure there is always an active user
   If txtNik = "SysMgr" Then
      Me.fraUser.Enabled = False
   Else
      Me.fraUser.Enabled = True
   End If
End Sub
