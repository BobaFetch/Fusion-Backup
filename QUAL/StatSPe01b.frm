VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Team Members For A Key"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "StatSPe01b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<=>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3300
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Highlite Selection And Press To Move"
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   6120
      TabIndex        =   9
      ToolTipText     =   "Update Team Members And Apply Changes"
      Top             =   600
      Width           =   875
   End
   Begin VB.ListBox lstMem 
      Height          =   2790
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "Double Click Or Select And Press The Buddy To Add Or Remove"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ListBox lstTem 
      Height          =   2790
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Double Click Or Select And Press The Buddy To Add Or Remove"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtDmy 
      Height          =   285
      Left            =   6000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   4320
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4530
      FormDesignWidth =   7035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Team Members Selected For Key                    "
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
      Left            =   3840
      TabIndex        =   8
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblKey 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label z2 
      BackStyle       =   0  'Transparent
      Caption         =   "Team Members For:"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Team Members                                                "
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
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "StatSPe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte
Dim bNewData As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6301
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdLst_Click()
   Dim b As Byte
   Dim a As Integer
   Dim iList As Integer
   
   On Error Resume Next
   For iList = 0 To lstTem.ListCount - 1
      If lstTem.Selected(iList) Then
         For a = 0 To lstMem.ListCount - 1
            If lstTem.List(iList) = lstMem.List(a) Then b = 1
         Next
         If Not b Then
            bNewData = 1
            cmdUpd.Enabled = True
            lstMem.AddItem lstTem.List(iList)
         End If
         b = 0
      End If
   Next
   For iList = 0 To lstMem.ListCount - 1
      If lstMem.Selected(iList) Then
         cmdUpd.Enabled = True
         lstMem.RemoveItem (iList)
         bNewData = 1
      End If
   Next
   
End Sub

Private Sub cmdUpd_Click()
   UpdateTeam
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillLists
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   Move StatSPe01a.Left + 400, StatSPe01a.Top + 800
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   Dim sMsg As String
   
   If bNewData = 1 Then
      sMsg = "The Team Membership Has Changed And Not Updated." & vbCr _
             & "Are You Certain That You Want To Leave?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
   StatSPe01a.optTem.Value = vbUnchecked
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set StatSPe01b = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillLists()
   Dim RdoCmb As ADODB.Recordset
   ' On Error GoTo DiaErr1
   sSql = "SELECT TMMID,TMMLSTNAME,TMMMINIT,TMMFSTNAME " _
          & "FROM RjtmTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            If Not IsNull(!TMMID) Then
               lstTem.AddItem "" & !TMMID & Chr$(9) _
                  & Trim(!TMMLSTNAME) & ", " & Trim(!TMMFSTNAME) _
                  & " " & Trim(!TMMMINIT)
            End If
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   sSql = "SELECT MEMID,MEMNAME FROM RjmmTable " _
          & "WHERE (MEMREF='" & Compress(lblPrt) & "' " _
          & "AND MEMKEY='" & lblKey & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            lstMem.AddItem "" & !MEMID & Chr$(9) _
               & Trim(!MEMNAME)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub





Private Sub lstMem_Click()
   On Error Resume Next
   If lstMem.ListCount > 0 Then
      cmdLst.Enabled = True
   Else
      cmdLst.Enabled = False
   End If
   
End Sub

Private Sub lstMem_DblClick()
   On Error Resume Next
   If lstMem.ListIndex >= 0 Then
      lstMem.Selected(lstMem.ListIndex) = True
   End If
   If lstMem.Selected(lstMem.ListIndex) Then
      cmdLst.Enabled = True
      cmdLst_Click
   End If
   
End Sub


Private Sub lstMem_GotFocus()
   Dim iList As Integer
   For iList = 0 To lstTem.ListCount - 1
      lstTem.Selected(iList) = False
   Next
   
End Sub


Private Sub lstMem_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete Then cmdUpd_Click
   
End Sub


Private Sub lstMem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If lstMem.ListIndex >= 0 Then
      lstMem.Selected(lstMem.ListIndex) = True
   End If
   
End Sub

Private Sub lstTem_Click()
   On Error Resume Next
   If lstTem.ListCount > 0 Then
      cmdLst.Enabled = True
   Else
      cmdLst.Enabled = False
   End If
   
End Sub


Private Sub lstTem_DblClick()
   On Error Resume Next
   If lstTem.ListIndex >= 0 Then
      lstTem.Selected(lstTem.ListIndex) = True
   End If
   If lstTem.Selected(lstTem.ListIndex) Then
      cmdLst.Enabled = True
      cmdLst_Click
   End If
   
End Sub


Private Sub lstTem_GotFocus()
   Dim iList As Integer
   'cmdLst.Enabled = False
   'sOldPart = cmbPrt
   For iList = 0 To lstMem.ListCount - 1
      lstMem.Selected(iList) = False
   Next
   
End Sub


Private Sub lstTem_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then cmdLst_Click
   
End Sub


Private Sub lstTem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If lstTem.ListIndex >= 0 Then
      lstTem.Selected(lstTem.ListIndex) = True
   End If
   
End Sub


Private Sub txtDmy_Click()
   'dummy for errors
   
End Sub



Private Sub UpdateTeam()
   Dim bResponse As Byte
   Dim a As Integer
   Dim iList As Integer
   Dim sMemId As String
   Dim sMemName As String
   Dim sMsg As String
   
   If lstMem.ListCount = 0 Then
      sMsg = "There Are No Team Members Selected. " & vbCr _
             & "Set The Team Member List To Empty?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then Exit Sub
   End If
   sMsg = "Are You Ready To Update The Team Members?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      sSql = "DELETE FROM RjmmTable WHERE " _
             & "(MEMREF='" & Compress(lblPrt) & "' " _
             & "AND MEMKEY='" & lblKey & "')"
      clsADOCon.ExecuteSQL sSql
      
      For iList = 0 To lstMem.ListCount - 1
         a = Len(Trim(lstMem.List(iList)))
         If a > 0 Then
            sMemId = Trim(Left(lstMem.List(iList), 15))
            sMemName = Mid(lstMem.List(iList), 17, a)
            sSql = "INSERT INTO RjmmTable (MEMREF,MEMKEY," _
                   & "MEMID,MEMNAME) VALUES('" _
                   & Compress(lblPrt) & "','" _
                   & lblKey & "','" _
                   & sMemId & "','" _
                   & sMemName & "')"
            clsADOCon.ExecuteSQL sSql
         End If
      Next
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox "Team Member List Was Successfully Updated.", _
            vbInformation, Caption
         bNewData = 0
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         
         MsgBox "Could Not Successfully Update Team List.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub
