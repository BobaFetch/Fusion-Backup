VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ToolTLf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Tool List"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdDel 
      Cancel          =   -1  'True
      Caption         =   "&Delete"
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Delete This Tool List"
      Top             =   720
      Width           =   875
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ComboBox cmbTol 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   " Select From List"
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5640
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2190
      FormDesignWidth =   6585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Number"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1155
   End
End
Attribute VB_Name = "ToolTLf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'9/1/04 New
Option Explicit
Dim bCancel As Byte
Dim bGoodTool As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetThisToolList() As Byte
   Dim RdoTool As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetToolList '" & Compress(cmbTol) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTool, ES_FORWARD)
   If bSqlRows Then
      With RdoTool
         cmbTol = "" & Trim(!TOOLLIST_NUM)
         txtDesc = "" & Trim(!TOOLLIST_DESC)
      End With
      GetThisToolList = 1
   Else
      txtDesc = "*** Tool List Wasn't Found ***"
      GetThisToolList = 0
   End If
   Set RdoTool = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthisto"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbTol_Click()
   bGoodTool = GetThisToolList()
   
End Sub


Private Sub cmbTol_LostFocus()
   cmbTol = CheckLen(cmbTol, 30)
   If bCancel = 0 Then bGoodTool = GetThisToolList()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdDel_Click()
   If bGoodTool = 0 Then
      MsgBox "Requires A Valid Tool List.", _
         vbInformation, Caption
   Else
      DeleteThisToolList
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3451
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ToolTLf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDesc.BackColor = BackColor
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbTol.Clear
   sSql = "Qry_FillToolListCombo"
   LoadComboBox cmbTol
   If cmbTol.ListCount > 0 Then cmbTol = cmbTol.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtDesc_Change()
   If Left(txtDesc, 8) = "*** Tool" Then txtDesc.ForeColor = _
           ES_RED Else txtDesc.ForeColor = vbBlack
   
End Sub



Private Sub DeleteThisToolList()
   Dim RdoDel As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sToolRef As String
   
   sToolRef = Compress(cmbTol)
   On Error GoTo DiaErr1
   'Test Routings
   sSql = "SELECT OPTOOLLIST FROM RtopTable WHERE " _
          & "OPTOOLLIST='" & sToolRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDel, ES_FORWARD)
   If bSqlRows Then
      MsgBox "This Tool List Is Used On At Least One Routing" & vbCrLf _
         & "And Cannot Be Deleted Until Removed.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   'Test MO Ops
   sSql = "SELECT OPTOOLLIST FROM RnopTable WHERE " _
          & "OPTOOLLIST='" & sToolRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDel, ES_FORWARD)
   If bSqlRows Then
      MsgBox "This Tool List Is Used On At Least One MO Op" & vbCrLf _
         & "And Cannot Be Deleted Until Removed.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   'It's clear, ask it they want to
   sMsg = "This Function Removes All Records Of " & Trim(cmbTol) & vbCrLf _
          & "And Cannot Be Reversed.  Continue Deleting?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      'Tool Lists
      sSql = "DELETE FROM TlhdTable WHERE TOOLLIST_REF='" & sToolRef & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "DELETE FROM TlitTable WHERE TOOLLISTIT_REF='" & sToolRef & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "The Tool List Has Been Deleted.", True
         FillCombo
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "The Tool Could Not Be Deleted.", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   Set RdoDel = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "deletethistool"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
