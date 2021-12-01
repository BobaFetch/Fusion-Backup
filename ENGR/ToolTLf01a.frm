VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ToolTLf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy A Tool List"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtNewDesc 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Tag             =   "2"
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton cmdCpy 
      Cancel          =   -1  'True
      Caption         =   "C&opy"
      Height          =   315
      Left            =   5640
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Copy This Tool List"
      Top             =   1560
      Width           =   875
   End
   Begin VB.TextBox lblLst 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ComboBox cmbTol 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
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
      TabIndex        =   2
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
      FormDesignHeight=   2535
      FormDesignWidth =   6585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Tool List"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Tool List"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   1635
   End
End
Attribute VB_Name = "ToolTLf01a"
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
'10/3/06 Revamped the entire CopyThisToolList procedure (SQL Commands)
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
         lblLst = "" & Trim(!TOOLLIST_DESC)
         txtNewDesc = "" & Trim(!TOOLLIST_DESC)
      End With
      GetThisToolList = 1
   Else
      lblLst = "*** Tool List Wasn't Found ***"
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
   cmbTol = FindToolList(cmbTol, lblLst)
   
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



Private Sub cmdCpy_Click()
   If bGoodTool = 0 Then
      MsgBox "Requires A Valid Existing Tool List.", _
         vbInformation, Caption
   Else
      If Len(Trim(txtNew)) < 4 Then bGoodTool = 0
      If bGoodTool = 0 Then
         MsgBox "Requires A Valid New Tool List.", _
            vbInformation, Caption
      Else
         CopyThisToolList
      End If
   End If
   
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3450
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
   Set ToolTLf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblLst.BackColor = BackColor
   
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




Private Sub CopyThisToolList()
   Dim RdoCpy As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sListRef As String
   Dim sNewRef As String
   
   sListRef = Compress(cmbTol)
   sNewRef = Compress(txtNew)
   On Error GoTo DiaErr1
   sMsg = "This Function Copies All Records Of " & Trim(cmbTol) & vbCrLf _
          & "To " & Trim(txtNew) & ". Continue To Copy To The New List?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      'In case the temp tables remain
      sSql = "DROP TABLE #Tlhd"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "DROP TABLE #Tlit"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      Err.Clear
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      'Header
      sSql = "SELECT * INTO #Tlhd from TlhdTable where TOOLLIST_REF='" & sListRef & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "UPDATE #Tlhd SET TOOLLIST_REF='" & sNewRef & "',TOOLLIST_NUM='" _
             & txtNew & "',TOOLLIST_DESC='" & txtNewDesc & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "INSERT INTO TlhdTable SELECT * FROM #Tlhd"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      'Tool List
      sSql = "SELECT * INTO #Tlit from TlitTable where TOOLLISTIT_REF='" & sListRef & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "UPDATE #Tlit SET TOOLLISTIT_REF='" & sNewRef & "',TOOLLISTIT_NUM='" _
             & txtNew & "' "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "INSERT INTO TlitTable SELECT * FROM #Tlit"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      MsgBox Err.Description
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "The Tool List Has Been Copied.", True
         txtNew = ""
         FillCombo
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         
         MsgBox "The Tool Could Not Be Copied.", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   Set RdoCpy = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "copythistool"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblLst_Change()
   If Left(lblLst, 8) = "*** Tool" Then lblLst.ForeColor = _
           ES_RED Else lblLst.ForeColor = vbBlack
   
End Sub

Private Sub txtNew_Change()
   txtNew = CheckLen(txtNew, 30)
   
End Sub


Private Sub txtNewDesc_LostFocus()
   txtNewDesc = CheckLen(txtNewDesc, 30)
   txtNewDesc = StrCase(txtNewDesc)
   
End Sub
