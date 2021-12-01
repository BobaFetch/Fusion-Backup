VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ToolTLf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Tool"
   ClientHeight    =   2535
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
   ScaleHeight     =   2535
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
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
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Delete This Tool"
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
   Begin VB.CheckBox optExp 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   1440
      Width           =   715
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
      FormDesignHeight=   2535
      FormDesignWidth =   6585
   End
   Begin VB.Label lblCls 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Number"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Expendable"
      Height          =   255
      Index           =   16
      Left            =   3600
      TabIndex        =   4
      Top             =   1440
      Width           =   1275
   End
End
Attribute VB_Name = "ToolTLf03a"
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
'4/26/06 Corrected GetThisTool query
Option Explicit
Dim bCancel As Byte
Dim bGoodTool As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetThisTool() As Byte
   Dim RdoTool As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT TOOL_PARTREF,TOOL_NUM,TOOL_DESC,TOOL_CLASS,TOOL_EXPENDABLE " _
          & "FROM TohdTable WHERE TOOL_PARTREF='" & Compress(cmbTol) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTool, ES_FORWARD)
   If bSqlRows Then
      With RdoTool
         cmbTol = "" & Trim(!TOOL_NUM)
         txtDesc = "" & Trim(!TOOL_DESC)
         lblCls = "" & Trim(!TOOL_CLASS)
         optExp.Value = !TOOL_EXPENDABLE
      End With
      GetThisTool = 1
   Else
      txtDesc = "*** Tool Wasn't Found ***"
      GetThisTool = 0
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
   bGoodTool = GetThisTool()
   
End Sub


Private Sub cmbTol_LostFocus()
   cmbTol = CheckLen(cmbTol, 30)
   If bCancel = 0 Then bGoodTool = GetThisTool()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdDel_Click()
   If bGoodTool = 0 Then
      MsgBox "Requires A Valid Tool.", _
         vbInformation, Caption
   Else
      DeleteThisTool
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3452
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
   Set ToolTLf03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDesc.BackColor = BackColor
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbTol.Clear
   sSql = "Qry_FillToolCombo"
   LoadComboBox cmbTol, -1
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



Private Sub DeleteThisTool()
   Dim RdoDel As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sToolRef As String
   
   sToolRef = Compress(cmbTol)
   On Error GoTo DiaErr1
   'Test Tool Lists
   sSql = "Qry_GetToolList '" & sToolRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDel, ES_FORWARD)
   If bSqlRows Then
      MsgBox "This Tool Is Used On At Least One Tool List" & vbCrLf _
         & "And Cannot Be Deleted Until Removed.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   'Test PO's
   sSql = "SELECT DISTINCT PIPART FROM PoitTable WHERE " _
          & "PIPART='" & sToolRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDel, ES_FORWARD)
   If bSqlRows Then
      MsgBox "This Tool Is Used For At Least One PO Item" & vbCrLf _
         & "And Cannot Be Deleted Until Removed.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   'Test Run's
   sSql = "SELECT DISTINCT RUNREF FROM RunsTable WHERE " _
          & "RUNREF='" & sToolRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDel, ES_FORWARD)
   If bSqlRows Then
      MsgBox "This Tool Is Used On At Least One MO Run" & vbCrLf _
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
      
      'Inventory
      sSql = "DELETE FROM InvaTable WHERE INPART='" & sToolRef & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      'BOM Header
      sSql = "DELETE FROM BmhdTable WHERE BMHREF='" & sToolRef & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      'Parts
      sSql = "DELETE FROM PartTable WHERE PARTREF='" & sToolRef & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      'Tools
      sSql = "DELETE FROM TohdTable WHERE TOOL_PARTREF='" & sToolRef & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "The Tool Has Been Deleted.", True
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
