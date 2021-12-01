VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Routing Library Operation"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTf06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Delete This Routing Library Entry"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtCmt 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   2025
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "9"
      Text            =   "RoutRTf06a.frx":07AE
      Top             =   1560
      Width           =   4065
   End
   Begin VB.ComboBox cmbOpr 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Library Operation"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtDsc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "2"
      Top             =   1080
      Width           =   3075
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   3720
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4005
      FormDesignWidth =   6285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Name"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1185
   End
End
Attribute VB_Name = "RoutRTf06a"
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

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbOpr_Click()
   GetOperation
   
End Sub


Private Sub cmbOpr_LostFocus()
   cmbOpr = CheckLen(cmbOpr, 12)
   If Len(Trim(cmbOpr)) > 0 Then GetOperation
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   If txtDsc.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Selection.", _
         vbInformation, Caption
   Else
      DeleteOperation
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3155
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RoutRTf06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   cmbOpr.Clear
   sSql = "SELECT LIBREF,LIBNUM FROM RlbrTable "
   LoadComboBox cmbOpr
   If cmbOpr.ListCount > 0 Then cmbOpr = cmbOpr.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetOperation()
   Dim RdoOpr As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT LIBREF,LIBSHOP,LIBNUM,LIBDESC,LIBCOMT FROM " _
          & "RlbrTable WHERE LIBREF='" & Compress(cmbOpr) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOpr, ES_STATIC)
   If bSqlRows Then
      With RdoOpr
         cmbOpr = "" & Trim(!LIBNUM)
         txtDsc = "" & Trim(!LIBDESC)
         txtCmt = "" & Trim(!LIBCOMT)
         txtDsc.ForeColor = Es_TextForeColor
         ClearResultSet RdoOpr
      End With
   Else
      txtDsc = "*** Library Operation Wasn't Found ***"
      txtDsc.ForeColor = ES_RED
      txtCmt = ""
   End If
   Set RdoOpr = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getoperat"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub DeleteOperation()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "Are You Sure That You Wish To Permanently" & vbCrLf _
          & "Remove This Standard Library Operation?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      sSql = "DELETE FROM RlbrTable WHERE LIBREF='" _
             & Compress(cmbOpr) & "' "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      If clsADOCon.RowsAffected > 0 Then
         MsgBox "Operation Successfully Deleted." _
            & vbInformation, Caption
         FillCombo
      Else
         MsgBox "Could Not Successfully Delete The Operation." _
            & vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub
