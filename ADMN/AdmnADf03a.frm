VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form AdmnADf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Standard Comment"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnADf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Class From List"
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox cmbCid 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select ID Or Enter A New ID"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "40 Characters Max"
      Top             =   1440
      Width           =   3495
   End
   Begin VB.CommandButton cmdClass 
      Cancel          =   -1  'True
      Caption         =   "D&elete"
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Delete This Standard Comment"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5160
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   720
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2415
      FormDesignWidth =   6120
   End
   Begin VB.Label lblListIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment Class"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment ID"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "AdmnADf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte
Dim bGoodComment As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetComment() As Byte
   Dim RdoStd As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT COMMENT_REF,COMMENT_ID,COMMENT_DESC " _
          & "FROM StcdTable WHERE COMMENT_REF='" & Compress(cmbCid) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStd, ES_FORWARD)
   If bSqlRows Then
      With RdoStd
         GetComment = 1
         cmbCid = "" & Trim(!COMMENT_ID)
         txtDsc = "" & Trim(!COMMENT_DESC)
      End With
   Else
      GetComment = 0
   End If
   Set RdoStd = Nothing
   Exit Function
   
DiaErr1:
   GetComment = 0
   sProcName = "getcomment"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillComments()
   cmbCid.Clear
   txtDsc = ""
   On Error GoTo DiaErr1
   sSql = "Qry_GetCommentClass '" & Trim(cmbCls) & "'"
   LoadComboBox cmbCid
   If cmbCid.ListCount > 0 Then
      cmbCid = cmbCid.List(0)
      bGoodComment = GetComment
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetClass()
   Dim RdoCls As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT COMMENT_CLASS,COMMENT_WIDTH,COMMENT_TOOLTIP " _
          & "FROM StchTable " _
          & "WHERE COMMENT_CLASS='" & Trim(cmbCls) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCls, ES_FORWARD)
   If bSqlRows Then
      With RdoCls
         cmbCls.ToolTipText = "" & Trim(!COMMENT_TOOLTIP)
         ClearResultSet RdoCls
      End With
   End If
   FillComments
   Set RdoCls = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getclass"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbCid_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   If Trim(cmbCid) = "" Then
      bGoodComment = 0
      Exit Sub
   Else
      bGoodComment = GetComment()
   End If
   
End Sub


Private Sub cmbCls_Click()
   GetClass
   
End Sub


Private Sub cmbCls_LostFocus()
   GetClass
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdClass_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Trim(cmbCid) <> "" Then
      sMsg = "This Action Does Not Affect Existing Comments But" & vbCr _
             & "Cannot Be Reversed. Do You Wish To Continue?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         sSql = "DELETE FROM StcdTable WHERE COMMENT_CLASS='" _
                & cmbCls & "' AND COMMENT_REF='" _
                & Compress(cmbCid) & "'"
         clsADOCon.ExecuteSQL sSql
         If clsADOCon.ADOErrNum = 0 Then
            GetClass
            SysMsg "Comment Deleted.", True
         Else
            MsgBox "Could Not Delete The Comment ID.", _
               vbInformation, Caption
         End If
      Else
         CancelTrans
      End If
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1152
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
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdmnADf03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillCommentsByIndex"
   LoadComboBox cmbCls, -1
   If cmbCls.ListCount > 0 Then cmbCls = cmbCls.List(Val(lblListIndex))
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
