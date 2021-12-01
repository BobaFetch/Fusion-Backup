VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form AdmnADe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standard Comments"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "AdmnADe03a.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClass 
      Cancel          =   -1  'True
      Caption         =   "&Add"
      Height          =   315
      Left            =   4680
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Add Some Class"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "40 Characters Max"
      Top             =   1320
      Width           =   3495
   End
   Begin VB.ComboBox cmbCid 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select ID Or Enter A New ID"
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox cmbCls 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Class From List"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtCmt 
      Height          =   2055
      Index           =   1
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "9"
      ToolTipText     =   "Standard Narrow Allow Up To 2048 Chars, Others As High As 5376"
      Top             =   1800
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox txtCmt 
      Height          =   2085
      Index           =   0
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Tag             =   "9"
      ToolTipText     =   "Standard Narrow Allow Up To 2048 Chars, Others As High As 5376"
      Top             =   1800
      Width           =   5295
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4200
      FormDesignWidth =   5730
   End
   Begin VB.Label lblListIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Note: The Width Represents The Approximate For Non True Type Fonts.  The Actual Will Be Different."
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment ID"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment Class"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "AdmnADe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 7/4/03
'11/11/05 Added select on Click
'4/3/06 Added BuildComments (new systems)
Option Explicit
Dim RdoStd As ADODB.Recordset


Dim bCancel As Byte
Dim bGoodComment As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbCid_Click()
   bGoodComment = GetComment()
   
End Sub

Private Sub cmbCid_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   If bCancel Then Exit Sub
   If Trim(cmbCid) = "" Then
      bGoodComment = 0
      Exit Sub
   Else
      bGoodComment = GetComment()
   End If
   If bGoodComment = 0 Then AddComment
   
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



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   bCancel = True
   
End Sub


Private Sub cmdClass_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Add A New User Comment Class?", ES_YESQUESTION, _
               Caption)
   If bResponse = vbYes Then
      AdmnADe03b.lblListIndex = cmbCls.ListCount
      AdmnADe03b.z1(3) = z1(3)
      AdmnADe03b.txtCls = AdmnADe03a.cmbCls
      AdmnADe03b.Show
   Else
      CancelTrans
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      BuildComments '4/3/06
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
   Set RdoStd = Nothing
   Set AdmnADe03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT COMMENT_CLASS FROM StchTable " _
          & "ORDER BY COMMENT_CLASS"
   LoadComboBox cmbCls, -1
   If cmbCls.ListCount > 0 Then cmbCls = cmbCls.List(Val(lblListIndex))
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
          & "FROM StchTable WHERE COMMENT_CLASS='" & Trim(cmbCls) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCls, ES_FORWARD)
   If bSqlRows Then
      With RdoCls
         z1(3) = "Width" & str$(!COMMENT_WIDTH)
         cmbCls.ToolTipText = "" & Trim(!COMMENT_TOOLTIP)
         If !COMMENT_WIDTH = 40 Then
            txtCmt(0).Visible = False
            txtCmt(1).Visible = True
         Else
            txtCmt(1).Visible = False
            txtCmt(0).Visible = True
         End If
         ClearResultSet RdoCls
      End With
   End If
   FillComments
   Exit Sub
   
DiaErr1:
   sProcName = "getclass"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillComments()
   cmbCid.Clear
   txtDsc = ""
   txtCmt(0) = ""
   txtCmt(1) = ""
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

Private Function GetComment() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM StcdTable WHERE COMMENT_REF='" _
          & Compress(cmbCid) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStd, ES_KEYSET)
   
   If bSqlRows Then
      With RdoStd
         GetComment = 1
         cmbCid = "" & Trim(!COMMENT_ID)
         txtDsc = "" & Trim(!COMMENT_DESC)
         txtCmt(0) = "" & Trim(!COMMENT_80)
         txtCmt(1) = "" & Trim(!COMMENT_40)
      End With
   Else
      GetComment = 0
   End If
   
   Exit Function
   
DiaErr1:
   GetComment = 0
   sProcName = "getcomment"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddComment()
'   Dim RdoNew As ADODB.Recordset
   Dim bResponse As Byte
   Dim b As Boolean
   Dim sMsg As String
   
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sMsg = "Add A New Comment " & cmbCid & " ?.."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      If cmbCid = "ALL" Then
         MsgBox "ALL Is An Illegal Name.", _
            vbInformation, Caption
         Exit Sub
      End If
      b = CheckValidColumn(cmbCid)
      If b Then
          sSql = "INSERT INTO StcdTable (COMMENT_REF, COMMENT_ID, COMMENT_CLASS) VALUES " & _
            "('" & Compress(cmbCid) & "','" & Trim(cmbCid) & "','" & Trim(cmbCls) & "')"
          clsADOCon.ExecuteSQL sSql
          If clsADOCon.ADOErrNum = 0 Then
            SysMsg "Comment Added.", True
            cmbCid.AddItem cmbCid
            bGoodComment = GetComment
            txtDsc.SetFocus
          Else
            MsgBox "Couldn't Add The Comment.", vbInformation, Caption
          End If
       End If
    End If
'      If b Then
'         sSql = "SELECT * FROM StcdTable WHERE COMMENT_ID=''"
'         RdoNew = clsADOCon.GetRecordSet(sSql, ES_DYNAMIC)
'            With RdoNew
'               .AddNew
'               !COMMENT_REF = Compress(cmbCid)
'               !COMMENT_ID = Trim(cmbCid)
'               !COMMENT_CLASS = Trim(cmbCls)
'               .Update
'               ClearResultSet RdoNew
'               If Err = 0 Then
'                  SysMsg "Comment Added.", True
'                  cmbCid.AddItem cmbCid
'                  bGoodComment = GetComment
'                  txtDsc.SetFocus
'               Else
'                  MsgBox "Couldn't Add The Comment.", _
'                     vbInformation, Caption
'               End If
'            End With
'      End If
'
'   Else
'      CancelTrans
'   End If
'   Set RdoNew = Nothing
   
End Sub

Private Sub lblListIndex_Change()
   cmbCls.Clear
   FillCombo
   
End Sub

Private Sub txtCmt_LostFocus(Index As Integer)
   If Trim(cmbCls) = "PO Remarks" Then
      txtCmt(0) = CheckLen(txtCmt(0), 5376)
   Else
      If Trim(cmbCls) = "PO Items" Or Trim(cmbCls) = "SO Items" Then
         txtCmt(Index) = CheckLen(txtCmt(Index), 2048)
      Else
         txtCmt(Index) = CheckLen(txtCmt(Index), 5376)
      End If
   End If
   
   txtCmt(Index) = StrCase(txtCmt(Index), ES_FIRSTWORD)
   On Error Resume Next
   With RdoStd
      If Index = 0 Then
         !COMMENT_80 = Trim(txtCmt(0))
      Else
         !COMMENT_40 = Trim(txtCmt(1))
      End If
      .Update
   End With
   
   
End Sub

Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   On Error Resume Next
   With RdoStd
      !COMMENT_DESC = Trim(txtDsc)
      .Update
   End With
   
End Sub



'Standard Comments  Moved 4/3/06

Public Sub BuildComments()
   Dim RdoTest As ADODB.Recordset
   On Error Resume Next
   MouseCursor 13
   sSql = "SELECT COMMENT_CLASS FROM dbo.StchTable WHERE COMMENT_CLASS='PO Remarks'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTest, ES_DYNAMIC)
   If Not bSqlRows Then
      'Insert Defaults
      sSql = "INSERT INTO StchTable (COMMENT_CLASS,COMMENT_LISTINDEX," _
             & "COMMENT_WIDTH) " _
             & "VALUES('PO Remarks',0,80)"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO StchTable (COMMENT_CLASS, COMMENT_LISTINDEX," _
             & "COMMENT_WIDTH) " _
             & "VALUES('PO Items',1,40)"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO StchTable (COMMENT_CLASS, COMMENT_LISTINDEX," _
             & "COMMENT_WIDTH) " _
             & "VALUES('SO Remarks',2,80)"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO StchTable (COMMENT_CLASS, COMMENT_LISTINDEX," _
             & "COMMENT_WIDTH) " _
             & "VALUES('SO Items',3,40)"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO StchTable (COMMENT_CLASS, COMMENT_LISTINDEX," _
             & "COMMENT_WIDTH) " _
             & "VALUES('Inv Remarks',4,80)"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO StchTable (COMMENT_CLASS, COMMENT_LISTINDEX," _
             & "COMMENT_WIDTH) " _
             & "VALUES('PS Remarks',5,80)"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO StchTable (COMMENT_CLASS, COMMENT_LISTINDEX," _
             & "COMMENT_WIDTH) " _
             & "VALUES('RTE Operations',6,40)"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO StchTable (COMMENT_CLASS, COMMENT_LISTINDEX," _
             & "COMMENT_WIDTH) " _
             & "VALUES('EST Comments',7,80)"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO StchTable (COMMENT_CLASS, COMMENT_LISTINDEX," _
             & "COMMENT_WIDTH) " _
             & "VALUES('MO Comments',8,80)"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO StchTable (COMMENT_CLASS, COMMENT_LISTINDEX," _
             & "COMMENT_WIDTH) " _
             & "VALUES('Bom Comments',9,40)"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE StchTable SET COMMENT_TOOLTIP='" _
             & "Standard PO Remarks' WHERE COMMENT_LISTINDEX=0"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE StchTable SET COMMENT_TOOLTIP='" _
             & "Standard PO Item Comments' WHERE COMMENT_LISTINDEX=1"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE StchTable SET COMMENT_TOOLTIP='" _
             & "Standard SO Remarks' WHERE COMMENT_LISTINDEX=2"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE StchTable SET COMMENT_TOOLTIP='" _
             & "Standard SO Item Comments' WHERE COMMENT_LISTINDEX=3"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE StchTable SET COMMENT_TOOLTIP='" _
             & "Standard Invoice Remarks' WHERE COMMENT_LISTINDEX=4"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE StchTable SET COMMENT_TOOLTIP='" _
             & "Standard PS Remarks' WHERE COMMENT_LISTINDEX=5"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE StchTable SET COMMENT_TOOLTIP='" _
             & "Standard Routing Operations' WHERE COMMENT_LISTINDEX=6"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE StchTable SET COMMENT_TOOLTIP='" _
             & "Standard Estimating Comments' WHERE COMMENT_LISTINDEX=7"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE StchTable SET COMMENT_TOOLTIP='" _
             & "Standard Manufacturing Order Comments' WHERE COMMENT_LISTINDEX=8"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE StchTable SET COMMENT_TOOLTIP='" _
             & "Standard Parts List Comments' WHERE COMMENT_LISTINDEX=9"
      clsADOCon.ExecuteSQL sSql
   End If
   Set RdoTest = Nothing
   
End Sub

