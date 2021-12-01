VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SysComments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standard Comments"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   HelpContextID   =   1104
   Icon            =   "SysComments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SysComments.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.CheckBox optAppend 
      Caption         =   "To Current Comments"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Check To Append Comments.  Uncheck To Replace Existing"
      Top             =   1520
      Width           =   2535
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "SysComments.frx":1078
      DownPicture     =   "SysComments.frx":19EA
      Height          =   350
      Left            =   4080
      Picture         =   "SysComments.frx":235C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Standard Comments"
      Top             =   4080
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "P&aste"
      Height          =   315
      Left            =   4800
      TabIndex        =   7
      ToolTipText     =   "Paste To The Current Form (No Need To Copy First)"
      Top             =   840
      Width           =   875
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy"
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      ToolTipText     =   "Copy To The ClipBoard And Paste Anywhere"
      Top             =   480
      Width           =   875
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "8"
      ToolTipText     =   "Description"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.ComboBox cmbCid 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Select ID From List"
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Class From List"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtCmt 
      Height          =   2055
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "8"
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
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "4"
      Top             =   1800
      Width           =   5295
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   8
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
      FormDesignHeight=   4320
      FormDesignWidth =   5730
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Append Selection"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Check To Append Comments.  Uncheck To Replace Existing"
      Top             =   1515
      Width           =   1575
   End
   Begin VB.Label lblControl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   2
      Left            =   5400
      Picture         =   "SysComments.frx":295E
      Top             =   4080
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   1
      Left            =   4920
      Picture         =   "SysComments.frx":32D0
      Top             =   4080
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   0
      Left            =   4440
      Picture         =   "SysComments.frx":3C42
      Top             =   4080
      Visible         =   0   'False
      Width           =   315
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
      TabIndex        =   13
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment ID"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment Class"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "SysComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'7/5/03 New
'1/8/05 Added PS Ship To, PO Ship To
'1/8/05 Changed Paste. See ListIndexesDocumentation
'5/22/06 Added bIndex for PROPLA (Estimating stored custom)
'        Changed the sort order and added Indexing
Option Explicit
Dim bCopied As Byte
Dim bGoodComment As Byte
Dim bIndex As Byte
Dim bOnLoad As Byte
Dim bWidth As Byte
Dim bDontUnload As Byte

Dim bCommentIdx(200) As Byte
Dim sCommentCls(200) As String
Dim sStartClass As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function AddComment() As Byte
    Dim sItemAdded As String
    
    On Error Resume Next
    clsADOCon.ADOErrNum = 0
    AddComment = 0
    sItemAdded = Trim(cmbCid)
    
    sSql = "INSERT INTO StCdTable (COMMENT_REF, COMMENT_ID, COMMENT_CLASS, COMMENT_DESC, COMMENT_80, COMMENT_40) " & _
           " VALUES ('" & Compress(cmbCid) & "', '" & sItemAdded & "','" & Trim(cmbCls) & "','" & Trim(txtDsc) & "','" & _
           txtCmt(0) & "','" & txtCmt(1) & "')"
           
    clsADOCon.ExecuteSQL sSql
    
    If clsADOCon.ADOErrNum = 0 Then
        SysMsg "Comment Added.", True
        cmbCid.AddItem sItemAdded
        FillComments (cmbCid.NewIndex)
        AddComment = 1
    Else
        AddComment = 0
    End If
    
End Function

Private Sub HandleAddButtons(bAdding As Boolean)
    cmdAdd.Top = 480
    cmdCancel.Top = 840
    cmdAdd.Left = 4800
    cmdCancel.Left = 4800
    cmdAdd.Width = 875
    cmdCancel.Width = 875
    cmdAdd.Height = 315
    cmdCancel.Height = 315
    txtCmt(0).TabStop = bAdding
    txtCmt(1).TabStop = bAdding
    If bAdding Then
        SysComments.Caption = "Add Comment"
    Else
        SysComments.Caption = "Standard Comments"
    End If
    cmbCid.Enabled = Not (bAdding)
    cmbCls.Enabled = Not (bAdding)
    optAppend.Visible = Not (bAdding)
    z1(4).Visible = Not (bAdding)
    cmdCopy.Visible = Not (bAdding)
    cmdCan.Visible = Not (bAdding)
    cmdAdd.Visible = bAdding
    cmdCancel.Visible = bAdding
    cmdPaste.Visible = Not (bAdding)
    txtDsc.Locked = Not (bAdding)
End Sub

Private Sub cmbCid_Click()
   GetComment
   
End Sub


Private Sub cmbCid_LostFocus()
   bGoodComment = GetComment
   If (bGoodComment = 0) Then
        If (cmbCls = sStartClass) Then
            bDontUnload = 1
            txtDsc.SetFocus
            If MsgBox("Do you wish to add this comment?", vbYesNo) = vbYes Then
                HandleAddButtons (True)
                txtDsc = ""
                txtCmt(0) = ""
                txtCmt(1) = ""
                txtDsc.SetFocus
            End If
        End If  '(cmbCls= sStartClass)
   End If       '(bGoodCommment=0)
  
End Sub


Private Sub cmbCls_Click()
   If cmbCls.ListIndex > -1 Then bIndex = bCommentIdx(cmbCls.ListIndex)
   GetClass
End Sub


Private Sub cmbCls_LostFocus()
   If cmbCls.ListIndex > -1 Then bIndex = bCommentIdx(cmbCls.ListIndex)
   If bDontUnload = 0 Then GetClass
   
End Sub


Private Sub cmdAdd_Click()
    AddComment
    HandleAddButtons (False)
    GetComment
    bDontUnload = 0
End Sub

Private Sub cmdCan_Click()
   Form_Deactivate
End Sub



Private Sub cmdCancel_Click()
    HandleAddButtons (False)
    GetComment
    bDontUnload = 0
End Sub

Private Sub cmdComments_Click()
   'Add one of these to the form and go
   If cmdComments Then
      'The Default is txtCmt and need not be included
      'Use Select Case cmdCopy to add your own
      SysComments.lblControl = "txtCmt"
      'See List For Index
      SysComments.lblListIndex = 3
      SysComments.Show
      cmdComments = False
   End If
   
End Sub

Private Sub cmdCopy_Click()
   If bWidth = 40 Then
      Clipboard.SetText txtCmt(1)
   Else
      Clipboard.SetText txtCmt(0)
   End If
   bCopied = 1
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1104
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

'Modified 1/8/05.  Added ActiveControl and blocked others

Private Sub cmdPaste_Click()
   'See Query_Unload too
   If Left(lblControl, 3) <> "lbl" Then
      If bWidth = 40 Then
         If optAppend.Value = vbUnchecked Then
            MdiSect.ActiveForm.ActiveControl = txtCmt(1)
         Else
            MdiSect.ActiveForm.ActiveControl = MdiSect.ActiveForm.ActiveControl & txtCmt(1)
         End If
      Else
         If optAppend.Value = vbUnchecked Then
            MdiSect.ActiveForm.ActiveControl = txtCmt(0)
         Else
            MdiSect.ActiveForm.ActiveControl = MdiSect.ActiveForm.ActiveControl & txtCmt(0)
         End If
      End If
   Else
      Select Case Trim(lblControl)
         Case "foobar"
            'Add a new one
            If bWidth = 40 Then
               If optAppend.Value = vbUnchecked Then
                  MdiSect.ActiveForm.foobar = txtCmt(1)
               Else
                  MdiSect.ActiveForm.foobar = MdiSect.ActiveForm.foobar & txtCmt(1)
               End If
            Else
               If optAppend.Value = vbUnchecked Then
                  MdiSect.ActiveForm.foobar = txtCmt(0)
               Else
                  MdiSect.ActiveForm.foobar = MdiSect.ActiveForm.foobar & txtCmt(0)
               End If
            End If
            '*** PS Items 10/28/03
         Case "lblCmt(1)"
            If bWidth = 40 Then
               If optAppend.Value = vbUnchecked Then
                  MdiSect.ActiveForm.lblCmt(1) = txtCmt(1)
               Else
                  MdiSect.ActiveForm.lblCmt(1) = MdiSect.ActiveForm.lblCmt(1) & txtCmt(1)
               End If
            Else
               If optAppend.Value = vbUnchecked Then
                  MdiSect.ActiveForm.lblCmt(1) = txtCmt(0)
               Else
                  MdiSect.ActiveForm.lblCmt(1) = MdiSect.ActiveForm.lblCmt(1) & txtCmt(0)
               End If
            End If
         Case "lblCmt(2)"
            If bWidth = 40 Then
               If optAppend.Value = vbUnchecked Then
                  MdiSect.ActiveForm.lblCmt(2) = txtCmt(1)
               Else
                  MdiSect.ActiveForm.lblCmt(2) = MdiSect.ActiveForm.lblCmt(2) & txtCmt(1)
               End If
            Else
               If optAppend.Value = vbUnchecked Then
                  MdiSect.ActiveForm.lblCmt(2) = txtCmt(0)
               Else
                  MdiSect.ActiveForm.lblCmt(2) = MdiSect.ActiveForm.lblCmt(2) & txtCmt(0)
               End If
            End If
         Case "lblCmt(3)"
            If bWidth = 40 Then
               If optAppend.Value = vbUnchecked Then
                  MdiSect.ActiveForm.lblCmt(3) = txtCmt(1)
               Else
                  MdiSect.ActiveForm.lblCmt(3) = MdiSect.ActiveForm.lblCmt(3) & txtCmt(1)
               End If
            Else
               If optAppend.Value = vbUnchecked Then
                  MdiSect.ActiveForm.lblCmt(3) = txtCmt(0)
               Else
                  MdiSect.ActiveForm.lblCmt(3) = MdiSect.ActiveForm.lblCmt(3) & txtCmt(0)
               End If
            End If
            '            Case "txtSta" '1/8/05 PS Ship To
            '                If bWidth = 40 Then
            '                    If optAppend.Value = vbUnchecked Then
            '                        MdiSect.ActiveForm.txtSta = txtCmt(1)
            '                    Else
            '                        MdiSect.ActiveForm.txtSta = MdiSect.ActiveForm.txtSta & txtCmt(1)
            '                    End If
            '                Else
            '                    If optAppend.Value = vbUnchecked Then
            '                        MdiSect.ActiveForm.txtSta = txtCmt(0)
            '                    Else
            '                        MdiSect.ActiveForm.txtSta = MdiSect.ActiveForm.txtSta & txtCmt(0)
            '                    End If
            '                End If
            '            '*** End PS Items
            '            Case "txtShp" '1/8/05 PS Ship To
            '                If bWidth = 40 Then
            '                    If optAppend.Value = vbUnchecked Then
            '                        MdiSect.ActiveForm.txtShp = txtCmt(1)
            '                    Else
            '                        MdiSect.ActiveForm.txtShp = MdiSect.ActiveForm.txtShp & txtCmt(1)
            '                    End If
            '                Else
            '                    If optAppend.Value = vbUnchecked Then
            '                        MdiSect.ActiveForm.txtShp = txtCmt(0)
            '                    Else
            '                        MdiSect.ActiveForm.txtShp = MdiSect.ActiveForm.txtShp & txtCmt(0)
            '                    End If
            '                End If
            '            Case Else
            '                If bWidth = 40 Then
            '                    If optAppend.Value = vbUnchecked Then
            '                        MdiSect.ActiveForm.txtCmt = txtCmt(1)
            '                    Else
            '                        MdiSect.ActiveForm.txtCmt = MdiSect.ActiveForm.txtCmt & txtCmt(1)
            '                    End If
            '                Else
            '                    If optAppend.Value = vbUnchecked Then
            '                        MdiSect.ActiveForm.txtCmt = txtCmt(0)
            '                    Else
            '                        MdiSect.ActiveForm.txtCmt = MdiSect.ActiveForm.txtCmt & txtCmt(0)
            '                    End If
            '                End If
      End Select
   End If
   Form_Deactivate
   
End Sub



Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bDontUnload = 0
      If Trim(lblControl) = "" Then lblControl = "txtCmt"
      FillCombo
      bOnLoad = 0
      HandleAddButtons (False)
      sStartClass = cmbCls
   End If
   MouseCursor 0
End Sub

Private Sub Form_Deactivate()
   If bDontUnload <> 1 Then Unload Me
   
End Sub


Private Sub Form_Load()
   If iBarOnTop Then
      Move MdiSect.Left + 400, MdiSect.Top + 2600
   Else
      Move MdiSect.Left + 2400, MdiSect.Top + 2000
   End If
   optAppend.Value = GetSetting("Esi2000", "System", "AppendComments", Trim(str(optAppend.Value)))
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Add any that you need here and paste
   '1/8/05 Removed. See Documentation
   '    Select Case Trim(lblControl)
   '        Case "foobar"
   '            MdiSect.ActiveForm.foobar.SetFocus
   '        Case Else
   '            MdiSect.ActiveForm.txtCmt.SetFocus
   '    End Select
   On Error Resume Next
   SaveSetting "Esi2000", "System", "AppendComments", Trim(str(optAppend.Value))
   If sCurrForm = "Full Estimate" And RunningBeta Then
      If cmbCls.ListIndex > -1 Then
         bIndex = bCommentIdx(cmbCls.ListIndex)
         SaveSetting "Esi2000", "ProPla", "EstimatingIndex", Trim(bIndex)
      End If
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set SysComments = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim RdoCls As ADODB.Recordset
   Dim bListIndex As Byte
   'On Error GoTo DiaErr1
   cmbCls.Clear
   Erase bCommentIdx
   Erase sCommentCls
   bListIndex = 0
   sSql = "SELECT COMMENT_LISTINDEX,COMMENT_CLASS FROM StchTable " _
          & "ORDER BY COMMENT_CLASS"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCls, ES_FORWARD)
   If bSqlRows Then
      With RdoCls
         Do Until .EOF
            bCommentIdx(bListIndex) = !COMMENT_LISTINDEX
            sCommentCls(bListIndex) = "" & Trim(!COMMENT_CLASS)
            cmbCls.AddItem "" & Trim(!COMMENT_CLASS)
            bListIndex = bListIndex + 1
            .MoveNext
         Loop
      End With
      ClearResultSet RdoCls
   End If
   If cmbCls.ListCount > 0 Then
      For bListIndex = 0 To cmbCls.ListCount - 1
         If bCommentIdx(bListIndex) = Val(lblListIndex) Then
            cmbCls = sCommentCls(bListIndex)
            bIndex = bCommentIdx(bListIndex)
         End If
      Next
      GetClass
   End If
   Set RdoCls = Nothing
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
         bWidth = !COMMENT_WIDTH
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
   FillComments (0)
   Set RdoCls = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getclass"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillComments(intStartNo As Integer)
'   Dim RdoCmt As ADODB.Recordset
   Dim intIndex As Integer
   
   cmbCid.Clear
   txtDsc = ""
   txtCmt(0) = ""
   txtCmt(1) = ""
   On Error GoTo DiaErr1
   sSql = "Qry_GetCommentClass '" & Trim(cmbCls) & "'"
   LoadComboBox cmbCid
   If cmbCid.ListCount > 0 Then
        If intStartNo > cmbCid.ListCount Then intIndex = 0 Else intIndex = intStartNo
      cmbCid = cmbCid.List(intIndex)
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
   Dim RdoStd As ADODB.Recordset
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
   Set RdoStd = Nothing
   Exit Function
   
DiaErr1:
   GetComment = 0
   sProcName = "getcomment"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   
End Sub



Private Sub ListIndexesDocumentation()
   'Sample Code. Note the SetFocus call to set target
   '    If cmdComments Then
   '        'See List For Index
   '        txtCmt.SetFocus
   '        syssysComments.lblListIndex = 0
   '        syssysComments.show
   '        cmdComments(0) = False
   '    End If
   
   
   
   'For Comments - Aligns with the ComboBox
   'There is no provision for User Classes
   '0 PO Remarks
   '1 PO Item Comments
   '2 SO Remarks
   '3 SO Item Comments
   '4 IN Remarks
   '5 PS Remarks
   '6 RTE Operations
   '7 EST Comments
   '8 MO  Comments
   '9 MO  Comments
   'The Index increases as the User Adds Comments
   
End Sub
