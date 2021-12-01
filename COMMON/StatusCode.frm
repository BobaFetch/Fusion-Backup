VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form StatusCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internal Status Code"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "StatusCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbStatID 
      Height          =   315
      Left            =   5880
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkAct 
      Caption         =   "Check1"
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "StatusCode.frx":030A
      DownPicture     =   "StatusCode.frx":0C7C
      Height          =   350
      Left            =   4560
      Picture         =   "StatusCode.frx":15EE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Standard Comments"
      Top             =   3480
      Width           =   350
   End
   Begin VB.TextBox txtStatCmt 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Tag             =   "9"
      ToolTipText     =   "Comments (3072 Chars Max)"
      Top             =   3480
      Width           =   4335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Assign"
      Height          =   315
      Left            =   6720
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Assign New Status Code"
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&UnAssign"
      Height          =   315
      Left            =   6720
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "UnAssign "
      Top             =   960
      Width           =   915
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   120
      Picture         =   "StatusCode.frx":1BF0
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   4560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6720
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5640
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4800
      FormDesignWidth =   7785
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
   End
   Begin VB.Label LableRef1 
      Caption         =   "Item"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   120
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Active"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   23
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblSCTRef2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   555
   End
   Begin VB.Label txtSCTRef 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   22
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label txtLModDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3240
      TabIndex        =   21
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label txtLModUser 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label txtStatCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1515
   End
   Begin VB.Label txtCurUser 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5160
      TabIndex        =   20
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblSCTRef1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3600
      TabIndex        =   19
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "Status Code"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSysCommIndex 
      Caption         =   "2"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblStatType 
      Caption         =   "StatType"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Status Code Comments:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Last Modified Date"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   2880
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Last Modified User"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   2520
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblSCTypeRef 
      Caption         =   "SO Number"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7080
      Picture         =   "StatusCode.frx":239E
      Top             =   1800
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7080
      Picture         =   "StatusCode.frx":2728
      Top             =   1560
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "StatusCode"
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
Dim bChanged As Byte
Dim bLoading As Byte

Dim sCustomers(500, 2) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
'Dim RdoQry As rdoQuery
'Dim RdoItm As ADODB.Recordset
Dim bNew  As Boolean


Private Sub chkAct_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
       ActiveStatusCode
    End If
End Sub

Private Sub chkAct_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveStatusCode
End Sub

Private Sub ActiveStatusCode()
    Dim lrows, i, iActStat As Integer
    Dim bResponse As Byte
    Dim strStatID As String
    Dim strMsg As String
    Dim iRow As String
    
    Grd.row = Grd.RowSel
    Grd.Col = 1
    strStatID = Grd.Text

    Grd.Col = 0
    If (Grd.Text = "InActive") Then
        strMsg = "Do you want to Activate Internal Status Code : " & strStatID & " ?"
        iActStat = 1 ' activate the status code
    Else
        iActStat = 0 ' activate the status code
        strMsg = "Do you want to In-Activate Internal Status Code : " & strStatID & " ?"
    End If

    bResponse = MsgBox(strMsg, ES_YESQUESTION, Caption)
    If bResponse = vbYes Then
        AddStatusCode strStatID, iActStat
        iRow = FillGrid(Grd.row, strStatID)
        Grd.row = iRow
        Grd.SetFocus
        Grd.Col = 1
    End If

End Sub


'
Private Sub cmdAdd_Click()
    Dim strStatID As String
    Dim iActStat As Integer
    Dim iRow As Integer
        
    txtStatCmt.Text = ""
   
    AddNewStatCode.txtSCTRef = txtSCTRef
    AddNewStatCode.lblSCTRef1 = IIf(lblSCTRef1 = "", 0, lblSCTRef1)
    AddNewStatCode.lblSCTRef2 = lblSCTRef2
    AddNewStatCode.lblStatType = lblStatType
    AddNewStatCode.Show vbModal

    'AddNewStatCode
    ' Add new status
    iActStat = 1 ' activate the status code
    strStatID = cmbStatID
    If (strStatID <> "") Then
        Grd.Rows = Grd.Rows + 1
        Grd.row = Grd.Rows - 1
        Grd.Col = 0
        AddStatusCode strStatID, iActStat
        iRow = FillGrid(Grd.row, strStatID)
        Grd.row = iRow
        Grd.Col = 1
        txtStatCmt.Text = ""

    End If
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCancel_Click()
    Dim lrows, i, iActStat, iRowSel As Integer
    Dim bResponse As Byte
    Dim strStatID As String
    Dim strMsg As String
    Dim iRow As String

    Grd.Col = 1
    iRowSel = Grd.row
    strStatID = Grd.Text
    strMsg = "Do you want to unAssign Internal Status Code : " & strStatID & " ?"
    bResponse = MsgBox(strMsg, ES_YESQUESTION, Caption)
    If bResponse = vbYes Then
        RemoveStatusCode strStatID
        If (iRowSel = (Grd.Rows - 1)) Then
            iRowSel = iRowSel - 1
        End If
        iRow = FillGrid(iRowSel, "")
        Grd.row = iRow
        Grd.SetFocus
        Grd.Col = 1
    End If
    
End Sub

Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtStatCmt.SetFocus
      SysComments.lblListIndex = lblSysCommIndex
      SysComments.Show
      cmdComments = False
   End If

End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2104
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub Form_Activate()
    Dim iRow As String
    bLoading = 0
    
    If bOnLoad = 1 Then
        iRow = FillGrid(1, "")
        Grd.row = iRow
        Grd.Col = 1
        
        'FillStatusCode cmbStatID, Me
    End If
    ' Loaded completed
    bLoading = 1
    MouseCursor 0
   
End Sub

Private Sub Form_Load()
   
    FormatControls
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "Assigned"
      .Col = 1
      .Text = "StatusID"
      .Col = 2
      .Text = "Status Code"
      
      .ColWidth(0) = 750
      .ColWidth(1) = 1000
      .ColWidth(2) = 3900
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set StatusCode = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub




Function FillGrid(iRowSel As Integer, strStatID As String) As Integer
    Dim RdoGrd As ADODB.Recordset
    Dim sBookParts As String
   
    Dim strStatCmdRef As String
    Dim strStatCmdRef1 As String
    Dim strStatCmdRef2 As String
    Dim strTransType As String
    Dim sSql1 As String
    
   On Error Resume Next
   Grd.Rows = 1
   On Error GoTo DiaErr1
    
    strStatCmdRef = txtSCTRef
    strStatCmdRef1 = IIf(lblSCTRef1 = "", 0, lblSCTRef1)
    strStatCmdRef2 = lblSCTRef2
    strTransType = lblStatType
   
    
    sSql = "SELECT STATUS_ACT_STATE, StcodeTable.STATUS_REF, ISNULL(COMMENT, '') COMMENT, " & _
                " StcodeTable.STATUS_CODE, STATUS_CUR_USER, STATUS_CUR_DATE" & _
            " FROM StCmtTable, StcodeTable " & _
            " WHERE StCmtTable.STATUS_REF = StcodeTable.STATUS_REF"

   sSql1 = " AND STATCODE_TYPE_REF = '" & strTransType & "' AND " & _
                " STATUS_CMT_REF = '" & strStatCmdRef & "' " & _
                " AND STATUS_CMT_REF1 = '" & strStatCmdRef1 & "' " & _
                " AND STATUS_CMT_REF2 = '" & strStatCmdRef2 & "' " & _
                " Order By STATUS_ACT_STATE DESC"

   sSql = sSql & sSql1
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
      With RdoGrd
         Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            Grd.row = Grd.Rows - 1
            Grd.Col = 0
            If (!STATUS_ACT_STATE = 0) Then
                Grd.Text = "InActive"
            Else
                Grd.Text = "Active"
            End If
            
            Grd.Col = 1
            Grd.Text = "" & Trim(!STATUS_REF)
            Grd.Col = 2
            Grd.Text = "" & Trim(!STATUS_CODE)
            
            'Grd.Row = iRowSel Then
            If strStatID = Trim(!STATUS_REF) Then
                txtLModUser = "" & Trim(!STATUS_CUR_USER)
                txtLModDate = Format(!STATUS_CUR_DATE, "mm/dd/yy")
                txtStatCode = "" & Trim(!STATUS_CODE)
                txtStatCmt = "" & Trim(!Comment)
                If (!STATUS_ACT_STATE = 0) Then
                    chkAct = vbUnchecked
                Else
                    chkAct = vbChecked
                End If
                FillGrid = Grd.row
            ElseIf (strStatID = "" And Grd.row = iRowSel) Then
                txtLModUser = "" & Trim(!STATUS_CUR_USER)
                txtLModDate = Format(!STATUS_CUR_DATE, "mm/dd/yy")
                txtStatCode = "" & Trim(!STATUS_CODE)
                txtStatCmt = "" & Trim(!Comment)
                If (!STATUS_ACT_STATE = 0) Then
                    chkAct = vbUnchecked
                Else
                    chkAct = vbChecked
                End If
                FillGrid = Grd.row
            End If
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
   End If
   Set RdoGrd = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub grd_KeyPress(KeyAscii As Integer)
    Dim strStatID As String
    Dim iActStat As Integer
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        cmbStatID = ""
        'lblStatCd = ""
        Grd.Col = 1
        strStatID = Grd.Text
        
        GetStatusCodeInfo strStatID
    End If
   

End Sub


Private Sub grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim strStatID As String
    cmbStatID = ""
    'lblStatCd = ""
    
    Grd.Col = 1
    strStatID = Grd.Text
    
    GetStatusCodeInfo strStatID
End Sub


Private Sub GetStatusCodeInfo(strStatID As String)
     Dim RdoStat As ADODB.Recordset
    
     Dim strStatCmdRef As String
     Dim strStatCmdRef1 As String
     Dim strStatCmdRef2 As String
     Dim strTransType As String
     Dim sSql1 As String
     
    On Error GoTo DiaErr1
    
    strStatCmdRef = txtSCTRef
    strStatCmdRef1 = IIf(lblSCTRef1 = "", 0, lblSCTRef1)
    strStatCmdRef2 = lblSCTRef2
    strTransType = lblStatType
   
    
    sSql = "SELECT STATUS_ACT_STATE, STATUS_CUR_USER, StCmtTable.STATUS_REF, STATUS_CUR_DATE, " & _
            " StcodeTable.STATUS_CODE, ISNULL(COMMENT, '') COMMENT " & _
        " FROM StCmtTable, StcodeTable " & _
            " WHERE StCmtTable.STATUS_REF = StcodeTable.STATUS_REF"

    sSql1 = " AND STATCODE_TYPE_REF = '" & strTransType & "' AND " & _
                 " STATUS_CMT_REF = '" & strStatCmdRef & "' " & _
                 " AND STATUS_CMT_REF1 = '" & strStatCmdRef1 & "' " & _
                 " AND STATUS_CMT_REF2 = '" & strStatCmdRef2 & "' " & _
                 " AND StCmtTable.STATUS_REF = '" & strStatID & "'"
                 '" AND STATUS_ACT_STATE = 1 " & _

    sSql = sSql & sSql1
   
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoStat, ES_FORWARD)
    If bSqlRows Then
        With RdoStat
            txtLModUser = "" & Trim(!STATUS_CUR_USER)
            txtLModDate = Format(!STATUS_CUR_DATE, "mm/dd/yy")
            txtStatCode = "" & Trim(!STATUS_CODE)
            txtStatCmt = "" & Trim(!Comment)

            If (!STATUS_ACT_STATE = 0) Then
                chkAct.Value = vbUnchecked
            Else
                chkAct.Value = vbChecked
            End If

            ClearResultSet RdoStat
        End With
    End If
    Set RdoStat = Nothing
    Exit Sub
   
DiaErr1:
   sProcName = "GetStatusCodeInfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub AddStatusCode(strStatID As String, iActStat As Integer)
    
    Dim RdoStc As ADODB.Recordset
    On Error GoTo DiaErr1
    
    Dim strStatCmdRef As String
    Dim strStatCmdRef1 As String
    Dim strStatCmdRef2 As String
    Dim strTransType As String
    Dim strRes As String
    Dim strCurUser As String
    Dim strComments As String
     
    On Error GoTo DiaErr1
    
    strStatCmdRef = txtSCTRef
    strStatCmdRef1 = IIf(lblSCTRef1 = "", 0, lblSCTRef1)
    strStatCmdRef2 = lblSCTRef2
    strTransType = lblStatType
    strCurUser = txtCurUser
    strComments = txtStatCmt
    
    If strStatCmdRef = "" Then
      MsgBox "Please Select a Sales Order.", _
         vbInformation, Caption
        Exit Sub
    End If
    
    sSql = "Qry_AddInternStatCode '" & Trim(strStatCmdRef) & "','" & strStatCmdRef1 & _
                        "','" & strStatCmdRef2 & "','" & strStatID & _
                        "','" & strTransType & "','" & strCurUser & _
                        "','" & strComments & "'," & iActStat

    bSqlRows = clsADOCon.GetDataSet(sSql, RdoStc, ES_FORWARD)
    If bSqlRows Then
       With RdoStc
          strRes = "" & Trim(.Fields(0))
          ClearResultSet RdoStc
       End With
    End If

    Set RdoStc = Nothing
    Exit Sub
DiaErr1:
   sProcName = "AddStatusCode"
   CurrError.Number = Err.Number
    CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub RemoveStatusCode(strStatID As String)

     Dim strStatCmdRef As String
     Dim strStatCmdRef1 As String
     Dim strStatCmdRef2 As String
     Dim strTransType As String
     Dim sSql1 As String
     
    On Error GoTo DiaErr1
    
    strStatCmdRef = txtSCTRef
    strStatCmdRef1 = IIf(lblSCTRef1 = "", 0, lblSCTRef1)
    strStatCmdRef2 = lblSCTRef2
    strTransType = lblStatType
    
    sSql = "DELETE FROM StCmtTable WHERE STATCODE_TYPE_REF = '" & strTransType & "' " & _
                 " AND STATUS_CMT_REF = '" & strStatCmdRef & "' " & _
                 " AND STATUS_CMT_REF1 = '" & strStatCmdRef1 & "' " & _
                 " AND STATUS_CMT_REF2 = '" & strStatCmdRef2 & "' " & _
                 " AND StCmtTable.STATUS_REF = '" & strStatID & "'"

   
    clsADOCon.ExecuteSQL sSql
    Exit Sub
   
DiaErr1:
   sProcName = "RemoveStatusCode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtStatCmt_LostFocus()
    Dim iActStat As Integer
    Dim strStatID As String

    Grd.Col = 1
    strStatID = Grd.Text

    iActStat = 1 ' activate the status code
    AddStatusCode strStatID, iActStat

End Sub

