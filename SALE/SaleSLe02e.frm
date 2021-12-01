VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form SaleSLe02e 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internal Status Code"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "SaleSLe02e.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAcivate 
      Caption         =   "Activate"
      Height          =   315
      Left            =   6480
      TabIndex        =   28
      Top             =   960
      Width           =   915
   End
   Begin VB.TextBox txtStatCode 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   27
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "SaleSLe02e.frx":030A
      DownPicture     =   "SaleSLe02e.frx":0C7C
      Height          =   350
      Left            =   4560
      Picture         =   "SaleSLe02e.frx":15EE
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Standard Comments"
      Top             =   3240
      Width           =   350
   End
   Begin VB.TextBox txtStatCmt 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Tag             =   "9"
      ToolTipText     =   "Comments (3072 Chars Max)"
      Top             =   3240
      Width           =   4335
   End
   Begin VB.TextBox txtLModDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtLModUser 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   13
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtCurUser 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtSCTRef 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&New"
      Height          =   315
      Left            =   6480
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Add This Sales Order Item"
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   6480
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Cancel The Current PO Item"
      Top             =   1320
      Width           =   915
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   7200
      Picture         =   "SaleSLe02e.frx":1BF0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   2880
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7080
      Top             =   4920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6180
      FormDesignWidth =   7605
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   5655
      Begin VB.CommandButton CmdAssign 
         Caption         =   "&Assign"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4560
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Add This Sales Order Item"
         Top             =   240
         Width           =   915
      End
      Begin VB.ComboBox cmbStatID 
         Height          =   315
         Left            =   1440
         TabIndex        =   22
         Tag             =   "1"
         ToolTipText     =   "Select or Enter Sales Order Number (List Contains Last 3 Years Up To 500 Enties)"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblStatCd 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   1440
         TabIndex        =   21
         ToolTipText     =   "Status Code"
         Top             =   600
         Width           =   3015
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "New Status Code"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Status Code"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSysCommIndex 
      Caption         =   "2"
      Height          =   255
      Left            =   6720
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblStatType 
      Caption         =   "StatType"
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Comments:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Last Modified Date"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Last Modified User"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Current User"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblSCTRef2 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3660
      TabIndex        =   8
      ToolTipText     =   "Item Revision"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblSCTRef1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3285
      TabIndex        =   7
      ToolTipText     =   "Our Item Number"
      Top             =   120
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblSCTypeRef 
      Caption         =   "SO Number"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   6840
      Picture         =   "SaleSLe02e.frx":239E
      Top             =   2280
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   6720
      Picture         =   "SaleSLe02e.frx":2728
      Top             =   1800
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "SaleSLe02e"
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

Dim sCustomers(500, 2) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Dim RdoQry As rdoQuery
Dim RdoItm As rdoResultset
Dim bNew  As Boolean


Private Sub cmbStatID_Click()
    Dim strStatID As String
    Dim RdoStatCode As rdoResultset

    strStatID = cmbStatID
    sSql = "SELECT STATUS_CODE  FROM STCODETABLE " & _
                "WHERE STATUS_REF = '" & strStatID & "'"
    
    bSqlRows = GetDataSet(RdoStatCode, ES_FORWARD)
    If bSqlRows Then
       With RdoStatCode
          lblStatCd = "" & Trim(!STATUS_CODE)
          ClearResultSet RdoStatCode
       End With
    End If

End Sub

Private Sub cmdAcivate_Click()
    Dim lrows, i, iActStat As Integer
    Dim bResponse As Byte
    Dim strStatID As String
    Dim strMsg As String
    
    Grd.Col = 1
    strStatID = Grd.Text
    
    Grd.Col = 0
    If (Grd.CellPicture = Chkno.Picture) Then
        strMsg = "Do you want to Activate Internal Status Code : " & strStatID & " ?"
        iActStat = 1 ' activate the status code
    Else
        iActStat = 0 ' activate the status code
        strMsg = "Do you want to in Activate Internal Status Code : " & strStatID & " ?"
    End If

    bResponse = MsgBox(strMsg, ES_YESQUESTION, Caption)
    If bResponse = vbYes Then
        AddStatusCode strStatID, iActStat
        FillGrid
    End If
    
End Sub

Private Sub cmdAdd_Click()
    Grd.Rows = Grd.Rows + 1
    Grd.Row = Grd.Rows - 1
    Grd.Col = 0
    Set Grd.CellPicture = Chkyes.Picture
    CmdAssign.Enabled = True
    bNew = True
    txtStatCmt.Text = ""
End Sub

Private Sub CmdAssign_Click()
    
    Dim strStatID As String
    Dim iActStat As Integer
    
    strStatID = cmbStatID
    If strStatID = "" Then
      MsgBox "Please Select a Status Code.", _
         vbInformation, Caption
        Exit Sub
    End If
    ' Add new status
    iActStat = 1 ' activate the status code
    AddStatusCode strStatID, iActStat
    
    FillGrid
    CmdAssign.Enabled = False
    bNew = False
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCancel_Click()
    Dim lrows, i, iActStat As Integer
    Dim bResponse As Byte
    Dim strStatID As String
    Dim strMsg As String

    If bNew = True Then
        lrows = Grd.Rows
        For i = 0 To Grd.Rows - 1
            Grd.Row = i
            Grd.Col = 1
            If Grd.Text = "" Then
               Grd.RemoveItem (i)
            End If
        Next
        bNew = False
    Else
    
        Grd.Col = 1
        strStatID = Grd.Text
        strMsg = "Do you want to unAssign Internal Status Code : " & strStatID & " ?"
        bResponse = MsgBox(strMsg, ES_YESQUESTION, Caption)
        If bResponse = vbYes Then
            RemoveStatusCode strStatID
            FillGrid
        End If
        
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
   If bOnLoad = 1 Then
      FillGrid
      FillStatusCode cmbStatID, Me
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   
    FormatControls
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .Rows = 1
      .Row = 0
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
   CmdAssign.Enabled = False
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set SaleSLe02e = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub




Private Sub FillGrid()
    Dim RdoGrd As rdoResultset
    Dim sBookParts As String
   
    Dim strStatModRef As String
    Dim strStatModRef1 As String
    Dim strStatModRef2 As String
    Dim strTransType As String
    Dim sSql1 As String
    
   On Error Resume Next
   Grd.Rows = 1
   On Error GoTo DiaErr1
    
    strStatModRef = txtSCTRef
    strStatModRef1 = IIf(lblSCTRef1 = "", 0, lblSCTRef1)
    strStatModRef2 = lblSCTRef2
    strTransType = lblStatType
   
    
    sSql = "SELECT STATUS_ACT_STATE, StcodeTable.STATUS_REF, ISNULL(COMMENT, '') COMMENT, " & _
                " StcodeTable.STATUS_CODE, STATUS_CUR_USER, STATUS_CUR_DATE" & _
            " FROM StCmtTable, StcodeTable " & _
            " WHERE StCmtTable.STATUS_REF = StcodeTable.STATUS_REF"

   sSql1 = " AND MODULE_REF = '" & strTransType & "' AND " & _
                " STATUS_MODULE_REF = '" & strStatModRef & "' " & _
                " AND STATUS_MODULE_REF1 = '" & strStatModRef1 & "' " & _
                " AND STATUS_MODULE_REF2 = '" & strStatModRef2 & "' " & _
                " Order By STATUS_ACT_STATE DESC"

   sSql = sSql & sSql1
   
   bSqlRows = GetDataSet(RdoGrd, ES_FORWARD)
   If bSqlRows Then
      With RdoGrd
         Do Until .EOF
            Grd.Rows = Grd.Rows + 1
            Grd.Row = Grd.Rows - 1
            Grd.Col = 0
            If (!STATUS_ACT_STATE = 0) Then
                Set Grd.CellPicture = Chkno.Picture
            Else
                Set Grd.CellPicture = Chkyes.Picture
            End If
            
            Grd.Col = 1
            Grd.Text = "" & Trim(!STATUS_REF)
            Grd.Col = 2
            Grd.Text = "" & Trim(!STATUS_CODE)
            
            If Grd.Row = 1 Then
                txtLModUser = "" & Trim(!STATUS_CUR_USER)
                txtLModDate = Format(!STATUS_CUR_DATE, "mm/dd/yy")
                txtStatCode = "" & Trim(!STATUS_REF)
                txtStatCmt = "" & Trim(!Comment)
            End If
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
   End If
   Set RdoGrd = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
    Dim strStatID As String
    Dim iActStat As Integer
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        cmbStatID = ""
        lblStatCd = ""
        Grd.Col = 1
        strStatID = Grd.Text
        
        GetStatusCodeInfo strStatID
    End If
   

End Sub


Private Sub grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim strStatID As String
    cmbStatID = ""
    lblStatCd = ""
    
    Grd.Col = 1
    strStatID = Grd.Text
    
    GetStatusCodeInfo strStatID
End Sub


Private Sub GetStatusCodeInfo(strStatID As String)
     Dim RdoStat As rdoResultset
    
     Dim strStatModRef As String
     Dim strStatModRef1 As String
     Dim strStatModRef2 As String
     Dim strTransType As String
     Dim sSql1 As String
     
    On Error GoTo DiaErr1
    
    strStatModRef = txtSCTRef
    strStatModRef1 = IIf(lblSCTRef1 = "", 0, lblSCTRef1)
    strStatModRef2 = lblSCTRef2
    strTransType = lblStatType
   
    
    sSql = "SELECT STATUS_CUR_USER, StCmtTable.STATUS_REF, STATUS_CUR_DATE, " & _
            " ISNULL(COMMENT, '') COMMENT " & _
        " FROM StCmtTable, StcodeTable " & _
            " WHERE StCmtTable.STATUS_REF = StcodeTable.STATUS_REF"

    sSql1 = " AND MODULE_REF = '" & strTransType & "' AND " & _
                 " STATUS_MODULE_REF = '" & strStatModRef & "' " & _
                 " AND STATUS_MODULE_REF1 = '" & strStatModRef1 & "' " & _
                 " AND STATUS_MODULE_REF2 = '" & strStatModRef2 & "' " & _
                 " AND StCmtTable.STATUS_REF = '" & strStatID & "'"
                 '" AND STATUS_ACT_STATE = 1 " & _

    sSql = sSql & sSql1
   
    bSqlRows = GetDataSet(RdoStat, ES_FORWARD)
    If bSqlRows Then
        With RdoStat
            txtLModUser = "" & Trim(!STATUS_CUR_USER)
            txtLModDate = Format(!STATUS_CUR_DATE, "mm/dd/yy")
            txtStatCode = "" & Trim(!STATUS_REF)
            txtStatCmt = "" & Trim(!Comment)

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
    
    Dim RdoStc As rdoResultset
    On Error GoTo DiaErr1
    
    Dim strStatModRef As String
    Dim strStatModRef1 As String
    Dim strStatModRef2 As String
    Dim strTransType As String
    Dim strRes As String
    Dim strCurUser As String
    Dim strComments As String
     
    On Error GoTo DiaErr1
    
    strStatModRef = txtSCTRef
    strStatModRef1 = IIf(lblSCTRef1 = "", 0, lblSCTRef1)
    strStatModRef2 = lblSCTRef2
    strTransType = lblStatType
    strCurUser = txtCurUser
    strComments = txtStatCmt
    
    If strStatModRef = "" Then
      MsgBox "Please Select a Sales Order.", _
         vbInformation, Caption
        Exit Sub
    End If
    
    sSql = "Qry_AddInternStatCode '" & strStatModRef & "','" & strStatModRef1 & _
                        "','" & strStatModRef2 & "','" & strStatID & _
                        "','" & strTransType & "','" & strCurUser & _
                        "','" & strComments & "'," & iActStat

    bSqlRows = GetDataSet(RdoStc, ES_FORWARD)
    If bSqlRows Then
       With RdoStc
          strRes = "" & Trim(.rdoColumns(0))
          ClearResultSet RdoStc
       End With
    End If

    
    Exit Sub
DiaErr1:
   sProcName = "AddStatusCode"
   CurrError.Number = Err.Number
    CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub RemoveStatusCode(strStatID As String)

     Dim strStatModRef As String
     Dim strStatModRef1 As String
     Dim strStatModRef2 As String
     Dim strTransType As String
     Dim sSql1 As String
     
    On Error GoTo DiaErr1
    
    strStatModRef = txtSCTRef
    strStatModRef1 = IIf(lblSCTRef1 = "", 0, lblSCTRef1)
    strStatModRef2 = lblSCTRef2
    strTransType = lblStatType
    
    sSql = "DELETE FROM StCmtTable WHERE MODULE_REF = '" & strTransType & "' " & _
                 " AND STATUS_MODULE_REF = '" & strStatModRef & "' " & _
                 " AND STATUS_MODULE_REF1 = '" & strStatModRef1 & "' " & _
                 " AND STATUS_MODULE_REF2 = '" & strStatModRef2 & "' " & _
                 " AND StCmtTable.STATUS_REF = '" & strStatID & "'"

   
    RdoCon.Execute sSql, rdExecDirect
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
