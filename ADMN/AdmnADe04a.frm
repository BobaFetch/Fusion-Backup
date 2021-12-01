VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form AdmnADe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Status Code"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "AdmnADe04a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdClear 
      Cancel          =   -1  'True
      Caption         =   "Cl&ear"
      Height          =   315
      Left            =   5160
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Clear Status ID and Status Code "
      Top             =   960
      Width           =   875
   End
   Begin VB.TextBox txtStCode 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "Add Status Code"
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtStID 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Tag             =   "2"
      ToolTipText     =   "Status ID"
      Top             =   650
      Width           =   1150
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "&Close"
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   960
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4170
      FormDesignWidth =   6330
   End
   Begin MSFlexGridLib.MSFlexGrid grdStaCode 
      Height          =   2535
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Select to change Status Code"
      Top             =   1440
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   10
      Cols            =   6
      FixedCols       =   0
      Enabled         =   -1  'True
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   360
      Picture         =   "AdmnADe04a.frx":08CA
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   0
      Picture         =   "AdmnADe04a.frx":0C54
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Code"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status ID"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   650
      Width           =   1335
   End
End
Attribute VB_Name = "AdmnADe04a"
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
Dim bGoodStdID As Byte
Dim bOnLoad As Byte
Dim bChanges As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmdCan_Click()
    Unload Me
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdStCode_Click()
    
    If Trim(txtStID) = "" Then
        MsgBox "Please enter the Status ID.", _
           vbInformation, Caption
        'txtStID.SetFocus
        Exit Sub
    Else
        UpdateStdCode
    End If
End Sub

Private Sub UpdateStdCode()
    Dim RdoStc As ADODB.Recordset
    On Error GoTo DiaErr1
    
    Dim strStatCode As String
    Dim strStatID As String
    Dim strRes As String
    
    strStatCode = txtStCode
    strStatID = txtStID

    sSql = "Qry_UpdateStatusCode '" & strStatID & "','" & strStatCode & "'"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoStc, ES_FORWARD)
    If bSqlRows Then
       With RdoStc
          strRes = "" & Trim(.Fields(0))
          ClearResultSet RdoStc
       End With
    End If
    Set RdoStc = Nothing
    
    FillStatusCode
    ' Clear the fields
    txtStCode = ""
    txtStID = ""
    bChanges = False
    
    Exit Sub
DiaErr1:
   sProcName = "cmdStCode_Click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdClear_Click()
    ' Clear the fields
    txtStID.Enabled = True
    'txtStID.SetFocus
    txtStCode = ""
    txtStID = ""

End Sub

Private Sub Form_Activate()
    MDISect.lblBotPanel = Caption
    If bOnLoad Then
        bOnLoad = 0
        
        ' Clear the fields
        txtStID.Enabled = True
        txtStCode = ""
        txtStID = ""
    End If
    MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   With grdStaCode
      .Rows = 1
      .row = 0
      .Col = 0
      .Text = "StatusID"
      .Col = 1
      .Text = "Status Code"
      
      .ColWidth(0) = 1250
      .ColWidth(1) = 3450
   End With
   
   FillStatusCode
   bChanges = False
   bOnLoad = 1
   
End Sub


Private Sub FillStatusCode()
   
   Dim RdoStc As ADODB.Recordset
   Dim iRows As Integer
   
   On Error GoTo DiaErr1
   
   grdStaCode.Rows = 1
   sSql = "SELECT STATUS_REF, STATUS_CODE " & _
                " FROM StcodeTable ORDER BY STATUS_REF"
                
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStc, ES_FORWARD)
   If bSqlRows Then
      With RdoStc
         MouseCursor 13
         Do Until .EOF
            iRows = iRows + 1
            grdStaCode.Rows = grdStaCode.Rows + 1
            grdStaCode.row = grdStaCode.Rows - 1
            grdStaCode.Col = 0
            grdStaCode.Text = "" & Trim(!STATUS_REF)
            grdStaCode.Col = 1
            grdStaCode.Text = "" & Trim(!STATUS_CODE)
            .MoveNext
         Loop
         ClearResultSet RdoStc
      End With
      MouseCursor 0
   End If
    
   Set RdoStc = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "FillStatusCode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoStd = Nothing
   Set AdmnADe04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub grdStaCode_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        grdStaCode.Col = 0
        If grdStaCode.row = 0 Then grdStaCode.row = 1
      
        grdStaCode.Col = 0
        txtStID = grdStaCode.Text
        grdStaCode.Col = 1
        txtStCode = grdStaCode.Text
        grdStaCode.Col = 0
        ' disable editing the status ID field
        txtStID.Enabled = False
   End If
   
End Sub


Private Sub grdStaCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  grdStaCode.Col = 0
  If grdStaCode.row = 0 Then grdStaCode.row = 1
  
  grdStaCode.Col = 0
  txtStID = grdStaCode.Text
  grdStaCode.Col = 1
  txtStCode = grdStaCode.Text
  
  ' disable editing the status ID field
  txtStID.Enabled = False
   
End Sub

Private Function CheckStatusCode(strStdID As String) As Byte
   Dim RdoStc As ADODB.Recordset
   Dim iRows As Integer
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT STATUS_REF, STATUS_CODE " & _
                " FROM StcodeTable WHERE STATUS_REF = '" & strStdID & "'"
                
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStc, ES_FORWARD)
   If bSqlRows Then
      CheckStatusCode = 1
   Else
      CheckStatusCode = 0
   End If
   Set RdoStc = Nothing
   Exit Function
   
DiaErr1:
   CheckStatusCode = 0
   sProcName = "CheckStatusCode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function


Private Sub txtStCode_KeyDown(KeyCode As Integer, Shift As Integer)
   bChanges = True
End Sub

Private Sub txtStCode_LostFocus()
    If Trim(txtStID) = "" Then
        Exit Sub
    Else
        UpdateStdCode
    End If
End Sub

Private Sub txtStID_KeyPress(KeyAscii As Integer)
    KeyCase KeyAscii
End Sub

Private Sub txtStID_LostFocus()
    Dim sMsg  As String
    Dim bResponse  As Byte
    Dim sStdID As String
    Dim lRows As Long
    Dim j As Integer
    Dim tmpSID As String
    Dim tmpSCode As String
    
    If bCancel Then Exit Sub
    If Trim(txtStID) = "" Then
        MsgBox "Please enter the Status ID.", _
           vbInformation, Caption
        bGoodStdID = 0
        Exit Sub
    Else
        sStdID = txtStID
        bGoodStdID = CheckStatusCode(txtStID)
    End If
    
    If bGoodStdID = 0 Then
        sMsg = "Add A New " & txtStID & " statusCode ?"
        bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
        If bResponse = vbYes Then
            UpdateStdCode
            txtStID = sStdID
        Else
            txtStID = ""
            txtStID.SetFocus
        End If
    Else
        lRows = grdStaCode.Rows
        For j = 1 To lRows
            grdStaCode.row = j
            grdStaCode.Col = 0
            tmpSID = grdStaCode.Text
            grdStaCode.Col = 1
            tmpSCode = grdStaCode.Text
            If tmpSID = sStdID Then
                txtStCode = tmpSCode
                Exit For
            End If
        Next
    End If
End Sub
