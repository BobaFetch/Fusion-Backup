VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ViewTool 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tool List"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "viewTool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Double Click To Insert Part Number Or Select And Press Enter"
      Top             =   600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
   End
   Begin VB.Label lblLst 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool List"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "ViewTool"
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


Private Sub Command1_Click()
   Form_Deactivate
   
End Sub

Private Sub Form_Activate()
   If bOnLoad Then GetToolList
   bOnLoad = 0
   
End Sub

Private Sub Form_DblClick()
   Unload Me
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Unload Me
   
End Sub

Private Sub Form_Initialize()
   BackColor = ES_ViewBackColor
   
End Sub

Private Sub Form_Load()
   On Error Resume Next
   If MDISect.SideBar.Visible = False Then
      Move MDISect.Left + MDISect.ActiveForm.Left + 800, MDISect.Top + 3200
   Else
      Move MDISect.Left + MDISect.ActiveForm.Left + 2600, MDISect.Top + 3600
   End If
   With Grid1
      .Rows = 2
      .ColWidth(0) = 1400
      .ColWidth(1) = 2600
      .ColWidth(2) = 800
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .Row = 0
      .Col = 0
      .Text = "Tool Class"
      .Col = 1
      .Text = "Tool"
      .Col = 2
      .Text = "Quantity "
      .Col = 3
   End With
   bOnLoad = 1
   
End Sub


Private Sub GetToolList()
   Dim RdoGet As ADODB.Recordset
   Dim iRow As Integer
   Dim sClass As String
   On Error Resume Next
   Grid1.Rows = 1
   Grid1.Row = 1
   sSql = "SELECT TOOLLISTIT_REF,TOOLLISTIT_NUM,TOOLLISTIT_TOOLREF," _
          & "TOOLLISTIT_CLASS,TOOLLISTIT_QUANTITYUSED,TOOL_NUM,TOOL_PARTREF " _
          & "FROM TlitTable,TohdTable where (TOOLLISTIT_REF ='" & Compress(lblLst) _
          & "' AND TOOLLISTIT_TOOLREF=TOOL_PARTREF) ORDER BY TOOLLISTIT_CLASS,TOOL_NUM"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         Do Until .EOF
            iRow = iRow + 1
            If iRow > 300 Then Exit Do
            Grid1.Rows = iRow + 1
            Grid1.Col = 0
            Grid1.Row = iRow
            If sClass <> Trim(!TOOLLISTIT_CLASS) Then _
                              Grid1.Text = Trim(!TOOLLISTIT_CLASS) _
                              Else Grid1.Text = ""
            sClass = "" & Trim(!TOOLLISTIT_CLASS)
            Grid1.Col = 1
            Grid1.Text = "" & Trim(!TOOL_NUM)
            Grid1.Col = 2
            Grid1.Text = "" & Format(!TOOLLISTIT_QUANTITYUSED, "##0")
            .MoveNext
         Loop
         ClearResultSet RdoGet
      End With
   End If
   Set RdoGet = Nothing
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   MDISect.ActiveForm.optSle.value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub



Private Sub Form_Unload(Cancel As Integer)
   Set ViewTool = Nothing
   
End Sub

Private Sub Grid1_DblClick()
   On Error Resume Next
   Grid1.Col = 0
   Unload Me
   
End Sub
