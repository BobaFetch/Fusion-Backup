VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Default Routing Assignments"
   ClientHeight    =   4530
   ClientLeft      =   1950
   ClientTop       =   1455
   ClientWidth     =   7290
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4530
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTe03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "<"
      Height          =   285
      Index           =   8
      Left            =   6750
      TabIndex        =   13
      ToolTipText     =   "Assign Routing to Part Type 8"
      Top             =   4050
      Width           =   375
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "<"
      Height          =   285
      Index           =   6
      Left            =   6750
      TabIndex        =   11
      ToolTipText     =   "Assign Routing to Part Type 6"
      Top             =   3510
      Width           =   375
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "<"
      Height          =   285
      Index           =   5
      Left            =   6750
      TabIndex        =   9
      ToolTipText     =   "Assign Routing to Part Type 5"
      Top             =   2970
      Width           =   375
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "<"
      Height          =   285
      Index           =   4
      Left            =   6750
      TabIndex        =   7
      ToolTipText     =   "Assign Routing to Part Type 4"
      Top             =   2430
      Width           =   375
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "<"
      Height          =   285
      Index           =   3
      Left            =   6750
      TabIndex        =   5
      ToolTipText     =   "Assign Routing to Part Type 3"
      Top             =   1890
      Width           =   375
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "<"
      Height          =   285
      Index           =   2
      Left            =   6750
      TabIndex        =   3
      ToolTipText     =   "Assign Routing to Part Type 2"
      Top             =   1350
      Width           =   375
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "<"
      Height          =   285
      Index           =   1
      Left            =   6750
      TabIndex        =   1
      ToolTipText     =   "Assign Routing to Part Type 1"
      Top             =   810
      Width           =   375
   End
   Begin VB.ComboBox cmbPr8 
      Height          =   315
      Left            =   3330
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   4050
      Width           =   3345
   End
   Begin VB.ComboBox cmbPr6 
      Height          =   315
      Left            =   3330
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   3510
      Width           =   3345
   End
   Begin VB.ComboBox cmbPr5 
      Height          =   315
      Left            =   3330
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   2970
      Width           =   3345
   End
   Begin VB.ComboBox cmbPr4 
      Height          =   315
      Left            =   3330
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2430
      Width           =   3345
   End
   Begin VB.ComboBox cmbPr3 
      Height          =   315
      Left            =   3330
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1890
      Width           =   3345
   End
   Begin VB.ComboBox cmbPr2 
      Height          =   315
      Left            =   3330
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1350
      Width           =   3345
   End
   Begin VB.ComboBox cmbPr1 
      Height          =   315
      Left            =   3330
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   810
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6300
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   600
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4530
      FormDesignWidth =   7290
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assign"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   6660
      TabIndex        =   29
      Top             =   540
      Width           =   555
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   180
      TabIndex        =   28
      Top             =   4050
      Width           =   3075
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   180
      TabIndex        =   27
      Top             =   3510
      Width           =   3075
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   180
      TabIndex        =   26
      Top             =   2970
      Width           =   3075
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   25
      Top             =   2430
      Width           =   3075
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   180
      TabIndex        =   24
      Top             =   1890
      Width           =   3075
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   180
      TabIndex        =   23
      Top             =   1350
      Width           =   3075
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   22
      Top             =   810
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type 8                                                  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   180
      TabIndex        =   21
      Top             =   3780
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type 6                                                  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   20
      Top             =   3240
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type 5                                                  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   19
      Top             =   2700
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type 4                                                  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   18
      Top             =   2160
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type 3                                                  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   17
      Top             =   1620
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type 2                                                   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   16
      Top             =   1080
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type 1                                                   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   540
      Width           =   3075
   End
End
Attribute VB_Name = "RoutRTe03a"
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
Dim bGoodRout As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   
End Sub

Private Sub cmdAsn_Click(Index As Integer)
   Dim sCheckRout As String
   MouseCursor 13
   sSql = "UPDATE ComnTable SET "
   Select Case Index
      Case 1
         sSql = sSql & "RTEPART1='"
         sCheckRout = cmbPr1
      Case 2
         sSql = sSql & "RTEPART2='"
         sCheckRout = cmbPr2
      Case 3
         sSql = sSql & "RTEPART3='"
         sCheckRout = cmbPr3
      Case 4
         sSql = sSql & "RTEPART4='"
         sCheckRout = cmbPr4
      Case 5
         sSql = sSql & "RTEPART5='"
         sCheckRout = cmbPr5
      Case 6
         sSql = sSql & "RTEPART6='"
         sCheckRout = cmbPr6
      Case 8
         sSql = sSql & "RTEPART8='"
         sCheckRout = cmbPr8
   End Select
   
   sCheckRout = Compress(sCheckRout)
   If Trim(sCheckRout) = "" Or Trim(sCheckRout) = "NONE" Then Exit Sub
   bGoodRout = GetRouting(sCheckRout, Index)
   MouseCursor 0
   If bGoodRout Then
      sSql = sSql & sCheckRout & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected > 0 Then
         SysMsg "Routing Assigned.", True, Me
      Else
         MsgBox "Couldn't Update Default.", vbExclamation, Caption
      End If
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3103
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillCombos
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
   Set RoutRTe03a = Nothing
   
End Sub



Private Sub FillCombos()
   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   MouseCursor 13
   On Error GoTo DiaErr1
   sSql = "Qry_FillRoutings"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbPr1.hwnd, "" & Trim(!RTNUM)
            AddComboStr cmbPr2.hwnd, "" & Trim(!RTNUM)
            AddComboStr cmbPr3.hwnd, "" & Trim(!RTNUM)
            AddComboStr cmbPr4.hwnd, "" & Trim(!RTNUM)
            AddComboStr cmbPr5.hwnd, "" & Trim(!RTNUM)
            AddComboStr cmbPr6.hwnd, "" & Trim(!RTNUM)
            AddComboStr cmbPr8.hwnd, "" & Trim(!RTNUM)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   
   sSql = "SELECT RTEPART1,RTEPART2,RTEPART3,RTEPART4,RTEPART5,RTEPART6," _
          & "RTEPART8 FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         cmbPr1 = "" & Trim(!RTEPART1)
         cmbPr2 = "" & Trim(!RTEPART2)
         cmbPr3 = "" & Trim(!RTEPART3)
         cmbPr4 = "" & Trim(!RTEPART4)
         cmbPr5 = "" & Trim(!RTEPART5)
         cmbPr6 = "" & Trim(!RTEPART6)
         cmbPr8 = "" & Trim(!RTEPART8)
         ClearResultSet RdoCmb
      End With
      bGoodRout = GetRouting(cmbPr1, 1)
      bGoodRout = GetRouting(cmbPr2, 2)
      bGoodRout = GetRouting(cmbPr3, 3)
      bGoodRout = GetRouting(cmbPr4, 4)
      bGoodRout = GetRouting(cmbPr5, 5)
      bGoodRout = GetRouting(cmbPr6, 6)
      bGoodRout = GetRouting(cmbPr8, 8)
   End If
   For iList = 1 To 6
      If lblType(iList) = "" Then lblType(iList) = "NONE"
   Next
   If lblType(8) = "" Then lblType(8) = "NONE"
   Set RdoCmb = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombos"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetRouting(sRouting As String, bId As Integer) As Boolean
   Dim RdoRte As ADODB.Recordset
   Dim sSql2 As String
   On Error GoTo DiaErr1
   
   sSql2 = sSql
   sSql = "SELECT RTNUM FROM RthdTable WHERE RTREF='" & Compress(sRouting) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         Select Case bId
            Case 1
               cmbPr1 = "" & Trim(!RTNUM)
            Case 2
               cmbPr2 = "" & Trim(!RTNUM)
            Case 3
               cmbPr3 = "" & Trim(!RTNUM)
            Case 4
               cmbPr4 = "" & Trim(!RTNUM)
            Case 5
               cmbPr5 = "" & Trim(!RTNUM)
            Case 6
               cmbPr6 = "" & Trim(!RTNUM)
            Case 7
               cmbPr6 = "" & Trim(!RTNUM)
            Case Else
               cmbPr8 = "" & Trim(!RTNUM)
         End Select
         lblType(bId) = "" & Trim(!RTNUM)
         ClearResultSet RdoRte
      End With
      GetRouting = True
   Else
      GetRouting = False
      MsgBox "Routing Wasn't Found.", vbExclamation, Caption
   End If
   
   sSql = sSql2
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
