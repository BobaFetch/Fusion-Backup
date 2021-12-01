VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form VewAcct 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find An Account"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "VewAcct.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   3
      ToolTipText     =   "Select Accounts"
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "Select Accounts"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtAct 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Enter At Least One Character"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ListBox lstAct 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Double Click Or Select And Press Enter"
      Top             =   1440
      Width           =   4215
   End
   Begin VB.TextBox txtAct 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "Enter At Least One Character"
      Top             =   600
      Width           =   1935
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   3720
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3930
      FormDesignWidth =   4485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find Your Account Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Or"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descriptions"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Numbers"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter At Least (1) Leading Character"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   4215
   End
End
Attribute VB_Name = "VewAcct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bDescSelect As Byte

Private Sub cmdGo_Click(Index As Integer)
   If Index = 0 Then
      txtAct(1) = ""
      FillAccounts
   Else
      txtAct(0) = ""
      FillDescriptions
   End If
   
End Sub

Private Sub Form_DblClick()
   Unload Me
   
End Sub

Private Sub Form_Initialize()
   BackColor = ES_ViewBackColor
   
End Sub


Private Sub Form_Load()
   On Error Resume Next
   If MDISect.SideBar.Visible = False Then
      Move MDISect.Left + MDISect.ActiveForm.Left + 400, MDISect.Top + 2000
   Else
      Move MDISect.Left + MDISect.ActiveForm.Left + 800, MDISect.Top + 2600
   End If
   bUserAction = True
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   MDISect.ActiveForm.optVewAcct = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set VewAcct = Nothing
   
End Sub

Private Sub lstAct_DblClick()
   On Error Resume Next
   If lstAct.ListIndex >= 0 And lstAct.ListIndex < lstAct.ListCount Then
      If bDescSelect = 0 Then
         'diaGLe01a.cmbAct = Left(lstAct, 12)
      Else
         'diaGLe01a.cmbAct = Right(lstAct, 12)
      End If
      Unload Me
   End If
   
End Sub


Private Sub lstAct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error Resume Next
      If bDescSelect = 0 Then
         'diaGLe01a.cmbAct = Left(lstAct, 12)
      Else
         'diaGLe01a.cmbAct = Right(lstAct, 12)
      End If
      Unload Me
   End If
   
End Sub


Private Sub txtAct_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      KeyCase KeyAscii
   Else
      KeyCheck KeyAscii
   End If
   
End Sub


Private Sub txtAct_LostFocus(Index As Integer)
   txtAct(Index) = CheckLen(txtAct(Index), 12)
   If Index = 0 Then
      txtAct(1) = ""
      FillAccounts
   Else
      txtAct(0) = ""
      txtAct(Index) = StrCase(txtAct(Index))
      FillDescriptions
   End If
   
End Sub



Private Sub FillAccounts()
   Dim RdoRns As ADODB.Recordset
   Dim iList As Integer
   Dim sAcctRef As String
   
   lstAct.Clear
   On Error GoTo DiaErr1
   bDescSelect = 0
   sAcctRef = Compress(txtAct(0))
   If Len(sAcctRef) = 0 Then Exit Sub
   If Len(sAcctRef) > 0 Then
      sAcctRef = sAcctRef & "%"
      sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR " _
             & "FROM GlacTable WHERE GLACCTREF Like '" & sAcctRef & "%' " _
             & "ORDER BY GLACCTREF"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns)
      If bSqlRows Then
         With RdoRns
            Do Until .EOF
               iList = iList + 1
               If iList > 300 Then Exit Do
               lstAct.AddItem "" & !GLACCTNO & Chr(9) & Trim(!GLDESCR)
               .MoveNext
            Loop
            ClearResultSet RdoRns
            On Error Resume Next
            lstAct.SetFocus
         End With
      End If
   Else
      MsgBox "Enter At Least (1) Character.", vbInformation, Caption
   End If
   On Error Resume Next
   Set RdoRns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillacct"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillDescriptions()
   Dim RdoRns As ADODB.Recordset
   Dim iList As Integer
   Dim sAcctRef As String
   
   lstAct.Clear
   On Error GoTo DiaErr1
   bDescSelect = 1
   sAcctRef = Compress(txtAct(1))
   If Len(sAcctRef) = 0 Then Exit Sub
   If Len(sAcctRef) > 0 Then
      sAcctRef = sAcctRef & "%"
      sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR " _
             & "FROM GlacTable WHERE GLDESCR Like '" & sAcctRef & "%' " _
             & "ORDER BY GLDESCR,GLACCTREF"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns)
      If bSqlRows Then
         With RdoRns
            Do Until .EOF
               iList = iList + 1
               If iList > 300 Then Exit Do
               lstAct.AddItem "" & Left(!GLDESCR, 30) & Chr(9) & !GLACCTNO
               .MoveNext
            Loop
            ClearResultSet RdoRns
            On Error Resume Next
            lstAct.SetFocus
         End With
      End If
   Else
      MsgBox "Enter At Least (1) Character.", vbInformation, Caption
   End If
   On Error Resume Next
   Set RdoRns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillacct"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
