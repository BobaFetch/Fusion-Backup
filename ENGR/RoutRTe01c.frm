VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTe01c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routing Auto Increment"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   1200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1380
      FormDesignWidth =   4170
   End
   Begin VB.TextBox txtAut 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Increment Routing Steps (5 to 50)"
      Top             =   600
      Width           =   555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   875
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Increment Routing Steps By"
      Height          =   612
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1332
   End
End
Attribute VB_Name = "RoutRTe01c"
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
Dim iOldIncrement

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub Form_Load()
   SetFormSize Me
   Move 500, 500
   Dim RdoDef As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RTEINCREMENT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDef)
   If bSqlRows Then
      iAutoIncr = RdoDef!RTEINCREMENT
   Else
      iAutoIncr = 10
   End If
   If iAutoIncr <= 0 Then iAutoIncr = 10
   txtAut = Format(iAutoIncr, "##0")
   iOldIncrement = iAutoIncr
   Set RdoDef = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "form_load"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   If Val(txtAut) <> iAutoIncr Then
      bResponse = MsgBox("Change Routing Increment Settings?", 4 + 32 + 0, Caption)
      If bResponse = vbNo Then
         txtAut = Format(iAutoIncr, "##0")
         Cancel = True
      Else
         iAutoIncr = Val(txtAut)
         sSql = "UPDATE ComnTable SET RTEINCREMENT=" & iAutoIncr & " "
         clsADOCon.ExecuteSQL sSql ' rdExecDirect
      End If
   End If
   RoutRTe01a.optSet.value = vbUnchecked
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set RoutRTe01c = Nothing
   
End Sub


Private Sub txtAut_GotFocus()
   SelectFormat Me
   
End Sub


Private Sub txtAut_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtAut_KeyPress(KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtAut_LostFocus()
   Dim iList As Integer
   txtAut = CheckLen(txtAut, 3)
   txtAut = Format(Abs(Val(txtAut)), "##0")
   iList = Val(txtAut)
   If iList < 5 Or iList > 50 Then
      Beep
      txtAut = Format(iOldIncrement, "##0")
   End If
   
End Sub
