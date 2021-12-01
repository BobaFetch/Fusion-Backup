VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTe01f 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routing Change"
   ClientHeight    =   2715
   ClientLeft      =   2985
   ClientTop       =   2310
   ClientWidth     =   7155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2715
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRtByPrev 
      Height          =   285
      Left            =   5040
      TabIndex        =   11
      Tag             =   "3"
      ToolTipText     =   "Revision Of Routing"
      Top             =   960
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox txtRby 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   9
      Tag             =   "2"
      ToolTipText     =   "Engineer"
      Top             =   240
      Width           =   2340
   End
   Begin VB.TextBox txtPrevRev 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Revision Of Routing"
      Top             =   840
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txtRev 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Revision Of Routing"
      Top             =   810
      Width           =   465
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Tag             =   "4"
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Height          =   1185
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Tag             =   "9"
      Text            =   "RoutRTe01f.frx":0000
      ToolTipText     =   "Comment (5120 Chars Max)"
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton cmdAppy 
      Caption         =   "&Apply"
      Height          =   285
      Left            =   6000
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Fill The Current Operation"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   990
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2715
      FormDesignWidth =   7155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing By"
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   285
      Index           =   5
      Left            =   4080
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   810
      Width           =   675
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision Notes"
      Height          =   285
      Index           =   14
      Left            =   120
      TabIndex        =   5
      Top             =   1275
      Width           =   1185
   End
End
Attribute VB_Name = "RoutRTe01f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/2/04 Revised general structure and Fill button
'        Attempts to update Ops grid
'1/26/07 Undo
Option Explicit

Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub cmdAppy_Click()

   If ((txtRev = "") Or (txtRby = "") Or _
      (txtDte = "") Or (txtCmt = "")) Then
      MsgBox "Please enter all fields to update .", vbCritical
   Else
      RoutRTe01a.txtRev = txtRev
      RoutRTe01a.txtRby = txtRby
      RoutRTe01a.txtDte = txtDte
      RoutRTe01a.txtCmt = txtCmt & vbCrLf & RoutRTe01a.txtCmt
      RoutRTe01a.chkUpChild = 1
      
      Unload Me
      
   End If

End Sub

Private Sub cmdCan_Click()
   RoutRTe01a.txtRev = txtPrevRev
   RoutRTe01a.chkUpChild = 0
   Unload Me
   
End Sub



Private Sub Form_Activate()
   MouseCursor 13
   If bOnLoad Then
   
      sSql = "SELECT DISTINCT RTBY FROM RthdTable ORDER BY RTBY "
      LoadComboBox txtRby, -1
      
      If txtRby.ListCount > 0 Then txtRby = txtRtByPrev
      
      ES_TimeFormat = GetTimeFormat()
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   Move 2000, 2000
   FormatControls
   bOnLoad = 1
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   MouseCursor 0
   Set RoutRTe01f = Nothing
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtDte_LostFocus()
   If Trim(txtDte) <> "" Then txtDte = CheckDate(txtDte)
End Sub

