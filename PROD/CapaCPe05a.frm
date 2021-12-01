VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Calendar Template"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPe05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   3480
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3630
      FormDesignWidth =   6945
   End
   Begin VB.TextBox txtSt1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   52
      Top             =   2280
      Width           =   620
   End
   Begin VB.TextBox txtHr1 
      Height          =   285
      Index           =   7
      Left            =   2660
      TabIndex        =   53
      Top             =   2280
      Width           =   465
   End
   Begin VB.TextBox txtSt2 
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   54
      Top             =   2560
      Width           =   620
   End
   Begin VB.TextBox txtHr2 
      Height          =   285
      Index           =   7
      Left            =   2660
      TabIndex        =   55
      Top             =   2560
      Width           =   465
   End
   Begin VB.TextBox txtSt3 
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   56
      Top             =   2850
      Width           =   620
   End
   Begin VB.TextBox txtHr3 
      Height          =   285
      Index           =   7
      Left            =   2660
      TabIndex        =   57
      Top             =   2850
      Width           =   465
   End
   Begin VB.TextBox txtSt4 
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   58
      Top             =   3140
      Width           =   620
   End
   Begin VB.TextBox txtHr4 
      Height          =   285
      Index           =   7
      Left            =   2660
      TabIndex        =   59
      Top             =   3140
      Width           =   465
   End
   Begin VB.TextBox txtSt1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   840
      TabIndex        =   44
      Top             =   2280
      Width           =   620
   End
   Begin VB.TextBox txtHr1 
      Height          =   285
      Index           =   6
      Left            =   1460
      TabIndex        =   45
      Top             =   2280
      Width           =   465
   End
   Begin VB.TextBox txtSt2 
      Height          =   285
      Index           =   6
      Left            =   840
      TabIndex        =   46
      Top             =   2560
      Width           =   620
   End
   Begin VB.TextBox txtHr2 
      Height          =   285
      Index           =   6
      Left            =   1460
      TabIndex        =   47
      Top             =   2560
      Width           =   465
   End
   Begin VB.TextBox txtSt3 
      Height          =   285
      Index           =   6
      Left            =   840
      TabIndex        =   48
      Top             =   2850
      Width           =   620
   End
   Begin VB.TextBox txtHr3 
      Height          =   285
      Index           =   6
      Left            =   1460
      TabIndex        =   49
      Top             =   2850
      Width           =   465
   End
   Begin VB.TextBox txtSt4 
      Height          =   285
      Index           =   6
      Left            =   840
      TabIndex        =   50
      Top             =   3140
      Width           =   620
   End
   Begin VB.TextBox txtHr4 
      Height          =   285
      Index           =   6
      Left            =   1460
      TabIndex        =   51
      Top             =   3140
      Width           =   465
   End
   Begin VB.TextBox txtSt1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   36
      Top             =   840
      Width           =   620
   End
   Begin VB.TextBox txtHr1 
      Height          =   285
      Index           =   5
      Left            =   6260
      TabIndex        =   37
      Top             =   840
      Width           =   465
   End
   Begin VB.TextBox txtSt2 
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   38
      Top             =   1120
      Width           =   620
   End
   Begin VB.TextBox txtHr2 
      Height          =   285
      Index           =   5
      Left            =   6260
      TabIndex        =   39
      Top             =   1120
      Width           =   465
   End
   Begin VB.TextBox txtSt3 
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   40
      Top             =   1410
      Width           =   620
   End
   Begin VB.TextBox txtHr3 
      Height          =   285
      Index           =   5
      Left            =   6260
      TabIndex        =   41
      Top             =   1410
      Width           =   465
   End
   Begin VB.TextBox txtSt4 
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   42
      Top             =   1690
      Width           =   620
   End
   Begin VB.TextBox txtHr4 
      Height          =   285
      Index           =   5
      Left            =   6260
      TabIndex        =   43
      Top             =   1690
      Width           =   465
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.TextBox txtHr4 
      Height          =   285
      Index           =   4
      Left            =   5040
      TabIndex        =   35
      Top             =   1690
      Width           =   465
   End
   Begin VB.TextBox txtSt4 
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   34
      Top             =   1690
      Width           =   620
   End
   Begin VB.TextBox txtHr3 
      Height          =   285
      Index           =   4
      Left            =   5040
      TabIndex        =   33
      Top             =   1410
      Width           =   465
   End
   Begin VB.TextBox txtSt3 
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   32
      Top             =   1410
      Width           =   620
   End
   Begin VB.TextBox txtHr2 
      Height          =   285
      Index           =   4
      Left            =   5040
      TabIndex        =   31
      Top             =   1120
      Width           =   465
   End
   Begin VB.TextBox txtSt2 
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   30
      Top             =   1120
      Width           =   620
   End
   Begin VB.TextBox txtHr1 
      Height          =   285
      Index           =   4
      Left            =   5040
      TabIndex        =   29
      Top             =   840
      Width           =   465
   End
   Begin VB.TextBox txtSt1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   28
      Top             =   840
      Width           =   620
   End
   Begin VB.TextBox txtHr4 
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   27
      Top             =   1690
      Width           =   465
   End
   Begin VB.TextBox txtSt4 
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   26
      Top             =   1690
      Width           =   620
   End
   Begin VB.TextBox txtHr3 
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   25
      Top             =   1410
      Width           =   465
   End
   Begin VB.TextBox txtSt3 
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   24
      Top             =   1410
      Width           =   620
   End
   Begin VB.TextBox txtHr2 
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   23
      Top             =   1120
      Width           =   465
   End
   Begin VB.TextBox txtSt2 
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   22
      Top             =   1120
      Width           =   620
   End
   Begin VB.TextBox txtHr1 
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   21
      Top             =   840
      Width           =   465
   End
   Begin VB.TextBox txtSt1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   20
      Top             =   840
      Width           =   620
   End
   Begin VB.TextBox txtHr4 
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   19
      Top             =   1690
      Width           =   465
   End
   Begin VB.TextBox txtSt4 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   18
      Top             =   1690
      Width           =   620
   End
   Begin VB.TextBox txtHr3 
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   17
      Top             =   1410
      Width           =   465
   End
   Begin VB.TextBox txtSt3 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   16
      Top             =   1410
      Width           =   620
   End
   Begin VB.TextBox txtHr2 
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   15
      Top             =   1120
      Width           =   465
   End
   Begin VB.TextBox txtSt2 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   14
      Top             =   1120
      Width           =   620
   End
   Begin VB.TextBox txtHr1 
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   13
      Top             =   840
      Width           =   465
   End
   Begin VB.TextBox txtSt1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   12
      Top             =   840
      Width           =   620
   End
   Begin VB.TextBox txtHr4 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      Top             =   1690
      Width           =   465
   End
   Begin VB.TextBox txtSt4 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   10
      Top             =   1690
      Width           =   620
   End
   Begin VB.TextBox txtHr3 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   9
      Top             =   1410
      Width           =   465
   End
   Begin VB.TextBox txtSt3 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   8
      Top             =   1410
      Width           =   620
   End
   Begin VB.TextBox txtHr2 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   1120
      Width           =   465
   End
   Begin VB.TextBox txtSt2 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   6
      Top             =   1120
      Width           =   620
   End
   Begin VB.TextBox txtHr1 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   465
   End
   Begin VB.TextBox txtSt1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   620
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sunday"
      Height          =   195
      Index           =   10
      Left            =   840
      TabIndex        =   71
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Saturday"
      Height          =   195
      Index           =   9
      Left            =   2040
      TabIndex        =   70
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Friday"
      Height          =   195
      Index           =   8
      Left            =   840
      TabIndex        =   69
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thursday"
      Height          =   195
      Index           =   7
      Left            =   5640
      TabIndex        =   68
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday"
      Height          =   195
      Index           =   6
      Left            =   4440
      TabIndex        =   67
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tuesday"
      Height          =   195
      Index           =   5
      Left            =   3240
      TabIndex        =   66
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Monday"
      Height          =   195
      Index           =   4
      Left            =   2040
      TabIndex        =   65
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift 1:"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   64
      Top             =   2280
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift 2:"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   63
      Top             =   2560
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift 3:"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   62
      Top             =   2850
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift 4:"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   61
      Top             =   3140
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift 4:"
      Height          =   285
      Index           =   17
      Left            =   120
      TabIndex        =   3
      Top             =   1690
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift 3:"
      Height          =   285
      Index           =   16
      Left            =   120
      TabIndex        =   2
      Top             =   1410
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift 2:"
      Height          =   285
      Index           =   15
      Left            =   120
      TabIndex        =   1
      Top             =   1120
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift 1:"
      Height          =   285
      Index           =   14
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   585
   End
End
Attribute VB_Name = "CapaCPe05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/12/06 Revised ToolTipText
Option Explicit
Dim bGoodCal As Byte
Dim bChanged As Byte

Dim sShiftStart(8, 5) As String
Dim cShiftHours(8, 5) As Currency

Private Sub cmdCan_Click()
   If bChanged Then UpdateTemplate
   Unload Me
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4105
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   sCurrForm = ""
   bGoodCal = False
   GetTemplate
   Show
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set CapaCPe05a = Nothing
   
End Sub


Private Sub txtHr1_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtHr1_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtHr1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtHr1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtHr1_LostFocus(Index As Integer)
   txtHr1(Index) = CheckLen(txtHr1(Index), 4)
   If Val(txtHr1(Index)) > 12 Then txtHr1(Index) = "12.0"
   txtHr1(Index) = Format(Val(txtHr1(Index)), "#0.0")
   
End Sub


Private Sub txtHr2_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtHr2_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtHr2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtHr2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtHr2_LostFocus(Index As Integer)
   txtHr2(Index) = CheckLen(txtHr2(Index), 4)
   If Val(txtHr2(Index)) > 12 Then txtHr2(Index) = "12.0"
   txtHr2(Index) = Format(Val(txtHr2(Index)), "#0.0")
   
End Sub


Private Sub txtHr3_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtHr3_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtHr3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtHr3_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtHr3_LostFocus(Index As Integer)
   txtHr3(Index) = CheckLen(txtHr3(Index), 4)
   If Val(txtHr3(Index)) > 12 Then txtHr3(Index) = "12.0"
   txtHr3(Index) = Format(Val(txtHr3(Index)), "#0.0")
   
End Sub


Private Sub txtHr4_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtHr4_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtHr4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtHr4_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtHr4_LostFocus(Index As Integer)
   txtHr4(Index) = CheckLen(txtHr4(Index), 4)
   If Val(txtHr4(Index)) > 12 Then txtHr4(Index) = "12.0"
   txtHr4(Index) = Format(Val(txtHr4(Index)), "#0.0")
   
End Sub


Private Sub txtSt1_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtSt1_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtSt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtSt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtSt1_LostFocus(Index As Integer)
   txtSt1(Index) = CheckLen(txtSt1(Index), 6)
   Dim tc As New ClassTimeCharge
   txtSt1(Index) = tc.GetTime(txtSt1(Index))
   
End Sub


Private Sub txtSt2_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtSt2_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtSt2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtSt2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtSt2_LostFocus(Index As Integer)
   txtSt2(Index) = CheckLen(txtSt2(Index), 6)
   Dim tc As New ClassTimeCharge
   txtSt2(Index) = tc.GetTime(txtSt2(Index))
   
End Sub


Private Sub txtSt3_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtSt3_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtSt3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtSt3_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtSt3_LostFocus(Index As Integer)
   txtSt3(Index) = CheckLen(txtSt3(Index), 6)
   Dim tc As New ClassTimeCharge
   txtSt3(Index) = tc.GetTime(txtSt3(Index))
   
End Sub


Private Sub txtSt4_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtSt4_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtSt4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtSt4_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtSt4_LostFocus(Index As Integer)
   txtSt4(Index) = CheckLen(txtSt4(Index), 6)
   Dim tc As New ClassTimeCharge
   txtSt4(Index) = tc.GetTime(txtSt4(Index))
   
End Sub



Private Sub GetTemplate()
   Dim RdoCal As ADODB.Recordset
   Dim iList As Integer
   
   On Error GoTo DiaErr1
   bGoodCal = False
   sSql = "SELECT * FROM CctmTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCal)
   If bSqlRows Then
      With RdoCal
         sShiftStart(1, 1) = "" & !CALSUNST1
         sShiftStart(1, 2) = "" & !CALSUNST2
         sShiftStart(1, 3) = "" & !CALSUNST3
         sShiftStart(1, 4) = "" & !CALSUNST4
         
         sShiftStart(2, 1) = "" & !CALMONST1
         sShiftStart(2, 2) = "" & !CALMONST2
         sShiftStart(2, 3) = "" & !CALMONST3
         sShiftStart(2, 4) = "" & !CALMONST4
         
         sShiftStart(3, 1) = "" & !CALTUEST1
         sShiftStart(3, 2) = "" & !CALTUEST2
         sShiftStart(3, 3) = "" & !CALTUEST3
         sShiftStart(3, 4) = "" & !CALTUEST4
         
         sShiftStart(4, 1) = "" & !CALWEDST1
         sShiftStart(4, 2) = "" & !CALWEDST2
         sShiftStart(4, 3) = "" & !CALWEDST3
         sShiftStart(4, 4) = "" & !CALWEDST4
         
         sShiftStart(5, 1) = "" & !CALTHUST1
         sShiftStart(5, 2) = "" & !CALTHUST2
         sShiftStart(5, 3) = "" & !CALTHUST3
         sShiftStart(5, 4) = "" & !CALTHUST4
         
         sShiftStart(6, 1) = "" & !CALFRIST1
         sShiftStart(6, 2) = "" & !CALFRIST2
         sShiftStart(6, 3) = "" & !CALFRIST3
         sShiftStart(6, 4) = "" & !CALFRIST4
         
         sShiftStart(7, 1) = "" & !CALSATST1
         sShiftStart(7, 2) = "" & !CALSATST2
         sShiftStart(7, 3) = "" & !CALSATST3
         sShiftStart(7, 4) = "" & !CALSATST4
         
         cShiftHours(1, 1) = 0 + !CALSUNHR1
         cShiftHours(1, 2) = 0 + !CALSUNHR2
         cShiftHours(1, 3) = 0 + !CALSUNHR3
         cShiftHours(1, 4) = 0 + !CALSUNHR4
         
         cShiftHours(2, 1) = 0 + !CALMONHR1
         cShiftHours(2, 2) = 0 + !CALMONHR2
         cShiftHours(2, 3) = 0 + !CALMONHR3
         cShiftHours(2, 4) = 0 + !CALMONHR4
         
         cShiftHours(3, 1) = 0 + !CALTUEHR1
         cShiftHours(3, 2) = 0 + !CALTUEHR2
         cShiftHours(3, 3) = 0 + !CALTUEHR3
         cShiftHours(3, 4) = 0 + !CALTUEHR4
         
         cShiftHours(4, 1) = 0 + !CALWEDHR1
         cShiftHours(4, 2) = 0 + !CALWEDHR2
         cShiftHours(4, 3) = 0 + !CALWEDHR3
         cShiftHours(4, 4) = 0 + !CALWEDHR4
         
         cShiftHours(5, 1) = 0 + !CALTHUHR1
         cShiftHours(5, 2) = 0 + !CALTHUHR2
         cShiftHours(5, 3) = 0 + !CALTHUHR3
         cShiftHours(5, 4) = 0 + !CALTHUHR4
         
         cShiftHours(6, 1) = 0 + !CALFRIHR1
         cShiftHours(6, 2) = 0 + !CALFRIHR2
         cShiftHours(6, 3) = 0 + !CALFRIHR3
         cShiftHours(6, 4) = 0 + !CALFRIHR4
         
         cShiftHours(7, 1) = 0 + !CALSATHR1
         cShiftHours(7, 2) = 0 + !CALSATHR2
         cShiftHours(7, 3) = 0 + !CALSATHR3
         cShiftHours(7, 4) = 0 + !CALSATHR4
         ClearResultSet RdoCal
      End With
   End If
   For iList = 1 To 6
      txtSt1(iList) = "" & Trim(sShiftStart(iList, 1))
      txtSt2(iList) = "" & Trim(sShiftStart(iList, 2))
      txtSt3(iList) = "" & Trim(sShiftStart(iList, 3))
      txtSt4(iList) = "" & Trim(sShiftStart(iList, 4))
      
      txtSt1(iList).ToolTipText = "Shift Start-Enter As 8.00a"
      txtSt2(iList).ToolTipText = "Shift Start-Enter As 8.00a"
      txtSt3(iList).ToolTipText = "Shift Start-Enter As 8.00a"
      txtSt4(iList).ToolTipText = "Shift Start-Enter As 8.00a"
      
      txtHr1(iList) = Format(cShiftHours(iList, 1), "#0.0")
      txtHr2(iList) = Format(cShiftHours(iList, 2), "#0.0")
      txtHr3(iList) = Format(cShiftHours(iList, 3), "#0.0")
      txtHr4(iList) = Format(cShiftHours(iList, 4), "#0.0")
      
      txtHr1(iList).ToolTipText = "Shift Hours Enter As 2.5"
      txtHr2(iList).ToolTipText = "Shift Hours Enter As 2.5"
      txtHr3(iList).ToolTipText = "Shift Hours Enter As 2.5"
      txtHr4(iList).ToolTipText = "Shift Hours Enter As 2.5"
   Next
   txtSt1(iList) = "" & Trim(sShiftStart(iList, 1))
   txtSt2(iList) = "" & Trim(sShiftStart(iList, 2))
   txtSt3(iList) = "" & Trim(sShiftStart(iList, 3))
   txtSt4(iList) = "" & Trim(sShiftStart(iList, 4))
   
   txtSt1(iList).ToolTipText = "Shift Start-Enter As 8.00a"
   txtSt2(iList).ToolTipText = "Shift Start-Enter As 8.00a"
   txtSt3(iList).ToolTipText = "Shift Start-Enter As 8.00a"
   txtSt4(iList).ToolTipText = "Shift Start-Enter As 8.00a"
   
   txtHr1(iList) = Format(cShiftHours(iList, 1), "#0.0")
   txtHr2(iList) = Format(cShiftHours(iList, 2), "#0.0")
   txtHr3(iList) = Format(cShiftHours(iList, 3), "#0.0")
   txtHr4(iList) = Format(cShiftHours(iList, 4), "#0.0")
   
   txtHr1(iList).ToolTipText = "Shift Hours Enter As 2.5"
   txtHr2(iList).ToolTipText = "Shift Hours Enter As 2.5"
   txtHr3(iList).ToolTipText = "Shift Hours Enter As 2.5"
   txtHr4(iList).ToolTipText = "Shift Hours Enter As 2.5"
   
   bGoodCal = True
   bChanged = 0
   Erase sShiftStart
   Erase cShiftHours
   Set RdoCal = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "gettemplate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub UpdateTemplate()
   Dim iList As Integer
   Dim A As Integer
   For iList = 1 To 6
      A = A + Val(txtHr1(iList))
      A = A + Val(txtHr2(iList))
      A = A + Val(txtHr3(iList))
      A = A + Val(txtHr4(iList))
   Next
   A = A + Val(txtHr1(iList))
   A = A + Val(txtHr2(iList))
   A = A + Val(txtHr3(iList))
   A = A + Val(txtHr4(iList))
   On Error Resume Next
   sSql = "INSERT INTO CctmTable (CALREF) VALUES('TEMPLATE')"
   clsADOCon.ExecuteSQL sSql
   
   On Error GoTo DiaErr1
   sSql = "UPDATE CctmTable SET " _
          & "CALSUNST1='" & txtSt1(1) & "'," _
          & "CALSUNST2='" & txtSt2(1) & "'," _
          & "CALSUNST3='" & txtSt3(1) & "'," _
          & "CALSUNST4='" & txtSt4(1) & "'," _
          & "CALMONST1='" & txtSt1(2) & "'," _
          & "CALMONST2='" & txtSt2(2) & "'," _
          & "CALMONST3='" & txtSt3(2) & "'," _
          & "CALMONST4='" & txtSt4(2) & "'," _
          & "CALTUEST1='" & txtSt1(3) & "'," _
          & "CALTUEST2='" & txtSt2(3) & "'," _
          & "CALTUEST3='" & txtSt3(3) & "'," _
          & "CALTUEST4='" & txtSt4(3) & "'," _
          & "CALWEDST1='" & txtSt1(4) & "'," _
          & "CALWEDST2='" & txtSt2(4) & "'," _
          & "CALWEDST3='" & txtSt3(4) & "'," _
          & "CALWEDST4='" & txtSt4(4) & "'," _
          & "CALTHUST1='" & txtSt1(5) & "'," _
          & "CALTHUST2='" & txtSt2(5) & "'," _
          & "CALTHUST3='" & txtSt3(5) & "'," _
          & "CALTHUST4='" & txtSt4(5) & "'," _
          & "CALFRIST1='" & txtSt1(6) & "'," _
          & "CALFRIST2='" & txtSt2(6) & "',"
   sSql = sSql & "CALFRIST3='" & txtSt3(6) & "'," _
          & "CALFRIST4='" & txtSt4(6) & "'," _
          & "CALSATST1='" & txtSt1(7) & "'," _
          & "CALSATST2='" & txtSt2(7) & "'," _
          & "CALSATST3='" & txtSt3(7) & "'," _
          & "CALSATST4='" & txtSt4(7) & "'," _
          & "CALSUNHR1=" & Val(txtHr1(1)) & "," _
          & "CALSUNHR2=" & Val(txtHr2(1)) & "," _
          & "CALSUNHR3=" & Val(txtHr3(1)) & "," _
          & "CALSUNHR4=" & Val(txtHr4(1)) & "," _
          & "CALMONHR1=" & Val(txtHr1(2)) & "," _
          & "CALMONHR2=" & Val(txtHr2(2)) & "," _
          & "CALMONHR3=" & Val(txtHr3(2)) & "," _
          & "CALMONHR4=" & Val(txtHr4(2)) & "," _
          & "CALTUEHR1=" & Val(txtHr1(3)) & "," _
          & "CALTUEHR2=" & Val(txtHr2(3)) & "," _
          & "CALTUEHR3=" & Val(txtHr3(3)) & ","
   sSql = sSql & "CALTUEHR4=" & Val(txtHr4(3)) & "," _
          & "CALWEDHR1=" & Val(txtHr1(4)) & "," _
          & "CALWEDHR2=" & Val(txtHr2(4)) & "," _
          & "CALWEDHR3=" & Val(txtHr3(4)) & "," _
          & "CALWEDHR4=" & Val(txtHr4(4)) & "," _
          & "CALTHUHR1=" & Val(txtHr1(5)) & "," _
          & "CALTHUHR2=" & Val(txtHr2(5)) & "," _
          & "CALTHUHR3=" & Val(txtHr3(5)) & "," _
          & "CALTHUHR4=" & Val(txtHr4(5)) & "," _
          & "CALFRIHR1=" & Val(txtHr1(6)) & "," _
          & "CALFRIHR2=" & Val(txtHr2(6)) & "," _
          & "CALFRIHR3=" & Val(txtHr3(6)) & "," _
          & "CALFRIHR4=" & Val(txtHr4(6)) & "," _
          & "CALSATHR1=" & Val(txtHr1(7)) & "," _
          & "CALSATHR2=" & Val(txtHr2(7)) & "," _
          & "CALSATHR3=" & Val(txtHr3(7)) & "," _
          & "CALSATHR4=" & Val(txtHr4(7)) & "," _
          & "CALTOTHRS=" & str(A) & ""
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected > 0 Then
      SysMsg "Company Template Updated.", True, Me
   Else
      MsgBox "Couldn't Update Template.", vbExclamation, Caption
   End If
   sSql = ""
   Exit Sub
   
DiaErr1:
   sSql = ""
   sProcName = "updatetemplate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
