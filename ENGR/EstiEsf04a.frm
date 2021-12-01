VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiEsf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select A Routing"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox optFrom 
      Caption         =   "From ppi"
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Select This Routing For The Estimate"
      Top             =   720
      Width           =   852
   End
   Begin VB.ComboBox cmbRte 
      Height          =   288
      Left            =   1500
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Add/Edit Routing"
      Top             =   720
      WhatsThisHelpID =   100
      Width           =   3345
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "2"
      Text            =   " "
      ToolTipText     =   "(30) Char Maximun"
      Top             =   1080
      Width           =   3075
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1920
      FormDesignWidth =   6225
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1152
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1332
   End
End
Attribute VB_Name = "EstiEsf04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'4/11/06 New
Option Explicit
Dim bOnLoad As Byte
Dim bGoodRte As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Public Sub FillCombo()
   Dim RdoRtg As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_FillRoutings "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRtg, ES_FORWARD)
   If bSqlRows Then
      With RdoRtg
         cmbRte = "" & Trim(!RTNUM)
         Do Until .EOF
            AddComboStr cmbRte.hwnd, "" & Trim(!RTNUM)
            .MoveNext
         Loop
         ClearResultSet RdoRtg
      End With
   End If
   If cmbRte.ListCount > 0 Then GetRouting
   Set RdoRtg = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Set RdoRtg = Nothing
   DoModuleErrors MDISect.ActiveForm
   
End Sub

Private Sub GetRouting()
   Dim RdoRte As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RTREF,RTNUM,RTDESC FROM RthdTable " _
          & "WHERE RTREF='" & Compress(cmbRte) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         cmbRte = "" & Trim(!RTNUM)
         txtDsc = "" & Trim(!RTDESC)
         cmdSel.Enabled = True
         bGoodRte = 1
         ClearResultSet RdoRte
      End With
   Else
      bGoodRte = 0
      cmdSel.Enabled = False
      txtDsc = "*** Routing Wasn't Found ***"
   End If
   Set RdoRte = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getrouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbRte_Click()
   GetRouting
   
End Sub


Private Sub cmbRte_LostFocus()
   GetRouting
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdSel_Click()
   If bGoodRte = 1 Then
      AlwaysOnTop Me.hwnd, False
      If optFrom.value = vbChecked Then
         ' MM TODO: ppiESe02a.lblRouting = Compress(cmbRte)
      Else
         EstiESe02a.lblRouting = Compress(cmbRte)
      End If
      Unload Me
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   AlwaysOnTop Me.hwnd, True
   'Move 2000, 3000
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set EstiEsf04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub
