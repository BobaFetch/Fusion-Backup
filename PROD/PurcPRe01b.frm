VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order Terms"
   ClientHeight    =   2460
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe01b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPnet 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Tag             =   "1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtPdsc 
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   1
      Tag             =   "1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtPday 
      Height          =   285
      Left            =   5040
      TabIndex        =   2
      Tag             =   "1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtPdte 
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Tag             =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtPdsc 
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   4
      Tag             =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtPdue 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Tag             =   "1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2460
      FormDesignWidth =   6765
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   288
      Index           =   2
      Left            =   5760
      TabIndex        =   17
      Top             =   1320
      Width           =   552
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Day Of The Month"
      Height          =   288
      Index           =   1
      Left            =   3480
      TabIndex        =   16
      Top             =   1680
      Width           =   1872
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Not Enabled On Printed Purchase Orders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   320
      Width           =   4872
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Terms (Either Terms or Prox Terms):"
      Height          =   285
      Index           =   23
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   2835
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Terms: Net "
      Height          =   285
      Index           =   24
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Days.  Quick Payment Disc Of"
      Height          =   285
      Index           =   25
      Left            =   1680
      TabIndex        =   12
      Top             =   960
      Width           =   2355
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "% For"
      Height          =   285
      Index           =   26
      Left            =   4560
      TabIndex        =   11
      Top             =   960
      Width           =   555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Days Or Less"
      Height          =   285
      Index           =   27
      Left            =   5640
      TabIndex        =   10
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prox Terms:  Invoices Dated On Or Before The "
      Height          =   285
      Index           =   28
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   3555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Receive a"
      Height          =   285
      Index           =   29
      Left            =   4200
      TabIndex        =   8
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount If Paid By The"
      Height          =   285
      Index           =   30
      Left            =   1080
      TabIndex        =   7
      Top             =   1680
      Width           =   1875
   End
End
Attribute VB_Name = "PurcPRe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/2/05 Corrected KeySet Cursor
Option Explicit
Dim RdoTrm As ADODB.Recordset
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub
Private Sub cmdCan_Click()
   ' PurcPRe02a.optTrm.Value = vbUnchecked
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then GetTerms
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   cmdCan_Click
   
End Sub


Private Sub Form_Load()
   Dim iList As Integer
   Move PurcPRe02a.Left + 500, PurcPRe02a.Top + 1800
   FormatControls
   If Len(Trim(PurcPRe02a.lblPrn)) > 0 Then
      For iList = 0 To Controls.Count - 1
         If TypeOf Controls(iList) Is TextBox Then
            Controls(iList).Enabled = False
         End If
      Next
   Else
      For iList = 0 To Controls.Count - 1
         If TypeOf Controls(iList) Is TextBox Then
            Controls(iList).Enabled = True
         End If
      Next
   End If
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Set RdoTrm = Nothing
   PurcPRe02a.RefreshPoCursor
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set PurcPRe01b = Nothing
   
End Sub




Private Sub txtPday_LostFocus()
   txtPday = CheckLen(txtPday, 3)
   If Len(txtPday) > 0 Then txtPday = Format(Abs(Val(txtPday)), "##0")
   On Error Resume Next
   RdoTrm!PODDAYS = Val(txtPday)
   RdoTrm.Update
   If Err > 0 Then ValidateEdit
   
End Sub

Private Sub txtPdsc_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtPdsc_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtPdsc_LostFocus(Index As Integer)
   txtPdsc(Index) = CheckLen(txtPdsc(Index), 4)
   If Len(txtPdsc(Index)) > 0 Then txtPdsc(Index) = Format(Abs(Val(txtPdsc(Index))), "#0.0")
   On Error Resume Next
   RdoTrm!PODISCOUNT = Val(txtPdsc(Index))
   RdoTrm.Update
   If Err > 0 Then ValidateEdit
   
End Sub

Private Sub txtPdte_LostFocus()
   txtPdte = CheckLen(txtPdte, 2)
   If Len(txtPdte) > 0 Then txtPdte = Format(Abs(Val(txtPdte)), "#0")
   If Val(txtPdte) > 30 Then txtPdte = "30"
   On Error Resume Next
   RdoTrm!POPROXDT = Val(txtPdte)
   RdoTrm.Update
   If Err > 0 Then ValidateEdit
   
End Sub

Private Sub txtPdue_LostFocus()
   txtPdue = CheckLen(txtPdue, 2)
   If Len(txtPdue) > 0 Then txtPdue = Format(Abs(Val(txtPdue)), "#0")
   If Val(txtPdue) > 30 Then txtPdue = "30"
   On Error Resume Next
   RdoTrm!POPROXDUE = Val(txtPdue)
   RdoTrm.Update
   If Err > 0 Then ValidateEdit
   
End Sub

Private Sub txtPnet_LostFocus()
   txtPnet = CheckLen(txtPnet, 3)
   If Len(txtPnet) > 0 Then txtPnet = Format(Abs(Val(txtPnet)), "##0")
   On Error Resume Next
   RdoTrm!PONETDAYS = Val(txtPnet)
   RdoTrm.Update
   If Err > 0 Then ValidateEdit
   
End Sub





Private Sub GetTerms()
   sSql = "SELECT PONETDAYS,PODISCOUNT,PODDAYS,POPROXDT," _
          & "PODISCOUNT,POPROXDUE FROM PohdTable " _
          & "WHERE PONUMBER=" & Val(PurcPRe02a.cmbPon) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTrm, ES_KEYSET)
   
End Sub
