VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PurcPRe01c 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Current Purchase Order items"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3720
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3315
      FormDesignWidth =   6735
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Press ""Esc"" To Close"
      Top             =   360
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   10
      Cols            =   7
      FixedCols       =   0
      ForeColor       =   8404992
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   435
      Left            =   3720
      TabIndex        =   0
      Top             =   3240
      Width           =   915
   End
   Begin VB.Image Chkyes 
      Height          =   180
      Left            =   240
      Picture         =   "PurcPRe01c.frx":0000
      Top             =   3000
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   840
      Picture         =   "PurcPRe01c.frx":0492
      Top             =   3000
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "PurcPRe01c"
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

Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub


Private Sub Form_Click()
   Form_Deactivate
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Initialize()
   BackColor = ES_ViewBackColor
   
End Sub


Private Sub Form_Load()
   SetFormSize Me
   If iBarOnTop Then
      Move MDISect.Left + 700, PurcPRe02b.Top + 1900
   Else
      Move MDISect.Left + 3500, PurcPRe02b.Top + 1100
   End If
   With Grd
      .row = 0
      .Col = 0
      .Text = "Item"
      .Col = 1
      .Text = "Part Number"
      .Col = 2
      .Text = "Quantity"
      .Col = 3
      .Text = "Price"
      .Col = 4
      .Text = "Sch Date"
      .Col = 6
      .Text = "Canceled"
   End With
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim b As Byte
   b = PurcPRe02b.GetThisItem(False)
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   PurcPRe02b.txtQty.SetFocus
   WindowState = 1
   Set PurcPRe01c = Nothing
   
End Sub
