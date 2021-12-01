VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form diaNewJournal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New GL Joural Name"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "diaNewJournal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4200
   Begin VB.TextBox txtJrlName 
      Height          =   285
      Left            =   360
      MaxLength       =   12
      TabIndex        =   3
      Tag             =   "3"
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   405
      Left            =   1320
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Select new Jounal Name"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "C&ancel"
      Height          =   405
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4200
      Top             =   1800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1650
      FormDesignWidth =   4200
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select New Journal Name"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "diaNewJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOpen As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub cmdCan_Click()
   diaGLe02a.lblRevJrlName = ""
   Unload Me
   
End Sub

Private Sub cmdOk_Click()
   diaGLe02a.lblRevJrlName.Caption = txtJrlName.Text
   Unload Me
End Sub

Private Sub Form_Load()
   On Error Resume Next
   If MdiSect.Sidebar.Visible = False Then
      Move MdiSect.Left + MdiSect.ActiveForm.Left + 800, MdiSect.Top + 3200
   Else
      Move MdiSect.Left + MdiSect.ActiveForm.Left + 2600, MdiSect.Top + 3600
   End If
End Sub



'1/26/04


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set diaNewJournal = Nothing
   
End Sub



