VERSION 5.00
Begin VB.Form ParseHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parse Help"
   ClientHeight    =   3450
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   2640
      TabIndex        =   0
      Top             =   3480
      Width           =   875
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Press Esc or Click the other dialog to close Help."
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   11
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   5532
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Each entry requires an entry greater than zero."
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   10
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   5532
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "When you press Test, the Variable Labels and TextBoxes will be visible."
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   9
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   5532
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   8
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   5532
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "description."
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   5532
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "The formula will be named and a (20) char Index will be created with a"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   5532
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   5532
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Each formula will be designated VAR[somename]. See the example."
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   5532
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "If more are required, then more than one formula must be created."
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   5532
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "There can be no more than (4) formulae.  The text will remain open for editing."
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   5772
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "The formula must follow mathematical rules."
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   5532
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Parse Test:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5532
   End
End
Attribute VB_Name = "ParseHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub Form_Deactivate()
   cmdCan_Click
   
End Sub

Private Sub Form_Load()
   BackColor = ES_ViewBackColor
   
End Sub
