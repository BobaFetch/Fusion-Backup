VERSION 5.00
Begin VB.Form LotsLTf04b 
   Caption         =   "Print Invenory returned"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDontPrint 
      Cancel          =   -1  'True
      Caption         =   "&Don't Print"
      Height          =   435
      Left            =   2760
      TabIndex        =   5
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   435
      Left            =   1560
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblRMA 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1560
      TabIndex        =   23
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblVendorName 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1560
      TabIndex        =   21
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RMA Number"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   285
      Index           =   4
      Left            =   5400
      TabIndex        =   19
      Top             =   2520
      Width           =   465
   End
   Begin VB.Label lblPOItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   18
      ToolTipText     =   "Lot Location"
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev"
      Height          =   285
      Index           =   5
      Left            =   6840
      TabIndex        =   17
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label lblPORev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7245
      TabIndex        =   16
      ToolTipText     =   "Lot Location"
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblPORel 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4485
      TabIndex        =   15
      ToolTipText     =   "Lot Location"
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Release"
      Height          =   285
      Index           =   6
      Left            =   3840
      TabIndex        =   14
      Top             =   2520
      Width           =   585
   End
   Begin VB.Label Label7 
      Caption         =   "PO Number"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblPONum 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   2520
      Width           =   1875
   End
   Begin VB.Label lblLocation 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Lot Location"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblUserLotNum 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "User Lot Number"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Lot Number"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblLotNum 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label lblQty 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   915
   End
   Begin VB.Label lblPartNo 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity Received"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Part Number"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "LotsLTf04b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ParentForm As Form

Private Sub cmdDontPrint_Click()
   ParentForm.bPrint = 0
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   ParentForm.bPrint = 1
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Left = (Screen.Width - Me.Width) / 2
   Me.Top = (Screen.Height - Me.Height) / 2
End Sub

