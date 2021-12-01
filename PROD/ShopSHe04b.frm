VERSION 5.00
Begin VB.Form ShopSHe04b 
   Caption         =   "Print MO Labels"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDontPrint 
      Cancel          =   -1  'True
      Caption         =   "&Don't Print"
      Height          =   435
      Left            =   2640
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   435
      Left            =   1320
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtQtyPerLabel 
      Height          =   315
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   5
      Text            =   "########"
      Top             =   2040
      Width           =   915
   End
   Begin VB.Label lblRunNo 
      Caption         =   "lblRunNo"
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Run No"
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblLocation 
      Caption         =   "lblLocation"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Lot Location"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblUserLotNum 
      Caption         =   "lblUserLotNum"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "User Lot Number"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Lot Number"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblLotNum 
      Caption         =   "lblLotNum"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Quantity Per Label"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblQty 
      Caption         =   "lblQty"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label lblPartNo 
      Caption         =   "lblPartNo"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity Received"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
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
Attribute VB_Name = "ShopSHe04b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ParentForm As Form

Private Sub cmdDontPrint_Click()
   ParentForm.quantityPerLabel = 0
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   ParentForm.quantityPerLabel = CCur(Me.txtQtyPerLabel.Text)
   ParentForm.totalQuantity = CCur(Me.lblQty.Caption)
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Left = (Screen.Width - Me.Width) / 2
   Me.Top = (Screen.Height - Me.Height) / 2
End Sub

