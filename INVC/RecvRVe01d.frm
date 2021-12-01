VERSION 5.00
Begin VB.Form RecvRVe01d
   Caption = "Print Inventory Labels"
   ClientHeight = 2295
   ClientLeft = 60
   ClientTop = 450
   ClientWidth = 4680
   ControlBox = 0 'False
   LinkTopic = "Form1"
   ScaleHeight = 2295
   ScaleWidth = 4680
   StartUpPosition = 3 'Windows Default
   Begin VB.CommandButton cmdDontPrint
      Cancel = -1 'True
      Caption = "&Don't Print"
      Height = 435
      Left = 2520
      TabIndex = 10
      Top = 1680
      Width = 975
   End
   Begin VB.CommandButton cmdPrint
      Caption = "&Print"
      Height = 435
      Left = 1200
      TabIndex = 9
      Top = 1680
      Width = 975
   End
   Begin VB.TextBox txtQtyPerLabel
      Height = 315
      Left = 1920
      MaxLength = 8
      TabIndex = 8
      Text = "########"
      Top = 1200
      Width = 915
   End
   Begin VB.Label Label2
      Caption = "Quantity Per Label"
      Height = 255
      Left = 120
      TabIndex = 7
      Top = 1200
      Width = 1695
   End
   Begin VB.Label lblQty
      Caption = "lblQty"
      Height = 255
      Left = 1920
      TabIndex = 6
      Top = 840
      Width = 915
   End
   Begin VB.Label lblPartNo
      Caption = "lblPartNo"
      Height = 255
      Left = 1920
      TabIndex = 5
      Top = 480
      Width = 2415
   End
   Begin VB.Label lblRev
      Caption = "lblRev"
      Height = 255
      Left = 2460
      TabIndex = 4
      Top = 120
      Width = 495
   End
   Begin VB.Label lblItem
      Caption = "lblItem"
      Height = 255
      Left = 1920
      TabIndex = 3
      Top = 120
      Width = 435
   End
   Begin VB.Label Label4
      Caption = "Quantity Received"
      Height = 255
      Left = 120
      TabIndex = 2
      Top = 840
      Width = 1695
   End
   Begin VB.Label Label3
      Caption = "Part Number"
      Height = 255
      Left = 120
      TabIndex = 1
      Top = 480
      Width = 1695
   End
   Begin VB.Label Label1
      Caption = "Item"
      Height = 255
      Left = 120
      TabIndex = 0
      Top = 120
      Width = 1695
   End
End
Attribute VB_Name = "RecvRVe01d"
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
