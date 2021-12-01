VERSION 5.00
Begin VB.Form ShopSHe02a
   Caption = "Dummy Revise MO"
   ClientHeight = 2328
   ClientLeft = 60
   ClientTop = 348
   ClientWidth = 3732
   LinkTopic = "Form1"
   LockControls = -1 'True
   ScaleHeight = 2328
   ScaleWidth = 3732
   StartUpPosition = 3 'Windows Default
   Begin VB.TextBox txtQty
      Height = 285
      Left = 840
      TabIndex = 5
      Text = "Text1"
      Top = 1320
      Width = 615
   End
   Begin VB.TextBox cmbRun
      Appearance = 0 'Flat
      Height = 285
      Left = 840
      TabIndex = 4
      Text = "cmbRun"
      Top = 720
      Width = 1815
   End
   Begin VB.CheckBox optPick
      Caption = "optPick"
      Height = 255
      Left = 840
      TabIndex = 3
      Top = 1680
      Width = 1215
   End
   Begin VB.Label lblStat
      Caption = "Dummy"
      Height = 255
      Left = 840
      TabIndex = 6
      Top = 2040
      Width = 1335
   End
   Begin VB.Label Label1
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      Caption = "Dummy for compatibility with Prod"
      ForeColor = &H80000008&
      Height = 255
      Left = 840
      TabIndex = 2
      Top = 120
      Width = 3495
   End
   Begin VB.Label lblFrom
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Left = 840
      TabIndex = 1
      Top = 960
      Width = 1815
   End
   Begin VB.Label cmbPrt
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Left = 840
      TabIndex = 0
      Top = 480
      Width = 1575
   End
End
Attribute VB_Name = "ShopSHe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
   Set ShopSHe02a = Nothing
   
End Sub
