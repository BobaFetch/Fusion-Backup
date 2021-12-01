VERSION 5.00
Begin VB.Form ShopSHe02a
   Caption = "Dummy diaSrvmo"
   ClientHeight = 2496
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 3744
   LinkTopic = "Form1"
   ScaleHeight = 2496
   ScaleWidth = 3744
   StartUpPosition = 3 'Windows Default
   Begin VB.CheckBox optFrom
      Caption = "optFrom"
      Height = 252
      Left = 240
      TabIndex = 6
      Top = 1680
      Width = 1452
   End
   Begin VB.Label cmbRun
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BorderStyle = 1 'Fixed Single
      Caption = "cmbRun"
      ForeColor = &H80000008&
      Height = 252
      Left = 240
      TabIndex = 5
      Top = 1320
      Width = 1092
   End
   Begin VB.Label cmbPrt
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BorderStyle = 1 'Fixed Single
      Caption = "cmbPrt"
      ForeColor = &H80000008&
      Height = 252
      Left = 240
      TabIndex = 4
      Top = 1080
      Width = 1092
   End
   Begin VB.Label lblSch
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BorderStyle = 1 'Fixed Single
      Caption = "lblSch"
      ForeColor = &H80000008&
      Height = 252
      Left = 240
      TabIndex = 3
      Top = 840
      Width = 1092
   End
   Begin VB.Label txtQty
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BorderStyle = 1 'Fixed Single
      Caption = "txtQty"
      ForeColor = &H80000008&
      Height = 252
      Left = 240
      TabIndex = 2
      Top = 600
      Width = 1092
   End
   Begin VB.Label txtPri
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BorderStyle = 1 'Fixed Single
      Caption = "txtPri"
      ForeColor = &H80000008&
      Height = 252
      Left = 240
      TabIndex = 1
      Top = 360
      Width = 1092
   End
   Begin VB.Label lblFrom
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BorderStyle = 1 'Fixed Single
      Caption = "lblFrom"
      ForeColor = &H80000008&
      Height = 252
      Left = 240
      TabIndex = 0
      Top = 120
      Width = 1092
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
Option Explicit
