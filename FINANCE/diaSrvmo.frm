VERSION 5.00
Begin VB.Form diaSrvmo
   BorderStyle = 3 'Fixed Dialog
   Caption = "This Form Is A Dummy"
   ClientHeight = 1125
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 3855
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 1125
   ScaleWidth = 3855
   ShowInTaskbar = 0 'False
   StartUpPosition = 3 'Windows Default
   Begin VB.ComboBox cmbRun
      Height = 315
      Left = 1920
      TabIndex = 1
      Text = "Combo1"
      Top = 240
      Width = 615
   End
   Begin VB.ComboBox cmbPrt
      Height = 315
      Left = 360
      TabIndex = 0
      Text = "Combo1"
      Top = 240
      Width = 1335
   End
   Begin VB.Label Label1
      Caption = "See diaPsh01.frm"
      Height = 255
      Left = 240
      TabIndex = 2
      Top = 720
      Width = 2415
   End
End
Attribute VB_Name = "diaSrvmo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
