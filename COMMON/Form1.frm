VERSION 5.00
Begin VB.Form Form1
   Caption = "Form1"
   ClientHeight = 3195
   ClientLeft = 60
   ClientTop = 345
   ClientWidth = 4680
   LinkTopic = "Form1"
   ScaleHeight = 3195
   ScaleWidth = 4680
   StartUpPosition = 3 'Windows Default
   Begin VB.Image img05
      Height = 600
      Left = 120
      Picture = "Form1.frx":0000
      Top = 840
      Visible = 0 'False
      Width = 600
   End
   Begin VB.Image XPHelpUp
      Height = 300
      Left = 0
      Picture = "Form1.frx":0BCA
      Top = 360
      Visible = 0 'False
      Width = 330
   End
   Begin VB.Image XPHelpDn
      Height = 300
      Left = 360
      Picture = "Form1.frx":115C
      Top = 360
      Visible = 0 'False
      Width = 330
   End
   Begin VB.Image XPPrinterUp
      Height = 285
      Left = 0
      Picture = "Form1.frx":16EE
      Top = 0
      Visible = 0 'False
      Width = 330
   End
   Begin VB.Image XPPrinterDn
      Height = 300
      Left = 360
      Picture = "Form1.frx":1C3C
      Top = 0
      Visible = 0 'False
      Width = 330
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
