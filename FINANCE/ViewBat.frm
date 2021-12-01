VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ViewBat
   BackColor = &H8000000C&
   BorderStyle = 3 'Fixed Dialog
   Caption = "Selected Checks"
   ClientHeight = 2760
   ClientLeft = 1620
   ClientTop = 3645
   ClientWidth = 5010
   Icon = "ViewBat.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2760
   ScaleWidth = 5010
   ShowInTaskbar = 0 'False
   Begin MSFlexGridLib.MSFlexGrid Grid1
      Height = 2175
      Left = 120
      TabIndex = 0
      Top = 0
      Width = 4815
      _ExtentX = 8493
      _ExtentY = 3836
      _Version = 393216
      Cols = 4
      FixedCols = 0
      HighLight = 0
   End
   Begin VB.Label lblTotal
      Alignment = 1 'Right Justify
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 255
      Left = 3550
      TabIndex = 1
      Top = 2280
      Width = 1335
   End
End
Attribute VB_Name = "ViewBat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Move diaChbat.Left + 400, diaChbat.Top + 1400
   Grid1.ColWidth(0) = 1300
   Grid1.ColWidth(1) = 1200
   Grid1.ColWidth(2) = 1150
   Grid1.ColWidth(3) = 1050
   Grid1.Row = 0
   Grid1.Col = 0
   Grid1 = "Check Number"
   Grid1.Col = 1
   Grid1 = "Vendor"
   Grid1.Col = 2
   Grid1 = "Date"
   Grid1.Col = 3
   Grid1 = "Amount"
   
End Sub
