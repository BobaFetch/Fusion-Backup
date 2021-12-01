VERSION 5.00
Begin VB.Form Form1
   BorderStyle = 3 'Fixed Dialog
   Caption = "Form1"
   ClientHeight = 6096
   ClientLeft = 36
   ClientTop = 324
   ClientWidth = 8268
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 6096
   ScaleWidth = 8268
   ShowInTaskbar = 0 'False
   Begin VB.Label WCDescription
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 6
      Left = 120
      TabIndex = 27
      Top = 5400
      Visible = 0 'False
      Width = 3732
   End
   Begin VB.Label WorkCenter
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 6
      Left = 120
      TabIndex = 26
      Top = 5040
      Visible = 0 'False
      Width = 1572
   End
   Begin VB.Label UsedHours
      Appearance = 0 'Flat
      BackColor = &H000000FF&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 6
      Left = 1884
      TabIndex = 25
      Top = 5040
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label FreeHours
      Appearance = 0 'Flat
      BackColor = &H0000C000&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 6
      Left = 4920
      TabIndex = 24
      Top = 5040
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label WCDescription
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 5
      Left = 120
      TabIndex = 23
      Top = 4680
      Visible = 0 'False
      Width = 3732
   End
   Begin VB.Label WorkCenter
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 5
      Left = 120
      TabIndex = 22
      Top = 4320
      Visible = 0 'False
      Width = 1572
   End
   Begin VB.Label UsedHours
      Appearance = 0 'Flat
      BackColor = &H000000FF&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 5
      Left = 1884
      TabIndex = 21
      Top = 4320
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label FreeHours
      Appearance = 0 'Flat
      BackColor = &H0000C000&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 5
      Left = 4920
      TabIndex = 20
      Top = 4320
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label WCDescription
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 4
      Left = 120
      TabIndex = 19
      Top = 3960
      Visible = 0 'False
      Width = 3732
   End
   Begin VB.Label WorkCenter
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 4
      Left = 120
      TabIndex = 18
      Top = 3600
      Visible = 0 'False
      Width = 1572
   End
   Begin VB.Label UsedHours
      Appearance = 0 'Flat
      BackColor = &H000000FF&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 4
      Left = 1884
      TabIndex = 17
      Top = 3600
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label FreeHours
      Appearance = 0 'Flat
      BackColor = &H0000C000&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 4
      Left = 4920
      TabIndex = 16
      Top = 3600
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label WCDescription
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 3
      Left = 120
      TabIndex = 15
      Top = 3240
      Visible = 0 'False
      Width = 3732
   End
   Begin VB.Label WorkCenter
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 3
      Left = 120
      TabIndex = 14
      Top = 2880
      Visible = 0 'False
      Width = 1572
   End
   Begin VB.Label UsedHours
      Appearance = 0 'Flat
      BackColor = &H000000FF&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 3
      Left = 1884
      TabIndex = 13
      Top = 2880
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label FreeHours
      Appearance = 0 'Flat
      BackColor = &H0000C000&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 3
      Left = 4920
      TabIndex = 12
      Top = 2880
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label WCDescription
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 2
      Left = 120
      TabIndex = 11
      Top = 2520
      Visible = 0 'False
      Width = 3732
   End
   Begin VB.Label WorkCenter
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 2
      Left = 120
      TabIndex = 10
      Top = 2160
      Visible = 0 'False
      Width = 1572
   End
   Begin VB.Label UsedHours
      Appearance = 0 'Flat
      BackColor = &H000000FF&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 2
      Left = 1884
      TabIndex = 9
      Top = 2160
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label FreeHours
      Appearance = 0 'Flat
      BackColor = &H0000C000&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 2
      Left = 4920
      TabIndex = 8
      Top = 2160
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label WCDescription
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 1
      Left = 120
      TabIndex = 7
      Top = 1800
      Visible = 0 'False
      Width = 3732
   End
   Begin VB.Label WorkCenter
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 1
      Left = 120
      TabIndex = 6
      Top = 1440
      Visible = 0 'False
      Width = 1572
   End
   Begin VB.Label UsedHours
      Appearance = 0 'Flat
      BackColor = &H000000FF&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 1
      Left = 1884
      TabIndex = 5
      Top = 1440
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label FreeHours
      Appearance = 0 'Flat
      BackColor = &H0000C000&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 1
      Left = 4920
      TabIndex = 4
      Top = 1440
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label WCDescription
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 0
      Left = 120
      TabIndex = 3
      Top = 1080
      Visible = 0 'False
      Width = 3732
   End
   Begin VB.Label FreeHours
      Appearance = 0 'Flat
      BackColor = &H0000C000&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 0
      Left = 4920
      TabIndex = 2
      Top = 720
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label UsedHours
      Appearance = 0 'Flat
      BackColor = &H000000FF&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00FFFFFF&
      Height = 252
      Index = 0
      Left = 1920
      TabIndex = 1
      Top = 720
      Visible = 0 'False
      Width = 3000
   End
   Begin VB.Label WorkCenter
      Appearance = 0 'Flat
      BackColor = &H80000005&
      ForeColor = &H80000008&
      Height = 252
      Index = 0
      Left = 120
      TabIndex = 0
      Top = 720
      Visible = 0 'False
      Width = 1572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   MouseCursor 0
   Me.BackColor = vbWhite
   UsedHours.Width = 2200
   FreeHours.Width = 3800
   FreeHours.Left = UsedHours.Left + UsedHours.Width
   
End Sub
