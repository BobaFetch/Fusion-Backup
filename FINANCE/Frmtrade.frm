VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.0#0"; "RESIZE32.OCX"
Begin VB.Form frmTrade
   AutoRedraw = -1 'True
   BackColor = &H00800000&
   BorderStyle = 0 'None
   Caption = "   "
   ClientHeight = 6495
   ClientLeft = 840
   ClientTop = 525
   ClientWidth = 7920
   ControlBox = 0 'False
   FillColor = &H00FFFFFF&
   BeginProperty Font
   Name = "MS Sans Serif"
   Size = 9.75
   Charset = 0
   Weight = 700
   Underline = 0 'False
   Italic = 0 'False
   Strikethrough = 0 'False
   EndProperty
   ForeColor = &H00800000&
   LinkTopic = "Form2"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 433
   ScaleMode = 3 'Pixel
   ScaleWidth = 528
   ShowInTaskbar = 0 'False
   StartUpPosition = 2 'CenterScreen
   Visible = 0 'False
   WindowState = 2 'Maximized
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 6960
      Top = 4200
      _Version = 196608
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 64
      Enabled = -1 'True
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      AutoCenterFormOnLoad = -1 'True
      FormDesignHeight = 6495
      FormDesignWidth = 7920
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00FFFFFF&
      Height = 735
      Left = 1320
      TabIndex = 4
      Top = 5640
      Width = 5535
   End
   Begin VB.Image Image2
      Height = 3900
      Left = 1320
      Picture = "Frmtrade.frx":0000
      Stretch = -1 'True
      Top = 840
      Width = 5475
   End
   Begin VB.Image Image1
      Appearance = 0 'Flat
      Height = 720
      Left = 1260
      Picture = "Frmtrade.frx":D634
      Top = 120
      Width = 5760
   End
   Begin VB.Label lblRel
      Appearance = 0 'Flat
      BackColor = &H00C00000&
      BackStyle = 0 'Transparent
      Caption = "Beta Build 1.29.59"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = -1 'True
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00FFFFFF&
      Height = 345
      Left = 5160
      TabIndex = 0
      Top = 4800
      Width = 2175
   End
   Begin VB.Label Label5
      Appearance = 0 'Flat
      BackColor = &H00C00000&
      BackStyle = 0 'Transparent
      Caption = "Bothell, Washington"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = -1 'True
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00FFFFFF&
      Height = 255
      Left = 1320
      TabIndex = 1
      Top = 5280
      Width = 3015
   End
   Begin VB.Label Label4
      Appearance = 0 'Flat
      BackColor = &H00C00000&
      BackStyle = 0 'Transparent
      Caption = "Software Engineering Group,Inc        "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = -1 'True
      Italic = -1 'True
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00FFFFFF&
      Height = 255
      Left = 1320
      TabIndex = 3
      Top = 5040
      Width = 4335
   End
   Begin VB.Label Label3
      Appearance = 0 'Flat
      BackColor = &H00C00000&
      BackStyle = 0 'Transparent
      Caption = "Enterprise Systems, Inc"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = -1 'True
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00FFFFFF&
      Height = 255
      Left = 1320
      TabIndex = 2
      Top = 4800
      Width = 4215
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim sBuild As String
   SetFormSize Me
   Move (Screen.Width - Width) / 2, ((Screen.Height - Height) / 2)
   sBuild = GetSetting("Esi2000", "System", "LastBuild", sBuild)
   'lblRel = App.Major & "." & App.Minor & "." & App.Revision & " "
   lblRel = "Beta Build " & sBuild
   '   lblRel = "Demo Release " & App.Major & "." & App.Minor & "." & App.Revision & " "
   lblDsc.Height = 1000
   Label4 = Label4 & Chr(174) & "1994-" & Format(Now, "yyyy")
   lblDsc = "These computer programs are protected by copyright law " _
            & "and international treaties. Unauthorized reproduction or " _
            & "distribution of these programs, or any portion of them, " _
            & "may result in severe criminal and civil penalties, and " _
            & "will result in prosecution to the fullest extent of " _
            & "the law."
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmTrade = Nothing
   
End Sub
