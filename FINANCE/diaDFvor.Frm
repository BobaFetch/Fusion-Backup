VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "RESIZE32.OCX"
Begin VB.Form diaDFvor
   BorderStyle = 3 'Fixed Dialog
   Caption = "Delete Favorites"
   ClientHeight = 3675
   ClientLeft = 2850
   ClientTop = 1665
   ClientWidth = 4875
   Icon = "diaDFvor.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 3675
   ScaleWidth = 4875
   ShowInTaskbar = 0 'False
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 360
      Top = 0
      _Version = 131077
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3675
      FormDesignWidth = 4875
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Delete"
      Height = 285
      Index = 9
      Left = 90
      TabIndex = 8
      Top = 3150
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Delete"
      Height = 285
      Index = 8
      Left = 90
      TabIndex = 7
      Top = 2880
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Delete"
      Height = 285
      Index = 7
      Left = 90
      TabIndex = 6
      Top = 2610
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Delete"
      Height = 285
      Index = 6
      Left = 90
      TabIndex = 5
      Top = 2340
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Delete"
      Height = 285
      Index = 5
      Left = 90
      TabIndex = 4
      Top = 2070
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Delete"
      Height = 285
      Index = 4
      Left = 90
      TabIndex = 3
      Top = 1800
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Delete"
      Height = 285
      Index = 3
      Left = 90
      TabIndex = 2
      Top = 1530
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Delete"
      Height = 285
      Index = 2
      Left = 90
      TabIndex = 1
      Top = 1260
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Delete"
      Height = 285
      Index = 1
      Left = 90
      TabIndex = 0
      Top = 990
      Width = 825
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 3960
      TabIndex = 9
      TabStop = 0 'False
      Top = 90
      Width = 875
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 9
      Left = 1080
      TabIndex = 18
      TabStop = 0 'False
      Top = 3150
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 8
      Left = 1080
      TabIndex = 17
      TabStop = 0 'False
      Top = 2880
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 7
      Left = 1080
      TabIndex = 16
      TabStop = 0 'False
      Top = 2610
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 6
      Left = 1080
      TabIndex = 15
      TabStop = 0 'False
      Top = 2340
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 5
      Left = 1080
      TabIndex = 14
      TabStop = 0 'False
      Top = 2070
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 4
      Left = 1080
      TabIndex = 13
      TabStop = 0 'False
      Top = 1800
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 3
      Left = 1080
      TabIndex = 12
      TabStop = 0 'False
      Top = 1530
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 2
      Left = 1080
      TabIndex = 11
      TabStop = 0 'False
      Top = 1260
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 1
      Left = 1080
      TabIndex = 10
      TabStop = 0 'False
      Top = 990
      Width = 3345
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 20
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaDFvor.frx":030A
      PictureDn = "diaDFvor.frx":0450
   End
   Begin VB.Label z1
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "To Delete A Favorite, Press Delete Next To The Entry"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 9.75
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 555
      Left = 90
      TabIndex = 19
      Top = 360
      Width = 3705
   End
End
Attribute VB_Name = "diaDFvor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click(Index As Integer)
   txtFvr(Index) = ""
   
End Sub

Private Sub cmdAdd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Favorites"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Load()
   Dim i%
   SetFormSize Me
   Top = 600
   Left = 1980
   For i% = 1 To 9
      txtFvr(i) = sFavorites(i%)
   Next
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim i%
   For i% = 1 To 9
      sFavorites(i%) = "" & Trim(txtFvr(i))
   Next
   For i% = 1 To 9
      If sFavorites(i%) <> "" Then
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(i%))).Visible = True
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(i%))).Caption = sFavorites(i%)
      Else
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(i%))).Visible = False
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(i%))).Caption = ""
      End If
   Next
   SaveSetting "Esi2000", "EsiFina", "Favorite1", sFavorites(1)
   SaveSetting "Esi2000", "EsiFina", "Favorite2", sFavorites(2)
   SaveSetting "Esi2000", "EsiFina", "Favorite3", sFavorites(3)
   SaveSetting "Esi2000", "EsiFina", "Favorite4", sFavorites(4)
   SaveSetting "Esi2000", "EsiFina", "Favorite5", sFavorites(5)
   SaveSetting "Esi2000", "EsiFina", "Favorite6", sFavorites(6)
   SaveSetting "Esi2000", "EsiFina", "Favorite7", sFavorites(7)
   SaveSetting "Esi2000", "EsiFina", "Favorite8", sFavorites(8)
   SaveSetting "Esi2000", "EsiFina", "Favorite9", sFavorites(9)
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set diaFavor = Nothing
   
End Sub


Private Sub txtFvr_GotFocus(Index As Integer)
   On Error Resume Next
   cmdAdd(Index).SetFocus
   
End Sub

Private Sub txtFvr_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = 0
   
End Sub
