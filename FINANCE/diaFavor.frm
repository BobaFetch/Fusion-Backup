VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "RESIZE32.OCX"
Begin VB.Form diaFavor
   BorderStyle = 3 'Fixed Dialog
   Caption = "Add Favorites"
   ClientHeight = 3615
   ClientLeft = 2850
   ClientTop = 1665
   ClientWidth = 5310
   Icon = "diaFavor.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 3615
   ScaleWidth = 5310
   ShowInTaskbar = 0 'False
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 3840
      Top = 120
      _Version = 131077
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3615
      FormDesignWidth = 5310
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Add"
      Height = 285
      Index = 9
      Left = 90
      TabIndex = 8
      Top = 3030
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Add"
      Height = 285
      Index = 8
      Left = 90
      TabIndex = 7
      Top = 2760
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Add"
      Height = 285
      Index = 7
      Left = 90
      TabIndex = 6
      Top = 2490
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Add"
      Height = 285
      Index = 6
      Left = 90
      TabIndex = 5
      Top = 2220
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Add"
      Height = 285
      Index = 5
      Left = 90
      TabIndex = 4
      Top = 1950
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Add"
      Height = 285
      Index = 4
      Left = 90
      TabIndex = 3
      Top = 1680
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Add"
      Height = 285
      Index = 3
      Left = 90
      TabIndex = 2
      Top = 1410
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Add"
      Height = 285
      Index = 2
      Left = 90
      TabIndex = 1
      Top = 1140
      Width = 825
   End
   Begin VB.CommandButton cmdAdd
      Caption = "Add"
      Height = 285
      Index = 1
      Left = 90
      TabIndex = 0
      Top = 870
      Width = 825
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 4320
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
      MousePointer = 1 'Arrow
      TabIndex = 18
      TabStop = 0 'False
      Top = 3030
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 8
      Left = 1080
      MousePointer = 1 'Arrow
      TabIndex = 17
      TabStop = 0 'False
      Top = 2760
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 7
      Left = 1080
      MousePointer = 1 'Arrow
      TabIndex = 16
      TabStop = 0 'False
      Top = 2490
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 6
      Left = 1080
      MousePointer = 1 'Arrow
      TabIndex = 15
      TabStop = 0 'False
      Top = 2220
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 5
      Left = 1080
      MousePointer = 1 'Arrow
      TabIndex = 14
      TabStop = 0 'False
      Top = 1950
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 4
      Left = 1080
      MousePointer = 1 'Arrow
      TabIndex = 13
      TabStop = 0 'False
      Top = 1680
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 3
      Left = 1080
      MousePointer = 1 'Arrow
      TabIndex = 12
      TabStop = 0 'False
      Top = 1410
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 2
      Left = 1080
      MousePointer = 1 'Arrow
      TabIndex = 11
      TabStop = 0 'False
      Top = 1140
      Width = 3345
   End
   Begin VB.TextBox txtFvr
      BackColor = &H00FFFFFF&
      ForeColor = &H00800000&
      Height = 285
      Index = 1
      Left = 1080
      MousePointer = 1 'Arrow
      TabIndex = 10
      TabStop = 0 'False
      Top = 870
      Width = 3345
   End
   Begin Threed.SSCommand cmdDn
      Height = 375
      Left = 4560
      TabIndex = 20
      TabStop = 0 'False
      Top = 2730
      Width = 375
      _Version = 65536
      _ExtentX = 661
      _ExtentY = 661
      _StockProps = 78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Enabled = 0 'False
      RoundedCorners = 0 'False
      Outline = 0 'False
      Picture = "diaFavor.frx":030A
   End
   Begin Threed.SSCommand cmdUp
      Height = 375
      Left = 4560
      TabIndex = 21
      TabStop = 0 'False
      Top = 2280
      Width = 375
      _Version = 65536
      _ExtentX = 661
      _ExtentY = 661
      _StockProps = 78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Enabled = 0 'False
      RoundedCorners = 0 'False
      Outline = 0 'False
      Picture = "diaFavor.frx":080C
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 22
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
      PictureUp = "diaFavor.frx":0D0E
      PictureDn = "diaFavor.frx":0E54
   End
   Begin VB.Image Dsup
      Height = 300
      Left = 5040
      Picture = "diaFavor.frx":0F9A
      Top = 1080
      Visible = 0 'False
      Width = 285
   End
   Begin VB.Image Enup
      Height = 300
      Left = 5040
      Picture = "diaFavor.frx":148C
      Top = 720
      Visible = 0 'False
      Width = 285
   End
   Begin VB.Image Endn
      Height = 300
      Left = 5040
      Picture = "diaFavor.frx":197E
      Top = 1440
      Visible = 0 'False
      Width = 285
   End
   Begin VB.Image Dsdn
      Height = 300
      Left = 5040
      Picture = "diaFavor.frx":1E70
      Top = 1800
      Visible = 0 'False
      Width = 285
   End
   Begin VB.Label z1
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "To Add A Favorite, Open Any Form and Select Options, Add..."
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
      Top = 240
      Width = 3705
   End
End
Attribute VB_Name = "diaFavor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iIndex As Integer

Public Sub SaveFavorites()
   Dim i%
   Dim DataValue
   Dim DataValues As New Collection
   
   Erase sFavorites
   On Error Resume Next
   For i% = 1 To 9
      If Len(Trim$(txtFvr(i%))) > 0 Then
         DataValues.Add txtFvr(i%), txtFvr(i%)
      End If
   Next
   On Error GoTo 0
   i% = 0
   For Each DataValue In DataValues
      If DataValue <> "" Then
         i% = i% + 1
         sFavorites(i%) = DataValue
      End If
   Next
   For i% = 1 To 9
      txtFvr(i%) = sFavorites(i%)
   Next
   
End Sub

Private Sub cmdAdd_Click(Index As Integer)
   If Len(sCurrForm) > 0 Then
      txtFvr(Index) = sCurrForm
   Else
      MsgBox "Open A Form Or Form Doesn't Accept Favorites.", 64, Caption
   End If
   
End Sub

Private Sub cmdAdd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub cmdCan_Click()
   SaveFavorites
   Unload Me
   
End Sub


Private Sub cmdDn_Click()
   Dim sText As String
   sText = txtFvr(iIndex + 1)
   txtFvr(iIndex + 1) = txtFvr(iIndex)
   txtFvr(iIndex) = sText
   txtFvr_Click (iIndex + 1)
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Favorites"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdUp_Click()
   Dim sText As String
   sText = txtFvr(iIndex - 1)
   txtFvr(iIndex - 1) = txtFvr(iIndex)
   txtFvr(iIndex) = sText
   txtFvr_Click (iIndex - 1)
   
End Sub

Private Sub Form_Load()
   Dim i%
   SetFormSize Me
   Move 1980, 600
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


Private Sub txtFvr_Click(Index As Integer)
   Dim i%
   iIndex = Index
   For i% = 1 To 9
      txtFvr(i%).BackColor = QBColor(15)
      txtFvr(i%).ForeColor = QBColor(1)
   Next
   
   txtFvr(Index).BackColor = QBColor(1)
   txtFvr(Index).ForeColor = QBColor(15)
   If Index > 1 Then
      cmdUp.Enabled = True
      cmdUp.Picture = Enup
   Else
      cmdUp.Enabled = False
      cmdUp.Picture = Dsup
   End If
   If Index < 9 Then
      cmdDn.Enabled = True
      cmdDn.Picture = Endn
   Else
      cmdDn.Enabled = False
      cmdDn.Picture = Dsdn
   End If
   
End Sub

Private Sub txtFvr_GotFocus(Index As Integer)
   On Error Resume Next
   iIndex = Index
   cmdAdd(Index).SetFocus
   
End Sub

Private Sub txtFvr_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = 0
   
End Sub
