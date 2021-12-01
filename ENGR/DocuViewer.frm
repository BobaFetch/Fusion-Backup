VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form DocuViewer
   BorderStyle = 3 'Fixed Dialog
   Caption = "Change Viewers/Locations"
   ClientHeight = 2616
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 6348
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H8000000F&
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2616
   ScaleWidth = 6348
   ShowInTaskbar = 0 'False
   Begin VB.TextBox txtWeb
      Height = 285
      Left = 1080
      TabIndex = 3
      Tag = "2"
      ToolTipText = "Uses Default Brower If Not Set"
      Top = 2160
      Width = 4695
   End
   Begin VB.TextBox txtTxt
      Height = 285
      Left = 1080
      TabIndex = 2
      Tag = "2"
      ToolTipText = "Text Files (See Documents For Others)"
      Top = 1800
      Width = 4695
   End
   Begin VB.TextBox txtMMv
      Height = 285
      Left = 1080
      TabIndex = 1
      Tag = "2"
      ToolTipText = "Multimedia (Movies, Sound)"
      Top = 1320
      Width = 4695
   End
   Begin VB.TextBox txtPic
      Height = 285
      Left = 1080
      TabIndex = 0
      Tag = "2"
      ToolTipText = "Picture Files Except JPG (JPG Uses Windows Default)"
      Top = 840
      Width = 4695
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 5400
      TabIndex = 4
      TabStop = 0 'False
      Top = 90
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5880
      Top = 1080
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2616
      FormDesignWidth = 6348
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Web"
      Height = 255
      Index = 4
      Left = 240
      TabIndex = 9
      Top = 2160
      Width = 1095
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Set The Viewer And Location For Types Of Media Files"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 7.8
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00800000&
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 8
      Top = 240
      Width = 5055
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Pictures"
      Height = 255
      Index = 1
      Left = 240
      TabIndex = 7
      Top = 840
      Width = 1095
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Media"
      Height = 255
      Index = 2
      Left = 240
      TabIndex = 6
      Top = 1320
      Width = 1095
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Text"
      Height = 255
      Index = 3
      Left = 240
      TabIndex = 5
      Top = 1800
      Width = 1095
   End
End
Attribute VB_Name = "DocuViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnload As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbByr_Change()
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnload Then
      '   FillCombo
      bOnload = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnload = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Len(Trim(txtPic)) > 0 Then
      SaveSetting "Esi2000", "EsiEngr", "PicViewer", Trim(txtPic)
      DocuDCe05a.lblPic = txtPic
   End If
   If Len(Trim(txtMMv)) > 0 Then
      SaveSetting "Esi2000", "EsiEngr", "MMViewer", Trim(txtMMv)
      DocuDCe05a.lblMMV = txtMMv
   End If
   If Len(Trim(txtTxt)) > 0 Then
      SaveSetting "Esi2000", "EsiEngr", "TXTViewer", Trim(txtTxt)
      DocuDCe05a.lblTxt = txtTxt
   End If
   If Len(Trim(txtWeb)) > 0 Then
      SaveSetting "Esi2000", "EsiEngr", "WEBViewer", Trim(txtWeb)
      DocuDCe05a.lblWeb = txtWeb
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set DocuViewer = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub
