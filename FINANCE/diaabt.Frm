VERSION 5.00
Begin VB.Form diaAbt
   BorderStyle = 3 'Fixed Dialog
   Caption = "About Enterprise Systems"
   ClientHeight = 3045
   ClientLeft = 6000
   ClientTop = 1005
   ClientWidth = 3375
   ClipControls = 0 'False
   ControlBox = 0 'False
   BeginProperty Font
   Name = "MS Sans Serif"
   Size = 8.25
   Charset = 0
   Weight = 700
   Underline = 0 'False
   Italic = 0 'False
   Strikethrough = 0 'False
   EndProperty
   ForeColor = &H00C0C0C0&
   LinkTopic = "Form2"
   LockControls = -1 'True
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 3045
   ScaleWidth = 3375
   ShowInTaskbar = 0 'False
   Begin VB.TextBox txtDmy
      Appearance = 0 'Flat
      BackColor = &H8000000F&
      BorderStyle = 0 'None
      Height = 195
      Left = 240
      TabIndex = 0
      Top = 3000
      Width = 255
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 375
      Left = 1080
      TabIndex = 4
      TabStop = 0 'False
      Top = 2520
      Width = 1215
   End
   Begin VB.Image Image1
      BorderStyle = 1 'Fixed Single
      Height = 1575
      Index = 2
      Left = 0
      Picture = "diaabt.frx":0000
      Top = 4080
      Width = 3285
   End
   Begin VB.Image Image1
      BorderStyle = 1 'Fixed Single
      Height = 1575
      Index = 1
      Left = 0
      Picture = "diaabt.frx":FFEA
      Top = 3360
      Width = 3285
   End
   Begin VB.Label lblRel
      Alignment = 1 'Right Justify
      ForeColor = &H00800000&
      Height = 255
      Left = 1440
      TabIndex = 8
      Top = 2160
      Width = 1845
   End
   Begin VB.Label lblApp
      Alignment = 1 'Right Justify
      ForeColor = &H00800000&
      Height = 255
      Left = 1440
      TabIndex = 7
      Top = 1920
      Width = 1845
   End
   Begin VB.Label z1
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      Caption = "Build"
      ForeColor = &H00800000&
      Height = 255
      Index = 3
      Left = 120
      TabIndex = 6
      Top = 2160
      Width = 1215
   End
   Begin VB.Label z1
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      Caption = "Application"
      ForeColor = &H00800000&
      Height = 255
      Index = 2
      Left = 120
      TabIndex = 5
      Top = 1920
      Width = 1215
   End
   Begin VB.Image Image1
      BorderStyle = 1 'Fixed Single
      Height = 1575
      Index = 0
      Left = 40
      Picture = "diaabt.frx":1FFD4
      Top = 45
      Width = 3285
   End
   Begin VB.Label z1
      Alignment = 1 'Right Justify
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      Caption = "kb"
      ForeColor = &H00800000&
      Height = 255
      Index = 0
      Left = 2880
      TabIndex = 3
      Top = 1710
      Width = 400
   End
   Begin VB.Label lblMem
      Alignment = 1 'Right Justify
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      Caption = " "
      ForeColor = &H00800000&
      Height = 255
      Left = 1920
      TabIndex = 2
      Top = 1710
      Width = 1035
   End
   Begin VB.Label z1
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      Caption = "Physical Memory"
      ForeColor = &H00800000&
      Height = 255
      Index = 1
      Left = 120
      TabIndex = 1
      Top = 1710
      Width = 1815
   End
   Begin VB.Line Line3
      BorderColor = &H00C0C0C0&
      X1 = 0
      X2 = 3255
      Y1 = 0
      Y2 = 0
   End
End
Attribute VB_Name = "diaAbt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2004) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/30/04 Corrected lblApp (do not attempt to show the Database).
'        See Form_Load.

Option Explicit




Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub Form_Load()
   Dim bByte As Byte
   Dim FreeSpc As Long
   Dim sBuild As String
   Dim sYear As String
   Dim sAppCap As String
   Dim MyMemory As MEMORYSTATUS
   Top = 0
   If iBarOnTop Then
      Left = MdiSect.Width - (Width + 220)
   Else
      Left = MdiSect.Width - (Width + MdiSect.SideBar.Width + 220)
   End If
   GlobalMemoryStatus MyMemory
   FreeSpc = (MyMemory.dwTotalPhys) / 1000
   lblMem = Format(FreeSpc&, "###,##0")
   sAppCap = MdiSect.Caption
   bByte = InStr(sAppCap, "-")
   sAppCap = Left$(sAppCap, bByte - 2)
   lblApp = sAppCap
   sBuild = GetSetting("Esi2000", "System", "LastBuild", sBuild)
   'lblRel = App.Major & "." & App.Minor & "." & App.Revision & " "
   lblRel = sBuild
   
   sYear = Format$(Now, "yyyy")
   Select Case sYear
      Case "2003"
         Image1(0).Picture = Image1(1).Picture
      Case "2004"
         Image1(0).Picture = Image1(2).Picture
      Case Else
         Image1(0).Picture = Image1(0).Picture
   End Select
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set diaAbt = Nothing
   
End Sub
