VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CustomColorsMom
   BorderStyle = 1 'Fixed Single
   Caption = "Custom Section Backgrounds"
   ClientHeight = 4920
   ClientLeft = 36
   ClientTop = 324
   ClientWidth = 4128
   Icon = "CustomColorsMom.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 4920
   ScaleWidth = 4128
   StartUpPosition = 2 'CenterScreen
   Begin VB.CommandButton cmdColors
      Caption = "More"
      Height = 315
      Left = 2880
      TabIndex = 24
      ToolTipText = "More Colors"
      Top = 4200
      Width = 1000
   End
   Begin VB.TextBox txtBlue
      Enabled = 0 'False
      Height = 288
      Left = 1440
      TabIndex = 5
      ToolTipText = "Blue"
      Top = 4440
      Width = 492
   End
   Begin VB.TextBox txtGreen
      Enabled = 0 'False
      Height = 288
      Left = 1440
      TabIndex = 4
      ToolTipText = "Green"
      Top = 4080
      Width = 492
   End
   Begin VB.TextBox txtRed
      Enabled = 0 'False
      Height = 288
      Left = 1440
      TabIndex = 3
      ToolTipText = "Red"
      Top = 3720
      Width = 492
   End
   Begin ComctlLib.Slider SlideRed
      Height = 252
      Left = 1920
      TabIndex = 0
      ToolTipText = "Red Slider"
      Top = 3720
      Width = 732
      _ExtentX = 1291
      _ExtentY = 445
      _Version = 327682
      Max = 255
   End
   Begin VB.CommandButton cmdApply
      Caption = "Apply"
      Height = 315
      Left = 2880
      TabIndex = 19
      ToolTipText = "Apply Settings Or Leave"
      Top = 3720
      Width = 1000
   End
   Begin VB.Frame StandardColor
      Appearance = 0 'Flat
      BackColor = &H00FFFFFF&
      BorderStyle = 0 'None
      Caption = "Light Blue"
      ForeColor = &H80000005&
      Height = 1212
      Left = 2040
      TabIndex = 12
      ToolTipText = "Current Application Workspace"
      Top = 2400
      Width = 2052
      Begin VB.Label z1
         Alignment = 2 'Center
         BackStyle = 0 'Transparent
         Caption = "Current Windows Setting"
         ForeColor = &H00000000&
         Height = 252
         Index = 5
         Left = 0
         TabIndex = 18
         ToolTipText = "Current Application Workspace"
         Top = 480
         Width = 2052
      End
   End
   Begin VB.Frame lightGreen
      Appearance = 0 'Flat
      BackColor = &H00C0FFC0&
      BorderStyle = 0 'None
      Caption = "Light Blue"
      ForeColor = &H80000005&
      Height = 1212
      Left = 0
      TabIndex = 11
      ToolTipText = "Go Ahead And Click Me"
      Top = 2400
      Width = 2052
      Begin VB.Label z1
         Alignment = 2 'Center
         BackStyle = 0 'Transparent
         Caption = "Light Green"
         ForeColor = &H00400000&
         Height = 252
         Index = 4
         Left = 0
         TabIndex = 17
         ToolTipText = "Go Ahead And Click Me"
         Top = 480
         Width = 2052
      End
   End
   Begin VB.Frame zWhite
      Appearance = 0 'Flat
      BackColor = &H00FFFFFF&
      BorderStyle = 0 'None
      Caption = "Light Blue"
      ForeColor = &H80000005&
      Height = 1212
      Left = 2040
      TabIndex = 10
      ToolTipText = "Go Ahead And Click Me"
      Top = 1200
      Width = 2052
      Begin VB.Label z1
         Alignment = 2 'Center
         BackStyle = 0 'Transparent
         Caption = "White"
         ForeColor = &H00400000&
         Height = 252
         Index = 3
         Left = 0
         TabIndex = 16
         ToolTipText = "Go Ahead And Click Me"
         Top = 480
         Width = 2052
      End
   End
   Begin VB.Frame LightGray
      Appearance = 0 'Flat
      BackColor = &H00E0E0E0&
      BorderStyle = 0 'None
      Caption = "Light Blue"
      ForeColor = &H80000005&
      Height = 1212
      Left = 0
      TabIndex = 9
      ToolTipText = "Go Ahead And Click Me"
      Top = 1200
      Width = 2052
      Begin VB.Label z1
         Alignment = 2 'Center
         BackStyle = 0 'Transparent
         Caption = "Light Gray"
         ForeColor = &H00400000&
         Height = 252
         Index = 2
         Left = 0
         TabIndex = 15
         ToolTipText = "Go Ahead And Click Me"
         Top = 480
         Width = 2052
      End
   End
   Begin VB.Frame LightYellow
      Appearance = 0 'Flat
      BackColor = &H00C0FFFF&
      BorderStyle = 0 'None
      Caption = "Light Blue"
      ForeColor = &H80000005&
      Height = 1212
      Left = 2040
      TabIndex = 8
      ToolTipText = "Go Ahead And Click Me"
      Top = 0
      Width = 2052
      Begin VB.Label z1
         Alignment = 2 'Center
         BackStyle = 0 'Transparent
         Caption = "Light Yellow"
         ForeColor = &H00400000&
         Height = 252
         Index = 1
         Left = 0
         TabIndex = 14
         ToolTipText = "Go Ahead And Click Me"
         Top = 480
         Width = 2052
      End
   End
   Begin VB.Frame LightBlue
      Appearance = 0 'Flat
      BackColor = &H00FFC0C0&
      BorderStyle = 0 'None
      Caption = "Light Blue"
      ForeColor = &H80000005&
      Height = 1212
      Left = 0
      TabIndex = 7
      ToolTipText = "Go Ahead And Click Me"
      Top = 0
      Width = 2052
      Begin VB.Label z1
         Alignment = 2 'Center
         BackStyle = 0 'Transparent
         Caption = "Light Blue"
         ForeColor = &H00400040&
         Height = 252
         Index = 0
         Left = 0
         TabIndex = 13
         ToolTipText = "Go Ahead And Click Me"
         Top = 480
         Width = 2052
      End
   End
   Begin VB.CommandButton cmdClose
      Cancel = -1 'True
      Caption = "Close"
      Height = 372
      Left = 1560
      TabIndex = 6
      Top = 5280
      Width = 1092
   End
   Begin ComctlLib.Slider SlideGreen
      Height = 252
      Left = 1920
      TabIndex = 1
      ToolTipText = "Green Slider"
      Top = 4080
      Width = 732
      _ExtentX = 1291
      _ExtentY = 445
      _Version = 327682
      Max = 255
   End
   Begin ComctlLib.Slider SlideBlue
      Height = 252
      Left = 1920
      TabIndex = 21
      ToolTipText = "Blue Slider"
      Top = 4440
      Width = 732
      _ExtentX = 1291
      _ExtentY = 445
      _Version = 327682
      Max = 255
   End
   Begin MSComDlg.CommonDialog Cdi2
      Left = 0
      Top = 0
      _ExtentX = 677
      _ExtentY = 677
      _Version = 393216
   End
   Begin VB.Label z1
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "B"
      Height = 252
      Index = 8
      Left = 2640
      TabIndex = 2
      Top = 4440
      Width = 252
   End
   Begin VB.Label z1
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "G"
      Height = 252
      Index = 7
      Left = 2640
      TabIndex = 23
      Top = 4080
      Width = 252
   End
   Begin VB.Label z1
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "R"
      Height = 252
      Index = 6
      Left = 2640
      TabIndex = 22
      Top = 3720
      Width = 252
   End
   Begin VB.Label lblColor
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 1092
      Left = 0
      TabIndex = 20
      ToolTipText = "Test Colors - Click Any Of The Boxes To See How It Will Look"
      Top = 3720
      Width = 1332
   End
End
Attribute VB_Name = "CustomColorsMom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 2/27/06
Option Explicit
Dim iRed As Integer
Dim iGreen As Integer
Dim iBlue As Integer

Private Sub cmdApply_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Save The Current Changes.", ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      SaveSetting "Esi2000", "System", "SectionBackColorR", iRed
      SaveSetting "Esi2000", "System", "SectionBackColorG", iGreen
      SaveSetting "Esi2000", "System", "SectionBackColorB", iBlue
   Else
      MsgBox "Transaction Canceled By The User.", _
         vbInformation, sSysCaption
      
   End If
   Unload Me
   
End Sub

Private Sub cmdClose_Click()
   Unload Me
   
End Sub



Private Sub cmdColors_Click()
   On Error Resume Next
   Cdi2.Flags = cdlCCRGBInit
   Cdi2.Color = lblColor.BackColor
   Cdi2.ShowColor
   GetRGBValues Cdi2.Color
   
End Sub

Private Sub Form_Load()
   iRed = GetSetting("Esi2000", "System", "SectionBackColorR", iRed)
   iGreen = GetSetting("Esi2000", "System", "SectionBackColorG", iGreen)
   iBlue = GetSetting("Esi2000", "System", "SectionBackColorB", iBlue)
   
   SaveSetting "Esi2000", "System", "SectionBackColorR", iRed
   SaveSetting "Esi2000", "System", "SectionBackColorG", iGreen
   SaveSetting "Esi2000", "System", "SectionBackColorB", iBlue
   If iRed + iGreen + iBlue > 0 Then
      lblColor.BackColor = RGB(iRed, iGreen, iBlue)
      txtRed = Str(iRed)
      txtGreen = Str(iGreen)
      txtBlue = Str(iBlue)
      txtRed.Enabled = True
      txtGreen.Enabled = True
      txtBlue.Enabled = True
      SlideRed.Enabled = True
      SlideGreen.Enabled = True
      SlideBlue.Enabled = True
      SlideRed = iRed
      SlideGreen = iGreen
      SlideBlue = iBlue
   End If
   LightBlue.BackColor = RGB(214, 225, 254)
   LightYellow.BackColor = RGB(250, 253, 223)
   LightGray.BackColor = RGB(238, 238, 238)
   zWhite.BackColor = RGB(255, 255, 255)
   lightGreen.BackColor = RGB(227, 255, 255)
   StandardColor.BackColor = vbApplicationWorkspace
   
End Sub


Private Sub GetRGBValues(oClr As Long)
   On Error Resume Next
   Dim lRGB As Long
   lRGB = oClr
   If lRGB > 0 Then
      lblColor.BackColor = lRGB
      txtRed = lRGB And &HFF&
      txtGreen = (lRGB And &HFF00&) \ &H100
      txtBlue = (lRGB And &HFF0000) \ &H10000
      iRed = Val(txtRed)
      iGreen = Val(txtGreen)
      iBlue = Val(txtBlue)
   End If
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set CustomColorsMom = Nothing
   
End Sub


Private Sub Label1_Click()
   
End Sub

Private Sub lblColor_Change()
   iRed = Val(txtRed)
   iGreen = Val(txtRed)
   iBlue = Val(txtBlue)
   If (iRed + iGreen = iBlue) > 0 Then GetColor
   
End Sub

Private Sub LightBlue_Click()
   iRed = 214
   iGreen = 225
   iBlue = 254
   SlideRed = iRed
   SlideGreen = iGreen
   SlideBlue = iBlue
   GetColor
   
End Sub

Private Sub LightBlue_DblClick()
   ' SaveSetting "Esi2000", "System", "SectionBackColor", "RGB(214,225,254)"
   ' MsgBox "The Color Will Be The Restarted Section BackGround.", _
   vbInformation, Caption
   
End Sub

Private Sub LightGray_Click()
   iRed = 238
   iGreen = 238
   iBlue = 238
   SlideRed = iRed
   SlideGreen = iGreen
   SlideBlue = iBlue
   GetColor
   
End Sub

Private Sub LightGray_DblClick()
   'SaveSetting "Esi2000", "System", "SectionBackColor", "RGB(238,238,238)"
   'MsgBox "The Color Will Be The Restarted Section BackGround.", _
   '    vbInformation, Caption
   
End Sub

Private Sub lightGreen_Click()
   iRed = 227
   iGreen = 255
   iBlue = 255
   SlideRed = iRed
   SlideGreen = iGreen
   SlideBlue = iBlue
   GetColor
   
End Sub

Private Sub lightGreen_DblClick()
   'SaveSetting "Esi2000", "System", "SectionBackColor", "RGB(227,255,255)"
   'MsgBox "The Color Will Be The Restarted Section BackGround.", _
   vbInformation, Caption
   
End Sub

Private Sub LightYellow_Click()
   iRed = 250
   iGreen = 253
   iBlue = 223
   SlideRed = iRed
   SlideGreen = iGreen
   SlideBlue = iBlue
   GetColor
   
End Sub

Private Sub SlideBlue_Click()
   txtBlue = SlideBlue
   GetColor
   
End Sub

Private Sub SlideGreen_Scroll()
   txtGreen = SlideGreen
   GetColor
   
End Sub


Private Sub SlideRed_Scroll()
   txtRed = SlideRed
   GetColor
   
End Sub


Private Sub StandardColor_Click()
   iRed = 0
   iGreen = 0
   iBlue = 0
   txtRed = ""
   txtGreen = ""
   txtBlue = ""
   txtRed.Enabled = False
   txtGreen.Enabled = False
   txtBlue.Enabled = False
   SlideRed.Enabled = False
   SlideGreen.Enabled = False
   SlideBlue.Enabled = False
   
End Sub

Private Sub txtBlue_GotFocus()
   SelectFormat Me
   
End Sub

Private Sub txtBlue_LostFocus()
   txtBlue = Format(Abs(txtBlue), "#0")
   If Val(txtBlue) > 255 Then txtBlue = 255
   On Error GoTo Err1:
   SlideBlue = txtBlue
   GetColor
   Exit Sub
   Err1:
   
End Sub


Private Sub txtGreen_GotFocus()
   SelectFormat Me
   
End Sub


Private Sub txtGreen_LostFocus()
   txtGreen = Format(Abs(txtGreen), "#0")
   If Val(txtGreen) > 255 Then txtGreen = 255
   On Error GoTo Err1
   SlideGreen = txtGreen
   GetColor
   Exit Sub
   Err1:
   
End Sub


Private Sub txtRed_GotFocus()
   SelectFormat Me
   
End Sub


Private Sub txtRed_LostFocus()
   txtRed = Format(Abs(txtRed), "#0")
   If Val(txtRed) > 255 Then txtRed = 255
   On Error GoTo Err1
   SlideRed = txtRed
   GetColor
   Exit Sub
   Err1:
   
End Sub


Private Sub z1_Click(Index As Integer)
   Select Case Index
      Case 0
         iRed = 214
         iGreen = 225
         iBlue = 254
      Case 1
         iRed = 250
         iGreen = 253
         iBlue = 223
      Case 2
         iRed = 238
         iGreen = 238
         iBlue = 238
      Case 3
         iRed = 255
         iGreen = 255
         iBlue = 255
      Case 4
         iRed = 227
         iGreen = 255
         iBlue = 255
      Case 5
         iRed = 0
         iGreen = 0
         iBlue = 0
   End Select
   If Index < 5 Then
      SlideRed = iRed
      SlideGreen = iGreen
      SlideBlue = iBlue
      GetColor
   Else
      txtRed = ""
      txtGreen = ""
      txtBlue = ""
      txtRed.Enabled = False
      txtGreen.Enabled = False
      txtBlue.Enabled = False
      SlideRed.Enabled = False
      SlideGreen.Enabled = False
      SlideBlue.Enabled = False
   End If
   
End Sub

Private Sub zWhite_Click()
   iRed = 255
   iGreen = 255
   iBlue = 255
   SlideRed = iRed
   SlideGreen = iGreen
   SlideBlue = iBlue
   GetColor
   
End Sub

Public Sub GetColor()
   On Error Resume Next
   txtRed.Enabled = True
   txtGreen.Enabled = True
   txtBlue.Enabled = True
   SlideRed.Enabled = True
   SlideGreen.Enabled = True
   SlideBlue.Enabled = True
   txtRed = Str(SlideRed)
   txtGreen = Str(SlideGreen)
   txtBlue = Str(SlideBlue)
   iRed = Val(txtRed)
   iGreen = Val(txtGreen)
   iBlue = Val(txtBlue)
   lblColor.BackColor = RGB(SlideRed, SlideGreen, SlideBlue)
   
End Sub
