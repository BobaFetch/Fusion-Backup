VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form SysCustomColors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Section Backgrounds"
   ClientHeight    =   4935
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4140
   Icon            =   "CustomColors.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider SlideRed 
      Height          =   252
      Left            =   1920
      TabIndex        =   23
      Top             =   3720
      Width           =   732
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      Max             =   255
   End
   Begin MSComDlg.CommonDialog Cdi2 
      Left            =   360
      Top             =   4920
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton cmdColors 
      Caption         =   "More"
      Height          =   315
      Left            =   3000
      TabIndex        =   22
      ToolTipText     =   "More Colors"
      Top             =   4140
      Width           =   1000
   End
   Begin VB.TextBox txtBlue 
      Enabled         =   0   'False
      Height          =   288
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Blue"
      Top             =   4440
      Width           =   492
   End
   Begin VB.TextBox txtGreen 
      Enabled         =   0   'False
      Height          =   288
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "Green"
      Top             =   4080
      Width           =   492
   End
   Begin VB.TextBox txtRed 
      Enabled         =   0   'False
      Height          =   288
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Red"
      Top             =   3720
      Width           =   492
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   315
      Left            =   3000
      TabIndex        =   17
      ToolTipText     =   "Apply Settings Or Leave"
      Top             =   3720
      Width           =   1000
   End
   Begin VB.Frame StandardColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Light Blue"
      ForeColor       =   &H80000005&
      Height          =   1212
      Left            =   2040
      TabIndex        =   10
      ToolTipText     =   "Current Application Workspace"
      Top             =   2400
      Width           =   2052
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current Windows Setting"
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   5
         Left            =   0
         TabIndex        =   16
         ToolTipText     =   "Current Application Workspace"
         Top             =   480
         Width           =   2052
      End
   End
   Begin VB.Frame lightGreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Light Blue"
      ForeColor       =   &H80000005&
      Height          =   1212
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "Go Ahead And Click Me"
      Top             =   2400
      Width           =   2052
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Light Green"
         ForeColor       =   &H00400000&
         Height          =   252
         Index           =   4
         Left            =   0
         TabIndex        =   15
         ToolTipText     =   "Go Ahead And Click Me"
         Top             =   480
         Width           =   2052
      End
   End
   Begin VB.Frame zWhite 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Light Blue"
      ForeColor       =   &H80000005&
      Height          =   1212
      Left            =   2040
      TabIndex        =   8
      ToolTipText     =   "Go Ahead And Click Me"
      Top             =   1200
      Width           =   2052
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "White"
         ForeColor       =   &H00400000&
         Height          =   252
         Index           =   3
         Left            =   0
         TabIndex        =   14
         ToolTipText     =   "Go Ahead And Click Me"
         Top             =   480
         Width           =   2052
      End
   End
   Begin VB.Frame LightGray 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Light Blue"
      ForeColor       =   &H80000005&
      Height          =   1212
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Go Ahead And Click Me"
      Top             =   1200
      Width           =   2052
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Light Gray"
         ForeColor       =   &H00400000&
         Height          =   252
         Index           =   2
         Left            =   0
         TabIndex        =   13
         ToolTipText     =   "Go Ahead And Click Me"
         Top             =   480
         Width           =   2052
      End
   End
   Begin VB.Frame LightYellow 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Light Blue"
      ForeColor       =   &H80000005&
      Height          =   1212
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "Go Ahead And Click Me"
      Top             =   0
      Width           =   2052
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Light Yellow"
         ForeColor       =   &H00400000&
         Height          =   252
         Index           =   1
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "Go Ahead And Click Me"
         Top             =   480
         Width           =   2052
      End
   End
   Begin VB.Frame LightBlue 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Light Blue"
      ForeColor       =   &H80000005&
      Height          =   1212
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Go Ahead And Click Me"
      Top             =   0
      Width           =   2052
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Light Blue"
         ForeColor       =   &H00400040&
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   11
         ToolTipText     =   "Go Ahead And Click Me"
         Top             =   480
         Width           =   2052
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   372
      Left            =   1560
      TabIndex        =   4
      Top             =   5280
      Width           =   1092
   End
   Begin MSComctlLib.Slider SlideGreen 
      Height          =   252
      Left            =   1920
      TabIndex        =   24
      Top             =   4080
      Width           =   732
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      Max             =   255
   End
   Begin MSComctlLib.Slider SlideBlue 
      Height          =   252
      Left            =   1920
      TabIndex        =   25
      Top             =   4440
      Width           =   732
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      Max             =   255
   End
   Begin VB.Label lblHex 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   3000
      TabIndex        =   21
      ToolTipText     =   "Hex Value"
      Top             =   4560
      Width           =   996
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   252
      Index           =   8
      Left            =   2640
      TabIndex        =   0
      Top             =   4440
      Width           =   252
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   252
      Index           =   7
      Left            =   2640
      TabIndex        =   20
      Top             =   4080
      Width           =   252
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   252
      Index           =   6
      Left            =   2640
      TabIndex        =   19
      Top             =   3720
      Width           =   252
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1092
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "Test Colors - Click Any Of The Boxes To See How It Will Look"
      Top             =   3720
      Width           =   1332
   End
End
Attribute VB_Name = "SysCustomColors"
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
'3/16/06 Added CommonDialog Colors
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
      CancelTrans
   End If
   Unload Me
   
End Sub

Private Sub cmdClose_Click()
   Unload Me
   
End Sub



Private Sub cmdColors_Click()
   On Error Resume Next
   Cdi2.flags = cdlCCRGBInit
   Cdi2.Color = lblColor.BackColor
   Cdi2.ShowColor
   GetRGBValues Cdi2.Color
   
End Sub

Private Sub Form_Activate()
   lblColor.ToolTipText = "This Setting May Be Changed From Any Section Too"
   
End Sub

Private Sub Form_Initialize()
   AlwaysOnTop Me.hwnd, True
   
End Sub


Private Sub Form_Load()
   MDISect.Timer5.Enabled = False
   iRed = GetSetting("Esi2000", "System", "SectionBackColorR", iRed)
   iGreen = GetSetting("Esi2000", "System", "SectionBackColorG", iGreen)
   iBlue = GetSetting("Esi2000", "System", "SectionBackColorB", iBlue)
   
   SaveSetting "Esi2000", "System", "SectionBackColorR", iRed
   SaveSetting "Esi2000", "System", "SectionBackColorG", iGreen
   SaveSetting "Esi2000", "System", "SectionBackColorB", iBlue
   If iRed + iGreen + iBlue > 0 Then
      lblColor.BackColor = RGB(iRed, iGreen, iBlue)
      txtRed = str(iRed)
      txtGreen = str(iGreen)
      txtBlue = str(iBlue)
      txtRed.Enabled = True
      txtGreen.Enabled = True
      txtBlue.Enabled = True
      SlideRed.Enabled = True
      SlideGreen.Enabled = True
      SlideBlue.Enabled = True
      SlideRed = iRed
      SlideGreen = iGreen
      SlideBlue = iBlue
   Else
      lblColor.BackColor = MDISect.BackColor
   End If
   LightBlue.BackColor = RGB(214, 225, 254)
   LightYellow.BackColor = RGB(250, 253, 223)
   LightGray.BackColor = RGB(238, 238, 238)
   zWhite.BackColor = RGB(255, 255, 255)
   lightGreen.BackColor = RGB(227, 255, 255)
   StandardColor.BackColor = vbApplicationWorkspace
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   AlwaysOnTop Me.hwnd, False
   iRed = GetSetting("Esi2000", "System", "SectionBackColorR", iRed)
   iGreen = GetSetting("Esi2000", "System", "SectionBackColorG", iGreen)
   iBlue = GetSetting("Esi2000", "System", "SectionBackColorB", iBlue)
   If iRed + iGreen + iBlue = 0 Then
      'MDISect.BackColor = vbApplicationWorkspace
      MDISect.BackColor = GetBackgroundColor
   Else
      MDISect.BackColor = RGB(iRed, iGreen, iBlue)
   End If
   MDISect.Timer5.Enabled = True
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set CustomColors = Nothing
   
End Sub



Private Sub lblColor_Change()
   iRed = Val(txtRed)
   iGreen = Val(txtRed)
   iBlue = Val(txtBlue)
   If (iRed + iGreen = iBlue) > 0 Then GetColor
   
End Sub

Private Sub lblColor_Click()
   If (iRed + iGreen + iBlue) > 0 _
       Then MDISect.BackColor = RGB(iRed, iGreen, iBlue)
      
   End Sub
   
   
   Private Sub lblHex_Click()
      'MdiSect.BackColor = "&H" & lblHex
      
   End Sub
   
   
   Private Sub LightBlue_Click()
      iRed = 214
      iGreen = 225
      iBlue = 254
      SlideRed = iRed
      SlideGreen = iGreen
      SlideBlue = iBlue
      GetColor
      MDISect.BackColor = RGB(214, 225, 254)
      
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
      MDISect.BackColor = RGB(238, 238, 238)
      
   End Sub
   
   Private Sub lightGreen_Click()
      iRed = 227
      iGreen = 255
      iBlue = 255
      SlideRed = iRed
      SlideGreen = iGreen
      SlideBlue = iBlue
      GetColor
      MDISect.BackColor = RGB(227, 255, 255)
      
   End Sub
   
   Private Sub LightYellow_Click()
      iRed = 250
      iGreen = 253
      iBlue = 223
      SlideRed = iRed
      SlideGreen = iGreen
      SlideBlue = iBlue
      GetColor
      MDISect.BackColor = RGB(250, 253, 223)
      
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
      MDISect.BackColor = vbApplicationWorkspace
      txtRed = ""
      txtGreen = ""
      txtBlue = ""
      txtRed.Enabled = False
      txtGreen.Enabled = False
      txtBlue.Enabled = False
      SlideRed.Enabled = False
      SlideGreen.Enabled = False
      SlideBlue.Enabled = False
      lblHex = ""
      lblColor.BackColor = MDISect.BackColor
      
   End Sub
   
   Private Sub txtBlue_GotFocus()
      SelectFormat Me
      
   End Sub
   
   Private Sub txtBlue_LostFocus()
      txtBlue = Format(Abs(txtBlue), "#0")
      If Val(txtBlue) > 255 Then txtBlue = 255
      On Error GoTo ERR1:
      SlideBlue = txtBlue
      GetColor
      Exit Sub
ERR1:
      
   End Sub
   
   
   Private Sub txtGreen_GotFocus()
      SelectFormat Me
      
   End Sub
   
   
   Private Sub txtGreen_LostFocus()
      txtGreen = Format(Abs(txtGreen), "#0")
      If Val(txtGreen) > 255 Then txtGreen = 255
      On Error GoTo ERR1
      SlideGreen = txtGreen
      GetColor
      Exit Sub
ERR1:
      
   End Sub
   
   
   Private Sub txtRed_GotFocus()
      SelectFormat Me
      
   End Sub
   
   
   Private Sub txtRed_LostFocus()
      txtRed = Format(Abs(txtRed), "#0")
      If Val(txtRed) > 255 Then txtRed = 255
      On Error GoTo ERR1
      SlideRed = txtRed
      GetColor
      Exit Sub
ERR1:
      
   End Sub
   
   
   Private Sub z1_Click(Index As Integer)
      Select Case Index
         Case 0
            iRed = 214
            iGreen = 225
            iBlue = 254
            MDISect.BackColor = RGB(214, 225, 254)
         Case 1
            iRed = 250
            iGreen = 253
            iBlue = 223
            MDISect.BackColor = RGB(250, 253, 223)
         Case 2
            iRed = 238
            iGreen = 238
            iBlue = 238
            MDISect.BackColor = RGB(238, 238, 238)
         Case 3
            iRed = 255
            iGreen = 255
            iBlue = 255
            MDISect.BackColor = RGB(255, 255, 255)
         Case 4
            iRed = 227
            iGreen = 255
            iBlue = 255
            MDISect.BackColor = RGB(227, 255, 255)
         Case 5
            iRed = 0
            iGreen = 0
            iBlue = 0
            MDISect.BackColor = vbApplicationWorkspace
            lblColor.BackColor = MDISect.BackColor
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
      MDISect.BackColor = RGB(255, 255, 255)
      
   End Sub
   
   Public Sub GetColor()
      Dim HR As Variant
      Dim HG As Variant
      Dim HB As Variant
      
      On Error Resume Next
      txtRed.Enabled = True
      txtGreen.Enabled = True
      txtBlue.Enabled = True
      SlideRed.Enabled = True
      SlideGreen.Enabled = True
      SlideBlue.Enabled = True
      txtRed = str(SlideRed)
      txtGreen = str(SlideGreen)
      txtBlue = str(SlideBlue)
      iRed = Val(txtRed)
      iGreen = Val(txtGreen)
      iBlue = Val(txtBlue)
      
      HR = hex(iRed)
      HG = hex(iGreen)
      HB = hex(iBlue)
      If iRed < 16 Then HR = "0" & HR
      If iGreen < 16 Then HG = "0" & HG
      If iBlue < 16 Then HB = "0" & HB
      lblHex = HR & HG & HB
      lblColor.BackColor = RGB(SlideRed, SlideGreen, SlideBlue)
      
   End Sub
   
   Private Sub GetRGBValues(oClr As Long)
      Dim lRGB As Long
      lRGB = oClr
      On Error GoTo DiaErr1
      If lRGB > 0 Then
         lblColor.BackColor = lRGB
         txtRed = lRGB And &HFF&
         txtGreen = (lRGB And &HFF00&) \ &H100
         txtBlue = (lRGB And &HFF0000) \ &H10000
         iRed = Val(txtRed)
         iGreen = Val(txtGreen)
         iBlue = Val(txtBlue)
      End If
      Exit Sub
      
DiaErr1:
      Err.Clear
   End Sub
