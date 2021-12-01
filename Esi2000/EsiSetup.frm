VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EsiSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set User Preferences And Locations"
   ClientHeight    =   6075
   ClientLeft      =   2160
   ClientTop       =   2760
   ClientWidth     =   7350
   ClipControls    =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   925
   Icon            =   "EsiSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EsiSetup.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.CommandButton cmdColors 
      Height          =   192
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "Custom Colors"
      Top             =   2520
      Width           =   372
   End
   Begin VB.CheckBox optSoftBack 
      Alignment       =   1  'Right Justify
      Caption         =   "Use Soft Section Background (Soft Yellow )"
      Height          =   255
      Left            =   -120
      TabIndex        =   9
      ToolTipText     =   "Change The Background Section Color To Soft Yellow)"
      Top             =   5280
      Visible         =   0   'False
      Width           =   4100
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Index           =   1
      Left            =   3480
      TabIndex        =   26
      ToolTipText     =   "Effective Next Start - Double Click Manager Bar To Change"
      Top             =   4800
      Width           =   3135
      Begin VB.OptionButton optChm 
         Caption         =   ".chm (Recommended)"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Windows Me/2000/XP"
         Top             =   160
         Value           =   -1  'True
         Width           =   1900
      End
      Begin VB.OptionButton optHlp 
         Caption         =   ".hlp"
         Height          =   255
         Left            =   2200
         TabIndex        =   27
         ToolTipText     =   "Window 95/98 (16 Bit)"
         Top             =   160
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   360
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtWeb 
      Height          =   285
      Left            =   3360
      TabIndex        =   20
      ToolTipText     =   "Location Of Web Help (ESI Personel Setup Only). Enter Local Path If Help Is Server Installed"
      Top             =   5520
      Width           =   3615
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Remove Listed Server"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CheckBox optCaps 
      Alignment       =   1  'Right Justify
      Caption         =   "Turn Off Proper Casing Of Descriptions, Names, etc"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      ToolTipText     =   "Turns Off The Auto Capitalization Feature (Proper Casing)"
      Top             =   2160
      Width           =   4100
   End
   Begin VB.CheckBox optFrm 
      Alignment       =   1  'Right Justify
      Caption         =   "Open Last Form"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Re-opens The Function Last Used When Entering A Section"
      Top             =   1800
      Width           =   4100
   End
   Begin VB.CheckBox optSize 
      Alignment       =   1  'Right Justify
      Caption         =   "Turn Form Resizing On"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      ToolTipText     =   "When Checked, Forms Are Sized By Monitor Settings"
      Top             =   1440
      Width           =   4100
   End
   Begin VB.CheckBox optTab 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter Key Emulates Tab    "
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "Sets Enter Key To Act Like Tab (except multi-line Text Boxes)"
      Top             =   1080
      Width           =   4100
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5040
      Top             =   120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6075
      FormDesignWidth =   7350
   End
   Begin VB.CheckBox optFrom 
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Manager Bar Selection"
      Height          =   615
      Index           =   0
      Left            =   2400
      TabIndex        =   10
      ToolTipText     =   "Effective Next Start - Double Click Manager Bar To Change"
      Top             =   3000
      Width           =   4095
      Begin VB.CheckBox optSve 
         Caption         =   "Save Current"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         ToolTipText     =   "Save Current Bar When Closing"
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.OptionButton optVert 
         Caption         =   "Vertical"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         ToolTipText     =   "Open With A Vertical Bar"
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optHorz 
         Caption         =   "Horizontal"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Open With Horizontal Bar"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ComboBox cmbSrv 
      Height          =   288
      Left            =   3480
      TabIndex        =   19
      ToolTipText     =   "Select Server From List (2 Allowed)"
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CheckBox optMin 
      Alignment       =   1  'Right Justify
      Caption         =   "Minimize ES/2002 ERP Manager On Selection? "
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "Moves The Manager To The Task Bar After Program Selection"
      Top             =   720
      Width           =   4100
   End
   Begin VB.FileListBox File1 
      Enabled         =   0   'False
      Height          =   870
      Left            =   360
      Pattern         =   "esi*.exe;*.rpt"
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdRept 
      Caption         =   "&Reports"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Where Reports and Help Can Be Found"
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdPrg 
      Caption         =   "&Programs"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Where Programs And Help Can Be Found"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtRept 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   18
      ToolTipText     =   "Where Reports and Help Can Be Found"
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox txtPath 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   15
      ToolTipText     =   "Where Programs And Help Can Be Found"
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6360
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Section Background Colors"
      Height          =   252
      Index           =   2
      Left            =   2400
      TabIndex        =   29
      ToolTipText     =   "Custom Colors"
      Top             =   2520
      Width           =   3852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Help Style"
      Height          =   252
      Index           =   1
      Left            =   2400
      TabIndex        =   25
      ToolTipText     =   "Select Server From List (2 Allowed)"
      Top             =   4920
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Web Help"
      Height          =   252
      Left            =   2280
      TabIndex        =   24
      ToolTipText     =   "Location Of Web Help (ESI Personel Setup Only). Enter Local Path If Help Is Server Installed"
      Top             =   5520
      Width           =   972
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Id"
      Height          =   252
      Index           =   0
      Left            =   2400
      TabIndex        =   22
      ToolTipText     =   "Select Server From List (2 Allowed)"
      Top             =   4440
      Width           =   972
   End
End
Attribute VB_Name = "EsiSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of          ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'7/29/05 Help option added
'11/23/06 Added DeleteSettings to cmbSrv_LostFocus - Clear for Server Change
Option Explicit
Dim bEnterAsTab As Byte
Dim bOpenForm As Byte
Dim sOldServer As String
Dim sServers(3) As String

Private Sub cmbSrv_Click()
   bSQLOpen = 0
   
End Sub

Private Sub cmbSrv_GotFocus()
   SelectFormat Me
   
End Sub


Private Sub cmbSrv_KeyPress(KeyAscii As Integer)
   KeyCase KeyAscii, Me
   
End Sub


Private Sub cmbSrv_LostFocus()
   Dim b As Byte
   Dim a As Byte
   Dim Server0 As String
   Dim Server1 As String
   cmbSrv.AddItem cmbSrv
   'SaveSetting "Esi2000", "System", "ServerId", Trim(cmbSrv)
   SaveUserSetting USERSETTING_ServerName, Trim(cmbSrv)
   If cmbSrv.ListCount = 0 Then cmbSrv.AddItem cmbSrv
   sserver = Trim(cmbSrv)
   If sserver <> sOldServer Then
      '11/23/06 Deletes for new Server in UpdateTables
      On Error Resume Next 'may not be there
      'DeleteSetting "ES2000", "system", "eventlog"
      DeleteSetting "ES2000", "system", "channels"
      'DeleteSetting "ES2000", "system", "customreports"
      DeleteSetting "ES2000", "system", "systemmsgs"
      On Error GoTo 0
      bSQLOpen = 0
      sOldServer = sserver
      sDsn = RegisterSqlDsn(sDsn)
   End If
   SaveSetting "Esi2000", "System", "ServerId0", Trim(cmbSrv.List(0))
   If cmbSrv.ListCount > 1 Then
      SaveSetting "Esi2000", "System", "ServerId0", Trim(cmbSrv.List(0))
      SaveSetting "Esi2000", "System", "ServerId1", Trim(cmbSrv.List(1))
   Else
      SaveSetting "Esi2000", "System", "ServerId1", ""
   End If
   
End Sub


Private Sub cmdCan_Click()
   SaveSetting "Esi2000", "System", "FilePath", Trim(txtPath)
   
   'if running in VB, allow different report path
   If RunningInIDE Then
      SaveSetting "Esi2000", "System", "ReportPath", Trim(txtRept)
   End If
   Unload Me
   
End Sub

Private Sub cmdColors_Click()
   CustomColorsMom.Show
   
End Sub

Private Sub cmdDel_Click()
   Dim b As Byte
   b = MsgBox("Remove The Listed Server?", _
       ES_YESQUESTION, Caption)
   If b = vbYes Then
      For b = 0 To cmbSrv.ListCount - 1
         If cmbSrv = cmbSrv.List(b) Then
            cmbSrv.List(b) = ""
            cmbSrv = ""
         End If
      Next
      If cmbSrv.ListCount > 1 Then
         SaveSetting "Esi2000", "System", "ServerId0", Trim(cmbSrv.List(0))
         SaveSetting "Esi2000", "System", "ServerId1", Trim(cmbSrv.List(1))
      Else
         SaveSetting "Esi2000", "System", "ServerId1", ""
      End If
      'SaveSetting "Esi2000", "System", "ServerId", Trim(cmbSrv)
      SaveUserSetting USERSETTING_ServerName, Trim(cmbSrv)
      MsgBox "The Server Has Been Removed And You Should Set Another.", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub cmdHlp_Click()
   '    If cmdHlp Then
   '        MouseCursor 13
   '        OpenHelpContext 925
   '        cmdHlp = False
   '        MouseCursor 0
   '    End If
   
End Sub

Private Sub cmdPrg_Click()
   On Error Resume Next
   txtPath = Dir1 & "\"
   
End Sub

Private Sub cmdRept_Click()
   On Error Resume Next
   txtRept = Dir1 & "\"
   
End Sub


Private Sub Dir1_Change()
   On Error Resume Next
   File1 = Dir1
   
End Sub

Private Sub Dir1_Click()
   On Error Resume Next
   File1 = Dir1
   
End Sub


Private Sub Dir1_Scroll()
   On Error Resume Next
   File1 = Dir1
   
End Sub


Private Sub Drive1_Change()
   On Error Resume Next
   Dir1 = Drive1
   
End Sub

Private Sub Form_Activate()
   If optFrom.Value = vbUnchecked Then Move EsiLogon.Left - 1800, EsiLogon.Top - 400
   txtPath.BackColor = vbWhite
   txtRept.BackColor = vbWhite
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   Dim sCaps As Variant
   
   If RunningInIDE Then
      Me.cmdRept.Enabled = True
      Me.txtRept.Enabled = True
   End If
   
   txtPath = sFilePath
   txtRept = sReportPath
   If sHelpType = "hlp" Then optHlp.Value = True Else optChm.Value = True
   Drive1 = "c:"
   sOldServer = sserver
   optSve.Value = iSaveCurrent
   If bShowVertical = 1 Then
      optVert.Value = True
   Else
      optHorz.Value = True
   End If
   bOpenForm = GetSetting("Esi2000", "System", "ReOpenForms", bOpenForm)
   iMinimize = GetSetting("Esi2000", "mngr", "MinOnOpen", iMinimize)
   bEnterAsTab = GetSetting("Esi2000", "System", "EnterAsTab", bEnterAsTab)
   bNoResize = GetSetting("Esi2000", "System", "ResizeForm", bNoResize)
   
   If iMinimize Then optMin.Value = vbChecked Else optMin.Value = vbUnchecked
   If bEnterAsTab Then optTab.Value = vbChecked Else optTab.Value = vbUnchecked
   If bNoResize Then optSize.Value = vbChecked Else optSize.Value = vbUnchecked
   If bOpenForm Then optFrm.Value = vbChecked Else optFrm.Value = vbUnchecked
   
   'caps
   sCaps = GetSetting("Esi2000", "mngr", "AutoCaps", sCaps)
   ' DNS sserver = UCase(GetUserSetting(USERSETTING_ServerName))
   sserver = UCase(GetConfUserSetting(USERSETTING_ServerName))
   
   sServers(0) = GetSetting("Esi2000", "System", "ServerId0", Trim(cmbSrv.List(0)))
   sServers(1) = GetSetting("Esi2000", "System", "ServerId1", Trim(cmbSrv.List(1)))
   If Trim(sServers(0)) <> "" Then cmbSrv.AddItem sServers(0)
   If Trim(sServers(1)) <> "" Then cmbSrv.AddItem sServers(1)
   'cmbSrv = GetSetting("Esi2000", "System", "ServerId", Trim(cmbSrv))
   cmbSrv = sserver
   optMin.Caption = "Minimize " & sSysCaption & " Manager On Selection?"
   optCaps.Value = Val(sCaps)
   sHelpType = GetSetting("Esi2000", "System", "HelpType", sHelpType)
   If sHelpType = "" Then sHelpType = "chm"
   If sHelpType = "chm" Then optChm.Value = True _
                  Else optHlp.Value = True
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim b As Byte
   On Error Resume Next
   Err = 0
   If optSize.Value = vbChecked Then b = 1 Else b = 0
   SaveSetting "Esi2000", "mngr", "MinOnOpen", optMin.Value
   iMinimize = optMin.Value
   SaveSetting "Esi2000", "System", "EnterAsTab", bEnterAsTab
   SaveSetting "Esi2000", "System", "ReOpenForms", bOpenForm
   'new
   SaveSetting "Esi2000", "System", "ResizeForm", b
   '1/17/02
   SaveSetting "Esi2000", "mngr", "AutoCaps", optCaps.Value
   SaveSetting "Esi2000", "mngr", "ShowVertical", Abs(optVert.Value)
   '4/6/04
   If optChm.Value = True Then sHelpType = "chm" _
                     Else sHelpType = "hlp"
   SaveSetting "Esi2000", "System", "HelpType", sHelpType
   If Right(Trim(txtPath), 1) <> "\" Then txtPath = txtPath & "\"
   If Right(Trim(txtRept), 1) <> "\" Then txtRept = txtRept & "\"
   sFilePath = txtPath
   sReportPath = txtRept
   SaveSetting "Esi2000", "System", "FilePath", sFilePath
   SaveSetting "Esi2000", "System", "ReportPath", sReportPath
   If optHorz.Value = True Then bShowVertical = 0 Else bShowVertical = 1
   iSaveCurrent = optSve.Value
   SaveSetting "Esi2000", "mngr", "CurrentBar", iSaveCurrent
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set EsiSetup = Nothing
   
End Sub



Private Sub optFrm_Click()
   bOpenForm = optFrm.Value
   
End Sub

Private Sub optFrom_Click()
   'never visible - Check to see where it's loaded
   
End Sub

Private Sub optHorz_Click()
   bShowVertical = Abs(optVert.Value)
   
End Sub

Private Sub optHorz_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii, Me
   
End Sub


Private Sub optMin_Click()
   iMinimize = optMin.Value
   
End Sub


Private Sub optMin_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii, Me
   
End Sub


Private Sub optSize_Click()
   bNoResize = optSize.Value
   
End Sub

Private Sub optSize_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii, Me
   
End Sub


Private Sub optTab_Click()
   bEnterAsTab = optTab.Value
   
End Sub


Private Sub optTab_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii, Me
   
End Sub


Private Sub optVert_Click()
   bShowVertical = Abs(optVert.Value)
   
End Sub

Private Sub optVert_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii, Me
   
End Sub


Private Sub txtPath_GotFocus()
   SelectFormat Me
   
End Sub


Private Sub txtPath_KeyPress(KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub txtPath_LostFocus()
   txtPath = Trim(txtPath)
   If Len(txtPath) > 0 Then
      If Right(txtPath, 1) <> "\" Then txtPath = txtPath & "\"
   End If
   
End Sub

Private Sub txtRept_GotFocus()
   SelectFormat Me
   
End Sub


Private Sub txtRept_KeyPress(KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub txtRept_LostFocus()
   txtRept = Trim(txtRept)
   If Len(txtRept) > 0 Then
      If Right(txtRept, 1) <> "\" Then txtRept = txtRept & "\"
   End If
   
End Sub


Private Sub txtWeb_GotFocus()
   SelectFormat Me
   
End Sub


Private Sub txtWeb_KeyPress(KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub txtWeb_LostFocus()
   txtWeb = LTrim(txtWeb)
   If Left(txtWeb, 7) <> "http://" Then _
           txtWeb = "http://" & txtWeb
   
End Sub
