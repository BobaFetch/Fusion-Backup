VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form SysZoom
   BorderStyle = 3 'Fixed Dialog
   Caption = "Report Options"
   ClientHeight = 2112
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 5388
   ControlBox = 0 'False
   ForeColor = &H8000000F&
   HelpContextID = 907
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2112
   ScaleWidth = 5388
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdHlp
      Appearance = 0 'Flat
      Height = 250
      Left = 0
      Picture = "SysZoom.frx":0000
      Style = 1 'Graphical
      TabIndex = 10
      TabStop = 0 'False
      ToolTipText = "Subject Help"
      Top = 0
      UseMaskColor = -1 'True
      Width = 250
   End
   Begin VB.CheckBox optMax
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2040
      TabIndex = 1
      ToolTipText = "Reports Open Full Screen"
      Top = 1080
      Width = 735
   End
   Begin VB.ComboBox cmbZom
      ForeColor = &H00800000&
      Height = 315
      Left = 2040
      Sorted = -1 'True
      TabIndex = 2
      ToolTipText = "Set Default Zoom Level (Where Display Supported)"
      Top = 1440
      Width = 1335
   End
   Begin VB.CheckBox optBld
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2040
      TabIndex = 0
      ToolTipText = "Set For Reports (Where Printer Supported)"
      Top = 720
      Width = 735
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 4440
      TabIndex = 3
      TabStop = 0 'False
      Top = 0
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4800
      Top = 1920
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2112
      FormDesignWidth = 5388
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Maximized Reports"
      Height = 285
      Index = 5
      Left = 120
      TabIndex = 9
      Top = 1080
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Set Defaults for Crystal Reports:"
      Height = 285
      Index = 2
      Left = 120
      TabIndex = 8
      Top = 360
      Width = 2625
   End
   Begin VB.Label z1
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "%"
      Height = 285
      Index = 1
      Left = 3480
      TabIndex = 7
      Top = 1440
      Width = 225
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Displayed Reports)"
      Height = 285
      Index = 0
      Left = 3840
      TabIndex = 6
      Top = 1440
      Width = 1545
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Bold Printed Reports"
      Height = 285
      Index = 3
      Left = 120
      TabIndex = 5
      Top = 720
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Report Zoom Level"
      Height = 285
      Index = 4
      Left = 120
      TabIndex = 4
      ToolTipText = "Reports Open Full Screen"
      Top = 1440
      Width = 1665
   End
End
Attribute VB_Name = "SysZoom"
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

Private Sub cmbZom_Click()
   iZoomLevel = Val(cmbZom)
   
End Sub

Private Sub cmbZom_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cmbZom_LostFocus()
   iZoomLevel = Val(cmbZom)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub Form_Activate()
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   Move 0, 0
   optMax.Value = GetSetting("Esi2000", "System", "ReportMax", optMax.Value)
   If bBold Then optBld.Value = vbChecked Else optBld.Value = vbUnchecked
   If iZoomLevel > 0 Then
      cmbZom = Trim(Str(iZoomLevel))
   Else
      cmbZom = "Whole Page"
   End If
   cmbZom.AddItem " 25"
   cmbZom.AddItem " 32"
   cmbZom.AddItem " 50"
   cmbZom.AddItem " 75"
   cmbZom.AddItem "100"
   cmbZom.AddItem "150"
   cmbZom.AddItem "200"
   cmbZom.AddItem "300"
   cmbZom.AddItem "400"
   cmbZom.AddItem "Whole Page"
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Left(cmbZom, 2) = "Wh" Then
      SaveSetting "Esi2000", "System", "ReportZoom", 0
   Else
      SaveSetting "Esi2000", "System", "ReportZoom", iZoomLevel
   End If
   SaveSetting "Esi2000", "System", "ReportBold", Trim(Str(bBold))
   SaveSetting "Esi2000", "System", "ReportMax", optMax.Value
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   WindowState = 1
   Set SysZoom = Nothing
   
End Sub


Private Sub optBld_Click()
   If optBld.Value = vbChecked Then
      bBold = 1
   Else
      bBold = 0
   End If
   
End Sub

Private Sub optBld_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optMax_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
