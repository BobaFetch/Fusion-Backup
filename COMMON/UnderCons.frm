VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form UnderCons
   BorderStyle = 3 'Fixed Dialog
   Caption = "Under Construction"
   ClientHeight = 1260
   ClientLeft = 2550
   ClientTop = 1830
   ClientWidth = 3930
   ControlBox = 0 'False
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 1260
   ScaleWidth = 3930
   ShowInTaskbar = 0 'False
   Begin VB.Timer Timer1
      Interval = 3000
      Left = 240
      Top = 720
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "OK"
      Height = 375
      Left = 1440
      TabIndex = 2
      Top = 720
      Width = 1215
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 3240
      Top = 720
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 1260
      FormDesignWidth = 3930
   End
   Begin VB.Label Label2
      BackStyle = 0 'Transparent
      Caption = "We Apologize For Any Inconvenience....."
      Height = 255
      Left = 720
      TabIndex = 1
      Top = 360
      Width = 3015
   End
   Begin VB.Label Label1
      BackStyle = 0 'Transparent
      Caption = "This Site Is Currently Being Revised"
      Height = 255
      Left = 720
      TabIndex = 0
      Top = 120
      Width = 3135
   End
   Begin VB.Image Image1
      Height = 480
      Left = 120
      Picture = "UnderCons.frx":0000
      ToolTipText = "Under Construction"
      Top = 120
      Width = 480
   End
End
Attribute VB_Name = "UnderCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub Form_Activate()
   'SysSysSysBeep
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   Move MdiSect.Left + 2500, MdiSect.Top + 1500
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set UnderCons = Nothing
   
End Sub


Private Sub Timer1_Timer()
   Unload Me
   
End Sub
