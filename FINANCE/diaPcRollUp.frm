VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form diaPcRollUp
   BorderStyle = 3 'Fixed Dialog
   Caption = "Update Proposed Standard Cost For All Parts"
   ClientHeight = 3990
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 7335
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 3990
   ScaleWidth = 7335
   ShowInTaskbar = 0 'False
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5040
      Top = 240
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3990
      FormDesignWidth = 7335
   End
   Begin VB.CheckBox Check5
      Caption = "___"
      Height = 255
      Left = 3480
      TabIndex = 16
      Top = 3120
      Width = 975
   End
   Begin VB.CheckBox Check4
      Caption = "___"
      Height = 255
      Left = 3480
      TabIndex = 14
      Top = 2160
      Width = 855
   End
   Begin VB.CheckBox Check3
      Caption = "___"
      Height = 195
      Left = 3480
      TabIndex = 13
      Top = 1560
      Width = 855
   End
   Begin VB.CheckBox Check2
      Caption = "___"
      Height = 255
      Left = 3480
      TabIndex = 12
      Top = 1200
      Width = 735
   End
   Begin VB.CheckBox Check1
      Caption = "___"
      Height = 255
      Left = 3480
      TabIndex = 11
      Top = 840
      Width = 735
   End
   Begin VB.TextBox txtPer
      Height = 285
      Left = 3480
      TabIndex = 9
      Top = 2520
      Width = 675
   End
   Begin ComctlLib.ProgressBar PB
      Height = 255
      Left = 240
      TabIndex = 5
      Top = 3600
      Visible = 0 'False
      Width = 4455
      _ExtentX = 7858
      _ExtentY = 450
      _Version = 327682
      Appearance = 1
   End
   Begin VB.CommandButton cmdUpd
      Caption = "&Update"
      Height = 315
      Left = 6360
      TabIndex = 3
      ToolTipText = "Update Standard Cost To Calculated Total"
      Top = 3480
      Width = 875
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 6360
      TabIndex = 2
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 4
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
      PictureUp = "diaPcRollUp.frx":0000
      PictureDn = "diaPcRollUp.frx":0146
   End
   Begin Threed.SSRibbon SSRibbon1
      Height = 255
      Left = 0
      TabIndex = 18
      ToolTipText = "Show System Printers"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 450
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaPcRollUp.frx":028C
      PictureDn = "diaPcRollUp.frx":03D2
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 19
      ToolTipText = "Show System Printers"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 450
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaPcRollUp.frx":0524
      PictureDn = "diaPcRollUp.frx":066A
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Print Only)"
      Height = 285
      Index = 10
      Left = 4440
      TabIndex = 23
      Top = 3120
      Width = 2265
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Treated Like Raw Material Otherwise)"
      Height = 285
      Index = 9
      Left = 4440
      TabIndex = 22
      Top = 1560
      Width = 2865
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Primary & Secondary Shops Otherwise)"
      Height = 285
      Index = 8
      Left = 4440
      TabIndex = 21
      Top = 1200
      Width = 2865
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 20
      Top = 0
      Width = 2760
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Standard Exp Used Otherwise)"
      Height = 285
      Index = 7
      Left = 4440
      TabIndex = 17
      Top = 840
      Width = 2625
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Print This Update"
      Height = 285
      Index = 6
      Left = 240
      TabIndex = 15
      Top = 3120
      Width = 3105
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Percent (Zero To Ignore)"
      Height = 285
      Index = 5
      Left = 4440
      TabIndex = 10
      Top = 2520
      Width = 2385
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "List Variances Greater Than Or Equal To"
      Height = 285
      Index = 4
      Left = 240
      TabIndex = 8
      Top = 2520
      Width = 3345
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Set Standard Equal to The Proposed?"
      Height = 285
      Index = 3
      Left = 240
      TabIndex = 7
      Top = 2160
      Width = 3465
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Update Based ON BOM for ""B"" Parts?"
      Height = 285
      Index = 2
      Left = 240
      TabIndex = 6
      Top = 1560
      Width = 3465
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Use Labor Cost From The Routings"
      Height = 285
      Index = 0
      Left = 240
      TabIndex = 1
      Top = 1200
      Width = 3465
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Use Expense Cost From Routings"
      Height = 285
      Index = 1
      Left = 240
      TabIndex = 0
      Top = 840
      Width = 2505
   End
End
Attribute VB_Name = "diaPcRollUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
' diaPcRollUp
'
' Created: 12/11/01 (nth)
' Revisions:
'
'
'*********************************************************************************
Option Explicit

Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
      
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   bOnLoad = True
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaPcRollUp = Nothing
   
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub
