VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form SaleSLview
   BackColor = &H80000018&
   BorderStyle = 1 'Fixed Single
   Caption = "Current Sales Order items"
   ClientHeight = 3564
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 5964
   Icon = "SaleSLview.frx":0000
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 3564
   ScaleWidth = 5964
   ShowInTaskbar = 0 'False
   StartUpPosition = 3 'Windows Default
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4920
      Top = 3360
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3564
      FormDesignWidth = 5964
   End
   Begin MSFlexGridLib.MSFlexGrid grd
      Height = 2535
      Left = 120
      TabIndex = 1
      ToolTipText = "Press ""Esc"" To Close"
      Top = 360
      Width = 5745
      _ExtentX = 10139
      _ExtentY = 4466
      _Version = 393216
      Rows = 10
      Cols = 6
      FixedCols = 0
      ForeColor = 8404992
      FocusRect = 0
      HighLight = 0
      GridLinesFixed = 1
      ScrollBars = 2
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "&Close"
      Height = 435
      Left = 2040
      TabIndex = 0
      Top = 3600
      Width = 915
   End
   Begin VB.Image Chkyes
      Height = 168
      Left = 0
      Picture = "SaleSLview.frx":030A
      Top = 3360
      Visible = 0 'False
      Width = 228
   End
   Begin VB.Image Chkno
      Height = 168
      Left = 240
      Picture = "SaleSLview.frx":0694
      Top = 3360
      Visible = 0 'False
      Width = 228
   End
   Begin VB.Label lblTot
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Height = 255
      Left = 3960
      TabIndex = 5
      Top = 3120
      Width = 975
   End
   Begin VB.Label lblQty
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      Height = 255
      Left = 1320
      TabIndex = 4
      Top = 3120
      Width = 975
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Total Sales"
      Height = 255
      Index = 1
      Left = 2760
      TabIndex = 3
      Top = 3120
      Width = 975
   End
   Begin VB.Label Z1
      BackStyle = 0 'Transparent
      Caption = "Total Quantity"
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 2
      Top = 3120
      Width = 1215
   End
End
Attribute VB_Name = "SaleSLview"
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

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub Form_Click()
   cmdCan_Click
   
End Sub

Private Sub Form_Initialize()
   BackColor = ES_ViewBackColor
   
End Sub


Private Sub Form_Load()
   SetFormSize Me
   If iBarOnTop Then
      Move MdiSect.Left + 800, SaleSLe02b.Top + 1900
   Else
      Move MdiSect.Left + 3600, SaleSLe02b.Top + 1100
   End If
   With grd
      .Row = 0
      .Col = 0
      .Text = "Item"
      .Col = 1
      .Text = "Part Number"
      .Col = 2
      .Text = "Quantity"
      .Col = 3
      .Text = "Schd Date"
      .Col = 4
      .Text = "Act Date"
      .Col = 5
      .Text = "Canceled"
   End With
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   SaleSLe02b.txtQty.SetFocus
   WindowState = 1
   Set SaleSLview = Nothing
   
End Sub
