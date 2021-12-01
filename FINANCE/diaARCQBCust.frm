VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARe09
   BorderStyle = 3 'Fixed Dialog
   Caption = "Map ES/2002 Customers To QuickBooks ®"
   ClientHeight = 2145
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 6240
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H80000007&
   LinkTopic = "Form1"
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 2145
   ScaleWidth = 6240
   ShowInTaskbar = 0 'False
   Begin VB.Frame Frame1
      Height = 30
      Left = 120
      TabIndex = 10
      Top = 1080
      Width = 6015
   End
   Begin VB.TextBox txtQBNum
      Height = 285
      Left = 1920
      TabIndex = 2
      Tag = "1"
      Top = 1680
      Width = 735
   End
   Begin VB.TextBox txtQBName
      Height = 285
      Left = 1920
      TabIndex = 1
      Top = 1320
      Width = 3375
   End
   Begin VB.ComboBox cmbCst
      Height = 315
      Left = 1920
      Sorted = -1 'True
      TabIndex = 0
      Top = 360
      Width = 1560
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 5280
      TabIndex = 3
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5640
      Top = 1440
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2145
      FormDesignWidth = 6240
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 9
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaARCQBCust.frx":0000
      PictureDn = "diaARCQBCust.frx":0146
   End
   Begin VB.Label lblnme
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Left = 1920
      TabIndex = 8
      Top = 720
      Width = 3375
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "ES/2002 Customer"
      Height = 285
      Index = 3
      Left = 120
      TabIndex = 7
      Top = 375
      Width = 1425
   End
   Begin VB.Label lblNum
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 315
      Left = 3600
      TabIndex = 6
      Top = 360
      Width = 450
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "QuickBook Customer #"
      Height = 285
      Index = 1
      Left = 120
      TabIndex = 5
      Top = 1680
      Width = 1785
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "QuickBook Customer"
      Height = 285
      Index = 0
      Left = 120
      TabIndex = 4
      Top = 1320
      Width = 1785
   End
End
Attribute VB_Name = "diaARe09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
' diaARCQBCust - Map ES/2002 Customers To QuickBooks
'
' Created: 6/18/02 (nth)
' Revisions:
'
'
'*********************************************************************************

Option Explicit
Dim bOnLoad As Byte
Dim bGoodCust As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Dim rdoCst As rdoResultset

Private Sub cmbCst_Click()
   FindThisCustomer Me
   bGoodCust = GetCustomer
End Sub

Private Sub cmbCst_LostFocus()
   FindThisCustomer Me
   bGoodCust = GetCustomer
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   bOnLoad = True
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set rdoCst = Nothing
   Set diaARe09 = Nothing
End Sub

Private Sub FillCombo()
   FillCustomers Me
   If Cur.CurrentCustomer <> "" Then cmbCst = Cur.CurrentCustomer
   bGoodCust = GetCustomer()
End Sub

Private Function GetCustomer() As Byte
   sSql = "SELECT CUQBNUM,CUQBNAME FROM CustTable " _
          & "WHERE CUNICKNAME = '" & Compress(cmbCst) & "'"
   bSqlRows = GetDataSet(rdoCst, ES_KEYSET)
   If bSqlRows Then
      txtQBName = "" & Trim(rdoCst!CUQBNAME)
      txtQBNum = "" & Trim(rdoCst!CUQBNum)
      GetCustomer = 1
   Else
      txtQBName = ""
      txtQBNum = ""
      GetCustomer = 0
   End If
End Function

Private Sub txtQBName_LostFocus()
   txtQBName = CheckLen(txtQBName, 30)
   If bGoodCust Then
      On Error Resume Next
      rdoCst.Edit
      rdoCst!CUQBNAME = "" & Trim(txtQBName)
      rdoCst.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub txtQBNum_LostFocus()
   If bGoodCust Then
      On Error Resume Next
      rdoCst.Edit
      rdoCst!CUQBNum = Val(txtQBNum)
      rdoCst.Update
      If Err > 0 Then ValidateEdit Me
   End If
End Sub
