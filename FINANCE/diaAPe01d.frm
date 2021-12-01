VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPe01d
   BorderStyle = 3 'Fixed Dialog
   Caption = "Invoice Due Date (Posting)"
   ClientHeight = 2190
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 4305
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H80000007&
   LinkTopic = "Form1"
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 2190
   ScaleWidth = 4305
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdPst
      Caption = "&Proceed"
      Enabled = 0 'False
      Height = 315
      Left = 3360
      TabIndex = 2
      ToolTipText = "Save and Post This Item"
      Top = 600
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 3720
      Top = 960
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2190
      FormDesignWidth = 4305
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 3360
      TabIndex = 0
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 1
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
      PictureUp = "diaAPe01d.frx":0000
      PictureDn = "diaAPe01d.frx":0146
   End
End
Attribute VB_Name = "diaAPe01d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaAPe01d - Invoice Due Date
'
' Notes:
'
' Created:
' Revisions:
'
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaTemp = Nothing
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   optPrn.Enabled = False
   optDis.Enabled = False
   
   SetMdiReportsize MdiSect
   
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   
   
   SetCrystalAction Me
   
   optPrn.Enabled = True
   optDis.Enabled = True
   
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillCombo()
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
End Sub

Private Sub GetOptions()
   Dim sOptions As String
End Sub
