VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaJrnCopy
   BorderStyle = 3 'Fixed Dialog
   Caption = "Copy a Journal"
   ClientHeight = 2385
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 6285
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2385
   ScaleWidth = 6285
   ShowInTaskbar = 0 'False
   Begin VB.CheckBox chkAmts
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2160
      TabIndex = 8
      Top = 1800
      Width = 735
   End
   Begin VB.ComboBox cmbJrn
      Height = 315
      Left = 2160
      TabIndex = 5
      Top = 960
      Width = 1815
   End
   Begin VB.TextBox txtnew
      Height = 285
      Left = 2160
      TabIndex = 3
      Top = 1440
      Width = 1575
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 5280
      TabIndex = 1
      TabStop = 0 'False
      ToolTipText = "Save And Exit"
      Top = 0
      Width = 875
   End
   Begin VB.CommandButton cmdcpy
      Caption = "&Copy"
      Height = 315
      Left = 5280
      TabIndex = 0
      Top = 480
      Width = 855
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 7
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
      PictureUp = "diaJrnCopy.frx":0000
      PictureDn = "diaJrnCopy.frx":0146
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Copy Debits and Credits"
      Height = 255
      Index = 3
      Left = 120
      TabIndex = 9
      Top = 1800
      Width = 1815
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "All Transactions, Account Numbers, And Comments Will Be Copied.  You May Choose To Copy Debit And Credit Amounts."
      Height = 495
      Index = 1
      Left = 480
      TabIndex = 6
      Top = 120
      Width = 4455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "New Journal Name"
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 4
      Top = 1440
      Width = 1575
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Original Journal Name"
      Height = 255
      Index = 2
      Left = 120
      TabIndex = 2
      Top = 960
      Width = 1815
   End
End
Attribute VB_Name = "diaJrnCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
   bOnLoad = False
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diapgl08 = Nothing
   
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub






Private Sub Label1_Click()
   
End Sub

Private Sub z1_Click(Index As Integer)
   
End Sub
