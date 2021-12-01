VERSION 5.00
Begin VB.Form SysPrinters
   BackColor = &H8000000C&
   BorderStyle = 5 'Sizable ToolWindow
   ClientHeight = 3195
   ClientLeft = 60
   ClientTop = 60
   ClientWidth = 3990
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 3195
   ScaleWidth = 3990
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdClose
      Caption = "&Close"
      Height = 315
      Left = 1440
      TabIndex = 1
      Top = 2760
      Width = 1155
   End
   Begin VB.ListBox lstPrinter
      Height = 2205
      Left = 360
      TabIndex = 0
      Top = 480
      Width = 3255
   End
   Begin VB.Label Label1
      BackStyle = 0 'Transparent
      Caption = "System Printers"
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 9.75
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = -1 'True
      Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00800000&
      Height = 255
      Left = 360
      TabIndex = 2
      Top = 120
      Width = 3135
   End
End
Attribute VB_Name = "SysPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnLoad As Byte

Private Sub List1_Click()
   
End Sub


Private Sub cmdClose_Click()
   Unload Me
   
End Sub

Private Sub Form_Activate()
   Dim X As Printer
   If bOnLoad Then
      For Each X In Printers
         If Left(X.DeviceName, 9) <> "Rendering" Then _
                 lstPrinter.AddItem X.DeviceName
      Next
      bOnLoad = 0
   End If
End Sub

Private Sub Form_Deactivate()
   Unload Me
End Sub

Private Sub Form_Load()
   Move MdiSect.Left + 5000, MdiSect.Top + 1000
   bOnLoad = 1
   Show
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set SysPrinters = Nothing
   
End Sub

Private Sub lstPrinter_Click()
   On Error Resume Next
   MdiSect.ActiveForm.lblPrinter = lstPrinter.List(lstPrinter.ListIndex)
End Sub
