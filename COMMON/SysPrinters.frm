VERSION 5.00
Begin VB.Form SysPrinters 
   BackColor       =   &H80000018&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ListBox lstPrinter 
      Height          =   2010
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "System Printers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "SysPrinters"
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
         If Left(X.DeviceName, 9) <> "Rendering" And InStr(1, X.DeviceName, "(redirected") = 0 Then _
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
   BackColor = ES_ViewBackColor
   bOnLoad = 1
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set SysPrinters = Nothing
   
End Sub

Private Sub lstPrinter_Click()
   On Error Resume Next
   If lstPrinter.ListCount > 0 Then
      MdiSect.ActiveForm.lblPrinter = lstPrinter.List(lstPrinter.ListIndex)
   End If
   
End Sub


Private Sub lstPrinter_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii = vbKeyReturn Then
      If lstPrinter.ListCount > 0 Then
         MdiSect.ActiveForm.lblPrinter = lstPrinter.List(lstPrinter.ListIndex)
      End If
      Unload Me
   End If
   
End Sub
